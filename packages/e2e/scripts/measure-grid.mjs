/**
 * Measures the grid on the demo's samples, in Chromium at a 2× pixel ratio:
 * how long a sample takes to open and draw its first sheet, main-thread work
 * and frame times while scrolling by wheel and in jumps, time per key while
 * an arrow key is held down, and memory after opening and after scrolling.
 * Each sample runs RUNS times, and the medians are printed and written to
 * packages/e2e/grid-measure/measurements.json. Needs the demo served, at
 * GRID_MEASURE_URL or http://localhost:4190:
 *
 *   npm run build -w @ruled/demo && npm run preview -w @ruled/demo -- --port 4190
 *   node packages/e2e/scripts/measure-grid.mjs [sample.xlsx ...]
 *
 * Headless Chromium draws without a GPU, so the times are for comparing one
 * build with another, not for what a person sees.
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from '@playwright/test';

const root = fileURLToPath(new URL('../../..', import.meta.url));
const out = join(root, 'packages/e2e/grid-measure');
const base = process.env.GRID_MEASURE_URL ?? 'http://localhost:4190';
const RUNS = Number(process.env.RUNS ?? 3);
const samples = process.argv.slice(2).length ? process.argv.slice(2) : ['big.xlsx', 'kitchen-sink.xlsx', 'formatting.xlsx'];
/** A frame this much longer than a 60 Hz frame counts as dropped. */
const SLOW_FRAME = 1000 / 60 + 8;

const median = (xs) => {
  const s = [...xs].sort((a, b) => a - b);
  return s.length ? (s.length % 2 ? s[(s.length - 1) / 2] : (s[s.length / 2 - 1] + s[s.length / 2]) / 2) : NaN;
};
const percentile = (xs, p) => {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.floor((p / 100) * s.length))] ?? NaN;
};

/** Records animation frames and long tasks in the page until stopped; returns the frames' lengths and the long tasks' time. */
const RECORD = () => {
  const w = window;
  w.__frames = [];
  w.__long = 0;
  w.__recording = true;
  const loop = (t) => {
    w.__frames.push(t);
    if (w.__recording) requestAnimationFrame(loop);
  };
  requestAnimationFrame(loop);
  w.__observer?.disconnect();
  w.__observer = new PerformanceObserver((list) => {
    for (const e of list.getEntries()) w.__long += e.duration;
  });
  w.__observer.observe({ type: 'longtask' });
};
const STOP = () => {
  const w = window;
  w.__recording = false;
  const f = w.__frames;
  // Two callbacks in one frame share its time; only the frames themselves count.
  return { frames: f.slice(1).map((t, i) => t - f[i]).filter((d) => d > 0), longTasks: w.__long };
};

/** The main thread's time so far, in ms: all its tasks, the script among them, and layout and style. */
async function work(cdp) {
  const { metrics } = await cdp.send('Performance.getMetrics');
  const m = Object.fromEntries(metrics.map((x) => [x.name, x.value]));
  return { task: m.TaskDuration * 1000, script: m.ScriptDuration * 1000, layout: (m.LayoutDuration + m.RecalcStyleDuration) * 1000 };
}
const since = (a, b, steps) => ({ taskMs: (b.task - a.task) / steps, scriptMs: (b.script - a.script) / steps, layoutMs: (b.layout - a.layout) / steps });

async function memory(cdp) {
  await cdp.send('HeapProfiler.collectGarbage');
  const { metrics } = await cdp.send('Performance.getMetrics');
  const m = Object.fromEntries(metrics.map((x) => [x.name, x.value]));
  return { heapMB: m.JSHeapUsedSize / 1048576, nodes: m.Nodes, listeners: m.JSEventListeners };
}

async function measure(browser, sample) {
  const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: 2 });
  const page = await context.newPage();
  const cdp = await context.newCDPSession(page);
  await cdp.send('Performance.enable');
  const result = { sample };

  // Opening: from the page's start until the first sheet's canvas has been sized and drawn.
  await page.goto(`${base}/?sample=${sample}`);
  await page.waitForFunction(() => {
    const c = document.querySelector('[data-testid="wbe-grid"] canvas');
    return !!c && c.width > 300;
  }, undefined, { timeout: 180_000 });
  await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
  result.openMs = await page.evaluate(() => performance.now());
  result.afterOpen = await memory(cdp);

  const surface = page.getByTestId('wbe-grid');
  const box = await surface.boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);

  // Scrolling by wheel: a turn every frame for two seconds.
  let before = await work(cdp);
  await page.evaluate(RECORD);
  for (let i = 0; i < 120; i++) {
    await page.mouse.wheel(0, 60);
    await page.waitForTimeout(16);
  }
  await page.waitForTimeout(300);
  const wheel = await page.evaluate(STOP);
  result.wheelWork = since(before, await work(cdp), 120);
  result.wheel = { medianMs: median(wheel.frames), p95Ms: percentile(wheel.frames, 95), slow: wheel.frames.filter((f) => f > SLOW_FRAME).length, frames: wheel.frames.length, longTaskMs: wheel.longTasks };

  // Scrolling in jumps: a long way each frame, as dragging a scrollbar's thumb does.
  before = await work(cdp);
  await page.evaluate(RECORD);
  await page.evaluate(async () => {
    const s = document.querySelector('[data-testid="wbe-grid-scroller"]');
    for (let i = 0; i < 60; i++) {
      s.scrollTop += 4000;
      await new Promise((r) => requestAnimationFrame(r));
    }
  });
  await page.waitForTimeout(300);
  const jumps = await page.evaluate(STOP);
  result.jumpWork = since(before, await work(cdp), 60);
  result.jumps = { medianMs: median(jumps.frames), p95Ms: percentile(jumps.frames, 95), slow: jumps.frames.filter((f) => f > SLOW_FRAME).length, frames: jumps.frames.length, longTaskMs: jumps.longTasks };
  result.afterScroll = await memory(cdp);

  // An arrow key held down: 100 presses from A2, until the active cell stops moving (merges and hidden rows decide where it ends).
  const nameBox = page.getByRole('combobox', { name: 'Go to cell' });
  await nameBox.click();
  await nameBox.fill('A2');
  await nameBox.press('Enter');
  await page.locator('[data-testid="wbe-grid"]:focus-within').first().waitFor();
  const start = Date.now();
  for (let i = 0; i < 100; i++) await page.keyboard.press('ArrowDown');
  let last = '';
  let settled = Date.now();
  for (let quiet = 0; quiet < 200; ) {
    const now = await nameBox.inputValue();
    if (now !== last) {
      last = now;
      settled = Date.now();
      quiet = 0;
    } else quiet = Date.now() - settled;
    await page.waitForTimeout(20);
  }
  result.keyMs = (settled - start) / 100;
  result.keyEnd = last;

  await context.close();
  return result;
}

const browser = await chromium.launch();
const all = [];
for (const sample of samples) {
  for (let run = 0; run < RUNS; run++) {
    const r = await measure(browser, sample);
    all.push(r);
    console.log(`${sample} run ${run + 1}: open ${r.openMs.toFixed(0)} ms, work per wheel turn ${r.wheelWork.taskMs.toFixed(2)} ms, per jump ${r.jumpWork.taskMs.toFixed(2)} ms, key ${r.keyMs.toFixed(1)} ms, heap ${r.afterScroll.heapMB.toFixed(1)} MB`);
  }
}
await browser.close();

const rows = [];
for (const sample of samples) {
  const runs = all.filter((r) => r.sample === sample);
  const m = (f) => median(runs.map(f));
  rows.push({
    sample,
    'open ms': m((r) => r.openMs).toFixed(0),
    'work per wheel turn ms (all / script)': `${m((r) => r.wheelWork.taskMs).toFixed(2)} / ${m((r) => r.wheelWork.scriptMs).toFixed(2)}`,
    'work per jump ms (all / script)': `${m((r) => r.jumpWork.taskMs).toFixed(2)} / ${m((r) => r.jumpWork.scriptMs).toFixed(2)}`,
    'slow frames (wheel, jumps)': `${m((r) => r.wheel.slow)}, ${m((r) => r.jumps.slow)}`,
    'long tasks ms': m((r) => r.wheel.longTaskMs + r.jumps.longTaskMs).toFixed(0),
    'ms per key': m((r) => r.keyMs).toFixed(1),
    'heap MB (opened / scrolled)': `${m((r) => r.afterOpen.heapMB).toFixed(1)} / ${m((r) => r.afterScroll.heapMB).toFixed(1)}`,
    'DOM nodes': m((r) => r.afterScroll.nodes),
    listeners: m((r) => r.afterScroll.listeners),
  });
}
console.table(rows);
mkdirSync(out, { recursive: true });
writeFileSync(join(out, 'measurements.json'), JSON.stringify({ runs: all, medians: rows }, null, 2));
console.log(`Wrote ${join(out, 'measurements.json')}`);
