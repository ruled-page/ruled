/**
 * Draws every sheet of the demo's samples in Glide's grid and in Ruled's,
 * and compares them pixel by pixel. Writes both pictures, a picture of where
 * they differ (red) and a page of them all, most different first, to
 * packages/e2e/grid-compare. Needs the demo served, at GRID_COMPARE_URL or
 * http://localhost:4190:
 *
 *   npm run build -w @ruled/demo && npm run preview -w @ruled/demo -- --port 4190
 *   node packages/e2e/scripts/compare-grids.mjs [sample.xlsx ...]
 */
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from '@playwright/test';
import sharp from 'sharp';

const root = fileURLToPath(new URL('../../..', import.meta.url));
const out = join(root, 'packages/e2e/grid-compare');
const base = process.env.GRID_COMPARE_URL ?? 'http://localhost:4190';
/** How far apart two pixels' colours may be, in any channel, and still count as the same. */
const TOLERANCE = 40;

const listed = [...readFileSync(join(root, 'examples/demo/src/Welcome.tsx'), 'utf8').matchAll(/file: '([^']+)'/g)].map((m) => m[1]);
const samples = process.argv.slice(2).length ? process.argv.slice(2) : listed;
mkdirSync(out, { recursive: true });

const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1280, height: 800 }, deviceScaleFactor: 1 });

/** Each sheet of a sample as a grid draws it: the sheet's name and a PNG of the sheet's area. */
async function draw(sample, grid) {
  const page = await context.newPage();
  await page.goto(`${base}/?sample=${sample}&grid=${grid}`);
  const tabs = page.getByRole('navigation', { name: 'Sheets' }).getByRole('tab');
  await tabs.first().waitFor({ timeout: 120_000 });
  const names = await tabs.allTextContents();
  const shots = [];
  for (const [i, name] of names.entries()) {
    await tabs.nth(i).click();
    // Charts, fonts and the screen readers' table settle.
    await page.waitForTimeout(1500);
    shots.push({ name, png: await page.locator('.wbe-sheet-area').screenshot() });
  }
  await page.close();
  return shots;
}

/** How many pixels differ, and a picture of where: red on a faded copy of the first. */
async function compare(a, b) {
  const [ia, ib] = await Promise.all([sharp(a).ensureAlpha().raw().toBuffer({ resolveWithObject: true }), sharp(b).ensureAlpha().raw().toBuffer({ resolveWithObject: true })]);
  const width = Math.max(ia.info.width, ib.info.width);
  const height = Math.max(ia.info.height, ib.info.height);
  const diff = Buffer.alloc(width * height * 4);
  let different = 0;
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const pa = x < ia.info.width && y < ia.info.height ? (y * ia.info.width + x) * 4 : -1;
      const pb = x < ib.info.width && y < ib.info.height ? (y * ib.info.width + x) * 4 : -1;
      const o = (y * width + x) * 4;
      const same = pa >= 0 && pb >= 0 && [0, 1, 2].every((c) => Math.abs(ia.data[pa + c] - ib.data[pb + c]) <= TOLERANCE);
      if (same) {
        const grey = 255 - (255 - (ia.data[pa] + ia.data[pa + 1] + ia.data[pa + 2]) / 3) * 0.25;
        diff.set([grey, grey, grey, 255], o);
      } else {
        different++;
        diff.set([230, 30, 30, 255], o);
      }
    }
  }
  return { share: different / (width * height), png: await sharp(diff, { raw: { width, height, channels: 4 } }).png().toBuffer() };
}

const rows = [];
for (const sample of samples) {
  const [glide, ruled] = [await draw(sample, 'glide'), await draw(sample, 'ruled')];
  for (const [i, g] of glide.entries()) {
    const r = ruled[i];
    if (!r) continue;
    const stem = `${sample.replace(/\.xlsx$/, '')}-${i + 1}`;
    const { share, png } = await compare(g.png, r.png);
    writeFileSync(join(out, `${stem}-glide.png`), g.png);
    writeFileSync(join(out, `${stem}-ruled.png`), r.png);
    writeFileSync(join(out, `${stem}-diff.png`), png);
    rows.push({ sample, sheet: g.name, stem, share });
    console.log(`${(share * 100).toFixed(2).padStart(6)}%  ${sample} / ${g.name}`);
  }
}
await browser.close();

rows.sort((a, b) => b.share - a.share);
const esc = (s) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
writeFileSync(
  join(out, 'index.html'),
  `<!doctype html><meta charset="utf-8"><title>Glide and Ruled's grids compared</title>
<style>body{font:14px system-ui,sans-serif;margin:16px}h2{font-size:15px;margin:24px 0 6px}figure{display:inline-block;margin:0 8px 8px 0;vertical-align:top}figcaption{color:#555;margin-bottom:4px}img{max-width:620px;border:1px solid #ccc}</style>
<h1>Glide's grid and Ruled's, sheet by sheet</h1>
<p>Each sheet as each grid draws it, and where they differ (red), most different first. Pixels count as the same within ${TOLERANCE} of each other in every channel.</p>
${rows
  .map(
    (r) => `<h2>${esc(r.sample)} / ${esc(r.sheet)}: ${(r.share * 100).toFixed(2)}% differs</h2>
<figure><figcaption>Glide</figcaption><img src="${r.stem}-glide.png" alt=""></figure><figure><figcaption>Ruled</figcaption><img src="${r.stem}-ruled.png" alt=""></figure><figure><figcaption>Difference</figcaption><img src="${r.stem}-diff.png" alt=""></figure>`,
  )
  .join('\n')}`,
);
console.log(`\nWrote ${join(out, 'index.html')}`);
