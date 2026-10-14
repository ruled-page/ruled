/**
 * Checks the built bundles against size budgets, and that the calculation
 * engine is only in the worker, never in code the page itself loads.
 *
 * Needs `npm run build` (the demo and portal) and `npm run build:lib` first. A budget is the
 * gzipped size, a little above today's, so growth is a decision rather than
 * an accident: raise it here when the growth is wanted.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { gzipSync } from 'node:zlib';

const root = fileURLToPath(new URL('..', import.meta.url));

/** [directory, file name pattern, gzipped kB]. */
const BUDGETS = [
  // The demo page, before a workbook is opened.
  ['examples/demo/dist/assets', /^index-.*\.js$/, 185],
  // The worker, the calculation engine included.
  ['examples/demo/dist/assets', /^workbook\.worker-.*\.js$/, 300],
  // Loaded only for workbooks with charts.
  ['examples/demo/dist/assets', /^ChartView-.*\.js$/, 75],
  ['packages/spreadsheet/dist/react', /^workbook-editor\.js$/, 110],
  ['packages/spreadsheet/dist/react/assets', /^workbook\.worker\.js$/, 300],
  // React and the worker included.
  ['packages/spreadsheet/dist/standalone', /^workbook-editor\.js$/, 380],
  // The engine on its own, its parser and number formatter apart.
  ['packages/formula-engine/dist', /^index\.js$/, 85],
];

/** Directories of code that runs on the page, where the engine must not be. */
const PAGE_CODE = ['examples/demo/dist/assets', 'examples/react-portal/dist/assets', 'packages/spreadsheet/dist/react'];

const problems = [];
for (const [dir, pattern, kB] of BUDGETS) {
  const path = join(root, dir);
  const files = existsSync(path) ? readdirSync(path).filter((f) => pattern.test(f)) : [];
  if (!files.length) problems.push(`${dir}: nothing matches ${pattern}; build first`);
  for (const f of files) {
    const size = gzipSync(readFileSync(join(path, f))).length / 1024;
    const line = `${join(dir, f)}: ${size.toFixed(0)} kB gzipped, budget ${kB} kB`;
    console.log(line);
    if (size > kB) problems.push(line);
  }
}
// React Server Components take the React build as client code only when its entry says so first.
const reactEntry = join(root, 'packages/spreadsheet/dist/react/workbook-editor.js');
if (existsSync(reactEntry) && !/^['"]use client['"]/.test(readFileSync(reactEntry, 'utf8'))) {
  problems.push("packages/spreadsheet/dist/react/workbook-editor.js: doesn't start with 'use client', which React Server Components need");
}
// The type declarations import nothing a host would have to install but React (the editor's) or nothing at all (the engine's).
for (const [dir, allowed] of [['packages/spreadsheet/dist/types', /^react(-dom)?(\/|$)/], ['packages/formula-engine/dist', /^$/]]) {
  const path = join(root, dir);
  for (const f of existsSync(path) ? readdirSync(path).filter((n) => n.endsWith('.d.ts')) : []) {
    const imported = [...readFileSync(join(path, f), 'utf8').matchAll(/(?:from|import\()\s*["']([^"'.][^"']*)["']/g)].map((m) => m[1]);
    for (const pkg of new Set(imported)) if (!allowed.test(pkg)) problems.push(`${dir}/${f}: imports ${pkg}, which hosts would have to install`);
  }
}
// The Node build works on its own: it recalculates a sample, choosing its engine itself.
const node = join(root, 'packages/spreadsheet/dist/node/index.js');
if (existsSync(node)) {
  try {
    const { recalculate } = await import(pathToFileURL(node).href);
    const result = await recalculate(readFileSync(join(root, 'examples/demo/public/samples/budget.xlsx')));
    if (!result.sheets.length || !(result.bytes.length > 0)) problems.push(`${node}: recalculating a sample gave no workbook`);
  } catch (e) {
    problems.push(`packages/spreadsheet/dist/node/index.js: recalculating a sample failed: ${e instanceof Error ? e.message : e}`);
  }
}
for (const dir of PAGE_CODE) {
  const path = join(root, dir);
  if (!existsSync(path)) continue;
  for (const f of readdirSync(path)) {
    // A constant only the engine's own code holds (the first of its gamma function's coefficients).
    if (f.endsWith('.js') && !f.includes('worker') && readFileSync(join(path, f), 'utf8').includes('9999999999999971')) {
      problems.push(`${join(dir, f)}: has the calculation engine in it, which belongs in the worker only`);
    }
  }
}
if (problems.length) {
  console.error(`\n${problems.join('\n')}`);
  process.exit(1);
}
