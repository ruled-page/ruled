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
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';

const root = fileURLToPath(new URL('..', import.meta.url));

/** [directory, file name pattern, gzipped kB]. */
const BUDGETS = [
  // The demo page, before a workbook is opened.
  ['examples/demo/dist/assets', /^index-.*\.js$/, 275],
  // The worker, the calculation engine included.
  ['examples/demo/dist/assets', /^workbook\.worker-.*\.js$/, 300],
  // Loaded only for workbooks with charts.
  ['examples/demo/dist/assets', /^ChartView-.*\.js$/, 75],
  ['packages/spreadsheet/dist/react', /^workbook-editor\.js$/, 175],
  ['packages/spreadsheet/dist/react/assets', /^workbook\.worker\.js$/, 300],
  // React and the worker included.
  ['packages/spreadsheet/dist/standalone', /^workbook-editor\.js$/, 565],
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
