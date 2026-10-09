/**
 * Checks the built bundles against size budgets, and that the calculation
 * engine is only in the worker, never in code the page itself loads.
 *
 * Needs `npm run build` and `npm run build:lib` first. A budget is the
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
  ['dist/assets', /^sites-.*\.js$/, 275],
  ['dist/assets', /^workbook\.worker-.*\.js$/, 300],
  // HyperFormula, which the worker loads when a workbook is opened with it.
  ['dist/assets', /^worker-calc-.*\.js$/, 210],
  // Loaded only for workbooks with charts.
  ['dist/assets', /^ChartView-.*\.js$/, 75],
  ['dist-lib/react', /^workbook-editor\.js$/, 175],
  ['dist-lib/react/assets', /^workbook\.worker\.js$/, 300],
  // React and the worker included.
  ['dist-lib/standalone', /^workbook-editor\.js$/, 565],
];

/** Directories of code that runs on the page, where the engine must not be. */
const PAGE_CODE = ['dist/assets', 'dist-lib/react'];

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
    // Its licence check's key, which only the engine's own code holds (the page may still name it, in the engine picker).
    if (f.endsWith('.js') && !f.includes('worker') && readFileSync(join(path, f), 'utf8').includes('gpl-v3')) {
      problems.push(`${join(dir, f)}: has the calculation engine in it, which belongs in the worker only`);
    }
  }
}
if (problems.length) {
  console.error(`\n${problems.join('\n')}`);
  process.exit(1);
}
