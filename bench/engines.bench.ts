import { existsSync, readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { gzipSync } from 'node:zlib';
import { it } from 'vitest';
import { hyperFormulaEngine } from '../src/core/calc';
import { chooseEngine, type EngineKind } from '../src/core/engine';
import { findCell } from '../src/core/sheetScanner';
import { WorkbookSession } from '../src/core/session';
import { CellFlag, VT } from '../src/core/types';

/**
 * Compares the calculation engines on the sample workbooks and the test
 * corpus, headless: `npm run engine:bench` (after `npm run build` for the
 * engines' download sizes). For each file and engine: opening, building the
 * engine and calculating, the time a typical edit takes, the heap the engine
 * adds, and how its results compare with the values Excel saved.
 */

/** The engines to compare, with the name of each one's chunk in the build. */
const ENGINES: { kind: EngineKind; chunk: RegExp }[] = [{ kind: hyperFormulaEngine, chunk: /^worker-calc-.*\.js$/ }];

const root = join(__dirname, '..');
const FILES = [
  ...readdirSync(join(root, 'public/samples')).map((f) => join('public/samples', f)),
  ...readdirSync(join(root, 'tests/corpus')).map((f) => join('tests/corpus', f)),
]
  .filter((f) => f.endsWith('.xlsx'))
  // Smallest first, so memory a large file leaves to be collected isn't counted against the next.
  .sort((a, b) => readFileSync(join(root, a)).length - readFileSync(join(root, b)).length);

/** How many cells holding plain numbers each file has edited, one at a time. */
const EDITS = 20;

const gc = (globalThis as { gc?: () => void }).gc;
const heapMB = () => {
  gc?.();
  return process.memoryUsage().heapUsed / 1048576;
};

function chunkKB(pattern: RegExp): string {
  const dir = join(root, 'dist/assets');
  const file = existsSync(dir) ? readdirSync(dir).find((f) => pattern.test(f)) : undefined;
  return file ? `${(gzipSync(readFileSync(join(dir, file))).length / 1024).toFixed(0)} kB` : 'build first';
}

/** The first cells holding plain numbers, which an edit changes. */
function numberCells(session: WorkbookSession): { sheet: number; row: number; col: number; value: number }[] {
  const out: { sheet: number; row: number; col: number; value: number }[] = [];
  session.model.sheets.forEach((m, sheet) => {
    const grid = m.scan?.grid;
    if (!grid || m.info.kind !== 'worksheet' || m.info.protected) return;
    for (let row = 0; row < grid.rowCount && out.length < EDITS; row++) {
      for (let i = grid.rowPtr[row]; i < grid.rowPtr[row + 1] && out.length < EDITS; i++) {
        if (grid.types[i] === VT.Number && !(grid.flags[i] & (CellFlag.Formula | CellFlag.ArrayMember | CellFlag.Spilled))) {
          out.push({ sheet, row, col: grid.cols[i], value: grid.nums[findCell(grid, row, grid.cols[i])] });
        }
      }
    }
  });
  return out;
}

const median = (xs: number[]) => (xs.length ? [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)] : 0);

it('compares the calculation engines', { timeout: 900_000 }, () => {
  const rows: Record<string, string | number>[] = [];
  for (const { kind, chunk } of ENGINES) {
    chooseEngine(kind);
    for (const file of FILES) {
      const bytes = new Uint8Array(readFileSync(join(root, file)));
      const before = heapMB();
      let started = performance.now();
      const session = WorkbookSession.open(bytes, file, 'en-US', { recalc: 'on' });
      const openMs = performance.now() - started;
      started = performance.now();
      const report = session.ensureEngine();
      const calcMs = performance.now() - started;
      const heap = heapMB() - before;
      const times: number[] = [];
      for (const c of numberCells(session)) {
        started = performance.now();
        session.apply({ type: 'input', sheet: c.sheet, row: c.row, col: c.col, text: String(c.value + 1) });
        times.push(performance.now() - started);
      }
      rows.push({
        engine: kind.name,
        file: file.replace(/^.*\//, ''),
        formulas: report.formulaCells,
        'open ms': Math.round(openMs),
        'calc ms': Math.round(calcMs),
        'edit ms (median)': Number(median(times).toFixed(1)),
        'edit ms (max)': Number(Math.max(0, ...times).toFixed(1)),
        'heap MB': Number(heap.toFixed(1)),
        'not calculated': report.unsupported,
        'differ from Excel': report.mismatched,
      });
      session.releaseEngine();
    }
    rows.push({ engine: kind.name, file: '(download, gzipped)', formulas: chunkKB(chunk) });
  }
  console.table(rows);
  if (!gc) console.log('Heap figures are rough: run with node --expose-gc for exact ones (npm run engine:bench does).');
});
