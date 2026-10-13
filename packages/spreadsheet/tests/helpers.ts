import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { WorkbookSession } from '../src/core/session';
import { parseA1 } from '@ruled/formula/refs';
import { findCell } from '../src/core/sheetScanner';
import { CellFlag, VT, type CellUpdate } from '../src/core/types';

export const SAMPLES = join(__dirname, '..', '..', '..', 'examples', 'demo', 'public', 'samples');

export function sample(name: string): Uint8Array {
  return new Uint8Array(readFileSync(join(SAMPLES, name)));
}

export function open(name: string): WorkbookSession {
  return WorkbookSession.open(sample(name), name);
}

export function sheetIndex(s: WorkbookSession, name: string): number {
  const i = s.model.sheets.findIndex((m) => m.info.name === name);
  if (i < 0) throw new Error(`no sheet ${name}`);
  return i;
}

/** The value a cell shows after load (Excel's saved value), as the UI sees it. */
export function loadedValue(s: WorkbookSession, sheet: string, ref: string): unknown {
  const grid = s.model.sheets[sheetIndex(s, sheet)].scan!.grid;
  const { row, col } = parseA1(ref)!;
  const idx = findCell(grid, row, col);
  if (idx < 0) return null;
  switch (grid.types[idx]) {
    case VT.Number:
      return grid.nums[idx];
    case VT.String:
      return grid.strs[idx];
    case VT.Bool:
      return grid.nums[idx] === 1;
    case VT.Error:
      return { error: grid.strs[idx] };
    default:
      return null;
  }
}

export function loadedFormula(s: WorkbookSession, sheet: string, ref: string): string | undefined {
  const grid = s.model.sheets[sheetIndex(s, sheet)].scan!.grid;
  const { row, col } = parseA1(ref)!;
  return grid.formulas.get(findCell(grid, row, col));
}

export function updateAt(updates: CellUpdate[], s: WorkbookSession, sheet: string, ref: string): CellUpdate | undefined {
  const i = sheetIndex(s, sheet);
  const { row, col } = parseA1(ref)!;
  return updates.find((u) => u.sheet === i && u.row === row && u.col === col);
}

export function valueOf(u: CellUpdate | undefined): unknown {
  if (!u) return undefined;
  if (u.t === VT.Number) return u.n;
  if (u.t === VT.String) return u.s;
  if (u.t === VT.Bool) return u.n === 1;
  if (u.t === VT.Error) return { error: u.s };
  return null;
}

export function has(flags: number | undefined, flag: number): boolean {
  return ((flags ?? 0) & flag) === flag;
}

export function parts(bytes: Uint8Array): Record<string, string> {
  const files = unzipSync(bytes);
  return Object.fromEntries(Object.entries(files).map(([k, v]) => [k, k.endsWith('.png') || k.endsWith('.bin') ? '<binary>' : strFromU8(v)]));
}

export { CellFlag };
