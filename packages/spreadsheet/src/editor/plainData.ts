import type { WorkbookReaders } from '../state/readers';
import type { SavedWorkbook } from '../state/store';

/** A saved workbook as plain data, without the functions that read its cells: what a server function can be sent. */
export type SavedWorkbookData = Omit<SavedWorkbook, keyof WorkbookReaders>;

/** A saved workbook without its readers. */
export function plainSaved(saved: SavedWorkbook): SavedWorkbookData {
  const { cell: _cell, range: _range, ...data } = saved;
  return data;
}

/** Whether two copies of a file hold the same bytes, as each render of a Server Component sends a new copy. */
export function sameBytes(a: ArrayBuffer | Uint8Array, b: ArrayBuffer | Uint8Array): boolean {
  if (a === b) return true;
  const x = a instanceof Uint8Array ? a : new Uint8Array(a);
  const y = b instanceof Uint8Array ? b : new Uint8Array(b);
  if (x.length !== y.length) return false;
  for (let i = 0; i < x.length; i++) if (x[i] !== y[i]) return false;
  return true;
}
