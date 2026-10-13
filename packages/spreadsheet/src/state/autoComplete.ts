import { MAX_ROWS } from '@ruled/formula/refs';
import { CellFlag, VT } from '../core/types';
import type { CellRecord } from './format';

/**
 * The entries Excel's AutoComplete offers in a cell: the text typed into the
 * cells above and below it in its column, as far as the first empty cell
 * either way. Numbers, dates and the results of formulas aren't offered, but
 * don't end the list.
 */
export function columnEntries(get: (row: number, col: number) => CellRecord | null, row: number, col: number): string[] {
  const entries: string[] = [];
  const take = (r: number) => {
    const rec = get(r, col);
    if (!rec || (rec.t === VT.None && !(rec.flags & CellFlag.Formula))) return false;
    if (rec.t === VT.String && rec.s && !(rec.flags & (CellFlag.Formula | CellFlag.ArrayMember))) entries.push(rec.s);
    return true;
  };
  for (let r = row - 1; r >= 0 && take(r); r--);
  for (let r = row + 1; r < MAX_ROWS && take(r); r++);
  return entries;
}

/**
 * The entry AutoComplete finishes typed text with: the one entry it is the
 * start of, ignoring case, as Excel does. None while it starts more than one
 * (or none), for a formula, or for text on more than one line.
 */
export function autoComplete(typed: string, entries: readonly string[]): string | undefined {
  if (!typed || typed.startsWith('=') || typed.includes('\n')) return undefined;
  const start = typed.toLowerCase();
  let found: string | undefined;
  for (const entry of entries) {
    const lower = entry.toLowerCase();
    if (!lower.startsWith(start)) continue;
    if (found === undefined) found = entry;
    else if (found.toLowerCase() !== lower) return undefined;
  }
  // The rest is added to what was typed, so the two must line up.
  return found !== undefined && found.length >= typed.length ? found : undefined;
}
