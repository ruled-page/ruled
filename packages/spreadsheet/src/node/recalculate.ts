import { nativeEngine } from '@ruled/formula/native';
import { parseRange, quoteSheetName, toA1, type Rect } from '@ruled/formula/refs';
import { chooseEngine, chosenEngine } from '../core/engine';
import { WorkbookSession } from '../core/session';
import { CellFlag } from '../core/types';
import { english } from '../i18n';

/** What a cell can be set to: a value, or a formula starting with "=" as it would be typed in Excel. A Date is read in UTC, and null or undefined empties the cell. */
export type CellInput = string | number | boolean | Date | null | undefined;

/** A cell's value: a number (dates as Excel's serial numbers), text, TRUE or FALSE, an error such as `{ error: '#DIV/0!' }`, or null when empty. */
export type CellValue = number | string | boolean | { error: string } | null;

export interface RecalculateOptions {
  /** Cells to set first, by reference with their sheet ("Inputs!B2"), or on the first sheet ("B2"). */
  set?: Record<string, CellInput>;
  /**
   * Works out every formula the engine can, rather than keeping the value
   * the file saved where the engine's differs. For files saved by tools
   * that don't calculate, whose saved values may be out of date; Excel's
   * own are kept by default.
   */
  ignoreSavedValues?: boolean;
  /** A password-protected workbook's password; the workbook comes back encrypted with it. */
  password?: string;
}

export interface Recalculated {
  /** The workbook saved again: the cells set, and each formula's value as worked out. Excel recalculates it once more on opening. */
  bytes: Uint8Array;
  /** The sheets' names, in order. */
  sheets: readonly string[];
  /** A cell's value: "Summary!B3", or "B3" on the first sheet. */
  get(cell: string): CellValue;
  /** A range's values, as rows: "Summary!A1:C10". */
  getRange(range: string): CellValue[][];
  /**
   * Formulas that keep the value the file saved, though a cell they read
   * was set, so it is now out of date: ones that can't be calculated here
   * (a function this engine doesn't have, a link to another workbook), and
   * ones in `differs`.
   */
  outOfDate: string[];
  /** Formulas whose value worked out here differs from the one the file saved, which they keep. Empty with `ignoreSavedValues`. */
  differs: string[];
}

/**
 * Opens a saved workbook (.xlsx or .xlsm), sets the cells given, works out
 * its formulas as Excel does, and saves it again, so that whatever reads
 * the file next, Excel or not, finds the values up to date.
 */
export async function recalculate(xlsx: Uint8Array | ArrayBuffer, options: RecalculateOptions = {}): Promise<Recalculated> {
  // In a browser the worker chooses the engine; here there's only the one.
  if (!chosenEngine()) chooseEngine(nativeEngine);
  const session = WorkbookSession.open(new Uint8Array(xlsx), 'workbook.xlsx', 'en-US', { recalc: 'on', savedValues: options.ignoreSavedValues ? 'ignore' : 'compare', password: options.password });
  session.ensureEngine();
  const sheets = session.loadResult().sheets.map((s) => s.name);

  /** The sheet and cells a reference names: "B2", "Inputs!B2:C4" or "'Q1 Sales'!A1". */
  const locate = (ref: string): { sheet: number; rect: Rect } => {
    const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(ref.trim());
    const name = m ? (m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]) : undefined;
    const sheet = name === undefined ? 0 : sheets.findIndex((n) => n.toLowerCase() === name.toLowerCase());
    if (sheet < 0) throw new Error(`There is no sheet named "${name}".`);
    const rect = parseRange(m ? m[3] : ref.trim());
    if (!rect) throw new Error(`"${ref}" isn't a cell or range reference.`);
    return { sheet, rect };
  };

  for (const [ref, input] of Object.entries(options.set ?? {})) {
    const { sheet, rect } = locate(ref);
    const text = typed(input);
    const result =
      text === undefined
        ? session.apply({ type: 'clear', sheet, rects: [{ r1: rect.r1, c1: rect.c1, r2: rect.r1, c2: rect.c1 }] })
        : session.apply({ type: 'input', sheet, row: rect.r1, col: rect.c1, text });
    const refused = result.rejected ?? (result.skipped ? result.skippedReasons[0] : undefined);
    if (refused) throw new Error(`${ref} can't be set: ${english.text(refused)}`);
  }

  const name = (sheet: number, row: number, col: number) => `${quoteSheetName(sheets[sheet])}!${toA1(row, col)}`;
  const flagged = (flag: number) =>
    session
      .stateUpdates()
      .filter((u) => u.flags & flag)
      .sort((a, b) => a.sheet - b.sheet || a.row - b.row || a.col - b.col)
      .map((u) => name(u.sheet, u.row, u.col));
  const value = (sheet: number, row: number, col: number): CellValue => {
    const v = session.valueAt(sheet, row, col);
    return v !== null && typeof v === 'object' ? { error: v.error } : v;
  };
  return {
    bytes: session.saveRecalculated(),
    sheets,
    get(cell) {
      const { sheet, rect } = locate(cell);
      if (rect.r1 !== rect.r2 || rect.c1 !== rect.c2) throw new Error(`"${cell}" is a range; getRange reads ranges.`);
      return value(sheet, rect.r1, rect.c1);
    },
    getRange(range) {
      const { sheet, rect } = locate(range);
      return Array.from({ length: rect.r2 - rect.r1 + 1 }, (_r, r) => Array.from({ length: rect.c2 - rect.c1 + 1 }, (_c, c) => value(sheet, rect.r1 + r, rect.c1 + c)));
    },
    outOfDate: flagged(CellFlag.Stale),
    differs: flagged(CellFlag.Mismatch),
  };
}

/** A value as typed into Excel in English, so it is stored as given: text marked as text, a date in ISO form. */
function typed(input: CellInput): string | undefined {
  if (input === null || input === undefined) return undefined;
  if (typeof input === 'string') return input.startsWith('=') ? input : `'${input}`;
  if (typeof input === 'number') return String(input);
  if (typeof input === 'boolean') return input ? 'TRUE' : 'FALSE';
  const iso = input.toISOString();
  const time = iso.slice(11, 19);
  return time === '00:00:00' ? iso.slice(0, 10) : `${iso.slice(0, 10)} ${time}`;
}
