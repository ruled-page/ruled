import { MAX_COLS, MAX_ROWS, parseA1, parseRange, toA1 } from '../core/refs';
import { CellFlag, type DisplayStyle, type Rect } from '../core/types';
import type { JsonValue } from '../schema/types';
import { cellJson } from './format';
import type { SheetView } from './sheetView';

export interface WorkbookReaders {
  /** Any cell's value as JSON, e.g. cell('Data', 'B4'); dates come back as ISO text. */
  cell(sheet: string, ref: string): JsonValue;
  /** A range's values row by row, e.g. range('Data', 'A1:F200'); without a range, the sheet's used range. */
  range(sheet: string, ref?: string): JsonValue[][];
}

/**
 * Formulas whose value may be out of date: ones this editor can't calculate,
 * still showing Excel's last result although a cell they depend on changed,
 * and every formula that depends on one of those.
 */
export function outOfDateCells(sheets: SheetView[]): { sheet: string; ref: string }[] {
  return sheets.flatMap((view) =>
    view
      .flagged(CellFlag.Stale)
      .sort((a, b) => a.row - b.row || a.col - b.col)
      .map(({ row, col }) => ({ sheet: view.info.name, ref: toA1(row, col) })),
  );
}

/** The A1 parser takes row 0 and columns past XFD; neither is on a sheet. */
const onSheet = (r: Rect) => r.r1 >= 0 && r.r2 < MAX_ROWS && r.c2 < MAX_COLS;

/** Readers over sheets as they are now, by sheet name and A1 reference. */
export function workbookReaders(sheets: SheetView[], styles: DisplayStyle[], date1904: boolean): WorkbookReaders {
  const find = (name: string) => {
    const view = sheets.find((v) => v.info.name.toLowerCase() === name.toLowerCase());
    if (!view) throw new Error(`There is no sheet called "${name}".`);
    return view;
  };
  const value = (view: SheetView, row: number, col: number) => {
    const rec = view.get(row, col);
    return cellJson(rec, styles[rec?.style ?? 0] ?? styles[0], date1904);
  };
  return {
    cell: (sheet, ref) => {
      const at = parseA1(ref);
      if (!at || !onSheet({ r1: at.row, c1: at.col, r2: at.row, c2: at.col })) throw new Error(`"${ref}" isn't a cell reference.`);
      return value(find(sheet), at.row, at.col);
    },
    range: (sheet, ref) => {
      const view = find(sheet);
      // Cells typed or pasted past the file's used range count too.
      const { maxRow, maxCol } = view.extent();
      const rect = ref ? parseRange(ref) : maxRow < 0 ? null : { r1: 0, c1: 0, r2: maxRow, c2: maxCol };
      if (ref && !(rect && onSheet(rect))) throw new Error(`"${ref}" isn't a range.`);
      if (!rect) return [];
      const rows: JsonValue[][] = [];
      const r2 = Math.min(rect.r2, Math.max(maxRow, rect.r1));
      const c2 = Math.min(rect.c2, Math.max(maxCol, rect.c1));
      for (let r = rect.r1; r <= r2; r++) {
        const line: JsonValue[] = [];
        for (let c = rect.c1; c <= c2; c++) line.push(value(view, r, c));
        rows.push(line);
      }
      return rows;
    },
  };
}
