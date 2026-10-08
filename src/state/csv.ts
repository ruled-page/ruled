import { writeCsv, type CsvOptions } from '../core/csv';
import { listSeparator } from '../core/locale';
import { MAX_COLS, MAX_ROWS } from '../core/refs';
import { VT, type DisplayStyle } from '../core/types';
import { displayText } from './format';
import type { SheetView } from './sheetView';

/**
 * A sheet as CSV, the way Excel's "CSV UTF-8" saves one: each cell's value
 * as its number format shows it, with the locale's decimal and thousands
 * separators; formulas as their results; from A1 to the last row and column
 * with a value; separated by the locale's list separator unless the options
 * say otherwise. Values that would run as formulas are escaped unless
 * `escapeFormulas` is false.
 */
export function sheetCsv(view: SheetView, styles: DisplayStyle[], date1904: boolean, locale: string, options: CsvOptions = {}): string {
  const values: string[][] = [];
  let lastRow = -1;
  let lastCol = -1;
  view.forEach({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, (row, col, rec) => {
    if (rec.t === VT.None) return;
    (values[row] ??= [])[col] = displayText(rec, styles[rec.style] ?? styles[0], date1904, false, locale).text;
    lastRow = Math.max(lastRow, row);
    lastCol = Math.max(lastCol, col);
  });
  const rows: string[][] = [];
  for (let r = 0; r <= lastRow; r++) rows.push(Array.from({ length: lastCol + 1 }, (_, c) => values[r]?.[c] ?? ''));
  return writeCsv(rows, options.delimiter ?? listSeparator(locale), options.escapeFormulas !== false);
}

/** The CSV file's name: the workbook's, and the sheet's when the workbook has several. */
export function csvFileName(workbookName: string, sheetName: string, sheets: number): string {
  const base = workbookName.replace(/\.\w+$/, '') || sheetName;
  return sheets > 1 ? `${base} - ${sheetName}.csv` : `${base}.csv`;
}
