/**
 * Excel's calculation on its own: `createWorkbook` makes a workbook of
 * sheets, values and formulas and keeps its results up to date. Also the
 * contract between an editor and a calculation engine, and the values that
 * cross it; the engine itself is "@ruled/formula/native", and Excel's cell
 * references and formula text are "./refs" and "./formula".
 */
export { createWorkbook, dateToSerial, serialToDate } from './book';
export type { CellChange, CellInput, CellValue, Workbook, WorkbookOptions } from './book';
export * from './api';
export * from './engineValue';
export { fifteenDigits } from './excelNumbers';
