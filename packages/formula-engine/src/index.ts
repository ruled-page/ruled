/**
 * Excel's calculation on its own: `createWorkbook` makes a workbook of
 * sheets, values and formulas and keeps its results up to date, in a
 * browser, a worker or Node.
 */
export { createWorkbook, dateToSerial, serialToDate } from './book';
export type { CellChange, CellInput, CellValue, Workbook, WorkbookOptions } from './book';
