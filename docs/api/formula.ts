/**
 * Excel's calculation on its own, with no editor: `createWorkbook` makes a
 * workbook of sheets, values and formulas and keeps its results up to date.
 * `import { createWorkbook } from '@ruled/formula'`.
 *
 * @module @ruled/formula
 */
export { createWorkbook, dateToSerial, serialToDate } from '../../packages/formula-engine/src/book';
export type { CellChange, CellInput, CellValue, Workbook, WorkbookOptions } from '../../packages/formula-engine/src/book';
