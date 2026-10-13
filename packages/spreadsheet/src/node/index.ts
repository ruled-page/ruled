/**
 * Ruled without a browser: `recalculate` opens a saved workbook, sets the
 * cells given, works out its formulas and saves it again, in Node.
 * `import { recalculate } from '@ruled/sheet/node'`.
 */
export { recalculate } from './recalculate';
export type { CellInput, CellValue, Recalculated, RecalculateOptions } from './recalculate';
