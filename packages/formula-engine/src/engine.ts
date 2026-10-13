/**
 * The contract between an editor and a calculation engine, and the values
 * that cross it: what @ruled/sheet uses to drive the engine in its worker.
 * The engine itself is "./native", and Excel's cell references and formula
 * text are "./refs" and "./formula". Not part of the published package.
 */
export * from './api';
export * from './engineValue';
export { fifteenDigits } from './excelNumbers';
