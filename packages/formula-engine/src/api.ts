import type { EngineValue } from './engineValue';

/**
 * What an editor needs from a calculation engine: the contract an engine
 * keeps (FormulaEngine, made by an EngineKind), the shape of the cells and
 * changes passed across it, and how its results are compared with the
 * values Excel saved. Engines themselves are loaded only when wanted (the
 * native one from "@ruled/formula/native"), so this module holds
 * none of the calculation.
 */

/** A cell's content as an engine takes it: a value, or a formula with "=". */
export type RawContent = Date | string | number | boolean | null | undefined;

/** The function a spilling formula is wrapped in for the engine (see spillFormula). */
export const SPILL_FUNCTION = 'WBE.SPILL';

/** The most cells one formula's results may spill into here. */
export const MAX_SPILL_CELLS = 1_000_000;

export interface ArrayDims {
  rows: number;
  cols: number;
}

/** Why the engine can't calculate a formula. */
export type Unsupported =
  | { kind: 'function'; name: string }
  | { kind: 'tableRefs' | 'externalLinks' | 'syntax' | 'array' | 'sharedFormula' | 'tooLarge' };

export interface EngineSheet {
  name: string;
  rows: RawContent[][];
}

export interface EngineName {
  name: string;
  /** File-form formula text, e.g. "Data!$A$1:$A$10". */
  formula: string;
  localSheet?: number;
}

export interface EngineWrite {
  sheet: number;
  row: number;
  col: number;
  raw: RawContent;
}

export interface EngineChange {
  sheet: number;
  row: number;
  col: number;
  value: EngineValue;
}

interface Pos {
  sheet: number;
  row: number;
  col: number;
}

/** A workbook's formulas worked out, kept up to date as cells are written. */
export interface FormulaEngine {
  /** Whether a sheet is in the engine (sheets nothing refers to are left out). */
  has(sheet: number): boolean;
  /** Puts a sheet left out at first into the engine; returns the values that changed. */
  addSheet(sheet: number, content: EngineSheet): EngineChange[];
  value(sheet: number, row: number, col: number): EngineValue;
  /** Where a HYPERLINK formula in a cell links to. */
  hyperlink(sheet: number, row: number, col: number): string | undefined;
  /** Whether a formula's own result is an array. */
  isArray(sheet: number, row: number, col: number): boolean;
  /** The size a spilling formula's result had when last worked out. */
  measured(sheet: number, row: number, col: number): ArrayDims | undefined;
  /** The size a formula's result has, worked out on its own. */
  resultSize(formula: string, sheet: number): ArrayDims;
  /** Writes cells; returns every value that changed as a result. */
  write(writes: EngineWrite[]): EngineChange[];
  /** Defines, changes or (with null) removes a name; `held` when it couldn't be, and is kept as it was. */
  setName(name: string, localSheet: number | undefined, formula: string | null): { changes: EngineChange[]; held: boolean };
  /** Evaluates a formula (with "=") on a sheet without placing it in a cell. */
  calculate(formula: string, sheet: number): EngineValue | EngineValue[][];
  /** Formula cells that depend on any of the given cells, directly or indirectly. */
  dependentsOf(cells: Pos[], limit?: number): Pos[];
  /** The formula cells where a circular reference closes (a formula that reads itself, directly or through others), by sheet, row and column. */
  circular(): Pos[];
  destroy(): void;
}

/**
 * Iterative calculation, as a workbook's calculation settings turn it on: a
 * formula in a circular reference reads the others' values from the pass
 * before, and the circle is worked round until no value moves by more than
 * `delta` or `count` passes are done.
 */
export interface Iteration {
  count: number;
  delta: number;
}

/** What the editor tells an engine about cells beyond their contents. */
export interface EngineLookups {
  /** A cell's formula as the formula bar shows it ("=SUM(A1:A3)"), for FORMULATEXT; undefined when it holds none. */
  formulaText(sheet: number, row: number, col: number): string | undefined;
}

/** One kind of engine: what it can calculate, and a new one for a workbook. */
export interface EngineKind {
  readonly id: string;
  readonly name: string;
  /**
   * Why this engine can't calculate a formula (file form, without "="), or
   * undefined when it can; `lambdas` are the workbook's names defined as
   * LAMBDAs, in capitals, which formulas call as functions.
   */
  unsupportedReason(fileFormula: string, lambdas?: ReadonlySet<string>): Unsupported | undefined;
  create(sheets: (EngineSheet | null)[], names: EngineName[], date1904: boolean, lookups?: EngineLookups, iteration?: Iteration): FormulaEngine;
}

/**
 * A formula (in the engine's form, with "=") wrapped so the engine gives its
 * result rows × cols cells; `fixed` fills them as an array formula entered
 * with Ctrl+Shift+Enter does.
 */
export function spillFormula(formula: string, rows: number, cols: number, fixed = false): string {
  return `=${SPILL_FUNCTION}(${rows},${cols},${formula.replace(/^=/, '')}${fixed ? ',1' : ''})`;
}

/** Whether a value is the #SPILL! error, shown by a formula whose results have no room. */
export function isSpillError(v: EngineValue): boolean {
  return v !== null && typeof v === 'object' && v.error === '#SPILL!';
}

/**
 * Functions Excel works out by searching for a result rather than by a
 * formula, which it documents as accurate "within 0.000001 percent": their
 * results can differ from an exact answer in about the eighth significant
 * digit, so they are compared with that allowance.
 */
const SEARCHED = new Set(['IRR', 'XIRR', 'RATE', 'YIELD']);

/**
 * Functions Excel works out from fits good to about eight digits (the Bessel
 * functions): the same fits give Excel's digits, but the route it takes for
 * some arguments isn't known, so they have the same allowance.
 */
const APPROXIMATED = new Set(['BESSELI', 'BESSELJ', 'BESSELK', 'BESSELY']);

/** How closely a formula's result must match Excel's saved value, relative to its size. */
export function tolerance(functions: readonly string[]): number {
  return functions.some((f) => SEARCHED.has(f) || APPROXIMATED.has(f)) ? 1e-7 : 1e-9;
}

export function valuesMatch(excel: EngineValue, engine: EngineValue, within = 1e-9): boolean {
  if (excel === null || excel === '') return engine === null || engine === '' || engine === 0;
  if (typeof excel === 'number' && typeof engine === 'number') {
    const scale = Math.max(1, Math.abs(excel), Math.abs(engine));
    return Math.abs(excel - engine) <= within * scale;
  }
  if (typeof excel === 'object' && typeof engine === 'object' && engine) return excel.error === engine.error;
  return excel === engine;
}
