import type { RawCellContent } from 'hyperformula';
import { msg, type Message, type MessageKey, type Text } from '../i18n/message';
import type { EngineValue } from './engineValue';
import { engineFormula } from './formula';
import type { CellContent } from './types';

/**
 * What the session needs from a calculation engine, whichever one works the
 * formulas out: HyperFormula (./calc), or another with the same contract.
 * Engines are loaded only when chosen (see src/worker/engines.ts), so this
 * module holds what the session uses without loading any of them.
 */

/** A cell's content as an engine takes it: a value, or a formula with "=". */
export type RawContent = RawCellContent;

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
  | { kind: 'tableRefs' | 'externalLinks' | 'threeD' | 'syntax' | 'array' | 'sharedFormula' | 'tooLarge' };

const LABELS: Record<Exclude<Unsupported['kind'], 'function'>, MessageKey> = {
  tableRefs: 'core.unsupported.tableRefs',
  externalLinks: 'core.unsupported.externalLinks',
  threeD: 'core.unsupported.threeD',
  syntax: 'core.unsupported.syntax',
  array: 'core.unsupported.array',
  sharedFormula: 'core.unsupported.sharedFormula',
  tooLarge: 'core.unsupported.tooLarge',
};

const EXPLANATIONS: Record<Exclude<Unsupported['kind'], 'function'>, MessageKey> = {
  tableRefs: 'core.cellReason.tableRefs',
  externalLinks: 'core.cellReason.externalLinks',
  threeD: 'core.cellReason.threeD',
  syntax: 'core.cellReason.syntax',
  array: 'core.cellReason.array',
  sharedFormula: 'core.cellReason.sharedFormula',
  tooLarge: 'core.cellReason.tooLarge',
};

/** What isn't calculated, as the recalculation report lists it: "LET function". */
export function unsupportedLabel(u: Unsupported): Message {
  return u.kind === 'function' ? msg('core.unsupported.function', { name: u.name }) : msg(LABELS[u.kind]);
}

/** Why a formula isn't calculated, in a sentence about its cell. */
export function unsupportedExplanation(u: Unsupported): Message {
  return u.kind === 'function' ? msg('core.cellReason.function', { name: u.name }) : msg(EXPLANATIONS[u.kind]);
}

export interface EngineSheet {
  name: string;
  rows: RawCellContent[][];
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
  raw: RawCellContent;
}

export interface EngineChange {
  sheet: number;
  row: number;
  col: number;
  value: EngineValue;
}

export function toRaw(content: CellContent): RawCellContent {
  switch (content.kind) {
    case 'empty':
      return null;
    case 'number':
      return content.value;
    case 'bool':
      return content.value;
    // A leading apostrophe stops the engine from reading "007" or "1/2" as numbers.
    case 'string':
      return "'" + content.value;
    case 'error':
      return content.value;
    case 'formula':
      return engineFormula(content.formula);
  }
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
  destroy(): void;
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
  create(sheets: (EngineSheet | null)[], names: EngineName[], date1904: boolean, lookups?: EngineLookups): FormulaEngine;
}

let fallback: EngineKind | undefined;

/** Sets the engine new sessions calculate with. */
export function chooseEngine(kind: EngineKind): void {
  fallback = kind;
}

/** The engine chosen for new sessions, if one has been. */
export function chosenEngine(): EngineKind | undefined {
  return fallback;
}

/** The engine new sessions calculate with. */
export function currentEngine(): EngineKind {
  if (!fallback) throw new Error('No calculation engine has been chosen: call chooseEngine first.');
  return fallback;
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

/** A value as a message shows it: a number, "text" in quotes, an error code, or "(empty)". */
export function describeValue(v: EngineValue): Text {
  if (v === null) return msg('core.value.empty');
  if (typeof v === 'object') return v.error;
  if (typeof v === 'string') return JSON.stringify(v);
  return String(v).toUpperCase() === 'TRUE' || String(v).toUpperCase() === 'FALSE' ? String(v).toUpperCase() : String(v);
}
