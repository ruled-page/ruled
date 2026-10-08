import { parseA1Ref, tokenize } from '@borgar/fx';
import { DetailedCellError, HyperFormula, type RawCellContent, type SerializedNamedExpression } from 'hyperformula';
import { msg, type Message, type MessageKey, type Text } from '../i18n/message';
import { installEngineLimits, tooLargeToCalculate } from './engineLimits';
import type { EngineValue } from './engineValue';
import { registerExcelFunctions } from './excelFunctions';
import { engineFormula, expandLets, stripFunctionPrefixes } from './formula';
import { MAX_RESULT_CELLS } from './functions/common';
import { measuredSizes, SPILL_FUNCTION, type ArrayDims } from './functions/spill';
import type { CellContent } from './types';

let registered: Set<string> | undefined;
let probe: HyperFormula | undefined;

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

/**
 * Why the engine cannot calculate this formula, or undefined when it can.
 * Checked up front so a formula is judged on its own content, not on errors
 * it inherits from the cells it references.
 */
export function unsupportedReason(fileFormula: string): Unsupported | undefined {
  registerExcelFunctions();
  installEngineLimits();
  registered ??= new Set(HyperFormula.getRegisteredFunctionNames('enGB'));
  // LET is written out for the engine; one that can't be is left for Excel.
  const expanded = expandLets(fileFormula);
  if (expanded === undefined) return { kind: 'function', name: 'LET' };
  let tokens;
  try {
    tokens = tokenize(expanded, { mergeRefs: true, allowTernary: false });
  } catch {
    return { kind: 'syntax' };
  }
  if (tooLargeToCalculate(tokens)) return { kind: 'tooLarge' };
  for (const t of tokens) {
    if (t.type === 'structured') return { kind: 'tableRefs' };
    if (t.type === 'func') {
      const name = stripFunctionPrefixes(t.value).toUpperCase();
      if (!registered.has(name) || name === SPILL_FUNCTION) return { kind: 'function', name: name.replace(/^_XLUDF\./, '') };
    }
    if (t.type === 'range' || t.type === 'range_beam' || t.type === 'range_ternary' || t.type === 'unknown') {
      if (t.value.includes('[')) return { kind: 'externalLinks' };
      const ref = parseA1Ref(t.value, { allowTernary: false });
      // The parser gives Jan:Mar!B2 one context, "Jan:Mar"; sheet names can't contain ":".
      const context = ref && 'context' in ref ? (ref.context ?? []) : [];
      if (context.length > 1 || context[0]?.includes(':')) return { kind: 'threeD' };
    }
  }
  probe ??= HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
  if (!probe.validateFormula(engineFormula(fileFormula))) return { kind: 'syntax' };
  return undefined;
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

/**
 * A name's formula as the engine takes it, or undefined for names it leaves
 * out: Excel's own (print areas and the like), broken references, and
 * ranges so large that formulas using them would get past the size limits.
 */
function nameExpression(n: EngineName): string | undefined {
  if (n.name.startsWith('_xlnm.') || /#REF!/.test(n.formula)) return undefined;
  if (unsupportedReason(n.formula)?.kind === 'tooLarge') return undefined;
  return engineFormula(n.formula);
}

function convert(v: unknown): EngineValue {
  if (v instanceof DetailedCellError) return { error: v.value, message: v.message };
  if (v === undefined || v === null) return null;
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  return null;
}

/**
 * HyperFormula configured to behave like Excel: 1900 date system with the
 * Lotus leap-year bug, empty cells as zero in arithmetic, full sheet size.
 */
export class CalcEngine {
  private readonly hf: HyperFormula;
  /** Our sheet index → engine sheet id (-1 when the sheet is not in the engine). */
  private readonly ids: number[];
  private readonly sheetOf = new Map<number, number>();

  constructor(sheets: (EngineSheet | null)[], names: EngineName[], date1904: boolean) {
    const config = {
      licenseKey: 'gpl-v3',
      maxRows: 1_048_576,
      maxColumns: 16_384,
      // The 1899-12-30 base date already absorbs Excel's phantom 29 Feb 1900.
      leapYear1900: false,
      nullDate: date1904 ? { year: 1904, month: 1, day: 1 } : { year: 1899, month: 12, day: 30 },
      evaluateNullToZero: true,
      // Smart rounding gives Excel's comparisons (0.1+0.2=0.3) and its snapping of
      // near-zero sums to 0. It also rounds every result, by default to 11
      // significant digits, which loses cents on large amounts; 15 is Excel's own
      // precision, and still hides float noise such as 0.30000000000000004.
      smartRounding: true,
      precisionRounding: 14,
      useArrayArithmetic: true,
      useWildcards: true,
      useRegularExpressions: false,
      matchWholeCell: true,
      caseSensitive: false,
      undoLimit: 0,
    };

    registerExcelFunctions();
    installEngineLimits();
    // Sheet names come from the file; one called __proto__ must be a sheet, not the object's prototype.
    const data: Record<string, RawCellContent[][]> = Object.create(null);
    for (const s of sheets) if (s) data[s.name] = s.rows;

    // Probe names against an empty copy of the workbook so one bad name
    // cannot stop the whole engine from building.
    const probe = HyperFormula.buildFromSheets(
      Object.fromEntries(Object.keys(data).map((n) => [n, []])),
      config,
    );
    const valid: SerializedNamedExpression[] = [];
    for (const n of names) {
      const expression = nameExpression(n);
      if (expression === undefined) continue;
      const scopeName = n.localSheet !== undefined ? sheets[n.localSheet]?.name : undefined;
      if (n.localSheet !== undefined && !scopeName) continue;
      const scope = scopeName !== undefined ? probe.getSheetId(scopeName) : undefined;
      try {
        if (probe.isItPossibleToAddNamedExpression(n.name, expression, scope)) {
          probe.addNamedExpression(n.name, expression, scope);
          valid.push({ name: n.name, expression, scope });
        }
      } catch {
        // Names the engine cannot represent are skipped; formulas using them
        // will be reported as not calculated here.
      }
    }
    probe.destroy();

    this.hf = HyperFormula.buildFromSheets(data, config, valid);
    this.ids = sheets.map((s) => (s ? this.hf.getSheetId(s.name) ?? -1 : -1));
    this.ids.forEach((id, i) => {
      if (id >= 0) this.sheetOf.set(id, i);
    });
  }

  /** Whether a sheet's cells are in the engine. */
  has(sheet: number): boolean {
    return (this.ids[sheet] ?? -1) >= 0;
  }

  /**
   * Adds a sheet that was left out when the engine was built because no
   * formula used it, once one does. Returns the cells whose values changed.
   */
  addSheet(sheet: number, content: EngineSheet): EngineChange[] {
    if (this.has(sheet)) return [];
    // Changes made inside a batch are all returned by the batch.
    const changes = this.hf.batch(() => {
      const name = this.hf.addSheet(content.name);
      const id = this.hf.getSheetId(name)!;
      this.ids[sheet] = id;
      this.sheetOf.set(id, sheet);
      this.hf.setSheetContent(id, content.rows);
    });
    const out: EngineChange[] = [];
    for (const c of changes) {
      if (!('address' in c) || !c.address) continue;
      const s = this.sheetOf.get(c.address.sheet);
      if (s !== undefined && s !== sheet) out.push({ sheet: s, row: c.address.row, col: c.address.col, value: convert(c.newValue) });
    }
    return out;
  }

  value(sheet: number, row: number, col: number): EngineValue {
    const id = this.ids[sheet];
    if (id === undefined || id < 0) return null;
    return convert(this.hf.getCellValue({ sheet: id, row, col }));
  }

  /** The address a HYPERLINK formula in this cell links to, as last calculated. */
  hyperlink(sheet: number, row: number, col: number): string | undefined {
    const id = this.ids[sheet];
    if (id === undefined || id < 0) return undefined;
    const url: unknown = this.hf.getCellHyperlink({ sheet: id, row, col });
    return url === undefined || url === null ? undefined : String(url);
  }

  /** Whether the engine worked out an array for the formula in this cell, which Excel would spill. */
  isArray(sheet: number, row: number, col: number): boolean {
    const id = this.ids[sheet];
    if (id === undefined || id < 0) return false;
    try {
      return this.hf.isCellPartOfArray({ sheet: id, row, col });
    } catch {
      return false;
    }
  }

  /** The size a fitted formula's result had when the engine last worked it out. */
  measured(sheet: number, row: number, col: number): ArrayDims | undefined {
    const id = this.ids[sheet];
    if (id === undefined || id < 0) return undefined;
    return measuredSizes(this.hf.dependencyGraph).get(`${id}:${row}:${col}`);
  }

  /** The size of a formula's result (in the engine's form, with "="), worked out on its own: 1 × 1 for a single value. */
  resultSize(formula: string, sheet: number): ArrayDims {
    const v = this.calculate(formula, sheet);
    return Array.isArray(v) && v.length && v[0].length ? { rows: v.length, cols: v[0].length } : { rows: 1, cols: 1 };
  }

  /** Applies writes in one batch and returns every cell whose value changed. */
  write(writes: EngineWrite[]): EngineChange[] {
    if (!writes.length) return [];
    return this.cellChanges(
      this.hf.batch(() => {
        for (const w of writes) {
          const id = this.ids[w.sheet];
          if (id === undefined || id < 0) continue;
          try {
            this.hf.setCellContents({ sheet: id, row: w.row, col: w.col }, w.raw);
          } catch {
            // Content the engine rejects (e.g. an unparseable formula) is stored as text.
            this.hf.setCellContents({ sheet: id, row: w.row, col: w.col }, "'" + String(w.raw));
          }
        }
      }),
    );
  }

  /**
   * Defines a name, replacing the one of that name and scope, or removes it
   * (formula null). Formulas that use it recalculate, showing #NAME? once it
   * is gone, as in Excel. `held` is false when the engine can't take the
   * name, which is then left out as names are when the engine is built.
   */
  setName(name: string, localSheet: number | undefined, formula: string | null): { changes: EngineChange[]; held: boolean } {
    const scope = localSheet === undefined ? undefined : this.ids[localSheet];
    if (scope !== undefined && scope < 0) return { changes: [], held: false };
    const expression = formula === null ? undefined : nameExpression({ name, formula, localSheet });
    let held = false;
    const changes = this.hf.batch(() => {
      const exists = this.hf.getNamedExpression(name, scope) !== undefined;
      if (expression !== undefined && (exists ? this.hf.isItPossibleToChangeNamedExpression(name, expression, scope) : this.hf.isItPossibleToAddNamedExpression(name, expression, scope))) {
        if (exists) this.hf.changeNamedExpression(name, expression, scope);
        else this.hf.addNamedExpression(name, expression, scope);
        held = true;
      } else if (exists) {
        this.hf.removeNamedExpression(name, scope);
      }
    });
    return { changes: this.cellChanges(changes), held };
  }

  /** The cells among the engine's changes, by our sheet index. */
  private cellChanges(changes: ReturnType<HyperFormula['batch']>): EngineChange[] {
    const out: EngineChange[] = [];
    for (const c of changes) {
      if (!('address' in c) || !c.address) continue;
      const sheet = this.sheetOf.get(c.address.sheet);
      if (sheet === undefined) continue;
      out.push({ sheet, row: c.address.row, col: c.address.col, value: convert(c.newValue) });
    }
    return out;
  }

  /** Evaluates a formula (with "=") on a sheet without placing it in a cell. */
  calculate(formula: string, sheet: number): EngineValue | EngineValue[][] {
    const id = this.ids[sheet];
    if (id === undefined || id < 0) return { error: '#REF!' };
    // Only a range or an array function can be too large; most formulas here are neither.
    if (/:|SEQUENCE|MUNIT|RANDARRAY/i.test(formula) && unsupportedReason(formula.replace(/^=/, ''))?.kind === 'tooLarge') return { error: '#NUM!' };
    try {
      const v = this.hf.calculateFormula(formula, id);
      return Array.isArray(v) ? v.map((row) => row.map(convert)) : convert(v);
    } catch (e) {
      return { error: '#ERROR!', message: (e as Error).message };
    }
  }

  /**
   * Formula cells that depend on any of the given cells, directly or
   * indirectly. Used to mark results the engine cannot vouch for.
   */
  dependentsOf(cells: { sheet: number; row: number; col: number }[], limit = 200_000): { sheet: number; row: number; col: number }[] {
    type Node = { sheet: number; row: number; col: number } | { start: { sheet: number; row: number; col: number }; end: { sheet: number; row: number; col: number } };
    const seen = new Set<string>();
    const out: { sheet: number; row: number; col: number }[] = [];
    const queue: Node[] = [];
    for (const c of cells) {
      const id = this.ids[c.sheet];
      if (id !== undefined && id >= 0) queue.push({ sheet: id, row: c.row, col: c.col });
    }
    for (let qi = 0; qi < queue.length && out.length < limit; qi++) {
      const node = queue[qi];
      let deps: Node[];
      try {
        deps = this.hf.getCellDependents(node as never) as Node[];
      } catch {
        continue;
      }
      for (const d of deps) {
        const id = 'start' in d ? `r${d.start.sheet}:${d.start.row}:${d.start.col}:${d.end.row}:${d.end.col}` : `c${d.sheet}:${d.row}:${d.col}`;
        if (seen.has(id)) continue;
        seen.add(id);
        queue.push(d);
        if (!('start' in d)) {
          const sheet = this.sheetOf.get(d.sheet);
          if (sheet !== undefined) out.push({ sheet, row: d.row, col: d.col });
        }
      }
    }
    return out;
  }

  destroy(): void {
    this.hf.destroy();
  }
}

/** A formula (in the engine's form, with "=") wrapped so the engine gives its result rows × cols cells. */
export function spillFormula(formula: string, rows: number, cols: number): string {
  return `=${SPILL_FUNCTION}(${rows},${cols},${formula.replace(/^=/, '')})`;
}

/** The most cells one formula's results may spill into here. */
export const MAX_SPILL_CELLS = MAX_RESULT_CELLS;

/** Whether a value is the #SPILL! error, shown by a formula whose results have no room. */
export function isSpillError(v: EngineValue): boolean {
  return v !== null && typeof v === 'object' && v.error === '#SPILL!';
}

export function valuesMatch(excel: EngineValue, engine: EngineValue): boolean {
  if (excel === null || excel === '') return engine === null || engine === '' || engine === 0;
  if (typeof excel === 'number' && typeof engine === 'number') {
    const scale = Math.max(1, Math.abs(excel), Math.abs(engine));
    return Math.abs(excel - engine) <= 1e-9 * scale;
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
