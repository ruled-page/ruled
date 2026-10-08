import { isMessage, msg, TextError, type MessageKey, type Text } from '../i18n/message';
import { parseA1Ref, tokenize } from '@borgar/fx';
import type { RawCellContent } from 'hyperformula';
import { format as formatNumber, isDateFormat } from 'numfmt';
import {
  CalcEngine,
  describeValue,
  toRaw,
  unsupportedExplanation,
  unsupportedLabel,
  unsupportedReason,
  valuesMatch,
  type Unsupported,
  type EngineSheet,
  type EngineValue,
  type EngineWrite,
} from './calc';
import { evaluateConditionalFormats } from './conditionalFormat';
import { checkRule, evaluateListFormula, type Evaluator } from './dataValidation';
import { mergeEditResults } from './editResult';
import { planFill, type FillSource } from './fill';
import { displayFormula, engineFormula, isVolatile, makeShifter } from './formula';
import { parseInput } from './input';
import { cellKey, colToLetters, inRect, MAX_COLS, MAX_ROWS, parseA1, rangeToA1, toA1 } from './refs';
import { findCell } from './sheetScanner';
import {
  CellFlag,
  EMPTY,
  VT,
  type CalcReport,
  type CalcSample,
  type CellChange,
  type CellContent,
  type CellSnapshot,
  type CellUpdate,
  type DisplayStyle,
  type EditOp,
  type CfResult,
  type EditResult,
  type Finding,
  type ListSource,
  type LoadResult,
  type Rect,
  type SheetGrid,
} from './types';
import { joinTexts, readWorkbook, type SheetModel, type WorkbookModel } from './workbook';
import { buildPackage, removeMacros, writeWorkbook, type SaveInput } from './writer';
import { restructure } from './restructure';
import { refMapFor, type StructuralOp } from './refTransform';
import { OoxmlPackage, type PackageLimits } from './zip';

export interface OpenOptions {
  limits?: PackageLimits & {
    /** Most cells across all sheets, counted as they are read. Default 2,000,000. */
    maxCells?: number;
  };
  /**
   * What to do with macros, VBA or Excel 4.0 macro sheets: keep them
   * (default), remove them on save, or refuse the file.
   */
  macros?: 'keep' | 'remove' | 'block';
  /** ActiveX controls and embedded objects: keep them (default), or refuse the file. */
  activeContent?: 'keep' | 'block';
  /**
   * Live recalculation: always, never, or (the default) unless the workbook
   * is larger than LIVE_RECALC_LIMITS, when it starts off and can be turned on.
   */
  recalc?: 'auto' | 'on' | 'off';
}

const DEFAULT_MAX_CELLS = 2_000_000;

/**
 * Above these, building the calculation engine takes many seconds and a
 * gigabyte or more of memory, so live recalculation starts off.
 */
export const LIVE_RECALC_LIMITS = { formulas: 200_000, cells: 2_000_000 };

export interface CellState {
  content: CellContent;
  style: number;
}

interface Write extends CellState {
  sheet: number;
  row: number;
  col: number;
}

interface Pos {
  sheet: number;
  row: number;
  col: number;
}

/**
 * A formula the engine does not calculate (it holds Excel's saved value
 * instead), with the cells it reads. When any of those change, the formula
 * and everything downstream of it is marked stale.
 */
interface UntrustedFormula {
  target: Pos;
  /** Ranges the formula reads; null when they cannot be worked out (names, tables). */
  reads: { sheet: number; rect: Rect }[] | null;
}

const MERGED = '\u0000merged';

export function rowOfIndex(grid: SheetGrid, idx: number): number {
  let lo = 0;
  let hi = grid.rowCount - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (grid.rowPtr[mid] <= idx) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

/** Excel's saved value for a cell, from the grid arrays. */
export function gridValue(grid: SheetGrid, idx: number): EngineValue {
  switch (grid.types[idx]) {
    case VT.Number:
      return grid.nums[idx];
    case VT.String:
      return grid.strs[idx] ?? '';
    case VT.Bool:
      return grid.nums[idx] === 1;
    case VT.Error:
      return { error: grid.strs[idx] ?? '#VALUE!' };
    default:
      return null;
  }
}

function gridContent(grid: SheetGrid, idx: number): CellContent {
  const f = grid.formulas.get(idx);
  if (f !== undefined && grid.flags[idx] & CellFlag.Formula) return { kind: 'formula', formula: f };
  const v = gridValue(grid, idx);
  return valueContent(v);
}

function valueContent(v: EngineValue): CellContent {
  if (v === null) return EMPTY;
  if (typeof v === 'number') return { kind: 'number', value: v };
  if (typeof v === 'string') return { kind: 'string', value: v };
  if (typeof v === 'boolean') return { kind: 'bool', value: v };
  return { kind: 'error', value: v.error };
}

function valueRaw(v: EngineValue): RawCellContent {
  return toRaw(valueContent(v));
}

/**
 * A cell's content as people read it: the formula, or the value as text.
 * Numbers stay exact; dates (by the cell's format) read as ISO dates.
 */
export function contentText(c: CellContent, numFmt = 'General', date1904 = false): string {
  switch (c.kind) {
    case 'empty':
      return '';
    case 'formula':
      return displayFormula(c.formula);
    case 'bool':
      return c.value ? 'TRUE' : 'FALSE';
    case 'number': {
      let isDate = false;
      try {
        isDate = numFmt !== 'General' && isDateFormat(numFmt);
      } catch {
        isDate = false;
      }
      if (isDate) {
        const serial = c.value + (date1904 ? 1462 : 0);
        return formatNumber(Number.isInteger(serial) ? 'yyyy-mm-dd' : 'yyyy-mm-dd hh:mm', serial, { throws: false });
      }
      return String(parseFloat(c.value.toPrecision(15)));
    }
    default:
      return c.value;
  }
}

export function sameContent(a: CellContent, b: CellContent): boolean {
  if (a.kind !== b.kind) return false;
  switch (a.kind) {
    case 'empty':
      return true;
    case 'formula':
      return a.formula === (b as typeof a).formula;
    default:
      return Object.is(a.value, (b as typeof a).value);
  }
}

function setValue(u: CellUpdate, v: EngineValue): void {
  if (v === null) {
    u.t = VT.None;
  } else if (typeof v === 'number') {
    u.t = VT.Number;
    u.n = v;
  } else if (typeof v === 'string') {
    u.t = VT.String;
    u.s = v;
  } else if (typeof v === 'boolean') {
    u.t = VT.Bool;
    u.n = v ? 1 : 0;
  } else {
    u.t = VT.Error;
    u.s = v.error;
  }
}

/** Ranges read by a formula, resolved to sheet indexes; null if not all can be resolved. */
function formulaReads(formula: string, ownSheet: number, sheetIndex: Map<string, number>): UntrustedFormula['reads'] {
  let tokens;
  try {
    tokens = tokenize(formula, { mergeRefs: true, allowTernary: false });
  } catch {
    return null;
  }
  const reads: { sheet: number; rect: Rect }[] = [];
  for (const t of tokens) {
    if (t.type === 'range_named' || t.type === 'structured' || t.type === 'unknown') return null;
    if (t.type !== 'range' && t.type !== 'range_beam') continue;
    const ref = parseA1Ref(t.value, { allowTernary: false });
    if (!ref || !('range' in ref)) return null;
    const ctx = ref.context ?? [];
    if (ctx.length > 1) return null;
    const sheet = ctx.length ? sheetIndex.get(ctx[0].toLowerCase()) : ownSheet;
    if (sheet === undefined) return null;
    const r = ref.range;
    reads.push({
      sheet,
      rect: {
        r1: r.top ?? 0,
        c1: r.left ?? 0,
        r2: r.bottom ?? r.top ?? MAX_ROWS - 1,
        c2: r.right ?? r.left ?? MAX_COLS - 1,
      },
    });
  }
  return reads;
}

/**
 * The editing session for one workbook. It keeps the file's parts and the
 * scanned sheets untouched, tracks edits as an overlay, keeps the
 * calculation engine in step, and on save rewrites only what changed.
 */
export class WorkbookSession {
  readonly fileName: string;
  readonly model: WorkbookModel;
  private readonly pkg: OoxmlPackage;
  /** The file as opened; null after a structural change, when the session is built from a package. */
  private readonly originalBytes: Uint8Array | null;
  private readonly timings: Record<string, number>;
  private engine?: CalcEngine;
  private report?: CalcReport;
  /** Per sheet: cells whose content or style differs from the file. */
  private readonly edits: Map<number, CellState>[];
  /** Per sheet: latest engine value of formula cells (edited or recalculated). */
  private readonly values: Map<number, EngineValue>[];
  /** Per sheet: Unsupported/Mismatch flags for the file's own formulas. */
  private readonly calcFlags: Map<number, number>[];
  /** Per sheet: Unsupported flag for formulas typed in this session. */
  private readonly newFormulaFlags: Map<number, number>[];
  private readonly stale: Set<number>[];
  private untrusted: UntrustedFormula[] = [];
  private readonly sheetIndex: Map<string, number>;
  /**
   * After a change of rows or columns, the file's saved formula values are the
   * ones from before it, not Excel's. Only the formulas the previous session
   * found calculated differently from Excel ("sheet:row:col", where they are
   * now) keep their saved value; the rest are calculated.
   */
  private comparedOnly?: Set<string>;

  /** Locale used to read typed and pasted dates and numbers. */
  readonly locale: string;
  /** Remove the VBA project when saving (macro policy "remove"). */
  private stripMacros = false;
  /** Live recalculation is off: formulas keep Excel's saved values and no engine is built. */
  private recalcOff = false;
  /** What the structural change that made this session broke, to tell the user. */
  readonly structureWarnings: Text[] = [];
  /**
   * What changed before the last structural change, carried to where those
   * cells are now: the change list, the cells to mark as edited, and a
   * description of each structural change.
   */
  private baked: { changes: CellChange[]; keys: Set<number>[]; structure: CellChange[] };

  /** The name the saved file should have; macro-free files can't keep .xlsm. */
  get saveFileName(): string {
    return this.stripMacros ? this.fileName.replace(/\.xlsm$/i, '.xlsx') : this.fileName;
  }

  private constructor(pkg: OoxmlPackage, bytes: Uint8Array | null, fileName: string, model: WorkbookModel, timings: Record<string, number>, locale: string) {
    this.locale = locale;
    this.pkg = pkg;
    this.originalBytes = bytes;
    this.fileName = fileName;
    this.model = model;
    this.timings = timings;
    const n = model.sheets.length;
    this.edits = Array.from({ length: n }, () => new Map());
    this.values = Array.from({ length: n }, () => new Map());
    this.calcFlags = Array.from({ length: n }, () => new Map());
    this.newFormulaFlags = Array.from({ length: n }, () => new Map());
    this.stale = Array.from({ length: n }, () => new Set());
    this.sheetIndex = new Map(model.sheets.map((s, i) => [s.info.name.toLowerCase(), i]));
    this.baked = { changes: [], keys: Array.from({ length: n }, () => new Set()), structure: [] };
  }

  static open(bytes: Uint8Array, fileName: string, locale = 'en-US', options: OpenOptions = {}): WorkbookSession {
    const t0 = performance.now();
    const pkg = OoxmlPackage.open(bytes, options.limits);
    const t1 = performance.now();
    const maxCells = options.limits?.maxCells ?? DEFAULT_MAX_CELLS;
    const model = readWorkbook(pkg, locale, { left: maxCells, max: maxCells });
    const t2 = performance.now();
    const cells = model.sheets.reduce((n, m) => n + (m.scan?.grid.cols.length ?? 0), 0);
    const macros = options.macros ?? 'keep';
    if (model.hasMacros && macros === 'block') {
      throw new TextError(msg('core.open.macrosBlocked'));
    }
    if (model.hasActiveContent && options.activeContent === 'block') {
      throw new TextError(msg('core.open.activeContentBlocked'));
    }
    if (model.hasMacros && macros === 'remove') {
      for (const f of model.findings) {
        if (f.kind !== 'macro') continue;
        // The VBA project's finding has no sheet; a macro sheet's names it.
        if (!f.sheet) f.detail = msg('core.finding.macros.removedDetail');
        else if (isMessage(f.title) && f.title.key === 'core.finding.macroSheet.title') f.detail = msg('core.finding.macroSheet.removedDetail');
      }
    }
    const session = new WorkbookSession(
      pkg,
      bytes,
      fileName,
      model,
      { unzipMs: Math.round(t1 - t0), parseMs: Math.round(t2 - t1) },
      locale,
    );
    session.stripMacros = model.hasMacros && macros === 'remove';
    const formulas = model.sheets.reduce((n, m) => n + (m.scan?.grid.formulas.size ?? 0), 0);
    const recalc = options.recalc ?? 'auto';
    session.recalcOff = recalc === 'off' || (recalc === 'auto' && (formulas > LIVE_RECALC_LIMITS.formulas || cells > LIVE_RECALC_LIMITS.cells));
    return session;
  }

  loadResult(): LoadResult {
    return {
      fileName: this.fileName,
      sheets: this.model.sheets.map((s) => s.info),
      grids: this.model.sheets.map((s) => s.scan?.grid ?? null),
      styles: this.model.styles.display,
      findings: this.model.findings,
      date1904: this.model.date1904,
      timings: this.timings,
      properties: this.model.properties,
    };
  }

  get editedCount(): number {
    let n = this.baked.changes.length + this.baked.structure.length;
    this.edits.forEach((m, s) => {
      for (const key of m.keys()) if (!this.baked.keys[s].has(key)) n++;
    });
    return n;
  }

  /**
   * Every cell that differs from the file as opened (what it held, and what
   * it holds now), and every structural change, described.
   */
  changes(): CellChange[] {
    const own = this.ownChanges();
    if (!this.baked.changes.length && !this.baked.structure.length) return own;
    // A cell changed before a structural change and again since: it was what the file held, it is what it holds now.
    const merged = new Map(this.baked.changes.map((c) => [`${c.sheet}!${c.ref}`, c]));
    for (const c of own) {
      const key = `${c.sheet}!${c.ref}`;
      const earlier = merged.get(key);
      if (!earlier) merged.set(key, c);
      else if (earlier.before === c.after && !c.formatOnly) merged.delete(key);
      else merged.set(key, { ...c, before: earlier.before, formatOnly: earlier.before === c.after || undefined });
    }
    return [...this.baked.structure, ...merged.values()];
  }

  private ownChanges(): CellChange[] {
    const out: CellChange[] = [];
    this.edits.forEach((edits, s) => {
      const name = this.model.sheets[s].info.name;
      const keys = [...edits.keys()].sort((a, b) => a - b);
      for (const key of keys) {
        const row = Math.floor(key / MAX_COLS);
        const col = key % MAX_COLS;
        const now = edits.get(key)!;
        const was = this.original(s, row, col);
        const fmt = (style: number) => this.model.styles.get(style).numFmt;
        const before = contentText(was.content, fmt(was.style), this.model.date1904);
        const after = contentText(now.content, fmt(now.style), this.model.date1904);
        out.push({ sheet: name, ref: toA1(row, col), before, after, ...(before === after ? { formatOnly: true } : {}) });
      }
    });
    return out;
  }

  // ---------------------------------------------------------------------------
  // Cell state

  private sheet(i: number): SheetModel {
    const m = this.model.sheets[i];
    if (!m) throw new Error(`No sheet at index ${i}`);
    return m;
  }

  private defaultStyle(m: SheetModel, row: number, col: number): number {
    const rowStyle = m.scan?.rowStyles.get(row);
    if (rowStyle !== undefined) return rowStyle;
    return m.colStyles.find((c) => col >= c.min && col <= c.max)?.style ?? 0;
  }

  private original(sheet: number, row: number, col: number): CellState & { idx: number } {
    const m = this.sheet(sheet);
    const grid = m.scan?.grid;
    const idx = grid ? findCell(grid, row, col) : -1;
    if (!grid || idx < 0) return { content: EMPTY, style: this.defaultStyle(m, row, col), idx: -1 };
    return { content: gridContent(grid, idx), style: grid.styles[idx], idx };
  }

  current(sheet: number, row: number, col: number): CellState {
    return this.edits[sheet].get(cellKey(row, col)) ?? this.original(sheet, row, col);
  }

  private isUntrustedOriginal(sheet: number, row: number, col: number): { idx: number } | null {
    const key = cellKey(row, col);
    if (this.edits[sheet].has(key)) return null;
    const grid = this.sheet(sheet).scan?.grid;
    if (!grid) return null;
    const idx = findCell(grid, row, col);
    if (idx < 0) return null;
    const flags = grid.flags[idx] | (this.calcFlags[sheet].get(key) ?? 0);
    return flags & (CellFlag.Unsupported | CellFlag.Mismatch | CellFlag.ArrayMember) ? { idx } : null;
  }

  private flagsFor(sheet: number, row: number, col: number): number {
    const key = cellKey(row, col);
    const edit = this.edits[sheet].get(key);
    let f: number;
    if (edit) {
      f = CellFlag.Edited;
      if (edit.content.kind === 'formula') f |= CellFlag.Formula | (this.newFormulaFlags[sheet].get(key) ?? 0);
    } else {
      const grid = this.sheet(sheet).scan?.grid;
      const idx = grid ? findCell(grid, row, col) : -1;
      f = idx >= 0 ? grid!.flags[idx] : 0;
      f |= this.calcFlags[sheet].get(key) ?? 0;
    }
    if (this.stale[sheet].has(key)) f |= CellFlag.Stale;
    if (this.baked.keys[sheet].has(key)) f |= CellFlag.Edited;
    return f;
  }

  private formulaValue(sheet: number, row: number, col: number): EngineValue {
    const key = cellKey(row, col);
    const v = this.values[sheet].get(key);
    if (v !== undefined) return v;
    if (!this.edits[sheet].has(key)) {
      const grid = this.sheet(sheet).scan?.grid;
      const idx = grid ? findCell(grid, row, col) : -1;
      if (idx >= 0 && !(grid!.flags[idx] & CellFlag.NoCache)) return gridValue(grid!, idx);
    }
    return this.engine?.value(sheet, row, col) ?? null;
  }

  private update(sheet: number, row: number, col: number): CellUpdate {
    const st = this.current(sheet, row, col);
    const u: CellUpdate = { sheet, row, col, t: VT.None, style: st.style, flags: this.flagsFor(sheet, row, col) };
    if (st.content.kind === 'formula') {
      u.f = st.content.formula;
      setValue(u, this.formulaValue(sheet, row, col));
    } else if (st.content.kind !== 'empty') {
      setValue(u, st.content.kind === 'error' ? { error: st.content.value } : st.content.value);
    }
    return u;
  }

  private setCurrent(w: Write): void {
    const key = cellKey(w.row, w.col);
    const orig = this.original(w.sheet, w.row, w.col);
    if (sameContent(orig.content, w.content) && orig.style === w.style) this.edits[w.sheet].delete(key);
    else this.edits[w.sheet].set(key, { content: w.content, style: w.style });
    this.values[w.sheet].delete(key);
    this.newFormulaFlags[w.sheet].delete(key);
    this.stale[w.sheet].delete(key);
  }

  private blockedReason(sheet: number, row: number, col: number): Text | undefined {
    const info = this.sheet(sheet).info;
    if (info.kind !== 'worksheet') return msg('core.reason.notWorksheet');
    // A protected sheet's cells are locked unless their format unlocks them, as Excel's input cells are.
    if (info.protected && !this.model.styles.get(this.current(sheet, row, col).style).unlocked) return msg('core.reason.locked');
    for (const b of info.blocked) if (inRect(b, row, col)) return b.reason;
    for (const m of info.merges) {
      if (inRect(m, row, col) && (row !== m.r1 || col !== m.c1)) return MERGED;
    }
    return undefined;
  }

  // ---------------------------------------------------------------------------
  // Calculation engine

  /** Builds the engine on first use and checks it against Excel's saved values. */
  ensureEngine(): CalcReport {
    if (this.report) return this.report;
    if (this.recalcOff) return (this.report = this.offReport());
    const t0 = performance.now();
    const sheets = this.model.sheets;
    // What isn't calculated and how often, keyed by content (each formula makes its own reason).
    const reasons = new Map<string, { reason: Unsupported; count: number }>();
    const cellReasons = new Map<string, Text>();
    const samples: CalcSample[] = [];
    const sample = (s: CalcSample) => {
      if (samples.length < 50) samples.push(s);
    };
    let formulaCells = 0;
    let unsupported = 0;

    // Decide up front which formulas the engine cannot handle; it gets
    // Excel's saved value for those instead, so cells that depend on them
    // still calculate correctly.
    const constant: Set<number>[] = sheets.map(() => new Set());
    sheets.forEach((m, s) => {
      const grid = m.scan?.grid;
      if (!grid) return;
      for (const [idx, formula] of grid.formulas) {
        const flags = grid.flags[idx];
        if (!(flags & CellFlag.Formula)) continue;
        formulaCells++;
        const row = rowOfIndex(grid, idx);
        const col = grid.cols[idx];
        let reason: Unsupported | undefined;
        if (flags & CellFlag.ArrayMember) reason = { kind: 'array' };
        else if (flags & CellFlag.Unsupported) reason = { kind: 'sharedFormula' };
        else reason = unsupportedReason(formula);
        if (!reason) continue;
        unsupported++;
        const key = JSON.stringify(reason);
        reasons.set(key, { reason, count: (reasons.get(key)?.count ?? 0) + 1 });
        cellReasons.set(`${s}:${row}:${col}`, unsupportedExplanation(reason));
        constant[s].add(idx);
        this.calcFlags[s].set(cellKey(row, col), CellFlag.Unsupported);
        this.untrusted.push({ target: { sheet: s, row, col }, reads: formulaReads(formula, s, this.sheetIndex) });
        sample({ sheet: m.info.name, ref: toA1(row, col), formula: displayFormula(formula), reason: unsupportedLabel(reason) });
      }
    });

    this.constant = constant;
    const needed = this.sheetsFormulasUse();
    const engine = new CalcEngine(
      sheets.map((_m, s) => (needed.has(s) ? this.engineSheet(s) : null)),
      this.model.names,
      this.model.date1904,
    );
    this.engine = engine;

    // Compare everything else with Excel's saved values. A formula that
    // disagrees only because something it reads disagrees is not at fault,
    // so only the earliest disagreeing formulas are pinned to Excel's value
    // each round, then the rest are compared again.
    const updates: CellUpdate[] = [];
    let noCachedValue = 0;
    let candidates: Pos[] = [];
    // Cells edited before the engine was built (with live recalculation off)
    // hold the user's content, and formulas reading them can no longer match
    // Excel's saved values: those are calculated, not compared.
    const edited: Pos[] = [];
    this.edits.forEach((m, s) => {
      for (const key of m.keys()) edited.push({ sheet: s, row: Math.floor(key / MAX_COLS), col: key % MAX_COLS });
    });
    const affected = new Set((edited.length ? engine.dependentsOf(edited) : []).map((p) => `${p.sheet}:${p.row}:${p.col}`));
    sheets.forEach((m, s) => {
      const grid = m.scan?.grid;
      if (!grid) return;
      for (const [idx, formula] of grid.formulas) {
        if (!(grid.flags[idx] & CellFlag.Formula) || constant[s].has(idx)) continue;
        const row = rowOfIndex(grid, idx);
        const col = grid.cols[idx];
        if (this.edits[s].has(cellKey(row, col))) continue;
        if (affected.has(`${s}:${row}:${col}`) || (this.comparedOnly && !this.comparedOnly.has(`${s}:${row}:${col}`))) {
          this.values[s].set(cellKey(row, col), engine.value(s, row, col));
          continue;
        }
        if (grid.flags[idx] & CellFlag.NoCache) {
          noCachedValue++;
          this.values[s].set(cellKey(row, col), engine.value(s, row, col));
          continue;
        }
        if (!isVolatile(formula)) candidates.push({ sheet: s, row, col });
      }
    });

    let mismatched = 0;
    for (let round = 0; round < 6 && candidates.length; round++) {
      const differing = candidates.filter((p) => {
        const grid = sheets[p.sheet].scan!.grid;
        return !valuesMatch(gridValue(grid, findCell(grid, p.row, p.col)), engine.value(p.sheet, p.row, p.col));
      });
      if (!differing.length) break;
      const downstream = new Set(engine.dependentsOf(differing).map((p) => `${p.sheet}:${p.row}:${p.col}`));
      let roots = differing.filter((p) => !downstream.has(`${p.sheet}:${p.row}:${p.col}`));
      if (!roots.length) roots = differing;
      const pins: EngineWrite[] = [];
      for (const p of roots) {
        const grid = sheets[p.sheet].scan!.grid;
        const idx = findCell(grid, p.row, p.col);
        const excel = gridValue(grid, idx);
        const ours = engine.value(p.sheet, p.row, p.col);
        mismatched++;
        cellReasons.set(`${p.sheet}:${p.row}:${p.col}`, msg('core.cellReason.mismatch', { engine: describeValue(ours), excel: describeValue(excel) }));
        this.calcFlags[p.sheet].set(cellKey(p.row, p.col), CellFlag.Mismatch);
        this.untrusted.push({ target: p, reads: formulaReads(grid.formulas.get(idx) ?? '', p.sheet, this.sheetIndex) });
        sample({
          sheet: sheets[p.sheet].info.name,
          ref: toA1(p.row, p.col),
          formula: displayFormula(grid.formulas.get(idx) ?? ''),
          reason: msg('core.calc.mismatch'),
          excel: describeValue(excel),
          engine: describeValue(ours),
        });
        pins.push({ ...p, raw: valueRaw(excel) });
      }
      engine.write(pins);
      const pinned = new Set(roots);
      candidates = differing.filter((p) => !pinned.has(p));
    }

    // The UI only knows the file's own flags; send the engine's verdicts too.
    const sent = new Set<string>();
    const send = (s: number, key: number) => {
      if (sent.has(`${s}:${key}`)) return;
      sent.add(`${s}:${key}`);
      updates.push(this.update(s, Math.floor(key / MAX_COLS), key % MAX_COLS));
    };
    this.values.forEach((m, s) => {
      for (const key of m.keys()) send(s, key);
    });
    this.calcFlags.forEach((m, s) => {
      for (const key of m.keys()) send(s, key);
    });

    const findings: Finding[] = [];
    if (unsupported) {
      const list = [...reasons.values()]
        .sort((a, b) => b.count - a.count)
        .map(({ reason, count }) => msg('core.calc.reasonCount', { reason: unsupportedLabel(reason), n: count }));
      findings.push({
        kind: 'formulaEngine',
        severity: 'warn',
        title: msg('core.finding.notRecalculated.title'),
        detail: msg('core.finding.notRecalculated.detail', { unsupported, total: formulaCells, reasons: joinTexts(list) }),
      });
    }
    if (mismatched) {
      findings.push({
        kind: 'formulaEngine',
        severity: 'warn',
        title: msg('core.finding.calculatedDifferently.title'),
        detail: msg('core.finding.calculatedDifferently.detail', { mismatched }),
      });
    }

    this.report = {
      buildMs: Math.round(performance.now() - t0),
      formulaCells,
      liveCells: formulaCells - unsupported - mismatched,
      unsupported,
      mismatched,
      noCachedValue,
      unsupportedFunctions: [...reasons.values()].map((r) => unsupportedLabel(r.reason)),
      samples,
      updates,
      findings,
      cellReasons,
    };
    return this.report;
  }

  /** Formula cells whose saved value is kept instead of being calculated, per sheet, by grid index. */
  private constant: Set<number>[] = [];

  /** A sheet's cells as the engine takes them: the file's cells with this session's edits. */
  private engineSheet(s: number): EngineSheet | null {
    const m = this.model.sheets[s];
    const grid = m.scan?.grid;
    if (!grid) return null;
    const constant = this.constant[s] ?? new Set<number>();
    const rows: RawCellContent[][] = new Array(grid.rowCount);
    for (let r = 0; r < grid.rowCount; r++) {
      const start = grid.rowPtr[r];
      const end = grid.rowPtr[r + 1];
      if (start === end) {
        rows[r] = [];
        continue;
      }
      const row: RawCellContent[] = new Array(grid.cols[end - 1] + 1).fill(null);
      for (let i = start; i < end; i++) {
        const formula = grid.formulas.get(i);
        row[grid.cols[i]] =
          formula !== undefined && grid.flags[i] & CellFlag.Formula && !constant.has(i) ? toRaw({ kind: 'formula', formula }) : valueRaw(gridValue(grid, i));
      }
      rows[r] = row;
    }
    for (const [key, st] of this.edits[s]) {
      const row = Math.floor(key / MAX_COLS);
      const col = key % MAX_COLS;
      while (rows.length <= row) rows.push([]);
      const line = rows[row];
      while (line.length <= col) line.push(null);
      line[col] = toRaw(st.content);
    }
    return { name: m.info.name, rows };
  }

  /**
   * The sheets the engine needs: those with formulas, and those any formula,
   * defined name, validation or conditional format names. A sheet of plain
   * data nothing refers to is left out until a formula does.
   */
  private sheetsFormulasUse(): Set<number> {
    const sheets = this.model.sheets;
    const names = sheets.map((m) => m.info.name.toLowerCase());
    const needed = new Set<number>();
    const mention = (text: string) => {
      if (needed.size === sheets.length) return;
      const lower = text.toLowerCase();
      names.forEach((n, i) => {
        if (!needed.has(i) && lower.includes(n)) needed.add(i);
      });
    };
    sheets.forEach((m, s) => {
      if (m.info.kind !== 'worksheet') return;
      if (m.scan?.grid.formulas.size || [...this.edits[s].values()].some((e) => e.content.kind === 'formula')) needed.add(s);
    });
    for (const m of sheets) {
      for (const f of m.scan?.grid.formulas.values() ?? []) mention(f);
      for (const e of [...this.edits[sheets.indexOf(m)].values()]) if (e.content.kind === 'formula') mention(e.content.formula);
      for (const r of [...m.info.ruleValidations, ...m.info.conditionalFormats]) {
        const fs = 'formulas' in r ? r.formulas : [r.formula1, r.formula2 ?? ''];
        fs.forEach(mention);
      }
      for (const l of m.info.listValidations) if (l.source.kind === 'formula') mention(l.source.formula);
    }
    for (const n of this.model.names) mention(n.formula);
    return needed;
  }

  /**
   * Adds sheets the engine left out that a formula needs: the sheet it is on
   * (its unqualified references) and any sheet it names.
   */
  private ensureSheetsFor(formula: string, own: number): void {
    const engine = this.engine;
    if (!engine) return;
    const lower = formula.toLowerCase();
    this.model.sheets.forEach((m, s) => {
      if (m.info.kind !== 'worksheet' || engine.has(s) || (s !== own && !lower.includes(m.info.name.toLowerCase()))) return;
      const content = this.engineSheet(s);
      if (content) engine.addSheet(s, content);
    });
  }

  /** The report when live recalculation is off, with no engine built. */
  private offReport(): CalcReport {
    const formulaCells = this.model.sheets.reduce((n, m) => n + (m.scan?.grid.formulas.size ?? 0), 0);
    const cells = this.model.sheets.reduce((n, m) => n + (m.scan?.grid.cols.length ?? 0), 0);
    return {
      buildMs: 0,
      formulaCells,
      liveCells: 0,
      unsupported: 0,
      mismatched: 0,
      noCachedValue: 0,
      unsupportedFunctions: [],
      samples: [],
      updates: [],
      cellReasons: new Map(),
      off: true,
      findings: [
        {
          kind: 'formulaEngine',
          severity: 'warn',
          title: msg('core.finding.recalcOff.title'),
          detail: msg('core.finding.recalcOff.detail', { formulas: formulaCells, cells }),
        },
      ],
    };
  }

  /**
   * Lets go of the calculation engine, which for a large workbook holds
   * hundreds of megabytes, while this session waits in the undo history.
   * Its values stay; going back to it builds the engine again, with the
   * edits it holds.
   */
  releaseEngine(): void {
    if (!this.engine) return;
    this.engine.destroy();
    this.engine = undefined;
    this.report = undefined;
    // Worked out again when the engine is rebuilt.
    this.untrusted = [];
    for (const m of this.calcFlags) m.clear();
  }

  /** How many cells the workbook holds, to size how much undo history it can keep. */
  get cellCount(): number {
    return this.model.sheets.reduce((n, m) => n + (m.scan?.grid.cols.length ?? 0), 0);
  }

  /** Turns live recalculation on: builds the engine, brings in this session's edits, and reports it. */
  enableRecalc(): CalcReport {
    if (!this.recalcOff) return this.ensureEngine();
    this.recalcOff = false;
    this.report = undefined;
    for (const m of this.newFormulaFlags) m.clear();
    for (const s of this.stale) s.clear();
    const report = this.ensureEngine();
    // Formulas typed while it was off get their values now.
    this.edits.forEach((m, s) => {
      for (const [key, st] of m) {
        if (st.content.kind !== 'formula') continue;
        if (unsupportedReason(st.content.formula)) this.newFormulaFlags[s].set(key, CellFlag.Unsupported);
        else this.values[s].set(key, this.engine!.value(s, Math.floor(key / MAX_COLS), key % MAX_COLS));
      }
    });
    report.updates = this.stateUpdates();
    return report;
  }

  /**
   * Marks formulas the engine cannot vouch for after the given cells changed.
   * The engine holds formulas it can't calculate as Excel's saved values, so it
   * doesn't know what they read: a chain of them is followed here, round by
   * round, along with the engine's own formulas that read each one.
   */
  private markStale(changed: Pos[], extraRoots: Pos[] = []): Pos[] {
    const out: Pos[] = [];
    const mark = (p: Pos): boolean => {
      const key = cellKey(p.row, p.col);
      if (this.stale[p.sheet].has(key)) return false;
      if (this.current(p.sheet, p.row, p.col).content.kind !== 'formula') return false;
      this.stale[p.sheet].add(key);
      out.push(p);
      return true;
    };
    let frontier = changed;
    let roots: Pos[] = [...extraRoots];
    for (let round = 0; frontier.length || roots.length; round++) {
      if (this.untrusted.length && frontier.length) {
        for (const u of this.untrusted) {
          if (this.stale[u.target.sheet].has(cellKey(u.target.row, u.target.col))) continue;
          if (!this.isUntrustedOriginal(u.target.sheet, u.target.row, u.target.col)) continue;
          // A formula that could read anything is hit by the first change only.
          const hit =
            (u.reads === null && round === 0) ||
            (u.reads !== null && frontier.some((c) => u.reads!.some((r) => r.sheet === c.sheet && inRect(r.rect, c.row, c.col))));
          if (hit) roots.push(u.target);
        }
      }
      if (!roots.length) break;
      const next: Pos[] = [];
      for (const r of roots) if (!extraRoots.includes(r) && mark(r)) next.push(r);
      for (const d of this.engine?.dependentsOf(roots) ?? []) if (mark(d)) next.push(d);
      frontier = next;
      roots = [];
    }
    return out;
  }

  // ---------------------------------------------------------------------------
  // Conditional formatting

  /** A cell's current value, as formulas and rules see it. */
  private cellValue(sheet: number, row: number, col: number): EngineValue {
    const { content } = this.current(sheet, row, col);
    switch (content.kind) {
      case 'empty':
        return null;
      case 'formula':
        return this.formulaValue(sheet, row, col);
      case 'error':
        return { error: content.value };
      default:
        return content.value;
    }
  }

  /** What the sheet's conditional formatting does to each cell, for the current values. */
  conditionalFormats(sheet: number): CfResult {
    this.ensureEngine();
    const info = this.sheet(sheet).info;
    let maxRow = info.maxRow;
    let maxCol = info.maxCol;
    for (const key of this.edits[sheet].keys()) {
      maxRow = Math.max(maxRow, Math.floor(key / MAX_COLS));
      maxCol = Math.max(maxCol, key % MAX_COLS);
    }
    const ev = this.evaluator;
    return evaluateConditionalFormats(info.conditionalFormats, sheet, {
      value: (row, col) => this.cellValue(sheet, row, col),
      calculate: (formula) => ev.calculate(formula, sheet),
      maxRow,
      maxCol,
    });
  }

  // ---------------------------------------------------------------------------
  // Data validation

  private get evaluator(): Evaluator {
    const engine = this.engine;
    if (!engine) {
      // With recalculation off, rules that need formulas are left for Excel to judge (#ERROR! counts as not calculable).
      return { calculate: () => ({ error: '#ERROR!' }), withValue: (_s, _r, _c, _v, fn) => fn() };
    }
    return {
      calculate: (formula, sheet) => {
        this.ensureSheetsFor(formula, sheet);
        return engine.calculate(engineFormula(formula.replace(/^=/, '')), sheet);
      },
      withValue: (sheet, row, col, content, fn) => {
        this.ensureSheetsFor('', sheet);
        // What the engine holds for the cell now, so it can be put back exactly.
        const untrusted = this.isUntrustedOriginal(sheet, row, col);
        const grid = this.sheet(sheet).scan?.grid;
        const prev = untrusted && grid ? valueRaw(gridValue(grid, untrusted.idx)) : toRaw(this.current(sheet, row, col).content);
        engine.write([{ sheet, row, col, raw: toRaw(content) }]);
        try {
          return fn();
        } finally {
          engine.write([{ sheet, row, col, raw: prev }]);
        }
      },
    };
  }

  /** Whether typed text passes the sheet's validation rule at `ruleIndex`, as Excel would judge it. */
  checkValidation(sheet: number, row: number, col: number, text: string, ruleIndex: number): boolean {
    this.ensureEngine();
    const rule = this.sheet(sheet).info.ruleValidations[ruleIndex];
    if (!rule) return true;
    const cur = this.current(sheet, row, col);
    const parsed = parseInput(text, this.model.styles.get(cur.style).numFmt, this.locale, this.model.date1904);
    if (parsed.error) return false;
    return checkRule(rule, sheet, row, col, parsed.content, text, this.evaluator);
  }

  /** The values (or range) a formula-based list rule offers at a cell. */
  listSource(sheet: number, formula: string): ListSource | undefined {
    this.ensureEngine();
    return evaluateListFormula(formula, sheet, { sheetIndex: this.sheetIndex, ownSheet: sheet, names: this.model.names }, this.evaluator);
  }

  // ---------------------------------------------------------------------------
  // Editing

  apply(op: EditOp): EditResult {
    this.ensureEngine();
    if (op.type === 'batch') return mergeEditResults(op.ops.map((o) => this.apply(o)));
    const writes = new Map<string, Write>();
    const warnings: Text[] = [];
    const newStyles: { index: number; style: DisplayStyle }[] = [];
    const put = (w: Write) => writes.set(`${w.sheet}:${w.row}:${w.col}`, w);
    const styleFor = (base: number, numFmt?: string) => {
      if (!numFmt) return base;
      const r = this.model.styles.withNumFmt(base, numFmt);
      if (r.created) newStyles.push({ index: r.index, style: this.model.styles.get(r.index) });
      return r.index;
    };

    switch (op.type) {
      case 'input': {
        const cur = this.current(op.sheet, op.row, op.col);
        const parsed = parseInput(op.text, this.model.styles.get(cur.style).numFmt, this.locale, this.model.date1904);
        if (parsed.error) return this.emptyResult(parsed.error);
        put({ sheet: op.sheet, row: op.row, col: op.col, content: parsed.content, style: styleFor(cur.style, parsed.numFmt) });
        break;
      }
      case 'pasteText': {
        let bad = 0;
        op.values.forEach((line, i) =>
          line.forEach((text, j) => {
            const row = op.row + i;
            const col = op.col + j;
            if (row >= MAX_ROWS || col >= MAX_COLS) return;
            const cur = this.current(op.sheet, row, col);
            const cellFmt = this.model.styles.get(cur.style).numFmt;
            const parsed = parseInput(text, cellFmt, this.locale, this.model.date1904);
            if (parsed.error) bad++;
            let content: CellContent = parsed.error ? { kind: 'string' as const, value: text } : parsed.content;
            // The text is rounded to what the source showed; its exact value wins, unless it is a
            // date from a workbook in the other date system.
            const exact = op.exact?.[i]?.[j];
            if (exact != null && content.kind === 'number') {
              const date = isDateFormat(parsed.numFmt ?? cellFmt);
              if (!date || Math.floor(exact) === Math.floor(content.value)) content = { kind: 'number', value: exact };
            }
            put({ sheet: op.sheet, row, col, content, style: styleFor(cur.style, parsed.numFmt) });
          }),
        );
        if (bad) warnings.push(msg('core.edit.pastedAsText', { count: bad }));
        break;
      }
      case 'clear': {
        for (const rect of op.rects) this.forEachCell(op.sheet, rect, (row, col, cur) => {
          if (cur.content.kind !== 'empty') put({ sheet: op.sheet, row, col, content: EMPTY, style: cur.style });
        });
        break;
      }
      case 'copy': {
        const h = op.src.r2 - op.src.r1 + 1;
        const w = op.src.c2 - op.src.c1 + 1;
        let lost = 0;
        for (let row = op.dst.r1; row <= op.dst.r2; row++) {
          for (let col = op.dst.c1; col <= op.dst.c2; col++) {
            const sr = op.src.r1 + ((row - op.dst.r1) % h);
            const sc = op.src.c1 + ((col - op.dst.c1) % w);
            const r = this.copiedContent(op.srcSheet, sr, sc, row, col);
            if (r.lost) lost++;
            put({ sheet: op.sheet, row, col, content: r.content, style: r.style });
          }
        }
        if (lost) warnings.push(msg('core.edit.copiedAsValues', { count: lost }));
        break;
      }
      case 'fill': {
        warnings.push(...this.planFillWrites(op.sheet, op.src, op.dst, put));
        break;
      }
      case 'sort': {
        const { rect } = op;
        const merged = this.sheet(op.sheet).info.merges.find((m) => m.r1 <= rect.r2 && m.r2 >= rect.r1 && m.c1 <= rect.c2 && m.c2 >= rect.c1);
        if (merged) return this.emptyResult(msg('core.edit.sortMerged', { range: rangeToA1(merged) }));
        const top = op.header ? rect.r1 + 1 : rect.r1;
        const rows = Array.from({ length: Math.max(0, rect.r2 - top + 1) }, (_, i) => top + i);
        const order = rows
          .map((row, i) => ({ row, i, keys: op.keys.map((k) => this.cellValue(op.sheet, row, k.col)) }))
          .sort((a, b) => {
            for (let k = 0; k < op.keys.length; k++) {
              const d = compareForSort(a.keys[k], b.keys[k], op.keys[k].descending);
              if (d) return d;
            }
            return a.i - b.i;
          });
        let lost = 0;
        order.forEach((from, i) => {
          const row = top + i;
          if (from.row === row) return;
          for (let col = rect.c1; col <= rect.c2; col++) {
            const r = this.copiedContent(op.sheet, from.row, col, row, col);
            if (r.lost) lost++;
            put({ sheet: op.sheet, row, col, content: r.content, style: r.style });
          }
        });
        if (lost) warnings.push(msg('core.edit.sortedAsValues', { count: lost }));
        break;
      }
      case 'restore': {
        for (const c of op.cells) put({ ...c });
        break;
      }
      case 'inputs': {
        for (const c of op.cells) {
          const cur = this.current(op.sheet, c.row, c.col);
          const parsed = parseInput(c.text, this.model.styles.get(cur.style).numFmt, this.locale, this.model.date1904);
          if (parsed.error) warnings.push(parsed.error);
          const content = parsed.error ? { kind: 'string' as const, value: c.text } : parsed.content;
          put({ sheet: op.sheet, row: c.row, col: c.col, content, style: styleFor(cur.style, parsed.numFmt) });
        }
        break;
      }
    }
    return this.commit([...writes.values()], warnings, newStyles);
  }

  private emptyResult(rejected: Text): EditResult {
    return { updates: [], undo: [], redo: [], skipped: 0, skippedReasons: [], warnings: [], newStyles: [], editedCount: this.editedCount, rejected };
  }

  /** Calls fn for every non-empty or edited cell inside rect. */
  private forEachCell(sheet: number, rect: Rect, fn: (row: number, col: number, cur: CellState) => void): void {
    const grid = this.sheet(sheet).scan?.grid;
    const seen = new Set<number>();
    if (grid) {
      for (let row = rect.r1; row <= Math.min(rect.r2, grid.rowCount - 1); row++) {
        for (let i = grid.rowPtr[row]; i < grid.rowPtr[row + 1]; i++) {
          const col = grid.cols[i];
          if (col < rect.c1 || col > rect.c2) continue;
          seen.add(cellKey(row, col));
          fn(row, col, this.current(sheet, row, col));
        }
      }
    }
    for (const [key, st] of this.edits[sheet]) {
      if (seen.has(key)) continue;
      const row = Math.floor(key / MAX_COLS);
      const col = key % MAX_COLS;
      if (inRect(rect, row, col)) fn(row, col, st);
    }
  }

  private copiedContent(srcSheet: number, sr: number, sc: number, row: number, col: number): CellState & { lost: boolean } {
    const src = this.current(srcSheet, sr, sc);
    if (src.content.kind !== 'formula') return { ...src, lost: false };
    const flags = this.flagsFor(srcSheet, sr, sc);
    if (flags & CellFlag.ArrayMember) {
      return { content: valueContent(this.formulaValue(srcSheet, sr, sc)), style: src.style, lost: true };
    }
    try {
      return { content: { kind: 'formula', formula: makeShifter(src.content.formula, sr, sc)(row, col) }, style: src.style, lost: false };
    } catch {
      return { content: valueContent(this.formulaValue(srcSheet, sr, sc)), style: src.style, lost: true };
    }
  }

  private planFillWrites(sheet: number, src: Rect, dst: Rect, put: (w: Write) => void): Text[] {
    // The destination may include the source; fill the part beyond it.
    const vertical = dst.r1 < src.r1 || dst.r2 > src.r2;
    const forward = vertical ? dst.r2 > src.r2 : dst.c2 > src.c2;
    const isDate = (style: number) => {
      try {
        return isDateFormat(this.model.styles.get(style).numFmt);
      } catch {
        return false;
      }
    };
    let lost = 0;
    const lines = vertical ? [src.c1, src.c2] : [src.r1, src.r2];
    for (let line = lines[0]; line <= lines[1]; line++) {
      const srcCells: { row: number; col: number }[] = [];
      const len = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
      for (let i = 0; i < len; i++) srcCells.push(vertical ? { row: src.r1 + i, col: line } : { row: line, col: src.c1 + i });
      const states = srcCells.map((p) => this.current(sheet, p.row, p.col));
      const sources: FillSource[] = states.map((s) => ({ content: s.content, isDate: isDate(s.style) }));
      const targets: { row: number; col: number; x: number }[] = [];
      if (vertical) {
        if (forward) for (let r = src.r2 + 1; r <= dst.r2; r++) targets.push({ row: r, col: line, x: r - src.r1 });
        else for (let r = src.r1 - 1; r >= dst.r1; r--) targets.push({ row: r, col: line, x: r - src.r1 });
      } else {
        if (forward) for (let c = src.c2 + 1; c <= dst.c2; c++) targets.push({ row: line, col: c, x: c - src.c1 });
        else for (let c = src.c1 - 1; c >= dst.c1; c--) targets.push({ row: line, col: c, x: c - src.c1 });
      }
      const plans = planFill(sources, targets.map((t) => t.x), this.model.date1904);
      plans.forEach((plan, i) => {
        const t = targets[i];
        const s = srcCells[plan.source];
        if (plan.kind === 'value') {
          put({ sheet, row: t.row, col: t.col, content: plan.content, style: states[plan.source].style });
        } else {
          const r = this.copiedContent(sheet, s.row, s.col, t.row, t.col);
          if (r.lost) lost++;
          put({ sheet, row: t.row, col: t.col, content: r.content, style: r.style });
        }
      });
    }
    return lost ? [msg('core.edit.filledAsValues', { count: lost })] : [];
  }

  private commit(writes: Write[], warnings: Text[], newStyles: { index: number; style: DisplayStyle }[]): EditResult {
    const engine = this.engine;
    const undo: CellSnapshot[] = [];
    const redo: CellSnapshot[] = [];
    // Keyed by content: the same message, made again for each cell, counts once.
    const reasons = new Map<string, Text>();
    let skipped = 0;
    const applied: Write[] = [];

    for (const w of writes) {
      const reason = this.blockedReason(w.sheet, w.row, w.col);
      if (reason) {
        if (reason !== MERGED) {
          skipped++;
          reasons.set(JSON.stringify(reason), reason);
        }
        continue;
      }
      const prev = this.current(w.sheet, w.row, w.col);
      if (sameContent(prev.content, w.content) && prev.style === w.style) continue;
      undo.push({ sheet: w.sheet, row: w.row, col: w.col, content: prev.content, style: prev.style });
      redo.push({ sheet: w.sheet, row: w.row, col: w.col, content: w.content, style: w.style });
      this.setCurrent(w);
      applied.push(w);
    }

    const engineWrites: EngineWrite[] = applied.map((w) => {
      const untrusted = this.isUntrustedOriginal(w.sheet, w.row, w.col);
      const grid = this.sheet(w.sheet).scan?.grid;
      // A formula over the engine's size limits never reaches it.
      const tooLarge = w.content.kind === 'formula' && unsupportedReason(w.content.formula)?.kind === 'tooLarge';
      return {
        sheet: w.sheet,
        row: w.row,
        col: w.col,
        raw: untrusted && grid ? valueRaw(gridValue(grid, untrusted.idx)) : tooLarge ? null : toRaw(w.content),
      };
    });
    for (const w of applied) if (w.content.kind === 'formula') this.ensureSheetsFor(w.content.formula, w.sheet);
    const changes = engine?.write(engineWrites) ?? [];

    // Formulas typed in this session that the engine can't calculate (all of them, with recalculation off).
    const newUnsupported: Pos[] = [];
    for (const w of applied) {
      if (w.content.kind !== 'formula' || this.isUntrustedOriginal(w.sheet, w.row, w.col)) continue;
      const key = cellKey(w.row, w.col);
      if (!engine || unsupportedReason(w.content.formula)) {
        this.newFormulaFlags[w.sheet].set(key, CellFlag.Unsupported);
        newUnsupported.push(w);
      } else {
        this.values[w.sheet].set(key, engine.value(w.sheet, w.row, w.col));
      }
    }

    const touched = new Map<string, Pos>();
    const touch = (p: Pos) => touched.set(`${p.sheet}:${p.row}:${p.col}`, { sheet: p.sheet, row: p.row, col: p.col });
    applied.forEach(touch);
    for (const c of changes) {
      const key = cellKey(c.row, c.col);
      if (touched.has(`${c.sheet}:${c.row}:${c.col}`)) continue;
      if (this.current(c.sheet, c.row, c.col).content.kind !== 'formula') continue;
      if (this.isUntrustedOriginal(c.sheet, c.row, c.col)) continue;
      if (this.newFormulaFlags[c.sheet].has(key)) continue;
      this.values[c.sheet].set(key, c.value);
      touch(c);
    }
    for (const p of this.markStale([...applied, ...changes], newUnsupported)) touch(p);

    return {
      updates: [...touched.values()].map((p) => this.update(p.sheet, p.row, p.col)),
      undo,
      redo,
      skipped,
      skippedReasons: [...reasons.values()],
      warnings,
      newStyles,
      editedCount: this.editedCount,
    };
  }

  // ---------------------------------------------------------------------------
  // Saving

  /**
   * The workbook as .xlsx bytes, with any custom document properties given.
   * An unchanged workbook comes back byte for byte.
   */
  save(properties?: Record<string, string>): Uint8Array {
    if (this.editedCount === 0) {
      if (!this.stripMacros && this.originalBytes) return this.originalBytes;
      const pkg = this.pkg.clone();
      if (this.stripMacros) removeMacros(pkg, this.model.workbookPath);
      return pkg.toBytes();
    }
    return writeWorkbook({ ...this.saveInput(), properties });
  }

  /** The workbook with every edit written in, as a package. */
  private savedPackage(): OoxmlPackage {
    return this.edits.some((m) => m.size) ? buildPackage({ ...this.saveInput(), stripMacros: false }) : this.pkg.clone();
  }

  private saveInput(): SaveInput {
    return {
      stripMacros: this.stripMacros,
      pkg: this.pkg.clone(),
      model: this.model,
      edits: this.edits,
      cachedValue: (sheet, row, col) => {
        const key = cellKey(row, col);
        if (this.newFormulaFlags[sheet].has(key) || this.isUntrustedOriginal(sheet, row, col)) return undefined;
        return this.values[sheet].get(key);
      },
      recalculated: this.values.map((m, s) =>
        new Set([...m.keys()].filter((k) => !this.edits[s].has(k) && !this.newFormulaFlags[s].has(k))),
      ),
      isArray: (sheet, row, col) => !!this.engine?.isArray(sheet, row, col),
      changedIn: (sheet, rect) => {
        // Rows or columns inserted or deleted on the sheet, cells edited before or since, and formulas recalculated.
        if (this.baked.structure.some((c) => c.sheet === this.model.sheets[sheet]?.info.name)) return true;
        for (const keys of [this.edits[sheet].keys(), this.baked.keys[sheet], this.values[sheet].keys()]) {
          for (const key of keys) if (!rect || inRect(rect, Math.floor(key / MAX_COLS), key % MAX_COLS)) return true;
        }
        return false;
      },
    };
  }

  // ---------------------------------------------------------------------------
  // Structural changes

  /**
   * The workbook after inserting or deleting rows or columns, or moving
   * cells, as a new session. The current state is written out, the change is
   * applied to every part that refers to the moved cells, and the result is
   * opened. This session is left as it was, so undo can return to it.
   */
  restructure(op: StructuralOp): WorkbookSession {
    const pkg = this.savedPackage();
    const { broken } = restructure(pkg, readWorkbook(pkg, this.locale), op);
    const next = new WorkbookSession(pkg, null, this.fileName, readWorkbook(pkg, this.locale), {}, this.locale);
    next.stripMacros = this.stripMacros;
    next.recalcOff = this.recalcOff;
    if (broken) {
      next.structureWarnings.push(msg('core.structure.brokenRefs', { count: broken, error: '#REF!' }));
    }

    // Changes so far follow their cells; changes to cells that were deleted go.
    const map = refMapFor(op);
    const names = this.model.sheets.map((m) => m.info.name);
    const index = (name: string) => names.findIndex((n) => n === name);
    for (const c of this.changes()) {
      const s = index(c.sheet);
      const at = parseA1(c.ref);
      if (c.note || s < 0 || !at) {
        if (c.note) next.baked.structure.push(c);
        continue;
      }
      const to = map.range(s, { r1: at.row, c1: at.col, r2: at.row, c2: at.col });
      if (to) next.baked.changes.push({ ...c, sheet: names[to.sheet], ref: toA1(to.rect.r1, to.rect.c1) });
    }
    this.model.sheets.forEach((_m, s) => {
      for (const key of new Set([...this.edits[s].keys(), ...this.baked.keys[s]])) {
        const row = Math.floor(key / MAX_COLS);
        const col = key % MAX_COLS;
        const to = map.range(s, { r1: row, c1: col, r2: row, c2: col });
        if (to) next.baked.keys[to.sheet].add(cellKey(to.rect.r1, to.rect.c1));
      }
    });
    // With the engine built, this session knows which saved values are Excel's own word on
    // formulas it calculates differently; those move with their cells, and the rest are recalculated.
    if (this.report && !this.recalcOff) {
      const compared = new Set<string>();
      this.calcFlags.forEach((flags, s) => {
        for (const [key, flag] of flags) {
          if (!(flag & CellFlag.Mismatch)) continue;
          const row = Math.floor(key / MAX_COLS);
          const col = key % MAX_COLS;
          const to = map.range(s, { r1: row, c1: col, r2: row, c2: col });
          if (to) compared.add(`${to.sheet}:${to.rect.r1}:${to.rect.c1}`);
        }
      });
      next.comparedOnly = compared;
    }
    next.baked.structure.push(describeStructure(op, names));
    return next;
  }

  /** Every cell whose state differs from the file the session was built from, for redrawing all of it. */
  stateUpdates(): CellUpdate[] {
    const out: CellUpdate[] = [];
    this.model.sheets.forEach((_m, s) => {
      const keys = new Set([...this.edits[s].keys(), ...this.values[s].keys(), ...this.calcFlags[s].keys(), ...this.stale[s], ...this.baked.keys[s]]);
      for (const key of keys) out.push(this.update(s, Math.floor(key / MAX_COLS), key % MAX_COLS));
    });
    return out;
  }

  /** The engine's report, if it has been built. */
  get calcReport(): CalcReport | undefined {
    return this.report;
  }
}

/**
 * Excel's sort order: numbers, then text (ignoring case), then logicals, then
 * errors, reversed when descending; blank cells always go last.
 */
export function compareForSort(a: EngineValue, b: EngineValue, descending: boolean): number {
  const blankA = a === null || a === '';
  const blankB = b === null || b === '';
  if (blankA || blankB) return blankA === blankB ? 0 : blankA ? 1 : -1;
  const rank = (v: EngineValue) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
  let d = rank(a) - rank(b);
  if (!d) {
    if (typeof a === 'number' && typeof b === 'number') d = a - b;
    else if (typeof a === 'string' && typeof b === 'string') d = a.localeCompare(b, undefined, { sensitivity: 'base', numeric: false });
    else if (typeof a === 'boolean' && typeof b === 'boolean') d = Number(a) - Number(b);
  }
  return descending ? -d : d;
}

/** A structural change in words, for the change list. */
function describeStructure(op: StructuralOp, names: string[]): CellChange {
  if (op.type === 'move') {
    const h = op.src.r2 - op.src.r1;
    const w = op.src.c2 - op.src.c1;
    const dst = rangeToA1({ r1: op.row, c1: op.col, r2: op.row + h, c2: op.col + w });
    const where = op.toSheet !== op.sheet ? `${names[op.toSheet]}!${dst}` : dst;
    return { sheet: names[op.sheet], ref: rangeToA1(op.src), before: '', after: '', note: msg('core.change.movedTo', { where }) };
  }
  const rows = op.type === 'insertRows' || op.type === 'deleteRows';
  const end = op.at + op.count - 1;
  const ref = rows ? `${op.at + 1}:${end + 1}` : `${colToLetters(op.at)}:${colToLetters(end)}`;
  const notes: Record<Exclude<StructuralOp['type'], 'move'>, MessageKey> = {
    insertRows: 'core.change.insertedRows',
    insertCols: 'core.change.insertedColumns',
    deleteRows: 'core.change.deletedRows',
    deleteCols: 'core.change.deletedColumns',
  };
  return { sheet: names[op.sheet], ref, before: '', after: '', note: msg(notes[op.type], { count: op.count }) };
}
