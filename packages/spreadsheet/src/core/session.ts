import { isMessage, msg, TextError, type MessageKey, type Text } from '../i18n/message';
import { parseA1Ref, tokenize } from '@borgar/fx';
import { format as formatNumber, isDateFormat } from 'numfmt';
import { isSpillError, MAX_SPILL_CELLS, spillFormula, tolerance, valuesMatch, type ArrayDims, type EngineChange, type EngineKind, type EngineSheet, type EngineWrite, type FormulaEngine, type RawContent, type Unsupported } from '@sitepen/formula-engine';
import { currentEngine, describeValue, toRaw, unsupportedExplanation, unsupportedLabel } from './engine';
import { renameInFormula, sameDefinition } from './definedNames';
import type { EngineValue } from '@sitepen/formula-engine';
import { evaluateConditionalFormats } from './conditionalFormat';
import { applyCommentChange, describeCommentChange, newGuid, noteText, sameNote } from './commentEdits';
import type { Person } from './comments';
import { checkRule, evaluateListFormula, type Evaluator } from './dataValidation';
import { mergeEditResults } from './editResult';
import { ENGLISH_LISTS, planFill, seriesLine, type FillSource } from './fill';
import { SheetLinks } from './hyperlinks';
import { copyLinks, isHyperlinkFormula, literalHyperlinkTarget, sortLinks } from './links';
import { displayFormula, engineFormula, functionNames, isVolatile, makeShifter, stripFunctionPrefixes } from '@sitepen/formula-engine/formula';
import { validateFormula } from './input';
import { parseInput } from './input';
import { SheetLayout } from './layout';
import { isBuiltinName, isName, nameProblem } from './names';
import { resolveTableRefs, type TablePlace } from './tableRefs';
import { cellKey, colToLetters, inRect, MAX_COLS, MAX_ROWS, parseA1, rangeToA1, toA1 } from '@sitepen/formula-engine/refs';
import { findCell } from './sheetScanner';
import {
  CellFlag,
  EMPTY,
  VT,
  type CalcReport,
  type CalcSample,
  type CellChange,
  type CellNote,
  type CellContent,
  type CellSnapshot,
  type CellUpdate,
  type DisplayStyle,
  type EditOp,
  type CfResult,
  type DefinedName,
  type EditResult,
  type BorderEdge,
  type BorderPlacement,
  type Finding,
  type FormatChange,
  type Hyperlink,
  type LayoutChange,
  type ListSource,
  type LoadResult,
  type NoteSnapshot,
  type NameChange,
  type ReadOnlyCells,
  type Rect,
  type SheetGrid,
  type SheetInfo,
} from './types';
import { joinTexts, readWorkbook, type SheetModel, type WorkbookModel } from './workbook';
import { buildPackage, removeMacros, writeWorkbook, type SaveInput } from './writer';
import { csvPackage, isCsvFile } from './csvWorkbook';
import { restructure } from './restructure';
import { refMapFor, type RefMap, type StructuralOp } from './refTransform';
import { applySheetOp } from './sheetOps';
import { isSheetOp, sheetIndexMap, type SheetOp, type WorkbookOp } from './sheetRules';
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

/**
 * A formula whose results spill, as a dynamic array: the room the engine
 * gives its result, and the cells it covers now.
 */
interface Spill {
  /** The size of its result, as the engine has been given it; 0 until it has been. */
  rows: number;
  cols: number;
  /** The cells it covers now, its own included; none while something is in its way (#SPILL!). */
  shown?: Rect;
  /** An array formula entered with Ctrl+Shift+Enter: it fills the range it was entered in, whatever its result's size. */
  fixed?: boolean;
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
  // Excel's saved value for a spill belongs to the formula it spills from.
  if (grid.flags[idx] & CellFlag.Spilled) return EMPTY;
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

function valueRaw(v: EngineValue): RawContent {
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

/** Calls fn for each cell of rect, row by row. */
function forEachIn(rect: Rect, fn: (row: number, col: number) => void): void {
  for (let row = rect.r1; row <= rect.r2; row++) for (let col = rect.c1; col <= rect.c2; col++) fn(row, col);
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
  private engine?: FormulaEngine;
  /** The kind of engine this workbook is calculated with, chosen before it was opened. */
  readonly engineKind: EngineKind = currentEngine();
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
  private places?: TablePlace[];
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
  /** Per sheet: its layout, once a row, column, pane or merge has been changed. */
  private layouts: (SheetLayout | undefined)[] = [];
  /** Per sheet: cells whose note or thread differs from the file, with what they hold now (null: neither). */
  private readonly noteEdits: Map<number, CellNote | null>[];
  /**
   * Per sheet, after a structural change: what the workbook as opened had
   * on the cells whose note or thread the file this session was built from
   * has changed. Changes to notes are listed against it.
   */
  private readonly noteBase: Map<number, CellNote | null>[];
  /** Everyone who has written a threaded comment, by person id: the file's people, then anyone since. */
  private readonly people: Map<string, Person>;
  /** Per sheet: the file's notes and threads by cell, made when first needed. */
  private fileNotes?: Map<number, CellNote>[];
  /** Where circular references closed after the last edit. */
  private circularCells: { sheet: number; row: number; col: number }[] = [];
  /** The names defined as LAMBDAs, in capitals, which formulas call as functions; worked out again when names change. */
  private lambdas?: Set<string>;
  /** The file's defined names, as opened; model.names holds them as they now are. */
  private readonly fileNames: DefinedName[];
  /** Per sheet: its hyperlinks, once one has been added, changed or removed. */
  private links: (SheetLinks | undefined)[] = [];
  /**
   * Per sheet: where each HYPERLINK formula's link went when the screen was
   * last told, so a link that changes while its value doesn't is told too.
   */
  private formulaLinks: Map<number, string | undefined>[] = [];
  /** Per sheet: formulas whose results spill, by cell key. */
  private spills: Map<number, Spill>[];
  /** Per sheet: the cells spills cover now, other than their formulas' own, with the formula's cell key. */
  private spilled: Map<number, number>[];

  /** The name the saved file should have; macro-free files can't keep .xlsm, and a CSV file is saved as a workbook. */
  get saveFileName(): string {
    if (isCsvFile(this.fileName)) return this.fileName.replace(/\.\w+$/, '.xlsx');
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
    this.noteEdits = Array.from({ length: n }, () => new Map());
    this.noteBase = Array.from({ length: n }, () => new Map());
    this.people = new Map(model.people.map((p) => [p.id, p]));
    this.fileNames = model.names.map((d) => ({ ...d }));
    this.spills = Array.from({ length: n }, () => new Map());
    this.spilled = Array.from({ length: n }, () => new Map());
  }

  static open(bytes: Uint8Array, fileName: string, locale = 'en-US', options: OpenOptions = {}): WorkbookSession {
    const t0 = performance.now();
    const maxCells = options.limits?.maxCells ?? DEFAULT_MAX_CELLS;
    // A CSV file becomes a workbook package of one sheet, so it is edited and saved like any other.
    const csv = isCsvFile(fileName);
    const pkg = csv ? csvPackage(bytes, fileName, locale, maxCells, options.limits) : OoxmlPackage.open(bytes, options.limits);
    const t1 = performance.now();
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
      csv ? null : bytes,
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
      // As they now are: a session gone back to after a change of rows or columns has edits the file it was built from doesn't.
      sheets: this.model.sheets.map((_s, i) => this.infoNow(i)),
      names: this.nameList(),
      grids: this.model.sheets.map((s) => s.scan?.grid ?? null),
      styles: this.model.styles.display,
      findings: this.model.findings,
      date1904: this.model.date1904,
      timings: this.timings,
      properties: this.model.properties,
      structureLocked: this.model.structureLocked,
    };
  }

  get editedCount(): number {
    let n = this.baked.changes.length + this.baked.structure.length + this.noteChanges().length + this.layoutNotes().length + this.nameNotes().length;
    this.edits.forEach((m, s) => {
      for (const key of m.keys()) if (!this.baked.keys[s].has(key)) n++;
    });
    return n;
  }

  /** How each sheet's layout and hyperlinks differ from the file's. */
  private layoutNotes(): CellChange[] {
    return this.model.sheets.flatMap((m, s) => [...(this.layouts[s]?.notes(m.info.name) ?? []), ...(this.links[s]?.notes(m.info.name) ?? [])]);
  }

  /** How the names differ from the file's, for the change list. */
  private nameNotes(): CellChange[] {
    const out: CellChange[] = [];
    const note = (n: DefinedName, text: Text) =>
      out.push({ sheet: n.localSheet !== undefined ? (this.model.sheets[n.localSheet]?.info.name ?? '') : '', ref: n.name, before: '', after: '', note: text });
    const now = new Map<number, DefinedName>();
    for (const n of this.model.names) {
      if (n.fileIndex !== undefined) now.set(n.fileIndex, n);
      else note(n, msg('names.change.defined', { name: n.name, formula: displayFormula(n.formula) }));
    }
    this.fileNames.forEach((was, i) => {
      const is = now.get(i);
      if (!is) note(was, msg('names.change.deleted', { name: was.name }));
      else if (sameDefinition(is, was)) return;
      else if (is.name !== was.name) note(is, msg('names.change.renamed', { from: was.name, name: is.name, formula: displayFormula(is.formula) }));
      else if (is.formula === was.formula && is.localSheet === was.localSheet) note(is, msg('names.change.commented', { name: is.name }));
      else note(is, msg('names.change.changed', { name: is.name, formula: displayFormula(is.formula) }));
    });
    return out;
  }

  /** The names as they now are, copied. */
  private nameList(): DefinedName[] {
    return this.model.names.map((d) => ({ ...d }));
  }

  /** A sheet's info with its notes, links and layout as they now are, not as the file had them. */
  private infoNow(i: number): SheetInfo {
    let info = this.model.sheets[i].info;
    if (this.noteEdits[i].size) info = { ...info, notes: this.notesOf(i) };
    const links = this.links[i];
    if (links) info = { ...info, links: [...links.list] };
    const layout = this.layouts[i];
    if (layout?.changed) info = { ...info, ...layout.infoNow(info) };
    return info;
  }

  private layoutOf(sheet: number): SheetLayout {
    return (this.layouts[sheet] ??= new SheetLayout(this.sheet(sheet).info));
  }

  private linksOf(sheet: number): SheetLinks {
    return (this.links[sheet] ??= new SheetLinks(this.sheet(sheet).info.links));
  }

  /** A change to a sheet's links, unless it changes nothing or the sheet is protected, whose links stay as they are. */
  private linkChange(sheet: number, change: { add: Hyperlink[]; remove: Hyperlink[] }): LayoutChange[] {
    if ((!change.add.length && !change.remove.length) || this.sheet(sheet).info.protected) return [];
    return [{ sheet, op: { kind: 'links', ...change } }];
  }

  private mergesOf(sheet: number): Rect[] {
    return this.layouts[sheet]?.mergeList ?? this.sheet(sheet).info.merges;
  }

  /**
   * Every cell that differs from the file as opened (what it held, and what
   * it holds now), every structural change, described, and every cell whose
   * note or comment thread changed.
   */
  changes(): CellChange[] {
    const notes = this.noteChanges().map(({ sheet, key, before, after }) => ({
      sheet: this.model.sheets[sheet].info.name,
      ref: toA1(Math.floor(key / MAX_COLS), key % MAX_COLS),
      before: noteText(before),
      after: noteText(after),
      comment: describeCommentChange(before, after),
    }));
    const own = this.ownChanges();
    if (!this.baked.changes.length && !this.baked.structure.length) return [...own, ...notes];
    // A cell changed before a structural change and again since: it was what the file held, it is what it holds now.
    const merged = new Map(this.baked.changes.map((c) => [`${c.sheet}!${c.ref}`, c]));
    for (const c of own) {
      const key = `${c.sheet}!${c.ref}`;
      const earlier = merged.get(key);
      if (!earlier) merged.set(key, c);
      else if (earlier.before === c.after && !c.formatOnly) merged.delete(key);
      else merged.set(key, { ...c, before: earlier.before, formatOnly: earlier.before === c.after || undefined });
    }
    return [...this.baked.structure, ...merged.values(), ...notes];
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
        // A spilled value reads as the value it is, before and after.
        const savedSpill = was.idx >= 0 && this.sheet(s).scan!.grid.flags[was.idx] & CellFlag.Spilled ? gridValue(this.sheet(s).scan!.grid, was.idx) : undefined;
        const spill = now.content.kind === 'empty' ? this.spilledValue(s, row, col) : undefined;
        const before = contentText(savedSpill !== undefined ? valueContent(savedSpill) : was.content, fmt(was.style), this.model.date1904);
        const after = contentText(spill !== undefined ? valueContent(spill) : now.content, fmt(now.style), this.model.date1904);
        out.push({ sheet: name, ref: toA1(row, col), before, after, ...(before === after ? { formatOnly: true } : {}) });
      }
    });
    out.push(...this.layoutNotes(), ...this.nameNotes());
    return out;
  }

  // ---------------------------------------------------------------------------
  // Notes and comment threads

  private fileNote(sheet: number, row: number, col: number): CellNote | null {
    this.fileNotes ??= this.model.sheets.map((m) => new Map(m.info.notes.map((n) => [cellKey(n.row, n.col), n])));
    return this.fileNotes[sheet]?.get(cellKey(row, col)) ?? null;
  }

  /** A cell's note or comment thread as it is now, or null. */
  noteAt(sheet: number, row: number, col: number): CellNote | null {
    const edits = this.noteEdits[sheet];
    const key = cellKey(row, col);
    return edits?.has(key) ? (edits.get(key) ?? null) : this.fileNote(sheet, row, col);
  }

  /** A sheet's notes and threads as they are now. */
  private notesOf(sheet: number): CellNote[] {
    const edits = this.noteEdits[sheet];
    const kept = this.model.sheets[sheet].info.notes.filter((n) => !edits.has(cellKey(n.row, n.col)));
    return [...kept, ...[...edits.values()].filter((n): n is CellNote => !!n)];
  }

  /**
   * Cells whose note or thread differs from the workbook as it was opened,
   * with what each had then and has now, in order.
   */
  private noteChanges(): { sheet: number; key: number; before: CellNote | null; after: CellNote | null }[] {
    const out: { sheet: number; key: number; before: CellNote | null; after: CellNote | null }[] = [];
    this.model.sheets.forEach((_m, s) => {
      const base = this.noteBase[s];
      for (const key of [...new Set([...this.noteEdits[s].keys(), ...base.keys()])].sort((a, b) => a - b)) {
        const row = Math.floor(key / MAX_COLS);
        const col = key % MAX_COLS;
        const before = base.has(key) ? (base.get(key) ?? null) : this.fileNote(s, row, col);
        const after = this.noteAt(s, row, col);
        if (!sameNote(before, after)) out.push({ sheet: s, key, before, after });
      }
    });
    return out;
  }

  /** Puts notes and threads as given: what they were, for undo, and are, for redo. */
  private putNotes(notes: NoteSnapshot[]): { undo: NoteSnapshot[]; redo: NoteSnapshot[] } {
    const undo: NoteSnapshot[] = [];
    const redo: NoteSnapshot[] = [];
    for (const n of notes) {
      const was = this.noteAt(n.sheet, n.row, n.col);
      if (sameNote(was, n.note)) continue;
      // Restoring applies them in order, so the earliest goes last.
      undo.unshift({ sheet: n.sheet, row: n.row, col: n.col, note: was });
      const note = n.note && { ...n.note, row: n.row, col: n.col };
      redo.push({ sheet: n.sheet, row: n.row, col: n.col, note });
      const key = cellKey(n.row, n.col);
      if (sameNote(this.fileNote(n.sheet, n.row, n.col), note)) this.noteEdits[n.sheet].delete(key);
      else this.noteEdits[n.sheet].set(key, note);
      // Someone a restored thread names who isn't in the file yet. (One the file names without listing stays unlisted.)
      for (const e of note?.entries ?? []) {
        if (e.personId && e.author !== undefined && !this.people.has(e.personId)) {
          this.people.set(e.personId, { id: e.personId, displayName: e.author, userId: e.author, providerId: 'None' });
        }
      }
    }
    return { undo, redo };
  }

  /** Whoever writes a threaded comment: the person the file has with that name and no account, or a new one. */
  private personFor(author: string): string {
    for (const p of this.people.values()) if (p.displayName === author && (p.providerId ?? 'None') === 'None') return p.id;
    const id = newGuid();
    this.people.set(id, { id, displayName: author, userId: author, providerId: 'None' });
    return id;
  }

  /**
   * Why a cell's note or comments can't be changed, if they can't. As in
   * Excel, a protected sheet keeps them as they are unless its protection
   * lets objects be edited; whether the cell itself is locked doesn't matter.
   */
  private commentBlockedReason(sheet: number, row: number, col: number): Text | undefined {
    const info = this.sheet(sheet).info;
    if (info.kind !== 'worksheet') return msg('core.reason.notWorksheet');
    if (info.protected && info.objectsLocked) return msg('core.reason.commentsLocked');
    if (info.merges.some((m) => inRect(m, row, col) && (row !== m.r1 || col !== m.c1))) return msg('sheet.mergedPart');
    return undefined;
  }

  private applyComment(op: Extract<EditOp, { type: 'comment' }>): EditResult {
    const reason = this.commentBlockedReason(op.sheet, op.row, op.col);
    if (reason) return this.emptyResult(reason);
    let note: CellNote | null;
    try {
      note = applyCommentChange(this.noteAt(op.sheet, op.row, op.col), op.row, op.col, op.change, { person: (author) => this.personFor(author) });
    } catch (err) {
      if (err instanceof TextError) return this.emptyResult(err.text);
      throw err;
    }
    const { undo, redo } = this.putNotes([{ sheet: op.sheet, row: op.row, col: op.col, note }]);
    return { ...this.emptyResult(), noteUndo: undo, noteRedo: redo };
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

  /** Why the engine can't calculate a formula in a cell, or undefined when it can. */
  private unsupported(formula: string, sheet: number, row: number, col: number): Unsupported | undefined {
    this.lambdas ??= new Set(this.model.names.filter((n) => /^\s*LAMBDA\s*\(/i.test(stripFunctionPrefixes(n.formula))).map((n) => n.name.toUpperCase()));
    return this.engineKind.unsupportedReason(this.engineText(formula, sheet, row, col), this.lambdas);
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
      // Whether Excel's spill still reaches it is said below.
      f = idx >= 0 ? grid!.flags[idx] & ~CellFlag.Spilled : 0;
      f |= this.calcFlags[sheet].get(key) ?? 0;
    }
    if (this.stale[sheet].has(key)) f |= CellFlag.Stale;
    if (this.baked.keys[sheet].has(key)) f |= CellFlag.Edited;
    const anchor = this.spilled[sheet].get(key);
    if (anchor !== undefined) {
      f |= CellFlag.Spilled;
      // Values spilled from a formula that is out of date are too.
      if (this.stale[sheet].has(anchor)) f |= CellFlag.Stale;
    }
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
    const key = cellKey(row, col);
    const anchor = this.spilled[sheet].get(key);
    if (st.content.kind === 'formula') {
      u.f = st.content.formula;
      setValue(u, this.formulaValue(sheet, row, col));
      const spill = this.spills[sheet].get(key);
      const rect = spill && this.spillRange(row, col, spill);
      if (rect && (rect.r2 > rect.r1 || rect.c2 > rect.c1)) u[spill.fixed ? 'array' : 'spill'] = rect;
    } else if (anchor !== undefined) {
      setValue(u, this.engine?.value(sheet, row, col) ?? null);
      const spill = this.spills[sheet].get(anchor);
      if (spill?.shown) u[spill.fixed ? 'array' : 'spill'] = spill.shown;
    } else if (st.content.kind !== 'empty') {
      setValue(u, st.content.kind === 'error' ? { error: st.content.value } : st.content.value);
    }
    const links = (this.formulaLinks[sheet] ??= new Map());
    if (st.content.kind === 'formula' && isHyperlinkFormula(st.content.formula)) {
      const link = this.formulaLink(sheet, row, col, st.content.formula);
      links.set(cellKey(row, col), link);
      if (link !== undefined) u.link = link;
    } else {
      links.delete(cellKey(row, col));
    }
    return u;
  }

  /** Where a HYPERLINK formula links to now: as the engine worked it out, else as the formula writes it. */
  private formulaLink(sheet: number, row: number, col: number, formula: string): string | undefined {
    const calculated = this.isUntrustedOriginal(sheet, row, col) || this.newFormulaFlags[sheet].has(cellKey(row, col)) ? undefined : this.engine?.hyperlink(sheet, row, col);
    return calculated ?? literalHyperlinkTarget(formula);
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
    for (const m of this.mergesOf(sheet)) {
      if (inRect(m, row, col) && (row !== m.r1 || col !== m.c1)) return MERGED;
    }
    return undefined;
  }

  // ---------------------------------------------------------------------------
  // Calculation engine

  /** The workbook's Excel tables, where structured references point. */
  private get tablePlaces(): TablePlace[] {
    return (this.places ??= this.model.sheets.flatMap((m, sheet) =>
      m.info.tables.map((t) => ({ name: t.name, sheet, sheetName: m.info.name, rect: t.rect, headerRows: t.headerRows, totalsRows: t.totalsRows, columns: t.columns })),
    ));
  }

  /**
   * A formula as the engine gets it, with table references written as the
   * cells they stand for. One that can't be resolved stays as it is, and the
   * engine treats the formula as one it can't calculate.
   */
  private engineText(formula: string, sheet: number, row: number, col: number): string {
    if (formula.indexOf('[') === -1) return formula;
    return resolveTableRefs(formula, { sheet, row, col }, this.tablePlaces) ?? formula;
  }

  private engineContent(content: CellContent, sheet: number, row: number, col: number): CellContent {
    return content.kind === 'formula' ? { kind: 'formula', formula: this.engineText(content.formula, sheet, row, col) } : content;
  }

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
        const spill = this.savedSpill(s, idx, row, col);
        if (spill) {
          // An array formula, dynamic or entered with Ctrl+Shift+Enter, is calculated here when the engine can calculate its formula.
          reason = this.unsupported(formula, s, row, col);
          this.markFileArray(grid, idx, spill, !reason);
        } else if (flags & CellFlag.ArrayMember) reason = { kind: 'array' };
        else if (flags & CellFlag.Unsupported) reason = { kind: 'sharedFormula' };
        else reason = this.unsupported(formula, s, row, col);
        if (!reason) continue;
        unsupported++;
        const key = JSON.stringify(reason);
        reasons.set(key, { reason, count: (reasons.get(key)?.count ?? 0) + 1 });
        cellReasons.set(`${s}:${row}:${col}`, unsupportedExplanation(reason));
        constant[s].add(idx);
        this.calcFlags[s].set(cellKey(row, col), CellFlag.Unsupported);
        this.untrusted.push({ target: { sheet: s, row, col }, reads: formulaReads(this.engineText(formula, s, row, col), s, this.sheetIndex) });
        sample({ sheet: m.info.name, ref: toA1(row, col), formula: displayFormula(formula), reason: unsupportedLabel(reason) });
      }
    });

    this.constant = constant;
    const needed = this.sheetsFormulasUse();
    const engine = this.engineKind.create(
      sheets.map((_m, s) => (needed.has(s) ? this.engineSheet(s) : null)),
      // Names can use table references too; formulas see them as cells, as everywhere else.
      this.model.names.map((n) => ({ ...n, formula: this.engineText(n.formula, n.localSheet ?? 0, -1, -1) })),
      this.model.date1904,
      {
        formulaText: (s, row, col) => {
          const c = this.current(s, row, col).content;
          return c.kind === 'formula' ? displayFormula(c.formula) : undefined;
        },
      },
      this.model.iteration,
    );
    this.engine = engine;

    // The file's dynamic arrays went in with the room Excel saved for them, and
    // those whose results are larger or smaller now get room for those. Formulas
    // typed before the engine was built get theirs.
    const typed: Pos[] = [];
    this.edits.forEach((m, s) => {
      for (const [key, st] of m) if (st.content.kind === 'formula' && engine.has(s)) typed.push({ sheet: s, row: Math.floor(key / MAX_COLS), col: key % MAX_COLS });
    });
    this.fitSpills(typed, this.spillAnchors());
    this.refreshSpills([...this.spillAnchors(), ...typed]);

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

    // Iterating, Excel stops working a circle round once no value moves by
    // more than the maximum change, so values in and after one agree to that.
    const closing = this.model.iteration ? engine.circular() : [];
    const looped = new Set([...closing, ...(closing.length ? engine.dependentsOf(closing) : [])].map((p) => `${p.sheet}:${p.row}:${p.col}`));
    let mismatched = 0;
    for (let round = 0; round < 6 && candidates.length; round++) {
      const differing = candidates.filter((p) => {
        const grid = sheets[p.sheet].scan!.grid;
        const idx = findCell(grid, p.row, p.col);
        const loose = looped.has(`${p.sheet}:${p.row}:${p.col}`) ? this.model.iteration!.delta : 0;
        const within = Math.max(tolerance(functionNames(grid.formulas.get(idx) ?? '')), loose);
        return !valuesMatch(gridValue(grid, idx), engine.value(p.sheet, p.row, p.col), within) || this.spillDiffers(p);
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
        this.untrusted.push({ target: p, reads: formulaReads(this.engineText(grid.formulas.get(idx) ?? '', p.sheet, p.row, p.col), p.sheet, this.sheetIndex) });
        sample({
          sheet: sheets[p.sheet].info.name,
          ref: toA1(p.row, p.col),
          formula: displayFormula(grid.formulas.get(idx) ?? ''),
          reason: msg('core.calc.mismatch'),
          excel: describeValue(excel),
          engine: describeValue(ours),
        });
        pins.push({ ...p, raw: valueRaw(excel) }, ...this.keepFileArray(p));
      }
      // Spills may now be in the way of these, or have room taken from them.
      this.fitSpills([], engine.write(pins));
      this.refreshSpills(this.spillAnchors());
      const pinned = new Set(roots);
      candidates = differing.filter((p) => !pinned.has(p));
    }

    // The file's dynamic arrays that spill here are no longer locked.
    const liveArrays: { sheet: number; rect: Rect }[] = [];
    sheets.forEach((m, s) => {
      const scan = m.scan;
      for (const a of scan?.arrayRanges ?? []) {
        if (!a.dynamic || scan!.grid.flags[a.master] & CellFlag.ArrayMember) continue;
        const rect = { r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 };
        liveArrays.push({ sheet: s, rect });
        m.info.blocked = m.info.blocked.filter((b) => !(b.kind === 'array' && b.r1 === a.r1 && b.c1 === a.c1 && b.r2 === a.r2 && b.c2 === a.c2));
      }
    });

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
    for (const p of this.spillCells()) send(p.sheet, cellKey(p.row, p.col));
    // Where HYPERLINK formulas link to is worked out by the engine.
    sheets.forEach((m, s) => {
      const grid = m.scan?.grid;
      if (!grid) return;
      for (const [idx, formula] of grid.formulas) if (isHyperlinkFormula(formula)) send(s, cellKey(rowOfIndex(grid, idx), grid.cols[idx]));
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

    this.circularCells = this.circles(engine);
    this.report = {
      circular: this.circularCells,
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
      ...(liveArrays.length && { liveArrays }),
    };
    return this.report;
  }

  /**
   * Every cell a spill covers, the spilling formulas' own, and the cells
   * holding Excel's saved values for the file's spills, which show what the
   * spills show now: those whose display the file's grid doesn't give.
   */
  private spillCells(): Pos[] {
    const out = this.spillAnchors();
    this.spilled.forEach((m, sheet) => {
      for (const key of m.keys()) out.push({ sheet, row: Math.floor(key / MAX_COLS), col: key % MAX_COLS });
    });
    this.model.sheets.forEach((m, sheet) => {
      const grid = m.scan?.grid;
      if (!grid) return;
      for (let row = 0; row < grid.rowCount; row++) {
        for (let i = grid.rowPtr[row]; i < grid.rowPtr[row + 1]; i++) if (grid.flags[i] & CellFlag.Spilled) out.push({ sheet, row, col: grid.cols[i] });
      }
    });
    return out;
  }

  /** Formula cells whose saved value is kept instead of being calculated, per sheet, by grid index. */
  private constant: Set<number>[] = [];

  /** A sheet's cells as the engine takes them: the file's cells with this session's edits. */
  private engineSheet(s: number): EngineSheet | null {
    const m = this.model.sheets[s];
    const grid = m.scan?.grid;
    if (!grid) return null;
    const constant = this.constant[s] ?? new Set<number>();
    const rows: RawContent[][] = new Array(grid.rowCount);
    for (let r = 0; r < grid.rowCount; r++) {
      const start = grid.rowPtr[r];
      const end = grid.rowPtr[r + 1];
      if (start === end) {
        rows[r] = [];
        continue;
      }
      const row: RawContent[] = new Array(grid.cols[end - 1] + 1).fill(null);
      for (let i = start; i < end; i++) {
        // Excel's saved values for a spill are left out, so the formula can spill there.
        if (grid.flags[i] & CellFlag.Spilled) continue;
        const col = grid.cols[i];
        const formula = grid.formulas.get(i);
        if (formula === undefined || !(grid.flags[i] & CellFlag.Formula) || constant.has(i)) {
          row[col] = valueRaw(gridValue(grid, i));
          continue;
        }
        const raw = String(toRaw({ kind: 'formula', formula: this.engineText(formula, s, r, col) }));
        // A dynamic array starts with the room Excel saved for its results; an older array formula keeps the range it was entered in.
        const saved = this.edits[s].has(cellKey(r, col)) ? undefined : this.savedSpill(s, i, r, col);
        const fixed = !!saved && !m.scan!.fmeta.get(i)?.dynamic;
        if (saved) this.spills[s].set(cellKey(r, col), { rows: saved.r2 - saved.r1 + 1, cols: saved.c2 - saved.c1 + 1, ...(fixed && { fixed }) });
        // An array formula entered with Ctrl+Shift+Enter in one cell works on whole ranges, where an ordinary older formula meets its row or column.
        const oneCellArray = !saved && !this.edits[s].has(cellKey(r, col)) && m.scan!.fmeta.get(i)?.kind === 'array';
        row[col] = saved ? spillFormula(raw, saved.r2 - saved.r1 + 1, saved.c2 - saved.c1 + 1, fixed) : oneCellArray ? spillFormula(raw, 1, 1, true) : raw;
      }
      rows[r] = row;
    }
    for (const [key, st] of this.edits[s]) {
      const row = Math.floor(key / MAX_COLS);
      const col = key % MAX_COLS;
      while (rows.length <= row) rows.push([]);
      const line = rows[row];
      while (line.length <= col) line.push(null);
      line[col] = toRaw(this.engineContent(st.content, s, row, col));
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
    // As formulas write them: an apostrophe in a sheet's name is doubled inside its quotes ('Bob''s'!A1).
    const names = sheets.map((m) => m.info.name.toLowerCase().replace(/'/g, "''"));
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
      for (const e of this.edits[sheets.indexOf(m)].values()) if (e.content.kind === 'formula') mention(e.content.formula);
      for (const r of [...m.info.ruleValidations, ...m.info.conditionalFormats]) {
        const fs = 'formulas' in r ? r.formulas : [r.formula1, r.formula2 ?? ''];
        fs.forEach(mention);
      }
      for (const l of m.info.listValidations) if (l.source.kind === 'formula') mention(l.source.formula);
    }
    for (const n of this.model.names) mention(n.formula);
    // A table reference names the table, not its sheet.
    for (const t of this.tablePlaces) {
      if (needed.has(t.sheet)) continue;
      const name = t.name.toLowerCase();
      if (sheets.some((m, s) => [...(m.scan?.grid.formulas.values() ?? []), ...[...this.edits[s].values()].map((e) => (e.content.kind === 'formula' ? e.content.formula : ''))].some((f) => f.toLowerCase().includes(name)))) needed.add(t.sheet);
    }
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
    for (const m of this.spills) m.clear();
    for (const m of this.spilled) m.clear();
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
        if (this.unsupported(st.content.formula, s, Math.floor(key / MAX_COLS), key % MAX_COLS)) this.newFormulaFlags[s].set(key, CellFlag.Unsupported);
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
      const kept: Pos[] = [];
      for (const r of roots) {
        if (extraRoots.includes(r) || !mark(r)) continue;
        next.push(r);
        // The other cells of an array formula kept as Excel saved it are as out of date as its own.
        for (const m of this.arrayCells(r)) {
          if (this.stale[m.sheet].has(cellKey(m.row, m.col))) continue;
          this.stale[m.sheet].add(cellKey(m.row, m.col));
          out.push(m);
          kept.push(m);
        }
      }
      for (const d of this.engine?.dependentsOf([...roots, ...kept]) ?? []) if (mark(d)) next.push(d);
      frontier = [...next, ...kept];
      roots = [];
    }
    return out;
  }

  /** The cells of the file's multi-cell array formula or data table whose formula is in `p`, other than that one. */
  private arrayCells(p: Pos): Pos[] {
    const scan = this.sheet(p.sheet).scan;
    const idx = scan ? findCell(scan.grid, p.row, p.col) : -1;
    const a = idx >= 0 ? scan!.arrayRanges.find((x) => x.master === idx) : undefined;
    const out: Pos[] = [];
    if (a) {
      forEachIn(a, (row, col) => {
        if (row !== p.row || col !== p.col) out.push({ sheet: p.sheet, row, col });
      });
    }
    return out;
  }

  // ---------------------------------------------------------------------------
  // Dynamic arrays

  /**
   * Whether the formula in a cell spills its results, as Excel 365's
   * dynamic arrays do: one typed here, or one the file marks as a dynamic
   * array, that the engine calculates; or fills the range it was entered in,
   * as one of the file's array formulas entered with Ctrl+Shift+Enter. The
   * file's other formulas came from an Excel that worked them out to one
   * value, and keep that meaning.
   */
  private canSpill(sheet: number, row: number, col: number): boolean {
    const key = cellKey(row, col);
    const edit = this.edits[sheet].get(key);
    if (edit) return edit.content.kind === 'formula' && !this.newFormulaFlags[sheet].has(key);
    const scan = this.sheet(sheet).scan;
    const idx = scan ? findCell(scan.grid, row, col) : -1;
    return idx >= 0 && !!(scan!.grid.flags[idx] & CellFlag.Formula) && !!this.savedSpill(sheet, idx, row, col) && !this.isUntrustedOriginal(sheet, row, col);
  }

  /** The range a spill covers, or for one that is blocked, the range its results need. */
  private spillRange(row: number, col: number, spill: Spill): Rect | undefined {
    if (spill.shown) return spill.shown;
    return spill.rows ? { r1: row, c1: col, r2: row + spill.rows - 1, c2: col + spill.cols - 1 } : undefined;
  }

  /** The value a cell shows because a formula's results spill into it; undefined when none does. */
  private spilledValue(sheet: number, row: number, col: number): EngineValue | undefined {
    if (!this.spilled[sheet].has(cellKey(row, col))) return undefined;
    return this.engine?.value(sheet, row, col) ?? null;
  }

  /** A cell's content, or for a cell a formula's results spill into, the value it shows: what copying it takes. */
  private contentOrSpilled(sheet: number, row: number, col: number): CellState {
    const st = this.current(sheet, row, col);
    const spilled = st.content.kind === 'empty' ? this.spilledValue(sheet, row, col) : undefined;
    return spilled === undefined ? st : { content: valueContent(spilled), style: st.style };
  }

  /** The range of a spill from a formula whose results cover any cell of `rect`, if one does. */
  private spillIn(sheet: number, rect: Rect): Rect | undefined {
    for (const spill of this.spills[sheet].values()) {
      const r = spill.shown;
      if (r && (r.r2 > r.r1 || r.c2 > r.c1) && r.r1 <= rect.r2 && r.r2 >= rect.r1 && r.c1 <= rect.c2 && r.c2 >= rect.c1) return r;
    }
    return undefined;
  }

  /**
   * Gives each spilling formula room in the engine for the result it has:
   * formulas just written whose results are arrays, results that have grown
   * or shrunk, and blocked formulas whose way may be clear now. The engine
   * sizes an array before working it out (FILTER as large as what it
   * filters, SEQUENCE(A1) not at all), and checks a blocked formula's room
   * only when the formula is put in again, so each is put in at the size its
   * result really has. Returns the engine's changes.
   */
  private fitSpills(written: Pos[], changes: (Pos & { value?: EngineValue })[]): EngineChange[] {
    const engine = this.engine;
    if (!engine) return [];
    for (const p of written) {
      const spill = this.spills[p.sheet].get(cellKey(p.row, p.col));
      // Put in again as written: it has no room of its own until it is given some.
      if (spill) spill.rows = spill.cols = 0;
    }
    const formulaOf = (p: Pos) => String(toRaw(this.engineContent(this.current(p.sheet, p.row, p.col).content, p.sheet, p.row, p.col)));
    const out: EngineChange[] = [];
    const refitted = new Set<string>();
    let fresh = written;
    let changed = changes;
    for (let round = 0; round < 8 && (fresh.length || changed.length); round++) {
      const fits = new Map<string, EngineWrite & ArrayDims>();
      const fit = (p: Pos, size: ArrayDims) => {
        if (this.spills[p.sheet].get(cellKey(p.row, p.col))?.fixed) return;
        if (size.rows * size.cols <= MAX_SPILL_CELLS) fits.set(`${p.sheet}:${p.row}:${p.col}`, { sheet: p.sheet, row: p.row, col: p.col, raw: spillFormula(formulaOf(p), size.rows, size.cols), ...size });
      };
      // Formulas not given room yet, just written or gone to an error, whose results turn out to be arrays.
      const probe = (p: Pos) => {
        if (this.spills[p.sheet].get(cellKey(p.row, p.col))?.rows || !this.canSpill(p.sheet, p.row, p.col)) return;
        // A formula's own cell is part of an array only when the array is its own.
        const array = engine.isArray(p.sheet, p.row, p.col);
        const v = engine.value(p.sheet, p.row, p.col);
        if (!array && !(v !== null && typeof v === 'object')) return;
        const size = engine.resultSize(formulaOf(p), p.sheet);
        if (array || size.rows * size.cols > 1) fit(p, size);
      };
      for (const p of fresh) probe(p);
      for (const p of changed) {
        const spill = this.spills[p.sheet].get(cellKey(p.row, p.col));
        if (!spill?.rows) {
          // A formula whose result is an array the engine had no size for shows an error.
          if (!('value' in p) || (p.value !== null && typeof p.value === 'object')) probe(p);
          continue;
        }
        // Results that have grown or shrunk since they were given room.
        const m = engine.measured(p.sheet, p.row, p.col);
        if (m && (m.rows !== spill.rows || m.cols !== spill.cols)) fit(p, m);
      }
      // Blocked formulas, when something in their way changed or they were worked out again.
      const moved = [...fresh, ...changed];
      this.spills.forEach((spills, sheet) => {
        for (const [key, spill] of spills) {
          const row = Math.floor(key / MAX_COLS);
          const col = key % MAX_COLS;
          if (!spill.rows || refitted.has(`${sheet}:${row}:${col}`) || !isSpillError(engine.value(sheet, row, col))) continue;
          const room = { r1: row, c1: col, r2: row + spill.rows - 1, c2: col + spill.cols - 1 };
          if (moved.some((p) => p.sheet === sheet && inRect(room, p.row, p.col))) fit({ sheet, row, col }, engine.resultSize(formulaOf({ sheet, row, col }), sheet));
        }
      });
      if (!fits.size) break;
      for (const [id, f] of fits) {
        refitted.add(id);
        const key = cellKey(f.row, f.col);
        const spill = this.spills[f.sheet].get(key);
        if (spill) {
          spill.rows = f.rows;
          spill.cols = f.cols;
        } else {
          this.spills[f.sheet].set(key, { rows: f.rows, cols: f.cols });
        }
      }
      // Put in wrapped in the spill function, which gives the engine the size.
      const fitted = engine.write([...fits.values()]);
      out.push(...fitted);
      changed = fitted;
      fresh = [];
    }
    return out;
  }

  /**
   * Works out again which cells the spills from formulas among `cells`
   * cover, and returns every cell whose display that changes: the cells
   * covered before or after, and the formulas' own.
   */
  private refreshSpills(cells: Iterable<Pos>): Pos[] {
    const out: Pos[] = [];
    const done = new Set<string>();
    for (const p of cells) {
      const key = cellKey(p.row, p.col);
      const spill = this.spills[p.sheet].get(key);
      if (!spill || done.has(`${p.sheet}:${key}`)) continue;
      done.add(`${p.sheet}:${key}`);
      const was = spill.shown;
      const live = spill.rows > 0 && this.canSpill(p.sheet, p.row, p.col);
      const now = live && !isSpillError(this.engine?.value(p.sheet, p.row, p.col) ?? null) ? this.spillRange(p.row, p.col, { rows: spill.rows, cols: spill.cols }) : undefined;
      if (live) spill.shown = now;
      else this.spills[p.sheet].delete(key);
      const members = this.spilled[p.sheet];
      if (was) {
        forEachIn(was, (row, col) => {
          const k = cellKey(row, col);
          if (members.get(k) === key) members.delete(k);
          if (k !== key) out.push({ sheet: p.sheet, row, col });
        });
      }
      if (now) {
        forEachIn(now, (row, col) => {
          const k = cellKey(row, col);
          if (k === key) return;
          members.set(k, key);
          if (!was || !inRect(was, row, col)) out.push({ sheet: p.sheet, row, col });
        });
      }
      out.push({ sheet: p.sheet, row: p.row, col: p.col });
    }
    return out;
  }

  /** Every spilling formula's cell. */
  private spillAnchors(): Pos[] {
    const out: Pos[] = [];
    this.spills.forEach((m, sheet) => {
      for (const key of m.keys()) out.push({ sheet, row: Math.floor(key / MAX_COLS), col: key % MAX_COLS });
    });
    return out;
  }

  /**
   * Marks one of the file's array formulas as calculated here, its other
   * cells holding only Excel's saved values for it; or as kept as saved,
   * locked like any array formula this editor can't calculate.
   */
  private markFileArray(grid: SheetGrid, idx: number, ref: Rect, live: boolean): void {
    const multi = ref.r2 > ref.r1 || ref.c2 > ref.c1;
    if (live) grid.flags[idx] &= ~(CellFlag.Unsupported | CellFlag.ArrayMember);
    else if (multi) grid.flags[idx] |= CellFlag.Unsupported | CellFlag.ArrayMember;
    if (!multi) return;
    for (let row = ref.r1; row <= Math.min(ref.r2, grid.rowCount - 1); row++) {
      for (let k = grid.rowPtr[row]; k < grid.rowPtr[row + 1]; k++) {
        if (k === idx || grid.cols[k] < ref.c1 || grid.cols[k] > ref.c2 || grid.flags[k] & CellFlag.Formula) continue;
        grid.flags[k] = live ? (grid.flags[k] & ~CellFlag.ArrayMember) | CellFlag.Spilled : (grid.flags[k] & ~CellFlag.Spilled) | CellFlag.ArrayMember;
      }
    }
  }

  /**
   * Where the results of one of the file's array formulas went when Excel
   * saved it: a dynamic array's spill, or the range an array formula entered
   * with Ctrl+Shift+Enter fills, when that is more than its own cell.
   */
  private savedSpill(sheet: number, idx: number, row: number, col: number): Rect | undefined {
    const meta = idx >= 0 ? this.sheet(sheet).scan?.fmeta.get(idx) : undefined;
    if (meta?.dynamic) return meta.ref ?? { r1: row, c1: col, r2: row, c2: col };
    const ref = meta?.kind === 'array' ? meta.ref : undefined;
    return ref && (ref.r2 > ref.r1 || ref.c2 > ref.c1) ? ref : undefined;
  }

  /** Whether the engine spills one of the file's array formulas differently from what Excel saved. */
  private spillDiffers(p: Pos): boolean {
    const grid = this.sheet(p.sheet).scan!.grid;
    const saved = this.savedSpill(p.sheet, findCell(grid, p.row, p.col), p.row, p.col);
    if (!saved) return false;
    const ours = this.spills[p.sheet].get(cellKey(p.row, p.col))?.shown ?? { r1: p.row, c1: p.col, r2: p.row, c2: p.col };
    const all = { r1: p.row, c1: p.col, r2: Math.max(saved.r2, ours.r2), c2: Math.max(saved.c2, ours.c2) };
    let differs = false;
    forEachIn(all, (row, col) => {
      if (differs || (row === p.row && col === p.col)) return;
      const idx = inRect(saved, row, col) ? findCell(grid, row, col) : -1;
      const excel = idx >= 0 && grid.flags[idx] & CellFlag.Spilled ? gridValue(grid, idx) : null;
      const engine = inRect(ours, row, col) ? (this.engine?.value(p.sheet, row, col) ?? null) : null;
      differs = !valuesMatch(excel, engine);
    });
    return differs;
  }

  /**
   * Keeps one of the file's array formulas as Excel saved it, when the
   * engine calculates it differently: it no longer spills here, and its
   * other cells go back to Excel's values. Returns those, for the engine.
   */
  private keepFileArray(p: Pos): EngineWrite[] {
    const grid = this.sheet(p.sheet).scan!.grid;
    const idx = findCell(grid, p.row, p.col);
    const saved = this.savedSpill(p.sheet, idx, p.row, p.col);
    if (!saved) return [];
    const spill = this.spills[p.sheet].get(cellKey(p.row, p.col));
    if (spill) {
      spill.rows = 0;
      this.refreshSpills([p]);
    }
    this.markFileArray(grid, idx, saved, false);
    const pins: EngineWrite[] = [];
    forEachIn(saved, (row, col) => {
      const k = findCell(grid, row, col);
      if (k >= 0 && k !== idx && grid.flags[k] & CellFlag.ArrayMember) pins.push({ sheet: p.sheet, row, col, raw: valueRaw(gridValue(grid, k)) });
    });
    return pins;
  }

  /**
   * After the engine's values changed: gives spills the room their results
   * need, takes the new values of the formulas it calculates, and works out
   * again where spills reach. `written` are the cells just put in the
   * engine. Returns the engine's changes, with those the spills made, and
   * the cells whose spill changed; every cell to redraw is noted in `touched`.
   */
  private afterEngineChanges(written: Pos[], changes: EngineChange[], touched: Map<string, Pos>): { changes: EngineChange[]; spills: Pos[] } {
    const fitted = this.fitSpills(written, changes);
    // A cell changed more than once ends with its last value.
    const all = fitted.length ? [...new Map([...changes, ...fitted].map((c) => [`${c.sheet}:${c.row}:${c.col}`, c])).values()] : changes;
    this.absorb(all, touched);
    const spills = this.refreshSpills([...written, ...all]);
    for (const p of spills) touched.set(`${p.sheet}:${p.row}:${p.col}`, p);
    return { changes: all, spills };
  }

  // ---------------------------------------------------------------------------
  // Conditional formatting

  /** A cell's current value, as formulas and rules see it. */
  private cellValue(sheet: number, row: number, col: number): EngineValue {
    const { content } = this.current(sheet, row, col);
    switch (content.kind) {
      case 'empty':
        return this.spilledValue(sheet, row, col) ?? null;
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
        const file = formula.replace(/^=/, '');
        // A rule the engine can't calculate is left for Excel to judge.
        if (this.unsupported(file, sheet, -1, -1)) return { error: '#ERROR!' };
        this.ensureSheetsFor(formula, sheet);
        return engine.calculate(engineFormula(file), sheet);
      },
      withValue: (sheet, row, col, content, fn) => {
        this.ensureSheetsFor('', sheet);
        // What the engine holds for the cell now, so it can be put back exactly.
        const untrusted = this.isUntrustedOriginal(sheet, row, col);
        const grid = this.sheet(sheet).scan?.grid;
        const prev = untrusted && grid ? valueRaw(gridValue(grid, untrusted.idx)) : toRaw(this.current(sheet, row, col).content);
        engine.write([{ sheet, row, col, raw: toRaw(this.engineContent(content, sheet, row, col)) }]);
        try {
          return fn();
        } finally {
          // A spill the value was in the way of, or one from this cell, is given its room again.
          this.fitSpills([{ sheet, row, col }], engine.write([{ sheet, row, col, raw: prev }]));
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
    const r = this.applyEdit(op);
    const engine = this.engine;
    if (!engine) return r;
    // Where circular references close now; a new one is warned of, as Excel warns.
    const now = this.circles(engine);
    const before = new Set(this.circularCells.map((p) => `${p.sheet}:${p.row}:${p.col}`));
    this.circularCells = now;
    if (now.some((p) => !before.has(`${p.sheet}:${p.row}:${p.col}`))) r.warnings.push(msg('core.calc.circular'));
    return { ...r, circular: now };
  }

  /** Where circular references close: none to speak of when the workbook iterates them, as Excel neither warns nor names them then. */
  private circles(engine: FormulaEngine): { sheet: number; row: number; col: number }[] {
    return this.model.iteration ? [] : engine.circular();
  }

  private applyEdit(op: EditOp): EditResult {
    this.ensureEngine();
    if (op.type === 'batch') return mergeEditResults(op.ops.map((o) => this.applyEdit(o)));
    if (op.type === 'comment') return this.applyComment(op);
    if (op.type === 'layout') return this.applyLayout([{ sheet: op.sheet, op: op.op }]);
    if (op.type === 'name') return this.applyName(op.change);
    // Names go back first, so the formulas put back with them find them.
    if (op.type === 'restore' && op.names?.length) return mergeEditResults([this.applyNames(op.names), this.applyEdit({ ...op, names: undefined })]);
    let notes: { undo: NoteSnapshot[]; redo: NoteSnapshot[] } | undefined;
    const writes = new Map<string, Write>();
    const warnings: Text[] = [];
    const newStyles: { index: number; style: DisplayStyle }[] = [];
    /** Changes to links that go with the cells written. */
    const links: LayoutChange[] = [];
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
        const what = op.what ?? 'all';
        let lost = 0;
        for (let row = op.dst.r1; row <= op.dst.r2; row++) {
          for (let col = op.dst.c1; col <= op.dst.c2; col++) {
            // Turned on its side, the destination's rows come from the source's columns.
            const [down, across] = op.transpose ? [col - op.dst.c1, row - op.dst.r1] : [row - op.dst.r1, col - op.dst.c1];
            const sr = op.src.r1 + (down % h);
            const sc = op.src.c1 + (across % w);
            const here = this.current(op.sheet, row, col);
            if (what === 'formats') {
              put({ sheet: op.sheet, row, col, content: here.content, style: this.current(op.srcSheet, sr, sc).style });
              continue;
            }
            const r = this.copiedContent(op.srcSheet, sr, sc, row, col);
            if (what === 'values') {
              const src = this.contentOrSpilled(op.srcSheet, sr, sc);
              const content = src.content.kind === 'formula' ? valueContent(this.formulaValue(op.srcSheet, sr, sc)) : src.content;
              put({ sheet: op.sheet, row, col, content, style: here.style });
              continue;
            }
            // A spilled value comes as a value, unless its formula comes too and spills it there again.
            const from = this.spilled[op.srcSheet].get(cellKey(sr, sc));
            const respills = from !== undefined && !this.spills[op.srcSheet].get(from)?.fixed;
            if (respills && inRect(op.src, Math.floor(from / MAX_COLS), from % MAX_COLS) && !op.transpose) r.content = EMPTY;
            if (r.lost) lost++;
            put({ sheet: op.sheet, row, col, content: r.content, style: what === 'formulas' ? here.style : r.style });
          }
        }
        if (lost) warnings.push(msg('core.edit.copiedAsValues', { count: lost }));
        // Links go with the cells, as Excel pastes them.
        links.push(...this.linkChange(op.sheet, copyLinks(this.linksOf(op.srcSheet).list, op.src, this.linksOf(op.sheet).list, op.dst)));
        break;
      }
      case 'fill': {
        warnings.push(...this.planFillWrites(op.sheet, op.src, op.dst, op.lists ?? ENGLISH_LISTS, put));
        break;
      }
      case 'series': {
        warnings.push(...this.planSeriesWrites(op, put));
        break;
      }
      case 'sort': {
        const { rect } = op;
        const merged = this.mergesOf(op.sheet).find((m) => m.r1 <= rect.r2 && m.r2 >= rect.r1 && m.c1 <= rect.c2 && m.c2 >= rect.c1);
        if (merged) return this.emptyResult(msg('core.edit.sortMerged', { range: rangeToA1(merged) }));
        const top = op.header ? rect.r1 + 1 : rect.r1;
        // As in Excel, a formula's spilled results can't be moved apart.
        const spill = this.spillIn(op.sheet, { ...rect, r1: top });
        if (spill) return this.emptyResult(msg('core.edit.sortSpill', { range: rangeToA1(spill) }));
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
        links.push(...this.linkChange(op.sheet, sortLinks(this.linksOf(op.sheet).list, rect, top, order.map((o) => o.row))));
        break;
      }
      case 'restore': {
        for (const c of op.cells) put({ ...c });
        if (op.notes?.length) notes = this.putNotes(op.notes);
        break;
      }
      case 'format': {
        const info = this.sheet(op.sheet).info;
        let maxRow = info.maxRow;
        let maxCol = info.maxCol;
        for (const key of this.edits[op.sheet].keys()) {
          maxRow = Math.max(maxRow, Math.floor(key / MAX_COLS));
          maxCol = Math.max(maxCol, key % MAX_COLS);
        }
        // Whole rows and columns are formatted as far as the sheet's cells go.
        const rects = op.rects.map((r) => ({
          ...r,
          r2: r.r2 >= MAX_ROWS - 1 ? Math.max(r.r1, maxRow) : r.r2,
          c2: r.c2 >= MAX_COLS - 1 ? Math.max(r.c1, maxCol) : r.c2,
        }));
        if (rects.reduce((n, r) => n + (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1), 0) > MAX_FORMAT_CELLS) {
          return this.emptyResult(msg('core.format.tooLarge', { max: MAX_FORMAT_CELLS }));
        }
        for (const rect of rects) {
          for (let row = rect.r1; row <= rect.r2; row++) {
            for (let col = rect.c1; col <= rect.c2; col++) {
              const cur = this.current(op.sheet, row, col);
              const change = op.borders ? { ...op.change, border: placedBorder(op.borders, rect, row, col) } : op.change;
              let style = op.clear ? 0 : cur.style;
              if (Object.keys(change).length) {
                const r = this.model.styles.withFormat(style, change);
                if (r.created) newStyles.push({ index: r.index, style: this.model.styles.get(r.index) });
                style = r.index;
              }
              put({ sheet: op.sheet, row, col, content: cur.content, style });
            }
          }
        }
        break;
      }
      case 'linkFormat': {
        for (const rect of op.rects) {
          this.forEachCell(op.sheet, rect, (row, col, cur) => {
            const r = this.model.styles.withLinkFont(cur.style, op.on);
            if (r.created) newStyles.push({ index: r.index, style: this.model.styles.get(r.index) });
            if (r.index !== cur.style) put({ sheet: op.sheet, row, col, content: cur.content, style: r.index });
          });
        }
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
    const committed = this.commit([...writes.values()], warnings, newStyles, op.type === 'restore', 'readOnly' in op ? op.readOnly : undefined);
    const result = notes ? { ...committed, noteUndo: notes.undo, noteRedo: notes.redo } : committed;
    const layout = op.type === 'restore' ? (op.layout ?? []) : links;
    return layout.length ? mergeEditResults([result, this.applyLayout(layout)]) : result;
  }

  /** Changes rows', columns', panes' and merges' layout, or links, after checking each can be made. */
  private applyLayout(changes: LayoutChange[]): EditResult {
    for (const { sheet, op } of changes) {
      const info = this.sheet(sheet).info;
      if (info.kind !== 'worksheet') return this.emptyResult(msg('core.reason.notWorksheet'));
      // Excel's protection allows freezing panes, which only changes the view, and hiding,
      // showing and grouping rows or columns where it lets them be formatted.
      const allowed = op.kind === 'freeze' || ((op.kind === 'hidden' || op.kind === 'outline') && info.formatAllowed?.[op.axis]);
      if (info.protected && !allowed) return this.emptyResult(msg('core.reason.protected'));
      if (op.kind === 'outline' && Object.values(op.levels).some((v) => !(v >= 0 && v <= 7))) return this.emptyResult(msg('outline.refuse.tooDeep'));
      if (op.kind === 'merges') {
        const table = info.tables.find((t) => op.add.some((r) => r.r1 <= t.rect.r2 && r.r2 >= t.rect.r1 && r.c1 <= t.rect.c2 && r.c2 >= t.rect.c1));
        if (table) return this.emptyResult(msg('core.layout.mergeTable', { table: table.name }));
      }
    }
    const undo: LayoutChange[] = [];
    for (const { sheet, op } of changes) undo.unshift({ sheet, op: op.kind === 'links' ? this.linksOf(sheet).apply(op) : this.layoutOf(sheet).apply(op) });
    return { ...this.emptyResult(''), rejected: undefined, layout: { applied: changes, undo } };
  }

  /**
   * Checks a change to a name and makes it. Renaming a name renames it in the
   * cell formulas and other names that use it, as Excel does.
   */
  private applyName(change: NameChange): EditResult {
    const names = this.model.names;
    const before = change.from && names.find((n) => isName(n, change.from!));
    if (change.from && !before) return this.emptyResult(msg('names.error.gone', { name: change.from.name }));
    const to: DefinedName | null = change.to && {
      name: change.to.name,
      formula: change.to.formula,
      ...(change.to.localSheet !== undefined && { localSheet: change.to.localSheet }),
      ...(change.to.comment && { comment: change.to.comment }),
      ...(before?.hidden && { hidden: true }),
      ...(before?.fileIndex !== undefined && { fileIndex: before.fileIndex }),
    };
    const problem = to && this.definitionProblem(to, before);
    if (problem) return this.emptyResult(problem);
    const changes: NameChange[] = [{ ...(before && { from: { name: before.name, localSheet: before.localSheet } }), to }];
    const writes: Write[] = [];
    if (before && to && before.name !== to.name && before.localSheet === to.localSheet) {
      const sheetNames = this.model.sheets.map((m) => m.info.name);
      for (const n of names) {
        const formula = n === before ? undefined : renameInFormula(n.formula, before, to.name, n.localSheet, names, sheetNames);
        if (formula !== undefined) changes.push({ from: { name: n.name, localSheet: n.localSheet }, to: { ...n, formula } });
      }
      this.model.sheets.forEach((_m, s) =>
        this.forEachFormula(s, (row, col, formula, style) => {
          const renamed = renameInFormula(formula, before, to.name, s, names, sheetNames);
          if (renamed !== undefined) writes.push({ sheet: s, row, col, content: { kind: 'formula', formula: renamed }, style });
        }),
      );
    }
    const result = this.applyNames(changes);
    return writes.length ? mergeEditResults([result, this.commit(writes, [], [])]) : result;
  }

  /** Why a name can't be defined as `to` (replacing `before`): its name, its scope, a clash, or its formula. */
  private definitionProblem(to: DefinedName, before: DefinedName | undefined): Text | undefined {
    // A name the file has keeps it, even where Excel's own rules are looser than these.
    const problem = before && before.name === to.name ? undefined : nameProblem(to.name);
    if (problem) return problem;
    if (to.localSheet !== undefined && !this.model.sheets[to.localSheet]) return msg('names.error.noScope');
    const clash = this.model.names.find((n) => n !== before && isName(n, to));
    if (clash) {
      return to.localSheet === undefined
        ? msg('names.error.takenWorkbook', { name: clash.name })
        : msg('names.error.takenSheet', { name: clash.name, sheet: this.model.sheets[to.localSheet].info.name });
    }
    return validateFormula(to.formula) === undefined ? undefined : msg('names.error.refersToSyntax');
  }

  /**
   * Makes changes to the names in order, and recalculates the formulas that
   * use them. A change to a name that has gone is skipped.
   */
  private applyNames(changes: NameChange[]): EditResult {
    this.lambdas = undefined;
    const names = this.model.names;
    const applied: NameChange[] = [];
    const undo: NameChange[] = [];
    const warnings: Text[] = [];
    const recalculated: EngineChange[] = [];
    const engine = this.engine;
    for (const c of changes) {
      const at = c.from ? names.findIndex((n) => isName(n, c.from!)) : -1;
      if (c.from && at < 0) continue;
      const before = at >= 0 ? names[at] : undefined;
      const to = c.to && { ...c.to };
      if (to && at >= 0) names[at] = to;
      else if (to) names.push(to);
      else names.splice(at, 1);
      applied.push({ ...(c.from && { from: c.from }), to });
      undo.unshift({ ...(to && { from: { name: to.name, localSheet: to.localSheet } }), to: before ?? null });
      if (!engine) continue;
      if (before && (!to || !isName(before, to))) recalculated.push(...engine.setName(before.name, before.localSheet, null).changes);
      if (!to) continue;
      this.ensureSheetsFor(to.formula, to.localSheet ?? -1);
      if (to.localSheet !== undefined) this.ensureSheetsFor('', to.localSheet);
      const set = engine.setName(to.name, to.localSheet, this.engineText(to.formula, to.localSheet ?? 0, -1, -1));
      recalculated.push(...set.changes);
      if (!set.held && !isBuiltinName(to.name) && !to.formula.includes('#REF!')) warnings.push(msg('names.warn.notCalculated', { name: to.name }));
    }
    // A cell recalculated by more than one change ends with the last one's value.
    const latest = [...new Map(recalculated.map((c) => [`${c.sheet}:${c.row}:${c.col}`, c])).values()];
    const touched = new Map<string, Pos>();
    const after = this.afterEngineChanges([], latest, touched);
    for (const p of this.markStale([...after.changes, ...after.spills])) touched.set(`${p.sheet}:${p.row}:${p.col}`, p);
    return {
      ...this.emptyResult(''),
      rejected: undefined,
      updates: [...touched.values()].map((p) => this.update(p.sheet, p.row, p.col)),
      warnings,
      names: { applied, undo, list: this.nameList() },
    };
  }

  /** Calls fn for every formula on a sheet as it now is, except array formulas, which stay as they are. */
  private forEachFormula(sheet: number, fn: (row: number, col: number, formula: string, style: number) => void): void {
    const grid = this.sheet(sheet).scan?.grid;
    if (grid) {
      for (const [idx, formula] of grid.formulas) {
        if (!(grid.flags[idx] & CellFlag.Formula) || grid.flags[idx] & CellFlag.ArrayMember) continue;
        const row = rowOfIndex(grid, idx);
        if (!this.edits[sheet].has(cellKey(row, grid.cols[idx]))) fn(row, grid.cols[idx], formula, grid.styles[idx]);
      }
    }
    for (const [key, st] of this.edits[sheet]) if (st.content.kind === 'formula') fn(Math.floor(key / MAX_COLS), key % MAX_COLS, st.content.formula, st.style);
  }

  private emptyResult(rejected?: Text): EditResult {
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
    const src = this.contentOrSpilled(srcSheet, sr, sc);
    if (src.content.kind !== 'formula') return { ...src, lost: false };
    const flags = this.flagsFor(srcSheet, sr, sc);
    // An array formula entered with Ctrl+Shift+Enter goes as the value its cell shows.
    if (flags & CellFlag.ArrayMember || this.spills[srcSheet].get(cellKey(sr, sc))?.fixed) {
      return { content: valueContent(this.formulaValue(srcSheet, sr, sc)), style: src.style, lost: true };
    }
    try {
      return { content: { kind: 'formula', formula: makeShifter(src.content.formula, sr, sc)(row, col) }, style: src.style, lost: false };
    } catch {
      return { content: valueContent(this.formulaValue(srcSheet, sr, sc)), style: src.style, lost: true };
    }
  }

  private planFillWrites(sheet: number, src: Rect, dst: Rect, lists: readonly (readonly string[])[], put: (w: Write) => void): Text[] {
    // The destination may include the source; fill the part beyond it.
    const vertical = dst.r1 < src.r1 || dst.r2 > src.r2;
    const forward = vertical ? dst.r2 > src.r2 : dst.c2 > src.c2;
    let lost = 0;
    const lines = vertical ? [src.c1, src.c2] : [src.r1, src.r2];
    for (let line = lines[0]; line <= lines[1]; line++) {
      const srcCells: { row: number; col: number }[] = [];
      const len = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
      for (let i = 0; i < len; i++) srcCells.push(vertical ? { row: src.r1 + i, col: line } : { row: line, col: src.c1 + i });
      const targets: { row: number; col: number; x: number }[] = [];
      if (vertical) {
        if (forward) for (let r = src.r2 + 1; r <= dst.r2; r++) targets.push({ row: r, col: line, x: r - src.r1 });
        else for (let r = src.r1 - 1; r >= dst.r1; r--) targets.push({ row: r, col: line, x: r - src.r1 });
      } else {
        if (forward) for (let c = src.c2 + 1; c <= dst.c2; c++) targets.push({ row: line, col: c, x: c - src.c1 });
        else for (let c = src.c1 - 1; c >= dst.c1; c--) targets.push({ row: line, col: c, x: c - src.c1 });
      }
      lost += this.fillLine(sheet, srcCells, targets, lists, put);
    }
    return lost ? [msg('core.edit.filledAsValues', { count: lost })] : [];
  }

  /**
   * Fills the targets from the source cells along one row or column, as
   * dragging the fill handle does; says how many formulas could only be
   * copied as their values.
   */
  private fillLine(
    sheet: number,
    srcCells: { row: number; col: number }[],
    targets: { row: number; col: number; x: number }[],
    lists: readonly (readonly string[])[],
    put: (w: Write) => void,
  ): number {
    const isDate = (style: number) => {
      try {
        return isDateFormat(this.model.styles.get(style).numFmt);
      } catch {
        return false;
      }
    };
    const states = srcCells.map((p) => this.contentOrSpilled(sheet, p.row, p.col));
    const sources: FillSource[] = states.map((s) => ({ content: s.content, isDate: isDate(s.style) }));
    const plans = planFill(sources, targets.map((t) => t.x), this.model.date1904, lists);
    let lost = 0;
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
    return lost;
  }

  /**
   * Excel's Fill Series: each row (or column) of the range is a series from
   * the cells it starts with. Rows that don't start with a number (or, for
   * AutoFill, with anything) are left as they are. Numbers written take the
   * first cell's format, so a date series shows as dates.
   */
  private planSeriesWrites(op: Extract<EditOp, { type: 'series' }>, put: (w: Write) => void): Text[] {
    const { sheet, rect, series } = op;
    const across = series.by === 'rows';
    const length = across ? rect.c2 - rect.c1 + 1 : rect.r2 - rect.r1 + 1;
    const [first, last] = across ? [rect.r1, rect.r2] : [rect.c1, rect.c2];
    let lost = 0;
    let started = 0;
    let notGrowth = 0;
    for (let line = first; line <= last; line++) {
      const cells = Array.from({ length }, (_, i) => (across ? { row: line, col: rect.c1 + i } : { row: rect.r1 + i, col: line }));
      if (series.type === 'autofill') {
        // What the line starts with is continued through the rest of it.
        let k = 0;
        while (k < length && this.contentOrSpilled(sheet, cells[k].row, cells[k].col).content.kind !== 'empty') k++;
        if (!k) continue;
        started++;
        lost += this.fillLine(sheet, cells.slice(0, k), cells.slice(k).map((c, i) => ({ ...c, x: k + i })), op.lists ?? ENGLISH_LISTS, put);
        continue;
      }
      const starts: number[] = [];
      for (const c of cells) {
        const v = this.cellValue(sheet, c.row, c.col);
        if (typeof v !== 'number') break;
        starts.push(v);
        if (!series.trend) break;
      }
      if (!starts.length) continue;
      started++;
      const planned = seriesLine(starts, length, series, this.model.date1904);
      if (!planned) {
        notGrowth++;
        continue;
      }
      const style = this.current(sheet, cells[0].row, cells[0].col).style;
      planned.values.forEach((value, i) => {
        const at = cells[planned.from + i];
        // Numbers a trend replaces keep their own format.
        const own = planned.from + i < starts.length ? this.current(sheet, at.row, at.col).style : style;
        put({ sheet, row: at.row, col: at.col, content: { kind: 'number', value }, style: own });
      });
    }
    const warnings: Text[] = [];
    if (!started) warnings.push(msg(series.type === 'autofill' ? 'fill.series.nothingToContinue' : 'fill.series.noNumber'));
    if (notGrowth) warnings.push(msg('fill.series.growthTrend', { count: notGrowth }));
    if (lost) warnings.push(msg('core.edit.filledAsValues', { count: lost }));
    return warnings;
  }

  /**
   * Applies writes; `restoring` puts back an earlier state, so cells a merge
   * covers then are written too. Cells in `readOnly` are skipped as locked ones are.
   */
  private commit(writes: Write[], warnings: Text[], newStyles: { index: number; style: DisplayStyle }[], restoring = false, readOnly?: ReadOnlyCells[]): EditResult {
    const engine = this.engine;
    const undo: CellSnapshot[] = [];
    const redo: CellSnapshot[] = [];
    // Keyed by content: the same message, made again for each cell, counts once.
    const reasons = new Map<string, Text>();
    let skipped = 0;
    const applied: Write[] = [];
    /** Applied writes the engine is given: all but new formats for a spill's cells, which would disturb it. */
    const content: Write[] = [];

    for (const w of writes) {
      const reason = this.blockedReason(w.sheet, w.row, w.col) ?? readOnly?.find((r) => inRect(r, w.row, w.col))?.reason;
      // The cells a merge covers take formats, as Excel gives them the merge's borders; and earlier states are put back whole.
      const styleOnly = w.content.kind === 'empty' && this.current(w.sheet, w.row, w.col).content.kind === 'empty';
      if (reason && !(reason === MERGED && (restoring || styleOnly))) {
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
      const key = cellKey(w.row, w.col);
      const formatOnly = sameContent(prev.content, w.content) && (this.spilled[w.sheet].has(key) || this.spills[w.sheet].has(key));
      this.setCurrent(w);
      applied.push(w);
      if (!formatOnly) content.push(w);
    }

    const engineWrites: EngineWrite[] = content.map((w) => {
      const untrusted = this.isUntrustedOriginal(w.sheet, w.row, w.col);
      const grid = this.sheet(w.sheet).scan?.grid;
      // A formula over the engine's size limits never reaches it.
      const tooLarge = w.content.kind === 'formula' && this.unsupported(w.content.formula, w.sheet, w.row, w.col)?.kind === 'tooLarge';
      return {
        sheet: w.sheet,
        row: w.row,
        col: w.col,
        raw: untrusted && grid ? valueRaw(gridValue(grid, untrusted.idx)) : tooLarge ? null : toRaw(this.engineContent(w.content, w.sheet, w.row, w.col)),
      };
    });
    for (const w of applied) if (w.content.kind === 'formula') this.ensureSheetsFor(this.engineText(w.content.formula, w.sheet, w.row, w.col), w.sheet);
    const written = engine?.write(engineWrites) ?? [];

    // Formulas typed in this session that the engine can't calculate (all of them, with recalculation off).
    const newUnsupported: Pos[] = [];
    const calculated: Write[] = [];
    for (const w of applied) {
      if (w.content.kind !== 'formula' || this.isUntrustedOriginal(w.sheet, w.row, w.col)) continue;
      if (!engine || this.unsupported(w.content.formula, w.sheet, w.row, w.col)) {
        this.newFormulaFlags[w.sheet].set(cellKey(w.row, w.col), CellFlag.Unsupported);
        newUnsupported.push(w);
      } else {
        calculated.push(w);
      }
    }

    const touched = new Map<string, Pos>();
    const touch = (p: Pos) => touched.set(`${p.sheet}:${p.row}:${p.col}`, { sheet: p.sheet, row: p.row, col: p.col });
    applied.forEach(touch);
    const { changes, spills } = this.afterEngineChanges(content, written, touched);
    for (const w of calculated) this.values[w.sheet].set(cellKey(w.row, w.col), engine!.value(w.sheet, w.row, w.col));
    for (const p of this.markStale([...applied, ...changes, ...spills], newUnsupported)) {
      touch(p);
      // What a formula that is out of date spills is too.
      const spill = this.spills[p.sheet].get(cellKey(p.row, p.col))?.shown;
      if (spill) forEachIn(spill, (row, col) => touch({ sheet: p.sheet, row, col }));
    }
    // A HYPERLINK formula whose address changed while its text didn't.
    this.formulaLinks.forEach((m, s) => {
      for (const [key, was] of m) {
        const row = Math.floor(key / MAX_COLS);
        const col = key % MAX_COLS;
        if (touched.has(`${s}:${row}:${col}`)) continue;
        const st = this.current(s, row, col);
        if (st.content.kind !== 'formula' || this.formulaLink(s, row, col, st.content.formula) !== was) touch({ sheet: s, row, col });
      }
    });

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

  /** Takes the engine's new values for the formulas it calculates, noting each cell in `touched`. */
  private absorb(changes: EngineChange[], touched: Map<string, Pos>): void {
    for (const c of changes) {
      const id = `${c.sheet}:${c.row}:${c.col}`;
      const key = cellKey(c.row, c.col);
      if (touched.has(id)) continue;
      if (this.current(c.sheet, c.row, c.col).content.kind !== 'formula') continue;
      if (this.isUntrustedOriginal(c.sheet, c.row, c.col)) continue;
      if (this.newFormulaFlags[c.sheet].has(key)) continue;
      this.values[c.sheet].set(key, c.value);
      touched.set(id, { sheet: c.sheet, row: c.row, col: c.col });
    }
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
    const edited =
      this.edits.some((m) => m.size) || this.noteEdits.some((m) => m.size) || this.layouts.some((l) => l?.changed) || this.links.some((l) => l?.changed) || this.nameNotes().length > 0;
    return edited ? buildPackage({ ...this.saveInput(), stripMacros: false }) : this.pkg.clone();
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
      isArray: (sheet, row, col) => this.spills[sheet].has(cellKey(row, col)) || !!this.engine?.isArray(sheet, row, col),
      spills: this.spills.map((m) => new Map([...m].map(([key, spill]) => [key, spill.shown ?? { r1: Math.floor(key / MAX_COLS), c1: key % MAX_COLS, r2: Math.floor(key / MAX_COLS), c2: key % MAX_COLS }]))),
      spilled: (sheet, row, col) => ({ value: this.spilledValue(sheet, row, col) ?? null, style: this.current(sheet, row, col).style }),
      layout: this.layouts,
      notes: this.noteEdits,
      people: this.people,
      ...(this.nameNotes().length && { names: { file: this.fileNames, now: this.model.names } }),
      links: this.links,
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
   * The workbook after inserting or deleting rows or columns, moving cells,
   * or adding, renaming, deleting, moving, hiding or unhiding a sheet, as a
   * new session. The current state is written out, the change is applied to
   * every part that refers to what changed, and the result is opened. This
   * session is left as it was, so undo can return to it.
   */
  restructure(op: WorkbookOp): WorkbookSession {
    const carried = op.type === 'move' && op.toSheet !== op.sheet ? this.carriedTo(op) : undefined;
    const pkg = this.savedPackage();
    const before = readWorkbook(pkg, this.locale);
    // Rows and columns can be inserted or deleted through a spill here, which spills again afterwards.
    const spilling = new Set(this.spillAnchors().map((p) => `${p.sheet}:${p.row}:${p.col}`));
    const { broken } = isSheetOp(op) ? applySheetOp(pkg, before, op) : restructure(pkg, before, op, spilling);
    const next = new WorkbookSession(pkg, null, this.fileName, readWorkbook(pkg, this.locale), {}, this.locale);
    next.stripMacros = this.stripMacros;
    next.recalcOff = this.recalcOff;
    if (broken) {
      next.structureWarnings.push(msg('core.structure.brokenRefs', { count: broken, error: '#REF!' }));
    }

    // Changes so far follow their cells, and their sheets; changes to cells that were deleted go.
    const names = this.model.sheets.map((m) => m.info.name);
    const now = next.model.sheets.map((m) => m.info.name);
    const sheets = isSheetOp(op) ? sheetIndexMap(op, names.length) : names.map((_n, i) => i);
    const map: RefMap = isSheetOp(op) ? { sheets: new Set(), range: (s, rect) => (sheets[s] === null ? null : { sheet: sheets[s], rect }) } : refMapFor(op);
    const index = (name: string) => names.findIndex((n) => n === name);
    for (const c of this.changes()) {
      // Notes and threads are listed from what the workbook had, carried below.
      if (c.comment) continue;
      const s = index(c.sheet);
      const at = parseA1(c.ref);
      if (c.note || s < 0 || !at) {
        // Earlier changes to rows, columns and sheets are listed under the sheet's name as it is now.
        if (c.note) next.baked.structure.push(s >= 0 && sheets[s] !== null ? { ...c, sheet: now[sheets[s]] } : c);
        continue;
      }
      const to = map.range(s, { r1: at.row, c1: at.col, r2: at.row, c2: at.col });
      if (to) next.baked.changes.push({ ...c, sheet: now[to.sheet], ref: toA1(to.rect.r1, to.rect.c1) });
    }
    for (const { sheet, key, before } of this.noteChanges()) {
      const row = Math.floor(key / MAX_COLS);
      const col = key % MAX_COLS;
      const to = map.range(sheet, { r1: row, c1: col, r2: row, c2: col });
      if (to) next.noteBase[to.sheet].set(cellKey(to.rect.r1, to.rect.c1), before && { ...before, row: to.rect.r1, col: to.rect.c1 });
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
    if (carried) {
      next.putNotes(carried.notes);
      if (carried.links.length) next.linksOf(carried.sheet).apply({ add: carried.links, remove: [] });
    }
    if (op.type === 'insertRows') next.fillInsertedTableRows(op.sheet, op.at, op.count);
    if (op.type === 'growTable') {
      const t = next.model.sheets[op.sheet].info.tables.find((x) => x.name === op.table);
      if (t) next.fillInsertedTableRows(op.sheet, t.rect.r2 - op.rows + 1, op.rows);
    }
    return next;
  }

  /**
   * The notes, comment threads and links of cells moving to another sheet,
   * placed where the cells land. A sheet's own parts carry them only within
   * the sheet, so they are put on the destination as edits.
   */
  private carriedTo(op: Extract<WorkbookOp, { type: 'move' }>): { sheet: number; notes: NoteSnapshot[]; links: Hyperlink[] } {
    const dr = op.row - op.src.r1;
    const dc = op.col - op.src.c1;
    const dst = { r1: op.row, c1: op.col, r2: op.src.r2 + dr, c2: op.src.c2 + dc };
    const notes = this.notesOf(op.sheet)
      .filter((n) => inRect(op.src, n.row, n.col))
      .map((n) => ({ sheet: op.toSheet, row: n.row + dr, col: n.col + dc, note: { ...n, row: n.row + dr, col: n.col + dc } }));
    return { sheet: op.toSheet, notes, links: copyLinks(this.linksOf(op.sheet).list, op.src, [], dst).add };
  }

  /**
   * Rows inserted inside an Excel table take the table's calculated column
   * formulas and the formats of the row next to them, as Excel fills them.
   */
  private fillInsertedTableRows(sheet: number, at: number, count: number): void {
    const info = this.model.sheets[sheet]?.info;
    for (const t of info?.tables ?? []) {
      const top = t.rect.r1 + t.headerRows;
      const bottom = t.rect.r2 - t.totalsRows;
      if (at < top || at + count - 1 > bottom) continue;
      // The row above, or below when the rows went in at the top.
      const like = at > top ? at - 1 : at + count <= bottom ? at + count : -1;
      for (let row = at; row < at + count; row++) {
        for (let j = 0; j < t.columns.length; j++) {
          const col = t.rect.c1 + j;
          const style = like >= 0 ? this.current(sheet, like, col).style : 0;
          const formula = t.calculated?.[j];
          // A table that grows over cells already filled keeps what they hold.
          if ((!formula && !style) || this.current(sheet, row, col).content.kind !== 'empty') continue;
          this.setCurrent({ sheet, row, col, content: formula ? { kind: 'formula', formula } : EMPTY, style });
        }
      }
    }
  }

  /** Every cell whose state differs from the file the session was built from, for redrawing all of it. */
  stateUpdates(): CellUpdate[] {
    const out: CellUpdate[] = [];
    this.model.sheets.forEach((_m, s) => {
      const keys = new Set([...this.edits[s].keys(), ...this.values[s].keys(), ...this.calcFlags[s].keys(), ...this.stale[s], ...this.baked.keys[s]]);
      for (const key of keys) out.push(this.update(s, Math.floor(key / MAX_COLS), key % MAX_COLS));
    });
    const sent = new Set(out.map((u) => `${u.sheet}:${u.row}:${u.col}`));
    for (const p of this.spillCells()) if (!sent.has(`${p.sheet}:${p.row}:${p.col}`)) out.push(this.update(p.sheet, p.row, p.col));
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
/** Cells formatted in one step at most. */
const MAX_FORMAT_CELLS = 1_000_000;

/** The edges a border placed on a range gives the cell at row, col: unset edges stay as they were. */
function placedBorder(b: { where: BorderPlacement; edge?: BorderEdge }, rect: Rect, row: number, col: number): NonNullable<FormatChange['border']> {
  const e = b.where === 'none' ? null : (b.edge ?? { width: 1, pattern: 'solid' });
  const on = (side: boolean) => (side ? e : undefined);
  switch (b.where) {
    case 'all':
    case 'none':
      return { left: e, right: e, top: e, bottom: e };
    case 'outside':
      return { top: on(row === rect.r1), bottom: on(row === rect.r2), left: on(col === rect.c1), right: on(col === rect.c2) };
    case 'inside':
      return { top: on(row > rect.r1), bottom: on(row < rect.r2), left: on(col > rect.c1), right: on(col < rect.c2) };
    case 'top':
      return { top: on(row === rect.r1) };
    case 'bottom':
      return { bottom: on(row === rect.r2) };
    case 'left':
      return { left: on(col === rect.c1) };
    case 'right':
      return { right: on(col === rect.c2) };
  }
}

function describeStructure(op: WorkbookOp, names: string[]): CellChange {
  if (isSheetOp(op)) return describeSheetChange(op, names);
  if (op.type === 'growTable') {
    return { sheet: names[op.sheet], ref: op.table, before: '', after: '', note: msg('core.change.tableGrew', { table: op.table, count: op.rows }) };
  }
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
  const notes: Record<Exclude<StructuralOp['type'], 'move' | 'growTable'>, MessageKey> = {
    insertRows: 'core.change.insertedRows',
    insertCols: 'core.change.insertedColumns',
    deleteRows: 'core.change.deletedRows',
    deleteCols: 'core.change.deletedColumns',
  };
  return { sheet: names[op.sheet], ref, before: '', after: '', note: msg(notes[op.type], { count: op.count }) };
}

/** A change to a sheet in words, listed under the sheet's name (its new name, for one renamed). */
function describeSheetChange(op: SheetOp, names: string[]): CellChange {
  const change = (sheet: string, note: CellChange['note']): CellChange => ({ sheet, ref: '', before: '', after: '', note });
  switch (op.type) {
    case 'addSheet':
      return change(op.name, msg('sheets.change.added'));
    case 'renameSheet':
      return change(op.name, msg('sheets.change.renamed', { from: names[op.sheet] }));
    case 'deleteSheet':
      return change(names[op.sheet], msg('sheets.change.deleted'));
    case 'moveSheet':
      return change(names[op.sheet], msg('sheets.change.moved', { position: op.to + 1 }));
    case 'hideSheet':
      return change(names[op.sheet], msg(op.hidden ? 'sheets.change.hidden' : 'sheets.change.unhidden'));
  }
}
