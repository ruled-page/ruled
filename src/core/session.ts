import { parseA1Ref, tokenize } from '@borgar/fx';
import type { RawCellContent } from 'hyperformula';
import { isDateFormat } from 'numfmt';
import {
  CalcEngine,
  describeValue,
  toRaw,
  unsupportedReason,
  valuesMatch,
  type EngineValue,
  type EngineWrite,
} from './calc';
import { mergeEditResults } from './editResult';
import { planFill, type FillSource } from './fill';
import { displayFormula, isVolatile, makeShifter } from './formula';
import { parseInput } from './input';
import { cellKey, inRect, MAX_COLS, MAX_ROWS, toA1 } from './refs';
import { findCell } from './sheetScanner';
import {
  CellFlag,
  EMPTY,
  VT,
  type CalcReport,
  type CalcSample,
  type CellContent,
  type CellSnapshot,
  type CellUpdate,
  type DisplayStyle,
  type EditOp,
  type EditResult,
  type Finding,
  type LoadResult,
  type Rect,
  type SheetGrid,
} from './types';
import { readWorkbook, type SheetModel, type WorkbookModel } from './workbook';
import { writeWorkbook } from './writer';
import { OoxmlPackage } from './zip';

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

function explainUnsupported(reason: string): string {
  if (reason.endsWith(' function')) return `Uses the ${reason}, which this editor can't calculate.`;
  if (reason === 'Multi-cell array formulas') return "A multi-cell array or spill formula, which this editor can't calculate.";
  if (reason === 'Formula syntax') return "Uses formula syntax this editor can't read.";
  if (reason.startsWith('Shared formula')) return 'Part of a damaged shared formula; Excel will repair it when the file is opened.';
  return `Uses ${reason.charAt(0).toLowerCase()}${reason.slice(1)}, which this editor can't calculate.`;
}

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
  private readonly originalBytes: Uint8Array;
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

  private constructor(pkg: OoxmlPackage, bytes: Uint8Array, fileName: string, model: WorkbookModel, timings: Record<string, number>) {
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
  }

  static open(bytes: Uint8Array, fileName: string, locale = 'en-US'): WorkbookSession {
    const t0 = performance.now();
    const pkg = OoxmlPackage.open(bytes);
    const t1 = performance.now();
    const model = readWorkbook(pkg, locale);
    const t2 = performance.now();
    return new WorkbookSession(pkg, bytes, fileName, model, {
      unzipMs: Math.round(t1 - t0),
      parseMs: Math.round(t2 - t1),
    });
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
    };
  }

  get editedCount(): number {
    return this.edits.reduce((n, m) => n + m.size, 0);
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

  private blockedReason(sheet: number, row: number, col: number): string | undefined {
    const info = this.sheet(sheet).info;
    if (info.kind !== 'worksheet') return 'This sheet cannot be edited here.';
    if (info.protected) return 'This sheet is protected in Excel, so it is read-only here.';
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
    const t0 = performance.now();
    const sheets = this.model.sheets;
    const reasons = new Map<string, number>();
    const cellReasons = new Map<string, string>();
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
        let reason: string | undefined;
        if (flags & CellFlag.ArrayMember) reason = 'Multi-cell array formulas';
        else if (flags & CellFlag.Unsupported) reason = 'Shared formula without its master cell';
        else reason = unsupportedReason(formula);
        if (!reason) continue;
        unsupported++;
        reasons.set(reason, (reasons.get(reason) ?? 0) + 1);
        cellReasons.set(`${s}:${row}:${col}`, explainUnsupported(reason));
        constant[s].add(idx);
        this.calcFlags[s].set(cellKey(row, col), CellFlag.Unsupported);
        this.untrusted.push({ target: { sheet: s, row, col }, reads: formulaReads(formula, s, this.sheetIndex) });
        sample({ sheet: m.info.name, ref: toA1(row, col), formula: displayFormula(formula), reason });
      }
    });

    const engine = new CalcEngine(
      sheets.map((m, s) => {
        const grid = m.scan?.grid;
        if (!grid) return null;
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
              formula !== undefined && grid.flags[i] & CellFlag.Formula && !constant[s].has(i)
                ? toRaw({ kind: 'formula', formula })
                : valueRaw(gridValue(grid, i));
          }
          rows[r] = row;
        }
        return { name: m.info.name, rows };
      }),
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
    sheets.forEach((m, s) => {
      const grid = m.scan?.grid;
      if (!grid) return;
      for (const [idx, formula] of grid.formulas) {
        if (!(grid.flags[idx] & CellFlag.Formula) || constant[s].has(idx)) continue;
        const row = rowOfIndex(grid, idx);
        const col = grid.cols[idx];
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
        cellReasons.set(
          `${p.sheet}:${p.row}:${p.col}`,
          `This editor calculated ${describeValue(ours)} but Excel saved ${describeValue(excel)}, so Excel's value is kept.`,
        );
        this.calcFlags[p.sheet].set(cellKey(p.row, p.col), CellFlag.Mismatch);
        this.untrusted.push({ target: p, reads: formulaReads(grid.formulas.get(idx) ?? '', p.sheet, this.sheetIndex) });
        sample({
          sheet: sheets[p.sheet].info.name,
          ref: toA1(p.row, p.col),
          formula: displayFormula(grid.formulas.get(idx) ?? ''),
          reason: 'Calculated differently from Excel',
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
      const list = [...reasons.entries()].sort((a, b) => b[1] - a[1]).map(([r, n]) => `${r} (${n.toLocaleString()})`);
      findings.push({
        kind: 'formulaEngine',
        severity: 'warn',
        title: 'Formulas not recalculated here',
        detail: `${unsupported.toLocaleString()} of ${formulaCells.toLocaleString()} formulas use features this editor can't calculate: ${list.join(', ')}. They show Excel's last value and are marked with a grey corner; if their inputs change they turn amber. Excel recalculates them when the file is opened.`,
      });
    }
    if (mismatched) {
      findings.push({
        kind: 'formulaEngine',
        severity: 'warn',
        title: 'Formulas calculated differently here',
        detail: `${mismatched.toLocaleString()} formulas gave a different result here than Excel's saved value, so this editor shows Excel's value and does not recalculate them. Excel recalculates them when the file is opened.`,
      });
    }

    this.report = {
      buildMs: Math.round(performance.now() - t0),
      formulaCells,
      liveCells: formulaCells - unsupported - mismatched,
      unsupported,
      mismatched,
      noCachedValue,
      unsupportedFunctions: [...reasons.keys()],
      samples,
      updates,
      findings,
      cellReasons,
    };
    return this.report;
  }

  /** Marks formulas the engine cannot vouch for after the given cells changed. */
  private markStale(changed: Pos[], extraRoots: Pos[] = []): Pos[] {
    const roots: Pos[] = [...extraRoots];
    if (this.untrusted.length && changed.length) {
      for (const u of this.untrusted) {
        if (this.stale[u.target.sheet].has(cellKey(u.target.row, u.target.col))) continue;
        if (!this.isUntrustedOriginal(u.target.sheet, u.target.row, u.target.col)) continue;
        const hit =
          u.reads === null ||
          changed.some((c) => u.reads!.some((r) => r.sheet === c.sheet && inRect(r.rect, c.row, c.col)));
        if (hit) roots.push(u.target);
      }
    }
    if (!roots.length) return [];
    const out: Pos[] = [];
    const mark = (p: Pos) => {
      const key = cellKey(p.row, p.col);
      if (this.stale[p.sheet].has(key)) return;
      if (this.current(p.sheet, p.row, p.col).content.kind !== 'formula') return;
      this.stale[p.sheet].add(key);
      out.push(p);
    };
    for (const r of roots) if (!extraRoots.includes(r)) mark(r);
    for (const d of this.engine!.dependentsOf(roots)) mark(d);
    return out;
  }

  // ---------------------------------------------------------------------------
  // Editing

  apply(op: EditOp): EditResult {
    this.ensureEngine();
    if (op.type === 'batch') return mergeEditResults(op.ops.map((o) => this.apply(o)));
    const writes = new Map<string, Write>();
    const warnings: string[] = [];
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
        const parsed = parseInput(op.text, this.model.styles.get(cur.style).numFmt);
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
            const parsed = parseInput(text, this.model.styles.get(cur.style).numFmt);
            if (parsed.error) bad++;
            const content = parsed.error ? { kind: 'string' as const, value: text } : parsed.content;
            put({ sheet: op.sheet, row, col, content, style: styleFor(cur.style, parsed.numFmt) });
          }),
        );
        if (bad) warnings.push(`${bad} pasted value${bad === 1 ? '' : 's'} looked like a formula but wasn't valid, so ${bad === 1 ? 'it was' : 'they were'} pasted as text.`);
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
        if (lost) warnings.push(`${lost} formula${lost === 1 ? '' : 's'} couldn't be adjusted for the new position and ${lost === 1 ? 'was' : 'were'} pasted as values.`);
        break;
      }
      case 'fill': {
        warnings.push(...this.planFillWrites(op.sheet, op.src, op.dst, put));
        break;
      }
      case 'restore': {
        for (const c of op.cells) put({ ...c });
        break;
      }
      case 'inputs': {
        for (const c of op.cells) {
          const cur = this.current(op.sheet, c.row, c.col);
          const parsed = parseInput(c.text, this.model.styles.get(cur.style).numFmt);
          if (parsed.error) warnings.push(parsed.error);
          const content = parsed.error ? { kind: 'string' as const, value: c.text } : parsed.content;
          put({ sheet: op.sheet, row: c.row, col: c.col, content, style: styleFor(cur.style, parsed.numFmt) });
        }
        break;
      }
    }
    return this.commit([...writes.values()], warnings, newStyles);
  }

  private emptyResult(rejected: string): EditResult {
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

  private planFillWrites(sheet: number, src: Rect, dst: Rect, put: (w: Write) => void): string[] {
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
      const plans = planFill(sources, targets.map((t) => t.x));
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
    return lost ? [`${lost} formula${lost === 1 ? '' : 's'} couldn't be adjusted and ${lost === 1 ? 'was' : 'were'} filled as values.`] : [];
  }

  private commit(writes: Write[], warnings: string[], newStyles: { index: number; style: DisplayStyle }[]): EditResult {
    const engine = this.engine!;
    const undo: CellSnapshot[] = [];
    const redo: CellSnapshot[] = [];
    const reasons = new Set<string>();
    let skipped = 0;
    const applied: Write[] = [];

    for (const w of writes) {
      const reason = this.blockedReason(w.sheet, w.row, w.col);
      if (reason) {
        if (reason !== MERGED) {
          skipped++;
          reasons.add(reason);
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
      return {
        sheet: w.sheet,
        row: w.row,
        col: w.col,
        raw: untrusted && grid ? valueRaw(gridValue(grid, untrusted.idx)) : toRaw(w.content),
      };
    });
    const changes = engine.write(engineWrites);

    // Formulas typed in this session that the engine can't calculate.
    const newUnsupported: Pos[] = [];
    for (const w of applied) {
      if (w.content.kind !== 'formula' || this.isUntrustedOriginal(w.sheet, w.row, w.col)) continue;
      const key = cellKey(w.row, w.col);
      if (unsupportedReason(w.content.formula)) {
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
      skippedReasons: [...reasons],
      warnings,
      newStyles,
      editedCount: this.editedCount,
    };
  }

  // ---------------------------------------------------------------------------
  // Saving

  /** The workbook as .xlsx bytes. Unchanged workbooks come back byte for byte. */
  save(): Uint8Array {
    if (this.editedCount === 0) return this.originalBytes;
    return writeWorkbook({
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
    });
  }
}
