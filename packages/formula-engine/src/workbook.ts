import type { ArrayDims, EngineChange, EngineLookups, EngineName, EngineSheet, EngineWrite, FormulaEngine, Iteration, RawContent } from './api';
import type { EngineValue } from './engineValue';
import { engineFormula } from './formula';
import { MAX_COLS, MAX_ROWS } from './refs';
import { compileAt, resolve, type Template } from './compile';
import { intersect, type CellNotes, type Ctx, type Host } from './operands';
import { Areas, CALC, err, ERROR_CODES, isError, Lambda, Matrix, REF, Ref, SPILL, VALUE, type Operand, type Scalar } from './values';

/**
 * The workbook as this engine holds it: each sheet's values and formulas,
 * which formulas read which cells, and the values formulas worked out.
 *
 * Values are worked out on demand: reading a formula whose inputs changed
 * works it out first, and so on down, so formulas are calculated in the
 * order they need each other without sorting them. A chain deeper than the
 * JavaScript stack allows is handed back to a stack of our own (Deferred).
 */

const key = (row: number, col: number) => row * MAX_COLS + col;

/** How deep formulas may read other formulas before the rest is put off, to stay well inside the JavaScript stack. */
const MAX_DEPTH = 400;
/** Ranges at most this many columns wide are indexed by column; wider ones are checked one by one. */
const BLOCK = 16;
const WIDE = BLOCK * 4;
/** Ranges of at least this many cells are kept once read, for the other formulas reading them in the same change. */
const KEPT_CELLS = 16;
/** How many such ranges are kept at once, the oldest let go first. */
const KEPT_RANGES = 32;

class Deferred {
  constructor(readonly formula: Formula) {}
}

class Formula implements CellNotes {
  value: Scalar = 0;
  /** Its whole result, when that is more than one value. */
  result?: Matrix;
  /** Where its result spills now. */
  spill?: Ref;
  measured?: ArrayDims;
  link?: string;
  dirty = true;
  busy = false;
  /** Its value before the change being worked out. */
  prev: Scalar | undefined;
  /** Its spill moved, changed size or changed a value since the change began: the session redraws it from its cell. */
  reshaped = false;
  /** Where it is registered as reading cells, to take it out again. */
  reads: (Formula[] | RangeRead)[] = [];

  constructor(
    readonly sheet: number,
    readonly row: number,
    readonly col: number,
    readonly template: Template,
    /** From the file, written before dynamic arrays: ranges where one value is wanted meet its row or column. */
    readonly legacy: boolean,
  ) {}

  /** The range its result is given room in (WBE.SPILL's size), when it spills. */
  room(): Ref | undefined {
    const s = this.template.spill;
    return s ? new Ref(this.sheet, this.row, this.col, this.row + s.rows - 1, this.col + s.cols - 1) : undefined;
  }
}

interface RangeRead {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  f: Formula | null;
}

class Sheet {
  /**
   * Values (never null) and formulas, column by column, each indexed by
   * row: reading down a range is walking an array, and a cell is found
   * without hashing its position.
   */
  readonly columns: (Scalar | Formula | undefined)[][] = [];
  /** Cells a formula's result spills into, with that formula. */
  readonly spilled = new Map<number, Formula>();
  /** Formulas whose results spill. */
  readonly anchors = new Set<Formula>();
  /** Formulas reading a cell, by its key. */
  readonly exact = new Map<number, Formula[]>();
  /** Formulas reading a range, by block of BLOCK columns; wider ranges in `wide`. */
  readonly blocks = new Map<number, RangeRead[]>();
  readonly wide: RangeRead[] = [];
  maxRow = -1;
  maxCol = -1;

  constructor(readonly name: string) {}

  get(row: number, col: number): Scalar | Formula | undefined {
    return this.columns[col]?.[row];
  }

  set(row: number, col: number, v: Scalar | Formula): void {
    let column = this.columns[col];
    if (!column) this.columns[col] = column = [];
    column[row] = v;
  }

  delete(row: number, col: number): void {
    const column = this.columns[col];
    if (column && row < column.length) column[row] = undefined;
  }

  grow(row: number, col: number): void {
    if (row > this.maxRow) this.maxRow = row;
    if (col > this.maxCol) this.maxCol = col;
  }
}

interface NameDef {
  template: Template;
}

/** Whether a value moved by more than the maximum change between passes: numbers by more than `delta`, anything else at all. */
function moved(a: Scalar, b: Scalar, delta: number): boolean {
  return typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) > delta : !same(a, b);
}

function same(a: Scalar | undefined, b: Scalar | undefined): boolean {
  if (isError(a) || isError(b)) return isError(a) && isError(b) && a.error === b.error;
  return Object.is(a, b);
}

/** A cell's content as the session hands it over: values, "'text", error codes, or "=formula". */
function rawValue(raw: RawContent): Scalar {
  if (raw === null || raw === undefined) return null;
  if (typeof raw === 'number' || typeof raw === 'boolean') return raw;
  if (typeof raw === 'string') {
    if (raw.startsWith("'")) return raw.slice(1);
    if (ERROR_CODES.has(raw.toUpperCase())) return err(raw.toUpperCase());
    return raw;
  }
  return null;
}

export class NativeEngine implements FormulaEngine, Host {
  readonly date1904: boolean;
  private readonly sheets: (Sheet | null)[];
  private readonly byName = new Map<string, number>();
  private readonly names = new Map<string, NameDef>();
  /** Formulas using a defined name, by the name in capitals. */
  private readonly nameUsers = new Map<string, Set<Formula>>();
  /** Formulas naming a sheet that isn't in the engine yet, by the sheet's name in lower case. */
  private readonly waiting = new Map<string, Set<Formula>>();
  /** Formulas read while being worked out: where a circular reference closes. */
  private readonly circles = new Set<Formula>();
  /** Reads of a formula while it was being worked out, counted so a range read round a circle isn't kept. */
  private circleReads = 0;
  /**
   * Ranges read during the change being worked out, by sheet and corners, so
   * that formulas looking up the same table share one read of it (and the
   * indexes built over it); let go when the change is done, or a formula
   * within one is worked out again.
   */
  private readonly kept = new Map<string, { area: Ref; values: Matrix }>();
  /** Formulas with three-dimensional references, which a sheet added within their run of sheets joins. */
  private readonly spans = new Set<Formula>();
  private readonly volatile = new Set<Formula>();
  private depth = 0;
  /** Formulas made dirty by the change being worked out. */
  private dirtied: Formula[] = [];
  /** Values changed by the change being worked out, by "sheet:row:col"; none while the engine is built. */
  private changes?: Map<string, EngineChange>;

  constructor(
    sheets: (EngineSheet | null)[],
    names: EngineName[],
    date1904: boolean,
    private readonly lookups?: EngineLookups,
    private readonly iteration?: Iteration,
  ) {
    this.date1904 = date1904;
    this.sheets = sheets.map((s, i) => {
      if (!s) return null;
      this.byName.set(s.name.toLowerCase(), i);
      return new Sheet(s.name);
    });
    for (const n of names) this.defineName(n.name, n.localSheet, n.formula);
    sheets.forEach((s, i) => {
      if (s) this.load(i, s.rows, false);
    });
    this.settle(this.dirtied);
    this.iterate();
    this.dirtied = [];
    this.kept.clear();
  }

  // ---------------------------------------------------------------------------
  // FormulaEngine

  has(sheet: number): boolean {
    return !!this.sheets[sheet];
  }

  addSheet(sheet: number, content: EngineSheet): EngineChange[] {
    if (this.sheets[sheet]) return [];
    this.sheets[sheet] = new Sheet(content.name);
    this.byName.set(content.name.toLowerCase(), sheet);
    this.kept.clear();
    this.load(sheet, content.rows, false);
    this.settle(this.dirtied);
    this.iterate();
    this.dirtied = [];
    this.kept.clear();
    // Formulas that named it before it was here read it now.
    const users = [...(this.waiting.get(content.name.toLowerCase()) ?? [])];
    this.waiting.delete(content.name.toLowerCase());
    for (const f of this.spans) {
      if (this.sheets[f.sheet]?.get(f.row, f.col) !== f) this.spans.delete(f);
      else if (!users.includes(f)) users.push(f);
    }
    return this.change(() => {
      for (const f of users) {
        this.reregister(f);
        this.dirty(f);
      }
      this.touch([], users);
    }).filter((c) => c.sheet !== sheet);
  }

  value(sheet: number, row: number, col: number): Scalar {
    const sh = this.sheets[sheet];
    if (!sh) return null;
    const c = sh.get(row, col);
    if (c !== undefined) {
      if (!(c instanceof Formula)) return c;
      if (c.dirty) this.compute(c);
      // A formula read while it is being worked out reads round in a circle: 0, as Excel shows, or iterating, its value from the pass before.
      if (c.busy) {
        this.circles.add(c);
        this.circleReads++;
        return this.iteration ? c.value : 0;
      }
      return c.value;
    }
    if (sh.anchors.size) {
      for (const a of sh.anchors) {
        if (a.dirty && !a.busy) {
          const room = a.room();
          if (room && row >= room.r1 && row <= room.r2 && col >= room.c1 && col <= room.c2) this.compute(a);
        }
      }
    }
    const owner = sh.spilled.get(key(row, col));
    return owner?.result ? owner.result.get(row - owner.row, col - owner.col) : null;
  }

  hyperlink(sheet: number, row: number, col: number): string | undefined {
    const c = this.sheets[sheet]?.get(row, col);
    return c instanceof Formula ? c.link : undefined;
  }

  isArray(sheet: number, row: number, col: number): boolean {
    const sh = this.sheets[sheet];
    if (!sh) return false;
    const c = sh.get(row, col);
    if (c instanceof Formula) return !!c.spill || (!!c.result && !c.legacy);
    return sh.spilled.has(key(row, col));
  }

  measured(sheet: number, row: number, col: number): ArrayDims | undefined {
    const c = this.sheets[sheet]?.get(row, col);
    return c instanceof Formula ? c.measured : undefined;
  }

  resultSize(formula: string, sheet: number): ArrayDims {
    const v = this.calculate(formula, sheet);
    return Array.isArray(v) && v.length && v[0].length ? { rows: v.length, cols: v[0].length } : { rows: 1, cols: 1 };
  }

  write(writes: EngineWrite[]): EngineChange[] {
    if (!writes.length) return [];
    return this.change(() => {
      const cells: { sheet: number; row: number; col: number }[] = [];
      const areas: Ref[] = [];
      for (const w of writes) {
        if (!this.sheets[w.sheet]) continue;
        const gone = this.put(w.sheet, w.row, w.col, w.raw, true);
        if (gone) areas.push(gone);
        cells.push(w);
      }
      this.touch(cells, [], areas);
    });
  }

  setName(name: string, localSheet: number | undefined, formula: string | null): { changes: EngineChange[]; held: boolean } {
    if (localSheet !== undefined && !this.sheets[localSheet]) return { changes: [], held: false };
    let held = false;
    const changes = this.change(() => {
      held = this.defineName(name, localSheet, formula);
      // Formulas using it, and using names defined with it, and so on.
      const names = new Set([name.toUpperCase()]);
      for (let grew = true; grew; ) {
        grew = false;
        for (const [k, def] of this.names) {
          const n = k.slice(k.indexOf(':') + 1);
          if (!names.has(n) && def.template.names.some((x) => names.has(x))) {
            names.add(n);
            grew = true;
          }
        }
      }
      const users = [...new Set([...names].flatMap((n) => [...(this.nameUsers.get(n) ?? [])]))];
      for (const f of users) {
        this.reregister(f);
        this.dirty(f);
      }
      this.touch([], users);
    });
    return { changes, held };
  }

  calculate(formula: string, sheet: number): EngineValue | EngineValue[][] {
    if (!this.sheets[sheet]) return REF;
    const t = compileAt(formula, 0, 0);
    let out: Operand;
    // Ranges kept while working it out are let go after, as after a change.
    this.kept.clear();
    try {
      out = t.run({ host: this, sheet, row: 0, col: 0, legacy: false });
    } catch (e) {
      if (e instanceof Deferred) throw e;
      return isError(e) ? e : VALUE;
    }
    if (out instanceof Ref) out = out.single ? this.value(out.sheet, out.r1, out.c1) : this.read(out);
    if (out instanceof Lambda || out instanceof Areas) return out instanceof Areas ? VALUE : CALC;
    if (out instanceof Matrix) return out.rows === 1 && out.cols === 1 ? out.get(0, 0) ?? 0 : out.rowsOf();
    // A formula showing an empty cell shows 0.
    return out ?? 0;
  }

  circular(): { sheet: number; row: number; col: number }[] {
    const out: { sheet: number; row: number; col: number }[] = [];
    for (const f of this.circles) {
      if (this.sheets[f.sheet]?.get(f.row, f.col) === f) out.push({ sheet: f.sheet, row: f.row, col: f.col });
      else this.circles.delete(f);
    }
    return out.sort((a, b) => a.sheet - b.sheet || a.row - b.row || a.col - b.col);
  }

  dependentsOf(cells: { sheet: number; row: number; col: number }[], limit = 200_000): { sheet: number; row: number; col: number }[] {
    return this.dependentFormulas(cells, limit).map((f) => ({ sheet: f.sheet, row: f.row, col: f.col }));
  }

  private dependentFormulas(cells: { sheet: number; row: number; col: number }[], limit = 200_000): Formula[] {
    const seen = new Set<Formula>();
    const queue: Formula[] = [];
    const add = (f: Formula) => {
      if (seen.has(f)) return;
      seen.add(f);
      queue.push(f);
    };
    for (const c of cells) this.readersOf(c.sheet, c.row, c.col, add);
    for (let i = 0; i < queue.length && queue.length < limit; i++) {
      const f = queue[i];
      this.readersOf(f.sheet, f.row, f.col, add);
      const area = f.spill ?? f.room();
      if (area) this.readersIn(area, add);
    }
    return queue.slice(0, limit);
  }

  destroy(): void {
    this.sheets.length = 0;
    this.names.clear();
    this.nameUsers.clear();
    this.volatile.clear();
  }

  // ---------------------------------------------------------------------------
  // Host: what formulas read

  sheetIndex(name: string): number {
    return this.byName.get(name.toLowerCase()) ?? -1;
  }

  bound(ref: Ref): Ref {
    const sh = this.sheets[ref.sheet];
    if (!sh || (ref.r2 < MAX_ROWS - 1 && ref.c2 < MAX_COLS - 1)) return ref;
    const r2 = ref.r2 === MAX_ROWS - 1 ? Math.max(ref.r1, Math.min(ref.r2, sh.maxRow)) : ref.r2;
    const c2 = ref.c2 === MAX_COLS - 1 ? Math.max(ref.c1, Math.min(ref.c2, sh.maxCol)) : ref.c2;
    return new Ref(ref.sheet, ref.r1, ref.c1, r2, c2);
  }

  read(ref: Ref): Matrix {
    const r = this.bound(ref);
    const k = `${r.sheet}:${r.r1}:${r.c1}:${r.r2}:${r.c2}`;
    const kept = this.kept.get(k);
    if (kept) return kept.values;
    const circles = this.circleReads;
    const values = Matrix.build(r.rows, r.cols, (i, j) => this.value(r.sheet, r.r1 + i, r.c1 + j));
    if (r.rows * r.cols >= KEPT_CELLS && this.circleReads === circles) {
      if (this.kept.size >= KEPT_RANGES) this.kept.delete(this.kept.keys().next().value!);
      this.kept.set(k, { area: r, values });
    }
    return values;
  }

  /** Lets go of kept ranges an area overlaps, as a formula within them is worked out. */
  private forget(area: Ref | undefined): void {
    if (!area || !this.kept.size) return;
    for (const [k, { area: a }] of this.kept) {
      if (a.sheet === area.sheet && a.r1 <= area.r2 && area.r1 <= a.r2 && a.c1 <= area.c2 && area.c1 <= a.c2) this.kept.delete(k);
    }
  }

  /** Each value in a reference that isn't blank. */
  each(ref: Ref, f: (v: Scalar, row: number, col: number) => void, skipSubtotals = false): void {
    const sh = this.sheets[ref.sheet];
    if (!sh) return;
    const r = this.bound(ref);
    const last = Math.min(r.c2, sh.columns.length - 1);
    for (let col = r.c1; col <= last; col++) {
      const column = sh.columns[col];
      if (!column) continue;
      const end = Math.min(r.r2, column.length - 1);
      for (let row = r.r1; row <= end; row++) {
        const c = column[row];
        if (c === undefined) continue;
        if (!(c instanceof Formula)) {
          f(c, row, col);
          continue;
        }
        if (skipSubtotals && c.template.subtotal) continue;
        if (c.dirty) this.compute(c);
        if (c.busy) this.circles.add(c);
        f(c.busy && !this.iteration ? 0 : c.value, row, col);
      }
    }
    if (!sh.spilled.size) return;
    // Values spilled from formulas, which aren't cells of their own.
    for (const [k, owner] of sh.spilled) {
      const row = Math.floor(k / MAX_COLS);
      const col = k % MAX_COLS;
      if (row < r.r1 || row > r.r2 || col < r.c1 || col > r.c2) continue;
      const v = this.value(r.sheet, row, col);
      if (v !== null && owner.result) f(v, row, col);
    }
  }

  name(name: string, ctx: Ctx): Operand {
    const def = this.names.get(`${ctx.sheet}:${name}`) ?? this.names.get(`-1:${name}`);
    if (!def) return err('#NAME?');
    return def.template.run({ host: this, sheet: ctx.sheet, row: ctx.row, col: ctx.col, legacy: false });
  }

  spillOf(sheet: number, row: number, col: number): Ref | undefined {
    const c = this.sheets[sheet]?.get(row, col);
    if (!(c instanceof Formula)) return undefined;
    if (c.dirty) this.compute(c);
    return c.spill;
  }

  sheetCount(): number {
    return this.sheets.length;
  }

  formulaText(sheet: number, row: number, col: number): string | undefined {
    return this.lookups?.formulaText(sheet, row, col);
  }

  isFormula(sheet: number, row: number, col: number): boolean {
    return this.sheets[sheet]?.get(row, col) instanceof Formula;
  }

  // ---------------------------------------------------------------------------
  // Cells and what reads them

  private load(sheet: number, rows: RawContent[][], typed: boolean): void {
    rows.forEach((line, row) => {
      if (!line) return;
      for (let col = 0; col < line.length; col++) if (line[col] !== null && line[col] !== undefined) this.put(sheet, row, col, line[col], typed);
    });
  }

  /**
   * Puts content in a cell. A formula is made dirty, to be worked out; a
   * value that differs from what the cell showed is a change. Returns where
   * a formula it replaces spilled, which is to be worked out again.
   */
  private put(sheet: number, row: number, col: number, raw: RawContent, typed: boolean): Ref | undefined {
    const sh = this.sheets[sheet]!;
    const old = sh.get(row, col);
    const before = this.changes ? this.peek(sheet, row, col) : undefined;
    let gone: Ref | undefined;
    if (old instanceof Formula) {
      this.unregister(old);
      this.volatile.delete(old);
      sh.anchors.delete(old);
      gone = old.spill;
      if (old.spill) this.clearSpill(old);
      old.dirty = false;
    }
    if (typeof raw === 'string' && raw.startsWith('=')) {
      const t = compileAt(raw, row, col);
      const f = new Formula(sheet, row, col, t, !typed && !t.spill);
      f.prev = before;
      sh.set(row, col, f);
      this.register(f);
      if (t.volatile) this.volatile.add(f);
      if (t.spill) sh.anchors.add(f);
      this.dirtied.push(f);
    } else {
      const v = rawValue(raw);
      if (v === null) sh.delete(row, col);
      else sh.set(row, col, v);
      if (this.changes && !same(before, v)) this.noteChange(sheet, row, col, v);
    }
    sh.grow(row, col);
    return gone;
  }

  /** What a cell shows now, without working anything out. */
  private peek(sheet: number, row: number, col: number): Scalar {
    const sh = this.sheets[sheet];
    if (!sh) return null;
    const c = sh.get(row, col);
    if (c !== undefined) return c instanceof Formula ? c.value : c;
    const owner = sh.spilled.get(key(row, col));
    return owner?.result ? owner.result.get(row - owner.row, col - owner.col) : null;
  }

  private register(f: Formula): void {
    const ctx = { host: this as Host, sheet: f.sheet, row: f.row, col: f.col };
    const add = (spec: Template['refs'][number]) => {
      const ref = resolve(spec, ctx);
      if (spec.lastSheet !== undefined) this.spans.add(f);
      if (ref === REF) {
        for (const name of [spec.sheet, spec.lastSheet]) {
          if (name === null || name === undefined || this.sheetIndex(name) >= 0) continue;
          const k = name.toLowerCase();
          if (!this.waiting.has(k)) this.waiting.set(k, new Set());
          this.waiting.get(k)!.add(f);
        }
        return;
      }
      for (const r of ref instanceof Areas ? ref.refs : [ref as Ref]) this.watch(f, r);
    };
    for (const spec of f.template.refs) add(spec);
    // A name's cells, and those of the names it is defined with.
    const seen = new Set<string>();
    const follow = (name: string) => {
      if (seen.has(name)) return;
      seen.add(name);
      if (!this.nameUsers.has(name)) this.nameUsers.set(name, new Set());
      this.nameUsers.get(name)!.add(f);
      const def = this.names.get(`${f.sheet}:${name}`) ?? this.names.get(`-1:${name}`);
      if (!def) return;
      for (const spec of def.template.refs) add(spec);
      for (const inner of def.template.names) follow(inner);
    };
    for (const name of f.template.names) follow(name);
  }

  /** Notes that a formula reads a cell or range, so changes there make it dirty. */
  private watch(f: Formula, ref: Ref): void {
    const sh = this.sheets[ref.sheet];
    if (!sh) return;
    if (ref.single) {
      const k = key(ref.r1, ref.c1);
      let list = sh.exact.get(k);
      if (!list) sh.exact.set(k, (list = []));
      list.push(f);
      f.reads.push(list);
      return;
    }
    const entry: RangeRead = { r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2, f };
    f.reads.push(entry);
    if (ref.cols > WIDE) {
      sh.wide.push(entry);
      return;
    }
    for (let b = Math.floor(ref.c1 / BLOCK); b <= Math.floor(ref.c2 / BLOCK); b++) {
      let list = sh.blocks.get(b);
      if (!list) sh.blocks.set(b, (list = []));
      list.push(entry);
    }
  }

  private unregister(f: Formula): void {
    for (const r of f.reads) {
      if (Array.isArray(r)) {
        const i = r.indexOf(f);
        if (i >= 0) r.splice(i, 1);
      } else r.f = null;
    }
    f.reads = [];
    for (const users of this.nameUsers.values()) users.delete(f);
    for (const users of this.waiting.values()) users.delete(f);
  }

  private reregister(f: Formula): void {
    this.unregister(f);
    this.register(f);
  }

  /** The formulas that read a cell. */
  private readersOf(sheet: number, row: number, col: number, f: (x: Formula) => void): void {
    const sh = this.sheets[sheet];
    if (!sh) return;
    for (const x of sh.exact.get(key(row, col)) ?? []) f(x);
    const inside = (e: RangeRead) => e.f && row >= e.r1 && row <= e.r2 && col >= e.c1 && col <= e.c2;
    for (const e of sh.blocks.get(Math.floor(col / BLOCK)) ?? []) if (inside(e)) f(e.f!);
    for (const e of sh.wide) if (inside(e)) f(e.f!);
  }

  /** The formulas that read any cell of a range. */
  private readersIn(area: Ref, f: (x: Formula) => void): void {
    const sh = this.sheets[area.sheet];
    if (!sh) return;
    if (area.rows * area.cols <= 4096) {
      for (let row = area.r1; row <= area.r2; row++) for (let col = area.c1; col <= area.c2; col++) this.readersOf(area.sheet, row, col, f);
      return;
    }
    for (const [k, list] of sh.exact) {
      const row = Math.floor(k / MAX_COLS);
      const col = k % MAX_COLS;
      if (row >= area.r1 && row <= area.r2 && col >= area.c1 && col <= area.c2) for (const x of list) f(x);
    }
    const meets = (e: RangeRead) => e.f && e.r1 <= area.r2 && e.r2 >= area.r1 && e.c1 <= area.c2 && e.c2 >= area.c1;
    for (const list of sh.blocks.values()) for (const e of list) if (meets(e)) f(e.f!);
    for (const e of sh.wide) if (meets(e)) f(e.f!);
  }

  // ---------------------------------------------------------------------------
  // Working values out

  private dirty(f: Formula): boolean {
    if (f.dirty) return false;
    f.dirty = true;
    f.prev = f.value;
    this.dirtied.push(f);
    return true;
  }

  /**
   * Makes dirty everything that reads the given cells, areas and formulas,
   * and what reads those, and so on: with the formulas whose spills cover a
   * written cell, and every volatile formula.
   */
  private touch(cells: { sheet: number; row: number; col: number }[], formulas: Formula[], areas: Ref[] = []): void {
    const queue: Formula[] = [...formulas];
    const mark = (f: Formula) => {
      if (this.dirty(f)) queue.push(f);
    };
    for (const c of cells) {
      this.readersOf(c.sheet, c.row, c.col, mark);
      const sh = this.sheets[c.sheet]!;
      const own = sh.get(c.row, c.col);
      if (own instanceof Formula) queue.push(own);
      // Content in a spill's way blocks it; taken away, it lets it spill again.
      for (const a of sh.anchors) {
        const room = a.room();
        if (room && a !== own && c.row >= room.r1 && c.row <= room.r2 && c.col >= room.c1 && c.col <= room.c2) mark(a);
      }
    }
    for (const area of areas) this.readersIn(area, mark);
    for (const v of this.volatile) mark(v);
    for (let i = 0; i < queue.length; i++) {
      const f = queue[i];
      this.readersOf(f.sheet, f.row, f.col, mark);
      const area = f.room() ?? f.spill;
      if (area) this.readersIn(area, mark);
    }
  }

  /** Runs a change and returns the values it changed. */
  private change(run: () => void): EngineChange[] {
    this.changes = new Map();
    this.dirtied = [];
    this.kept.clear();
    try {
      run();
      this.settle(this.dirtied);
      this.iterate();
      for (const f of this.dirtied) {
        if (this.sheets[f.sheet]?.get(f.row, f.col) !== f) continue;
        if (!same(f.prev, f.value) || f.reshaped) this.noteChange(f.sheet, f.row, f.col, f.value);
        f.prev = undefined;
        f.reshaped = false;
      }
      return [...this.changes.values()];
    } finally {
      this.changes = undefined;
      this.dirtied = [];
      this.kept.clear();
    }
  }

  private noteChange(sheet: number, row: number, col: number, value: Scalar): void {
    this.changes?.set(`${sheet}:${row}:${col}`, { sheet, row, col, value });
  }

  /**
   * With iterative calculation on, works the circles the last change closed
   * round again: every formula in or after one, pass after pass (the first
   * was the one just made), until none moves by more than the maximum change
   * or the passes run out, as Excel does. Excel makes one pass fewer than its
   * maximum iterations: a formula counting itself up from 0 reaches 99 in 100.
   */
  private iterate(): void {
    const iteration = this.iteration;
    if (!iteration) return;
    const closing = this.dirtied.filter((f) => this.circles.has(f) && this.sheets[f.sheet]?.get(f.row, f.col) === f);
    if (!closing.length) return;
    const round = [...new Set([...closing, ...this.dependentFormulas(closing)])];
    for (let pass = 1; pass < iteration.count - 1; pass++) {
      const before = round.map((f) => f.value);
      // Formulas the change hadn't reached yet keep the value they had before it, for the changes it reports.
      for (const f of round) if (f.prev === undefined && this.changes) this.dirty(f);
      else f.dirty = true;
      this.kept.clear();
      this.settle(round);
      if (round.every((f, i) => !moved(before[i], f.value, iteration.delta))) break;
    }
  }

  /** Works out every dirty formula among `list`, deepest first where reading one needs another worked out first. */
  private settle(list: Formula[]): void {
    const stack: Formula[] = [];
    for (const start of list) {
      if (!start.dirty) continue;
      stack.push(start);
      while (stack.length) {
        const f = stack[stack.length - 1];
        if (!f.dirty) {
          stack.pop();
          continue;
        }
        try {
          this.compute(f);
          stack.pop();
        } catch (e) {
          if (!(e instanceof Deferred)) throw e;
          stack.push(e.formula);
        }
      }
    }
  }

  private compute(f: Formula): void {
    if (!f.dirty || f.busy) return;
    if (this.depth >= MAX_DEPTH) throw new Deferred(f);
    f.busy = true;
    this.depth++;
    // It closes a circle again only if it is read while being worked out again.
    this.circles.delete(f);
    try {
      f.link = undefined;
      const measured = f.measured;
      const ctx: Ctx = { host: this, sheet: f.sheet, row: f.row, col: f.col, legacy: f.legacy, notes: f };
      let out: Operand;
      try {
        out = f.template.run(ctx);
      } catch (e) {
        if (e instanceof Deferred) throw e;
        out = isError(e) ? e : VALUE;
      }
      // Where a circle closes it shows 0, as the formulas in the circle read it, unless iterating.
      if (this.circles.has(f) && !this.iteration) out = 0;
      this.store(f, out, ctx);
      if (measured?.rows !== f.measured?.rows || measured?.cols !== f.measured?.cols) f.reshaped = true;
      f.dirty = false;
    } finally {
      f.busy = false;
      this.depth--;
    }
  }

  /** Keeps a formula's result: its value, and for a spilling one, the cells it spills into if they are free. */
  private store(f: Formula, out: Operand, ctx: Ctx): void {
    let whole: Matrix | undefined;
    let v: Scalar;
    if (out instanceof Ref) {
      if (out.single) v = this.value(out.sheet, out.r1, out.c1);
      else if (f.legacy) v = intersect(out, ctx);
      else {
        whole = this.read(out);
        v = whole.get(0, 0);
      }
    } else if (out instanceof Matrix) {
      if (out.rows === 1 && out.cols === 1) v = out.get(0, 0);
      else {
        whole = out;
        v = out.get(0, 0);
      }
    } else if (out instanceof Lambda) v = CALC;
    else if (out instanceof Areas) v = VALUE;
    else v = out;
    // A formula showing an empty cell shows 0.
    if (v === null) v = 0;
    if (f.legacy) whole = undefined;

    const oldArea = f.spill;
    const oldResult = f.result;
    this.forget(oldArea ?? new Ref(f.sheet, f.row, f.col, f.row, f.col));
    if (f.template.spill && whole) {
      if (this.placeSpill(f, whole)) {
        f.value = v;
        f.result = whole;
      } else {
        f.value = SPILL;
        f.result = undefined;
      }
    } else {
      if (f.spill) this.clearSpill(f);
      f.value = v;
      f.result = whole;
    }
    this.forget(f.spill);
    if (this.changes && (oldArea || f.spill)) this.spillChanges(f, oldArea, oldResult);
  }

  /** Spills a result from its formula's cell, if every other cell it needs is free; false when something is in its way. */
  private placeSpill(f: Formula, m: Matrix): boolean {
    const sh = this.sheets[f.sheet]!;
    const area = new Ref(f.sheet, f.row, f.col, f.row + m.rows - 1, f.col + m.cols - 1);
    let free = area.r2 < MAX_ROWS && area.c2 < MAX_COLS;
    for (let row = area.r1; free && row <= area.r2; row++) {
      for (let col = area.c1; col <= area.c2; col++) {
        if (row === f.row && col === f.col) continue;
        const k = key(row, col);
        const owner = sh.spilled.get(k);
        if (sh.get(row, col) !== undefined || (owner && owner !== f)) {
          free = false;
          break;
        }
      }
    }
    if (f.spill) this.clearSpill(f);
    if (!free) return false;
    for (let row = area.r1; row <= area.r2; row++) {
      for (let col = area.c1; col <= area.c2; col++) if (row !== f.row || col !== f.col) sh.spilled.set(key(row, col), f);
    }
    f.spill = area;
    sh.grow(area.r2, area.c2);
    return true;
  }

  private clearSpill(f: Formula): void {
    const sh = this.sheets[f.sheet];
    const area = f.spill;
    f.spill = undefined;
    if (!sh || !area) return;
    for (let row = area.r1; row <= area.r2; row++) {
      for (let col = area.c1; col <= area.c2; col++) {
        const k = key(row, col);
        if (sh.spilled.get(k) === f) sh.spilled.delete(k);
      }
    }
  }

  /** Notes the values of a spill's cells that changed, in the area it covered before and the one it covers now. */
  private spillChanges(f: Formula, oldArea: Ref | undefined, oldResult: Matrix | undefined): void {
    const before = (row: number, col: number): Scalar =>
      oldArea && oldResult && row >= oldArea.r1 && row <= oldArea.r2 && col >= oldArea.c1 && col <= oldArea.c2 ? oldResult.get(row - oldArea.r1, col - oldArea.c1) : null;
    const visit = (area: Ref) => {
      for (let row = area.r1; row <= area.r2; row++) {
        for (let col = area.c1; col <= area.c2; col++) {
          if (row === f.row && col === f.col) continue;
          const now = this.peek(f.sheet, row, col);
          if (!same(before(row, col), now)) {
            this.noteChange(f.sheet, row, col, now);
            f.reshaped = true;
          }
        }
      }
    };
    if (oldArea) visit(oldArea);
    if (f.spill) visit(f.spill);
  }

  // ---------------------------------------------------------------------------
  // Names

  /** Defines (or with null removes) a name; false when it can't be calculated here, and is left out. */
  private defineName(name: string, localSheet: number | undefined, formula: string | null): boolean {
    const k = `${localSheet ?? -1}:${name.toUpperCase()}`;
    if (formula === null || name.startsWith('_xlnm.') || /#REF!/.test(formula)) {
      this.names.delete(k);
      return false;
    }
    const template = compileAt(engineFormula(formula), 0, 0);
    this.names.set(k, { template });
    return true;
  }
}
