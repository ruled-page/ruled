import { isSpillError, MAX_SPILL_CELLS, spillFormula, type ArrayDims, type EngineChange, type EngineName, type EngineWrite, type RawContent } from './api';
import type { EngineValue } from './engineValue';
import { displayFormula, engineFormula, toFileFormula } from './formula';
import { dateOf, serialOf } from './functions';
import { cellKey, inRect, keyCol, keyRow, lettersToCol, MAX_COLS, MAX_ROWS, parseRange, quoteSheetName, toA1, type Rect } from './refs';
import { changeSheetRefs, lineMap, transformFormula, type LineOp, type SheetRefChange } from './reshape';
import { NativeEngine } from './workbook';

/**
 * A workbook worked out with no editor: sheets by name, cells by A1
 * reference, values and formulas in, Excel's results out, kept up to date
 * as cells change.
 */

/** What a cell can be given: a value, or a formula starting with "=" as it would be typed in Excel. A Date is read in UTC. */
export type CellInput = string | number | boolean | Date | null | undefined;

/** A cell's value: a number (dates as Excel's serial numbers), text, TRUE or FALSE, an error such as `{ error: '#DIV/0!' }`, or null when empty. */
export type CellValue = EngineValue;

/** A cell whose value changed, by its reference with its sheet ("Summary!B2"). */
export interface CellChange {
  cell: string;
  value: CellValue;
}

export interface WorkbookOptions {
  /** The sheets in order, by name, each as rows of values and formulas from A1. */
  sheets?: Record<string, CellInput[][]>;
  /** Defined names, each a reference or formula: `{ Rate: 'Prices!$B$2' }`. */
  names?: Record<string, string>;
  /** Counts dates from 1904, as workbooks made on older Macs do. */
  date1904?: boolean;
}

export interface Workbook {
  /** The sheets' names, in order. */
  readonly sheets: readonly string[];
  /** A cell's value: "Prices!B2", or "B2" on the first sheet. */
  get(cell: string): CellValue;
  /** A range's values, as rows: "Prices!A1:C10". */
  getRange(range: string): CellValue[][];
  /** The formula in a cell as it was given ("=SUM(B2:B10)"), or undefined when it holds none. */
  formula(cell: string): string | undefined;
  /** Puts a value or formula in a cell, or rows of them from that cell; returns every value that changed as a result. */
  set(cell: string, input: CellInput | CellInput[][]): CellChange[];
  /** Puts values and formulas in several cells at once. */
  set(cells: Record<string, CellInput>): CellChange[];
  /** Works out a formula without putting it in a cell; references without a sheet are to `sheet`, or the first. */
  evaluate(formula: string, sheet?: string): CellValue | CellValue[][];
  /** Adds a sheet after the others; formulas that named it already read it now. */
  addSheet(name: string, rows?: CellInput[][]): CellChange[];
  /** Defines or changes a name, or removes it with null. */
  setName(name: string, formula: string | null): CellChange[];
  /** Renames a sheet; formulas that refer to it follow it. */
  renameSheet(sheet: string, name: string): CellChange[];
  /** Moves a sheet to a place among the others, 0 being the first. */
  moveSheet(sheet: string, to: number): CellChange[];
  /** Deletes a sheet; references to it become #REF!. */
  deleteSheet(sheet: string): CellChange[];
  /** Inserts empty rows where these are, the rows from there on moving down: "Prices!3:4" puts two before row 3. */
  insertRows(rows: string): CellChange[];
  /** Deletes rows, those below moving up: "Prices!3:4", or "3". */
  deleteRows(rows: string): CellChange[];
  /** Inserts empty columns where these are, the columns from there on moving right: "Prices!B:C" puts two before column B. */
  insertColumns(columns: string): CellChange[];
  /** Deletes columns, those to their right moving left: "Prices!B:C", or "B". */
  deleteColumns(columns: string): CellChange[];
}

interface Pos {
  sheet: number;
  row: number;
  col: number;
}

/** Makes a workbook from sheets of values and formulas. */
export function createWorkbook(options: WorkbookOptions = {}): Workbook {
  return new Book(options);
}

interface Given {
  /** The formula as given, with "=". */
  given: string;
  /** In the engine's form. */
  raw: string;
}

class Book implements Workbook {
  private names: string[] = [];
  private engine: NativeEngine;
  /** Each sheet's formulas, by cell key. */
  private formulas: Map<number, Given>[] = [];
  /** The room given to each spilling formula, by sheet and cell key. */
  private spills: Map<number, ArrayDims>[] = [];
  /** The defined names in the file's form, by name in capitals. */
  private readonly defined = new Map<string, EngineName>();
  private readonly date1904: boolean;

  constructor(options: WorkbookOptions) {
    this.date1904 = options.date1904 ?? false;
    const entries = Object.entries(options.sheets ?? {});
    if (!entries.length) entries.push(['Sheet1', []]);
    const formulas: EngineWrite[] = [];
    const sheets = entries.map(([name, rows], sheet) => {
      this.addName(name);
      return { name, rows: this.load(sheet, rows, formulas) };
    });
    for (const [name, formula] of Object.entries(options.names ?? {})) this.defined.set(name.toUpperCase(), { name, formula: fileForm(formula) });
    this.engine = new NativeEngine(sheets, [...this.defined.values()], this.date1904);
    this.engine.write(formulas);
    this.fitSpills(formulas, []);
  }

  get sheets(): readonly string[] {
    return this.names;
  }

  get(cell: string): CellValue {
    const { sheet, rect } = this.locate(cell);
    if (rect.r1 !== rect.r2 || rect.c1 !== rect.c2) throw new Error(`"${cell}" is a range; getRange reads ranges.`);
    return this.engine.value(sheet, rect.r1, rect.c1);
  }

  getRange(range: string): CellValue[][] {
    const { sheet, rect } = this.locate(range);
    return Array.from({ length: rect.r2 - rect.r1 + 1 }, (_r, r) => Array.from({ length: rect.c2 - rect.c1 + 1 }, (_c, c) => this.engine.value(sheet, rect.r1 + r, rect.c1 + c)));
  }

  formula(cell: string): string | undefined {
    const { sheet, rect } = this.locate(cell);
    return this.formulas[sheet].get(cellKey(rect.r1, rect.c1))?.given;
  }

  set(target: string | Record<string, CellInput>, input?: CellInput | CellInput[][]): CellChange[] {
    const cells: [Pos, CellInput][] = [];
    if (typeof target === 'string') {
      const { sheet, rect } = this.locate(target);
      const rows = Array.isArray(input) ? input : [[input]];
      rows.forEach((row, r) => row.forEach((v, c) => cells.push([{ sheet, row: rect.r1 + r, col: rect.c1 + c }, v])));
    } else {
      for (const [ref, v] of Object.entries(target)) {
        const { sheet, rect } = this.locate(ref);
        cells.push([{ sheet, row: rect.r1, col: rect.c1 }, v]);
      }
    }
    const rooms = this.rooms(cells.map(([p]) => p));
    const writes: EngineWrite[] = cells.map(([p, v]) => ({ ...p, raw: this.prepare(p, v) }));
    const changes = this.engine.write(writes);
    const fitted = this.fitSpills(writes.filter((w) => this.formulas[w.sheet].has(cellKey(w.row, w.col))), changes);
    return this.report([...changes, ...fitted, ...this.left(rooms)]);
  }

  evaluate(formula: string, sheet?: string): CellValue | CellValue[][] {
    return this.engine.calculate(engineFormula(fileForm(formula)), sheet === undefined ? 0 : this.sheetIndex(sheet));
  }

  addSheet(name: string, rows: CellInput[][] = []): CellChange[] {
    const sheet = this.names.length;
    this.addName(name);
    const formulas: EngineWrite[] = [];
    const changes = this.engine.addSheet(sheet, { name, rows: this.load(sheet, rows, formulas) });
    changes.push(...this.engine.write(formulas));
    return this.report([...changes, ...this.fitSpills(formulas, changes)]);
  }

  setName(name: string, formula: string | null): CellChange[] {
    const file = formula === null ? null : fileForm(formula);
    const { changes } = this.engine.setName(name, undefined, file);
    if (file === null) this.defined.delete(name.toUpperCase());
    else this.defined.set(name.toUpperCase(), { name, formula: file });
    return this.report([...changes, ...this.fitSpills([], changes)]);
  }

  renameSheet(sheet: string, name: string): CellChange[] {
    const at = this.sheetIndex(sheet);
    this.checkName(name, at);
    const change: SheetRefChange = { sheet: this.names[at].toLowerCase(), to: name, order: [], tables: new Set() };
    return this.reshape(
      this.names.map((n, i) => (i === at ? name : n)),
      this.names.map((_n, i) => i),
      (_sheet, row, col) => ({ row, col }),
      (formula) => changeSheetRefs(formula, change),
    );
  }

  moveSheet(sheet: string, to: number): CellChange[] {
    const from = this.sheetIndex(sheet);
    if (!Number.isInteger(to) || to < 0 || to >= this.names.length) throw new Error(`A sheet can't be moved to ${to}: the sheets are numbered from 0 to ${this.names.length - 1}.`);
    const order = this.names.map((_n, i) => i);
    order.splice(to, 0, ...order.splice(from, 1));
    // References name sheets, so they stay as they are; those spanning sheets ("Jan:Mar!A1") take in the sheets now between their ends.
    return this.reshape(
      order.map((i) => this.names[i]),
      order,
      (_sheet, row, col) => ({ row, col }),
      (formula) => formula,
    );
  }

  deleteSheet(sheet: string): CellChange[] {
    const at = this.sheetIndex(sheet);
    if (this.names.length === 1) throw new Error(`"${this.names[at]}" is the workbook's only sheet, which can't be deleted.`);
    const change: SheetRefChange = { sheet: this.names[at].toLowerCase(), to: null, order: this.names, tables: new Set() };
    const order = this.names.map((_n, i) => i).filter((i) => i !== at);
    return this.reshape(
      order.map((i) => this.names[i]),
      order,
      (_sheet, row, col) => ({ row, col }),
      (formula) => changeSheetRefs(formula, change),
    );
  }

  insertRows(rows: string): CellChange[] {
    return this.lines(rows, 'insertRows');
  }

  deleteRows(rows: string): CellChange[] {
    return this.lines(rows, 'deleteRows');
  }

  insertColumns(columns: string): CellChange[] {
    return this.lines(columns, 'insertCols');
  }

  deleteColumns(columns: string): CellChange[] {
    return this.lines(columns, 'deleteCols');
  }

  /** Inserts or deletes whole rows or columns, references following the cells they name as Excel moves them. */
  private lines(ref: string, type: LineOp['type']): CellChange[] {
    const rows = type.endsWith('Rows');
    const { sheet, rest } = this.split(ref);
    const m = (rows ? /^\$?(\d+)(?::\$?(\d+))?$/ : /^\$?([A-Za-z]{1,3})(?::\$?([A-Za-z]{1,3}))?$/).exec(rest);
    const ends = m ? [m[1], m[2] ?? m[1]].map((e) => (rows ? Number(e) - 1 : lettersToCol(e))) : [];
    if (!m || ends.some((e) => e < 0 || e >= (rows ? MAX_ROWS : MAX_COLS))) throw new Error(`"${ref}" isn't a reference to ${rows ? 'rows, such as "3:4"' : 'columns, such as "B:C"'}.`);
    const map = lineMap({ type, sheet, at: Math.min(...ends), count: Math.abs(ends[1] - ends[0]) + 1 });
    const ctx = { names: this.names, indexOf: (name: string) => this.names.findIndex((n) => n.toLowerCase() === name.toLowerCase()) };
    return this.reshape(
      this.names,
      this.names.map((_n, i) => i),
      (s, row, col) => {
        const to = map.range(s, { r1: row, c1: col, r2: row, c2: col });
        return to && { row: to.rect.r1, col: to.rect.c1 };
      },
      // Names aren't on a sheet, so only their references with one move.
      (formula, s = -1) => transformFormula(formula, s, s, map, ctx),
    );
  }

  /**
   * Builds the engine again in the workbook's new shape: `names` are the
   * sheets now, `order` the sheet each was before, `place` where a cell on
   * a sheet as it was goes (null when it goes altogether), and `rewrite` a
   * formula's new text, in the file's form, on a sheet as it was (none for
   * a defined name). Returns the cells, where they are now, whose values
   * aren't what they were.
   */
  private reshape(
    names: string[],
    order: number[],
    place: (sheet: number, row: number, col: number) => { row: number; col: number } | null,
    rewrite: (formula: string, sheet?: number) => string,
  ): CellChange[] {
    const before = new Map<string, EngineChange>();
    const rows: RawContent[][][] = order.map(() => []);
    const formulas = order.map(() => new Map<number, Given>());
    const written: EngineWrite[] = [];
    order.forEach((old, sheet) => {
      const follow = (row: number, col: number) => {
        const to = place(old, row, col);
        if (to) before.set(`${sheet}:${to.row}:${to.col}`, { sheet, ...to, value: this.engine.value(old, row, col) });
        return to;
      };
      for (const c of this.engine.constants(old)) {
        const to = follow(c.row, c.col);
        // Text keeps its quote, so it is never read as an error or a formula.
        if (to) (rows[sheet][to.row] ??= [])[to.col] = typeof c.value === 'string' ? `'${c.value}` : c.value !== null && typeof c.value === 'object' ? c.value.error : c.value;
      }
      for (const [key, f] of this.formulas[old]) {
        const to = follow(keyRow(key), keyCol(key));
        if (!to) continue;
        const file = fileForm(f.given);
        const next = rewrite(file, old);
        const given = next === file ? f : { given: displayFormula(next), raw: engineFormula(next) };
        formulas[sheet].set(cellKey(to.row, to.col), given);
        written.push({ sheet, ...to, raw: given.raw });
      }
      for (const cell of this.spilling(old)) follow(cell.row, cell.col);
    });
    for (const n of this.defined.values()) n.formula = rewrite(n.formula);
    this.engine.destroy();
    this.names = names;
    this.formulas = formulas;
    this.spills = order.map(() => new Map());
    this.engine = new NativeEngine(
      names.map((name, i) => ({ name, rows: rows[i] })),
      [...this.defined.values()],
      this.date1904,
    );
    this.engine.write(written);
    this.fitSpills(written, []);
    const changes: EngineChange[] = [];
    const seen = new Set<string>();
    const check = (sheet: number, row: number, col: number) => {
      const id = `${sheet}:${row}:${col}`;
      if (seen.has(id)) return;
      seen.add(id);
      const value = this.engine.value(sheet, row, col);
      if (!same(value, before.get(id)?.value ?? null)) changes.push({ sheet, row, col, value });
    };
    for (const c of before.values()) check(c.sheet, c.row, c.col);
    for (const c of written) check(c.sheet, c.row, c.col);
    names.forEach((_n, sheet) => this.spilling(sheet).forEach((c) => check(sheet, c.row, c.col)));
    return this.report(changes);
  }

  /** Throws unless Excel allows the name for a sheet and no other sheet (but `self`) has it. */
  private checkName(name: string, self?: number) {
    if (!name || /[\\/?*[\]:]/.test(name) || /^'|'$/.test(name) || name.length > 31) throw new Error(`"${name}" isn't a sheet name Excel allows.`);
    if (this.names.some((n, i) => i !== self && n.toLowerCase() === name.toLowerCase())) throw new Error(`There is already a sheet named "${name}".`);
  }

  private addName(name: string) {
    this.checkName(name);
    this.names.push(name);
    this.formulas.push(new Map());
    this.spills.push(new Map());
  }

  private sheetIndex(name: string): number {
    const sheet = this.names.findIndex((n) => n.toLowerCase() === name.toLowerCase());
    if (sheet < 0) throw new Error(`There is no sheet named "${name}".`);
    return sheet;
  }

  /** A reference's sheet, the first when it names none, and the rest of it: "'Q1 Sales'!A1" is "Q1 Sales" and "A1". */
  private split(ref: string): { sheet: number; rest: string } {
    const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(ref.trim());
    return m ? { sheet: this.sheetIndex(m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]), rest: m[3] } : { sheet: 0, rest: ref.trim() };
  }

  /** The sheet and cells a reference names: "B2", "Prices!B2:C4" or "'Q1 Sales'!A1". */
  private locate(ref: string): { sheet: number; rect: Rect } {
    const { sheet, rest } = this.split(ref);
    const rect = parseRange(rest);
    if (!rect || [rect.r1, rect.r2].some((r) => r < 0 || r >= 1_048_576) || rect.c2 >= 16_384) throw new Error(`"${ref}" isn't a cell or range reference.`);
    return { sheet, rect };
  }

  /** What the engine is given for a cell's input, keeping its formula when it has one. */
  private prepare(p: Pos, input: CellInput): RawContent {
    const key = cellKey(p.row, p.col);
    this.spills[p.sheet].delete(key);
    if (typeof input === 'string' && input.startsWith('=')) {
      const raw = engineFormula(fileForm(input));
      this.formulas[p.sheet].set(key, { given: input, raw });
      return raw;
    }
    this.formulas[p.sheet].delete(key);
    return input instanceof Date ? dateToSerial(input, this.date1904) : (input ?? null);
  }

  /**
   * A sheet's values for the engine to load, with its formulas left out and
   * added to `formulas`: written in afterwards, they are worked out as
   * formulas typed into Excel today are, spilling, rather than as formulas
   * from an older file.
   */
  private load(sheet: number, rows: CellInput[][], formulas: EngineWrite[]): RawContent[][] {
    return rows.map((row, r) =>
      row.map((input, c) => {
        const raw = this.prepare({ sheet, row: r, col: c }, input);
        if (!this.formulas[sheet].has(cellKey(r, c))) return raw;
        formulas.push({ sheet, row: r, col: c, raw });
        return null;
      }),
    );
  }

  /** The cells beside a formula that its spill covers. */
  private room(p: Pos): Pos[] {
    const spill = this.spills[p.sheet].get(cellKey(p.row, p.col));
    const cells: Pos[] = [];
    for (let r = 0; spill && r < spill.rows; r++) {
      for (let c = r ? 0 : 1; c < spill.cols; c++) cells.push({ sheet: p.sheet, row: p.row + r, col: p.col + c });
    }
    return cells;
  }

  /** The cells a sheet's formulas spill into. */
  private spilling(sheet: number): Pos[] {
    return [...this.spills[sheet].keys()].flatMap((key) => this.room({ sheet, row: keyRow(key), col: keyCol(key) }));
  }

  /** The cells beside these formulas that their spills cover, with their values, to see which a spill leaves. */
  private rooms(anchors: Pos[]): (Pos & { value: EngineValue })[] {
    return anchors.flatMap((p) => this.room(p).map((c) => ({ ...c, value: this.engine.value(c.sheet, c.row, c.col) })));
  }

  /** Those of `rooms`' cells whose values are different now. */
  private left(rooms: (Pos & { value: EngineValue })[]): EngineChange[] {
    return rooms.flatMap(({ sheet, row, col, value }) => {
      const now = this.engine.value(sheet, row, col);
      return same(now, value) ? [] : [{ sheet, row, col, value: now }];
    });
  }

  /**
   * Gives each spilling formula room in the engine for the result it has, as
   * Excel spills it: formulas just written whose results are arrays, results
   * that have grown or shrunk, and blocked formulas whose way may be clear
   * now. The engine sizes an array before working it out and checks a
   * blocked formula's room only when it is put in again, so each is put in
   * at the size its result really has. Returns the engine's changes.
   */
  private fitSpills(written: Pos[], changes: (Pos & { value?: EngineValue })[]): EngineChange[] {
    const engine = this.engine;
    const out: EngineChange[] = [];
    const refitted = new Set<string>();
    let fresh = written;
    let changed = changes;
    for (let round = 0; round < 8 && (fresh.length || changed.length); round++) {
      const fits = new Map<string, EngineWrite & ArrayDims>();
      const fit = (p: Pos, size: ArrayDims) => {
        const raw = this.formulas[p.sheet].get(cellKey(p.row, p.col))?.raw;
        if (raw && size.rows * size.cols <= MAX_SPILL_CELLS) fits.set(`${p.sheet}:${p.row}:${p.col}`, { ...p, raw: spillFormula(raw, size.rows, size.cols), ...size });
      };
      // Formulas without room yet whose results turn out to be arrays.
      const probe = (p: Pos) => {
        const raw = this.formulas[p.sheet].get(cellKey(p.row, p.col))?.raw;
        if (!raw || this.spills[p.sheet].get(cellKey(p.row, p.col))) return;
        // Read first: whether a result is an array is known once it's worked out.
        const v = engine.value(p.sheet, p.row, p.col);
        const array = engine.isArray(p.sheet, p.row, p.col);
        if (!array && !(v !== null && typeof v === 'object')) return;
        const size = engine.resultSize(raw, p.sheet);
        if (array || size.rows * size.cols > 1) fit(p, size);
      };
      for (const p of fresh) probe(p);
      for (const p of changed) probe(p);
      const moved = [...fresh, ...changed];
      this.spills.forEach((spills, sheet) => {
        for (const [key, spill] of spills) {
          const row = keyRow(key);
          const col = keyCol(key);
          // Results that have grown or shrunk since they were given room, though their first value may be the same.
          const m = engine.measured(sheet, row, col);
          if (m && (m.rows !== spill.rows || m.cols !== spill.cols)) fit({ sheet, row, col }, m);
          // Blocked formulas, when something in their way changed.
          if (refitted.has(`${sheet}:${row}:${col}`) || !isSpillError(engine.value(sheet, row, col))) continue;
          const room = { r1: row, c1: col, r2: row + spill.rows - 1, c2: col + spill.cols - 1 };
          if (moved.some((p) => p.sheet === sheet && inRect(room, p.row, p.col))) fit({ sheet, row, col }, engine.resultSize(this.formulas[sheet].get(key)!.raw, sheet));
        }
      });
      if (!fits.size) break;
      const rooms = this.rooms([...fits.values()]);
      for (const [id, f] of fits) {
        refitted.add(id);
        this.spills[f.sheet].set(cellKey(f.row, f.col), { rows: f.rows, cols: f.cols });
      }
      const fitted = engine.write([...fits.values()].map(({ sheet, row, col, raw }) => ({ sheet, row, col, raw })));
      out.push(...fitted, ...this.left(rooms));
      changed = fitted;
      fresh = [];
    }
    return out;
  }

  /** The engine's changes by cell reference, each cell once with its latest value. */
  private report(changes: EngineChange[]): CellChange[] {
    const latest = new Map<string, CellValue>();
    for (const c of changes) {
      const cell = `${quoteSheetName(this.names[c.sheet])}!${toA1(c.row, c.col)}`;
      latest.delete(cell);
      latest.set(cell, c.value);
    }
    return [...latest].map(([cell, value]) => ({ cell, value }));
  }
}

function same(a: EngineValue, b: EngineValue): boolean {
  return a === b || (a !== null && b !== null && typeof a === 'object' && typeof b === 'object' && a.error === b.error);
}

const DAY_MS = 86_400_000;

/** A date's serial number, as Excel stores dates (the days since 1900 began, or 1904 for `date1904` workbooks), from its UTC date and time. */
export function dateToSerial(date: Date, date1904 = false): number {
  const day = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  return serialOf(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate(), date1904) + (date.getTime() - day) / DAY_MS;
}

/** The date, in UTC, that a serial number stands for. */
export function serialToDate(serial: number, date1904 = false): Date {
  const { y, m, d } = dateOf(serial, date1904);
  return new Date(Date.UTC(y, m - 1, d) + Math.round((serial - Math.floor(serial)) * DAY_MS));
}

/** A formula or reference as given ("=SUM(A1:A3)", or "Prices!$B$2" for a name) in the file's form, without "=". */
function fileForm(formula: string): string {
  return toFileFormula(formula.replace(/^=/, ''));
}

