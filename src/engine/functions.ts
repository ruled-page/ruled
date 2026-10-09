import { format } from 'numfmt';
import { parseA1Ref, parseR1C1Ref } from '@borgar/fx';
import { excelRound } from '../core/excelNumbers';
import { MAX_COLS, MAX_ROWS } from '../core/refs';
import { MAX_ARRAY_CELLS } from '../core/sizeLimits';
import { hasWildcards, matchesWhole, parsePattern, searchIn } from '../core/wildcard';
import { scalarOf, valuesOf, type Ctx } from './operands';
import { bool, compare, DIV0, err, isError, Lambda, Matrix, NA, NUM, num, REF, Ref, round15, text, VALUE, type Operand, type Scalar } from './values';

/**
 * Excel's functions, as this engine works them out. A function either takes
 * one value per parameter (`scalar`), and is run on each value of an array
 * given instead, or takes its arguments as they come (ranges, arrays), as
 * SUM does; IF and its kind (`lazy`) evaluate only the arguments they need.
 * A function throws an error value to give it as its result.
 */
export interface FunctionSpec {
  min: number;
  max: number;
  /** Parameters taken one value at a time: all of them, or these. */
  scalar?: true | number[];
  fn?: (args: Operand[], ctx: Ctx) => Operand;
  lazy?: (args: ((ctx: Ctx) => Operand)[], ctx: Ctx) => Operand;
  /** Worked out again whenever anything changes. */
  volatile?: boolean;
}

const MANY = 255;

/** A function of single values. */
const scalar = (min: number, max: number, fn: (a: Scalar[], ctx: Ctx) => Operand): FunctionSpec => ({ min, max, scalar: true, fn: fn as FunctionSpec['fn'] });

/** A function of one number. */
const math1 = (f: (x: number) => number | Scalar): FunctionSpec =>
  scalar(1, 1, ([x]) => {
    const r = f(num(x));
    return typeof r === 'number' && !Number.isFinite(r) ? NUM : r;
  });

const optional = (v: Scalar | undefined, d: number) => (v === undefined || v === null ? d : num(v));

// -----------------------------------------------------------------------------
// Numbers in arguments, as aggregates count them

interface Collect {
  /** Typed values (not from ranges or arrays) that aren't numbers: counted (COUNT), or #VALUE! (SUM). */
  typed: 'coerce' | 'count';
  skipSubtotals?: boolean;
}

/** Each number among the arguments: from ranges and arrays, only numbers; typed values read as numbers; errors stop it. */
function numbers(args: Operand[], ctx: Ctx, f: (n: number) => void, how: Collect = { typed: 'coerce' }): void {
  const fromCells = (v: Scalar) => {
    if (typeof v === 'number') f(v);
    else if (isError(v)) throw v;
  };
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof Ref) ctx.host.each(a, fromCells, how.skipSubtotals);
    else if (a instanceof Matrix) a.data.forEach(fromCells);
    else if (a instanceof Lambda) throw VALUE;
    else if (how.typed === 'count') {
      if (typeof a === 'number' || typeof a === 'boolean') f(Number(a));
      else if (typeof a === 'string') {
        try {
          f(num(a));
        } catch {
          // Text that isn't a number isn't counted.
        }
      }
    } else if (isError(a)) throw a;
    else f(num(a));
  }
}

function list(args: Operand[], ctx: Ctx, how?: Collect): number[] {
  const out: number[] = [];
  numbers(args, ctx, (n) => out.push(n), how);
  return out;
}

const sum = (xs: number[]) => {
  let s = 0;
  for (const x of xs) s += x;
  return s;
};

function variance(xs: number[], sample: boolean): number {
  const n = xs.length;
  if (n < (sample ? 2 : 1)) throw DIV0;
  const mean = sum(xs) / n;
  let ss = 0;
  for (const x of xs) ss += (x - mean) ** 2;
  return ss / (sample ? n - 1 : n);
}

function nth(xs: number[], k: number, largest: boolean): number {
  const i = Math.ceil(k) - 1;
  if (i < 0 || i >= xs.length) throw NUM;
  const sorted = [...xs].sort((a, b) => (largest ? b - a : a - b));
  return sorted[i];
}

function median(xs: number[]): number {
  if (!xs.length) throw NUM;
  const s = [...xs].sort((a, b) => a - b);
  const m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

const AGGREGATES: Record<number, (xs: number[], counted: Operand[], ctx: Ctx) => number> = {
  1: (xs) => (xs.length ? sum(xs) / xs.length : (() => { throw DIV0; })()),
  2: (xs) => xs.length,
  3: (_xs, args, ctx) => counta(args, ctx, true),
  4: (xs) => (xs.length ? Math.max(...xs) : 0),
  5: (xs) => (xs.length ? Math.min(...xs) : 0),
  6: (xs) => xs.reduce((p, x) => p * x, 1),
  7: (xs) => Math.sqrt(variance(xs, true)),
  8: (xs) => Math.sqrt(variance(xs, false)),
  9: sum,
  10: (xs) => variance(xs, true),
  11: (xs) => variance(xs, false),
};

function counta(args: Operand[], ctx: Ctx, skipSubtotals = false): number {
  let n = 0;
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof Ref) ctx.host.each(a, (v) => void (v !== null && n++), skipSubtotals);
    else if (a instanceof Matrix) for (const v of a.data) n += v === null ? 0 : 1;
    else n++;
  }
  return n;
}

// -----------------------------------------------------------------------------
// Criteria, as COUNTIF and SUMIFS take them

/** Whether a value meets a criterion such as 5, ">=10", "<>", "North" or "N*". */
function criterion(c: Scalar): (v: Scalar) => boolean {
  if (isError(c)) return (v) => isError(v) && v.error === c.error;
  if (typeof c === 'number') return (v) => (typeof v === 'number' ? round15(v) === round15(c) : typeof v === 'string' && v.trim() !== '' && Number(v) === c);
  if (typeof c === 'boolean') return (v) => v === c;
  if (c === null) return (v) => v === null || v === '';
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c)!;
  const op = m[1] ?? '';
  const rest = m[2];
  const n = rest.trim() === '' ? undefined : Number(rest);
  const asNumber = n !== undefined && Number.isFinite(n);
  const upper = rest.toUpperCase();
  const asBool = upper === 'TRUE' ? true : upper === 'FALSE' ? false : undefined;
  const equal = (v: Scalar): boolean => {
    if (rest === '') return op === '=' ? v === null : v === null || v === '';
    if (asNumber) return typeof v === 'number' && round15(v) === round15(n!);
    if (asBool !== undefined) return v === asBool;
    if (typeof v !== 'string') return false;
    return hasWildcards(rest) ? matchesWhole(parsePattern(rest.toLowerCase()), v.toLowerCase()) : v.toLowerCase() === rest.toLowerCase();
  };
  if (op === '' || op === '=') return equal;
  if (op === '<>') return rest === '' ? (v) => v !== null && v !== '' : (v) => !equal(v);
  const test = op === '<' ? (d: number) => d < 0 : op === '>' ? (d: number) => d > 0 : op === '<=' ? (d: number) => d <= 0 : (d: number) => d >= 0;
  if (asNumber) return (v) => typeof v === 'number' && test(compare(v, n!));
  return (v) => typeof v === 'string' && test(compare(v, rest));
}

/** The cells of a range (or the values of an array) by offset, the range cut to the sheet's used part. */
function grid(op: Operand, ctx: Ctx): { rows: number; cols: number; at: (r: number, c: number) => Scalar } {
  if (op instanceof Ref) {
    const ref = ctx.host.bound(op);
    return { rows: ref.rows, cols: ref.cols, at: (r, c) => ctx.host.value(ref.sheet, ref.r1 + r, ref.c1 + c) };
  }
  if (op instanceof Matrix) return { rows: op.rows, cols: op.cols, at: (r, c) => op.get(r, c) };
  const v = scalarOf(op, ctx);
  return { rows: 1, cols: 1, at: () => v };
}

/** The offsets of the cells meeting every (range, criterion) pair, for COUNTIFS and its kind; ranges must be the same size. */
function matching(pairs: Operand[], ctx: Ctx): { hits: [number, number][]; rows: number; cols: number } {
  const ranges = [];
  for (let i = 0; i < pairs.length; i += 2) ranges.push({ g: grid(pairs[i], ctx), test: criterion(scalarOf(pairs[i + 1], ctx)) });
  const { rows, cols } = ranges[0].g;
  if (ranges.some((r) => r.g.rows !== rows || r.g.cols !== cols)) throw VALUE;
  const hits: [number, number][] = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) if (ranges.every((x) => x.test(x.g.at(r, c)))) hits.push([r, c]);
  }
  return { hits, rows, cols };
}

function ifs(kind: 'sum' | 'avg' | 'min' | 'max', target: Operand, pairs: Operand[], ctx: Ctx): number {
  const t = grid(target, ctx);
  const { hits, rows, cols } = matching(pairs, ctx);
  if (t.rows !== rows || t.cols !== cols) throw VALUE;
  const xs: number[] = [];
  for (const [r, c] of hits) {
    const v = t.at(r, c);
    if (typeof v === 'number') xs.push(v);
    else if (isError(v)) throw v;
  }
  if (kind === 'sum') return sum(xs);
  if (kind === 'avg') {
    if (!xs.length) throw DIV0;
    return sum(xs) / xs.length;
  }
  if (!xs.length) return 0;
  return kind === 'min' ? Math.min(...xs) : Math.max(...xs);
}

/**
 * Runs COUNTIF and its kind once per criterion when criteria come as an
 * array (=SUMIFS(C:C,A:A,E2#)), giving an array of results, as Excel does.
 */
function perCriterion(args: Operand[], ctx: Ctx, first: number, f: (args: Operand[]) => number): Operand {
  const at: number[] = [];
  for (let i = first; i < args.length; i += 2) at.push(i);
  const arrays = at.map((i) => {
    const a = args[i];
    return a instanceof Matrix ? a : a instanceof Ref && !a.single ? ctx.host.read(a) : undefined;
  });
  if (!arrays.some(Boolean)) return f(args);
  const rows = Math.max(...arrays.map((m) => m?.rows ?? 1));
  const cols = Math.max(...arrays.map((m) => m?.cols ?? 1));
  return Matrix.build(rows, cols, (r, c) => {
    const one = [...args];
    at.forEach((i, k) => {
      const m = arrays[k];
      if (m) one[i] = m.at(r, c);
    });
    try {
      return f(one);
    } catch (e) {
      if (isError(e)) return e;
      throw e;
    }
  });
}

/** SUMIF and AVERAGEIF: the target range is the criterion range's size, from its first cell. */
function ifOne(kind: 'sum' | 'avg', args: Operand[], ctx: Ctx): number {
  const range = args[0];
  let target = args[2] ?? range;
  if (target instanceof Ref && range instanceof Ref) {
    const r = ctx.host.bound(range);
    target = new Ref(target.sheet, target.r1, target.c1, target.r1 + r.rows - 1, target.c1 + r.cols - 1);
    return ifs(kind, target, [r, args[1]], ctx);
  }
  return ifs(kind, target, [range, args[1]], ctx);
}

// -----------------------------------------------------------------------------
// Lookups

function matrixOf(op: Operand, ctx: Ctx): Matrix {
  if (op instanceof Ref) return ctx.host.read(ctx.host.bound(op));
  if (op instanceof Matrix) return op;
  return new Matrix(1, 1, [scalarOf(op, ctx)]);
}

/** Whether a value equals the one looked for, wildcards (* ? ~) allowed in text. */
function exactly(target: Scalar, wildcards: boolean): (v: Scalar) => boolean {
  if (typeof target === 'string' && wildcards && hasWildcards(target)) {
    const p = parsePattern(target.toLowerCase());
    return (v) => typeof v === 'string' && matchesWhole(p, v.toLowerCase());
  }
  return (v) => v !== null && !isError(v) && !isError(target) && typeof v === typeof target && compare(v, target) === 0;
}

/**
 * The position of the largest value not above the one looked for (or, with
 * `descending`, the smallest not below), by Excel's binary search over a
 * list it takes to be sorted: values of another type are passed over.
 */
function approximate(values: Scalar[], target: Scalar, descending = false): number {
  if (isError(target)) throw target;
  let lo = 0;
  let hi = values.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let m = mid;
    while (m <= hi && (values[m] === null || typeof values[m] !== typeof target || isError(values[m]))) m++;
    if (m > hi) {
      hi = mid - 1;
      continue;
    }
    const d = compare(values[m], target);
    if (descending ? d >= 0 : d <= 0) {
      found = m;
      lo = m + 1;
    } else hi = mid - 1;
  }
  return found;
}

function lookupIn(values: Scalar[], target: Scalar, mode: number, last = false): number {
  if (mode === 0 || mode === 2) {
    const eq = exactly(target, true);
    if (last) {
      for (let i = values.length - 1; i >= 0; i--) if (eq(values[i])) return i;
    } else for (let i = 0; i < values.length; i++) if (eq(values[i])) return i;
    return -1;
  }
  // XLOOKUP's -1 and 1: an exact match, else the next smaller or larger.
  let best = -1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (v === null || isError(v) || typeof v !== typeof target) continue;
    const d = compare(v, target);
    if (d === 0) return i;
    if (mode < 0 ? d < 0 : d > 0) {
      if (best < 0 || (mode < 0 ? compare(v, values[best]) > 0 : compare(v, values[best]) < 0)) best = i;
    }
  }
  return best;
}

function vlookup(args: Operand[], ctx: Ctx, horizontal: boolean): Scalar {
  const target = scalarOf(args[0], ctx);
  if (isError(target)) throw target;
  const table = matrixOf(args[1], ctx);
  const index = Math.trunc(num(scalarOf(args[2], ctx)));
  const approx = args[3] === undefined ? true : bool(scalarOf(args[3], ctx));
  const lines = horizontal ? table.rows : table.cols;
  if (index < 1) throw VALUE;
  if (index > lines) throw REF;
  const keys = horizontal ? Array.from({ length: table.cols }, (_, c) => table.get(0, c)) : Array.from({ length: table.rows }, (_, r) => table.get(r, 0));
  const at = approx ? approximate(keys, target) : lookupIn(keys, target, 0);
  if (at < 0) throw NA;
  return horizontal ? table.get(index - 1, at) : table.get(at, index - 1);
}

/** The cells of a reference or values of an array as one list, row by row. */
function flat(op: Operand, ctx: Ctx): { values: Scalar[]; rows: number; cols: number } {
  const m = matrixOf(op, ctx);
  return { values: m.data, rows: m.rows, cols: m.cols };
}

function match(args: Operand[], ctx: Ctx): number {
  const target = scalarOf(args[0], ctx);
  const { values, rows, cols } = flat(args[1], ctx);
  if (rows > 1 && cols > 1) throw NA;
  const type = args[2] === undefined ? 1 : Math.sign(num(scalarOf(args[2], ctx)));
  const at = type === 0 ? lookupIn(values, target, 0) : approximate(values, target, type < 0);
  if (at < 0) throw NA;
  return at + 1;
}

/** Part of a reference (a reference still) or of an array, for INDEX and XLOOKUP. */
function slice(op: Operand, ctx: Ctx, r1: number, c1: number, r2: number, c2: number): Operand {
  if (op instanceof Ref) return new Ref(op.sheet, op.r1 + r1, op.c1 + c1, op.r1 + r2, op.c1 + c2);
  const m = matrixOf(op, ctx);
  if (r1 === r2 && c1 === c2) return m.get(r1, c1);
  return Matrix.build(r2 - r1 + 1, c2 - c1 + 1, (r, c) => m.get(r1 + r, c1 + c));
}

function index(args: Operand[], ctx: Ctx): Operand {
  const area = args[0];
  const dims = area instanceof Ref ? { rows: area.rows, cols: area.cols } : matrixOf(area, ctx);
  let row = args[1] === undefined ? 0 : Math.trunc(num(scalarOf(args[1], ctx)));
  let col = args[2] === undefined ? 0 : Math.trunc(num(scalarOf(args[2], ctx)));
  // A single row or column takes its one number as the position along it.
  if (args[2] === undefined && dims.rows === 1 && dims.cols > 1) [row, col] = [1, row];
  if (row < 0 || col < 0 || row > dims.rows || col > dims.cols) throw REF;
  if (dims.cols === 1 && col === 0 && row > 0) col = 1;
  const r1 = row ? row - 1 : 0;
  const r2 = row ? row - 1 : dims.rows - 1;
  const c1 = col ? col - 1 : 0;
  const c2 = col ? col - 1 : dims.cols - 1;
  return slice(area, ctx, r1, c1, r2, c2);
}

function xlookup(args: Operand[], ctx: Ctx): Operand {
  const target = scalarOf(args[0], ctx);
  const look = flat(args[1], ctx);
  if (look.rows > 1 && look.cols > 1) throw VALUE;
  const ret = args[2] instanceof Ref ? { rows: args[2].rows, cols: args[2].cols } : matrixOf(args[2], ctx);
  const vertical = look.cols === 1 && look.rows > 1 ? true : look.rows === 1 && look.cols > 1 ? false : ret.rows >= ret.cols;
  if ((vertical ? ret.rows : ret.cols) !== look.values.length) throw VALUE;
  const mode = args[4] === undefined || args[4] === null ? 0 : num(scalarOf(args[4], ctx));
  const search = args[5] === undefined || args[5] === null ? 1 : num(scalarOf(args[5], ctx));
  let at: number;
  if (Math.abs(search) === 2) at = approximate(look.values, target, search < 0);
  else at = lookupIn(look.values, target, mode, search < 0);
  if (at < 0) {
    if (args[3] !== undefined) return valuesOf(args[3], ctx);
    throw NA;
  }
  return vertical ? slice(args[2], ctx, at, 0, at, ret.cols - 1) : slice(args[2], ctx, 0, at, ret.rows - 1, at);
}

/** A reference written as text, for INDIRECT: A1 or R1C1, with a sheet or not, or a defined name. */
function indirect(args: Operand[], ctx: Ctx): Operand {
  const t = text(scalarOf(args[0], ctx)).trim();
  const a1 = args[1] === undefined ? true : bool(scalarOf(args[1], ctx));
  const parsed = a1 ? parseA1Ref(t, { allowTernary: false }) : parseR1C1Ref(t);
  if (parsed && 'range' in parsed && parsed.range) {
    const context = parsed.context ?? [];
    const sheet = context.length ? ctx.host.sheetIndex(context[0]) : ctx.sheet;
    if (sheet < 0) throw REF;
    if (a1) {
      const r = parsed.range as { top: number | null; left: number | null; bottom?: number | null; right?: number | null };
      const top = r.top ?? 0;
      const left = r.left ?? 0;
      const bottom = r.top === null ? MAX_ROWS - 1 : (r.bottom ?? r.top);
      const right = r.left === null ? MAX_COLS - 1 : (r.right ?? r.left);
      return new Ref(sheet, Math.min(top, bottom), Math.min(left, right), Math.max(top, bottom), Math.max(left, right));
    }
    const r = parsed.range as { r0: number | null; c0: number | null; r1?: number | null; c1?: number | null; $r0?: boolean; $c0?: boolean; $r1?: boolean; $c1?: boolean };
    const at = (v: number | null | undefined, abs: boolean | undefined, here: number, all: number) => (v === null || v === undefined ? all : abs ? v : here + v);
    const r1 = at(r.r0, r.$r0, ctx.row, 0);
    const c1 = at(r.c0, r.$c0, ctx.col, 0);
    const r2 = at(r.r1 ?? r.r0, r.$r1 ?? r.$r0, ctx.row, MAX_ROWS - 1);
    const c2 = at(r.c1 ?? r.c0, r.$c1 ?? r.$c0, ctx.col, MAX_COLS - 1);
    if (Math.min(r1, r2, c1, c2) < 0) throw REF;
    return new Ref(sheet, Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2));
  }
  const named = ctx.host.name(t.toUpperCase(), ctx);
  if (named instanceof Ref) return named;
  throw REF;
}

function offset(args: Operand[], ctx: Ctx): Operand {
  const base = args[0];
  if (!(base instanceof Ref)) throw VALUE;
  const rows = Math.trunc(num(scalarOf(args[1], ctx)));
  const cols = Math.trunc(num(scalarOf(args[2], ctx)));
  const height = args[3] === undefined || args[3] === null ? base.rows : Math.trunc(num(scalarOf(args[3], ctx)));
  const width = args[4] === undefined || args[4] === null ? base.cols : Math.trunc(num(scalarOf(args[4], ctx)));
  if (height < 1 || width < 1) throw REF;
  const r1 = base.r1 + rows;
  const c1 = base.c1 + cols;
  if (r1 < 0 || c1 < 0 || r1 + height > MAX_ROWS || c1 + width > MAX_COLS) throw REF;
  return new Ref(base.sheet, r1, c1, r1 + height - 1, c1 + width - 1);
}

// -----------------------------------------------------------------------------
// Dynamic arrays

function sized(rows: number, cols: number): void {
  if (rows < 1 || cols < 1) throw VALUE;
  if (rows * cols > MAX_ARRAY_CELLS) throw NUM;
}

function sortRows(m: Matrix, keys: { index: number; order: number }[], byCol: boolean): Matrix {
  const lines = byCol ? m.cols : m.rows;
  const at = (line: number, k: number) => (byCol ? m.get(k, line) : m.get(line, k));
  const order = Array.from({ length: lines }, (_, i) => i).sort((a, b) => {
    for (const { index, order } of keys) {
      const x = at(a, index);
      const y = at(b, index);
      // Blanks and errors sort last, either way.
      const xb = x === null || isError(x);
      const yb = y === null || isError(y);
      if (xb || yb) {
        if (xb !== yb) return xb ? 1 : -1;
        continue;
      }
      const d = compare(x, y);
      if (d) return d * order;
    }
    return a - b;
  });
  return byCol ? Matrix.build(m.rows, m.cols, (r, c) => m.get(r, order[c])) : Matrix.build(m.rows, m.cols, (r, c) => m.get(order[r], c));
}

function unique(m: Matrix, byCol: boolean, once: boolean): Matrix {
  const lines = byCol ? m.cols : m.rows;
  const width = byCol ? m.rows : m.cols;
  const key = (line: number) => Array.from({ length: width }, (_, k) => {
    const v = byCol ? m.get(k, line) : m.get(line, k);
    return typeof v === 'string' ? `s${v.toLowerCase()}` : isError(v) ? `e${v.error}` : `${typeof v}${String(v)}`;
  }).join('\u0000');
  const counts = new Map<string, number>();
  const firsts: number[] = [];
  for (let i = 0; i < lines; i++) {
    const k = key(i);
    if (!counts.has(k)) firsts.push(i);
    counts.set(k, (counts.get(k) ?? 0) + 1);
  }
  const keep = once ? firsts.filter((i) => counts.get(key(i)) === 1) : firsts;
  if (!keep.length) throw err('#CALC!');
  return byCol ? Matrix.build(m.rows, keep.length, (r, c) => m.get(r, keep[c])) : Matrix.build(keep.length, m.cols, (r, c) => m.get(keep[r], c));
}

// -----------------------------------------------------------------------------
// Dates: serial numbers in the 1900 system (with its 29 February 1900) or the 1904 one

const DAY_MS = 86_400_000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

/** The serial number of a date; months and days past their end roll over, as DATE's do. */
export function serialOf(y: number, m: number, d: number, date1904: boolean): number {
  const date = new Date(0);
  date.setUTCFullYear(y, m - 1, d);
  date.setUTCHours(0, 0, 0, 0);
  const t = date.getTime();
  if (date1904) return Math.round((t - EPOCH_1904) / DAY_MS);
  const days = Math.round((t - EPOCH_1900) / DAY_MS);
  // Before 1 March 1900, Excel's serials don't count its 29 February.
  return days < 61 ? days - 1 : days;
}

/** The year, month and day of a serial number. */
export function dateOf(serial: number, date1904: boolean): { y: number; m: number; d: number } {
  let s = Math.floor(serial);
  if (!date1904) {
    if (s === 60) return { y: 1900, m: 2, d: 29 };
    if (s < 60) s += 1;
  }
  const t = new Date((date1904 ? EPOCH_1904 : EPOCH_1900) + s * DAY_MS);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}

function dateArg(v: Scalar, ctx: Ctx): number {
  const n = typeof v === 'string' ? dateValue(v, ctx) : num(v);
  if (n < 0) throw NUM;
  return n;
}

/** A date written as text, ISO (2026-10-09) or month first (10/9/2026), as a serial number. */
function dateValue(t: string, ctx: Ctx): number {
  const s = t.trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (m) return serialOf(+m[1], +m[2], +m[3], ctx.host.date1904);
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
  if (m) {
    const y = +m[3] < 100 ? (+m[3] < 30 ? 2000 : 1900) + +m[3] : +m[3];
    return serialOf(y, +m[1], +m[2], ctx.host.date1904);
  }
  const n = Number(s);
  if (s && Number.isFinite(n)) return n;
  throw VALUE;
}

function eomonth(start: number, months: number, ctx: Ctx, end: boolean): number {
  const { y, m, d } = dateOf(start, ctx.host.date1904);
  const target = m + Math.trunc(months);
  if (end) return serialOf(y, target + 1, 0, ctx.host.date1904);
  const last = dateOf(serialOf(y, target + 1, 0, ctx.host.date1904), ctx.host.date1904).d;
  return serialOf(y, target, Math.min(d, last), ctx.host.date1904);
}

function weekday(serial: number, type: number, ctx: Ctx): number {
  const { y, m, d } = dateOf(serial, ctx.host.date1904);
  // 0 is Sunday; the 1900 system's 29 February 1900 and the days before it are a day out, as in Excel.
  let day = new Date(Date.UTC(y, m - 1, d)).getUTCDay();
  if (!ctx.host.date1904 && serial < 61) day = (Math.floor(serial) + 6) % 7;
  switch (type) {
    case 1:
    case 17:
      return day + 1;
    case 2:
    case 11:
      return ((day + 6) % 7) + 1;
    case 3:
      return (day + 6) % 7;
    case 12:
    case 13:
    case 14:
    case 15:
    case 16:
      return ((day - (type - 10) + 7) % 7) + 1;
    default:
      throw NUM;
  }
}

function workdays(start: number, end: number, holidays: Set<number>, ctx: Ctx): number {
  const step = end >= start ? 1 : -1;
  let n = 0;
  for (let s = Math.floor(start); step > 0 ? s <= end : s >= end; s += step) {
    const w = weekday(s, 2, ctx);
    if (w <= 5 && !holidays.has(s)) n += step;
  }
  return n;
}

function holidaysOf(op: Operand | undefined, ctx: Ctx): Set<number> {
  const out = new Set<number>();
  if (op !== undefined) numbers([op], ctx, (n) => out.add(Math.floor(n)));
  return out;
}

// -----------------------------------------------------------------------------
// Text

const TRIM = /^ +| +$|( ) +/g;

function proper(s: string): string {
  return s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase());
}

function textFormat(value: Scalar, pattern: string, ctx: Ctx): string {
  const v = typeof value === 'number' && ctx.host.date1904 ? value + 1462 : value;
  try {
    return format(pattern, v as number | string | boolean, { leap1900: true, throws: true, nbsp: false });
  } catch {
    throw VALUE;
  }
}

// -----------------------------------------------------------------------------
// The functions

/** IF and IFERROR and their kind, run on each value when their test is an array: the branch `test` picks, or the first value itself. */
function lazyEach(args: ((ctx: Ctx) => Operand)[], ctx: Ctx, test: (v: Scalar) => number | undefined): Operand {
  const first = valuesOf(args[0](ctx), ctx);
  // IF without a value for FALSE gives FALSE.
  const branch = (i: number): Scalar | Matrix => (args[i] ? valuesOf(args[i](ctx), ctx) : false);
  if (!(first instanceof Matrix)) {
    const pick = test(first);
    return pick === undefined ? first : branch(pick);
  }
  const cache = new Map<number, Scalar | Matrix>();
  const get = (i: number) => {
    if (!cache.has(i)) cache.set(i, branch(i));
    return cache.get(i)!;
  };
  return Matrix.build(first.rows, first.cols, (r, c) => {
    const v = first.get(r, c);
    const pick = test(v);
    if (pick === undefined) return v;
    const b = get(pick);
    return b instanceof Matrix ? b.at(r, c) : b;
  });
}

const ifTest = (v: Scalar) => {
  if (isError(v)) return undefined;
  return bool(v) ? 1 : 2;
};

export const FUNCTIONS: Record<string, FunctionSpec> = {
  // Logic
  IF: { min: 2, max: 3, lazy: (a, ctx) => lazyEach(a, ctx, ifTest) },
  IFERROR: { min: 2, max: 2, lazy: (a, ctx) => lazyEach(a, ctx, (v) => (isError(v) ? 1 : undefined)) },
  IFNA: { min: 2, max: 2, lazy: (a, ctx) => lazyEach(a, ctx, (v) => (isError(v) && v.error === '#N/A' ? 1 : undefined)) },
  IFS: {
    min: 2,
    max: MANY,
    lazy: (a, ctx) => {
      for (let i = 0; i + 1 < a.length; i += 2) {
        const v = scalarOf(a[i](ctx), ctx);
        if (isError(v)) return v;
        if (bool(v)) return valuesOf(a[i + 1](ctx), ctx);
      }
      return NA;
    },
  },
  SWITCH: {
    min: 3,
    max: MANY,
    lazy: (a, ctx) => {
      const v = scalarOf(a[0](ctx), ctx);
      if (isError(v)) return v;
      let i = 1;
      for (; i + 1 < a.length; i += 2) {
        const c = scalarOf(a[i](ctx), ctx);
        if (!isError(c) && v !== null && c !== null && typeof v === typeof c && compare(v, c) === 0) return valuesOf(a[i + 1](ctx), ctx);
      }
      return i < a.length ? valuesOf(a[i](ctx), ctx) : NA;
    },
  },
  CHOOSE: {
    min: 2,
    max: MANY,
    lazy: (a, ctx) => {
      const i = Math.trunc(num(scalarOf(a[0](ctx), ctx)));
      if (i < 1 || i >= a.length) throw VALUE;
      return a[i](ctx);
    },
  },
  AND: { min: 1, max: MANY, fn: (a, ctx) => logical(a, ctx, (xs) => xs.every(Boolean)) },
  OR: { min: 1, max: MANY, fn: (a, ctx) => logical(a, ctx, (xs) => xs.some(Boolean)) },
  XOR: { min: 1, max: MANY, fn: (a, ctx) => logical(a, ctx, (xs) => xs.filter(Boolean).length % 2 === 1) },
  NOT: scalar(1, 1, ([x]) => !bool(x)),
  TRUE: { min: 0, max: 0, fn: () => true },
  FALSE: { min: 0, max: 0, fn: () => false },

  // Arithmetic and rounding
  SUM: { min: 1, max: MANY, fn: (a, ctx) => sum(list(a, ctx)) },
  PRODUCT: { min: 1, max: MANY, fn: (a, ctx) => list(a, ctx).reduce((p, x) => p * x, 1) },
  SUMSQ: { min: 1, max: MANY, fn: (a, ctx) => sum(list(a, ctx).map((x) => x * x)) },
  ABS: math1(Math.abs),
  INT: math1(Math.floor),
  SIGN: math1(Math.sign),
  SQRT: math1((x) => (x < 0 ? NUM : Math.sqrt(x))),
  EXP: math1(Math.exp),
  LN: math1((x) => (x <= 0 ? NUM : Math.log(x))),
  LOG10: math1((x) => (x <= 0 ? NUM : Math.log10(x))),
  LOG: scalar(1, 2, ([x, b]) => {
    const n = num(x);
    const base = optional(b, 10);
    if (n <= 0 || base <= 0 || base === 1) throw base === 1 ? DIV0 : NUM;
    return Math.log(n) / Math.log(base);
  }),
  POWER: scalar(2, 2, ([x, y]) => {
    const r = Math.pow(num(x), num(y));
    if (num(x) === 0 && num(y) < 0) throw DIV0;
    return Number.isFinite(r) ? r : NUM;
  }),
  PI: { min: 0, max: 0, fn: () => Math.PI },
  MOD: scalar(2, 2, ([n, d]) => {
    const a = num(n);
    const b = num(d);
    if (b === 0) throw DIV0;
    return Number((a - b * Math.floor(a / b)).toPrecision(15));
  }),
  QUOTIENT: scalar(2, 2, ([n, d]) => {
    if (num(d) === 0) throw DIV0;
    return Math.trunc(num(n) / num(d));
  }),
  ROUND: scalar(2, 2, ([x, d]) => excelRound(num(x), num(d), 'half')),
  ROUNDUP: scalar(2, 2, ([x, d]) => excelRound(num(x), num(d), 'up')),
  ROUNDDOWN: scalar(2, 2, ([x, d]) => excelRound(num(x), num(d), 'down')),
  TRUNC: scalar(1, 2, ([x, d]) => excelRound(num(x), optional(d, 0), 'down')),
  MROUND: scalar(2, 2, ([x, m]) => {
    const n = num(x);
    const k = num(m);
    if (k === 0) return 0;
    if (Math.sign(n) * Math.sign(k) < 0) throw NUM;
    return excelRound(Math.round(n / k) * k, 10, 'half');
  }),
  CEILING: scalar(2, 2, ([x, s]) => {
    const n = num(x);
    const k = num(s);
    if (k === 0) return 0;
    if (n > 0 && k < 0) throw NUM;
    return excelRound(Math.ceil(n / k) * k, 10, 'half');
  }),
  FLOOR: scalar(2, 2, ([x, s]) => {
    const n = num(x);
    const k = num(s);
    if (k === 0) throw DIV0;
    if (n > 0 && k < 0) throw NUM;
    return excelRound(Math.floor(n / k) * k, 10, 'half');
  }),
  'CEILING.MATH': scalar(1, 3, ([x, s, mode]) => {
    const n = num(x);
    const k = Math.abs(optional(s, 1));
    if (k === 0) return 0;
    const away = n < 0 && optional(mode, 0) !== 0;
    return excelRound((away ? Math.floor(n / k) : Math.ceil(n / k)) * k, 10, 'half');
  }),
  'FLOOR.MATH': scalar(1, 3, ([x, s, mode]) => {
    const n = num(x);
    const k = Math.abs(optional(s, 1));
    if (k === 0) return 0;
    const toward = n < 0 && optional(mode, 0) !== 0;
    return excelRound((toward ? Math.ceil(n / k) : Math.floor(n / k)) * k, 10, 'half');
  }),
  EVEN: math1((x) => {
    const r = Math.ceil(Math.abs(x) / 2) * 2;
    return x < 0 ? -r : r;
  }),
  ODD: math1((x) => {
    let r = Math.ceil(Math.abs(x));
    if (r % 2 === 0) r += 1;
    return x < 0 ? -r : r;
  }),
  RAND: { min: 0, max: 0, volatile: true, fn: () => Math.random() },
  RANDBETWEEN: {
    ...scalar(2, 2, ([a, b]) => {
      const lo = Math.ceil(num(a));
      const hi = Math.floor(num(b));
      if (lo > hi) throw NUM;
      return lo + Math.floor(Math.random() * (hi - lo + 1));
    }),
    volatile: true,
  },
  SUMPRODUCT: {
    min: 1,
    max: MANY,
    fn: (a, ctx) => {
      const ms = a.map((x) => matrixOf(x, ctx));
      const { rows, cols } = ms[0];
      if (ms.some((m) => m.rows !== rows || m.cols !== cols)) throw VALUE;
      let total = 0;
      for (let i = 0; i < rows * cols; i++) {
        let p = 1;
        for (const m of ms) {
          const v = m.data[i];
          if (isError(v)) throw v;
          p *= typeof v === 'number' ? v : 0;
        }
        total += p;
      }
      return total;
    },
  },

  // Statistics
  COUNT: { min: 1, max: MANY, fn: (a, ctx) => list(a, ctx, { typed: 'count' }).length },
  COUNTA: { min: 1, max: MANY, fn: (a, ctx) => counta(a, ctx) },
  COUNTBLANK: {
    min: 1,
    max: 1,
    fn: ([r], ctx) => {
      if (!(r instanceof Ref)) throw VALUE;
      let n = 0;
      for (let row = r.r1; row <= Math.min(r.r2, r.r1 + MAX_ARRAY_CELLS); row++) for (let col = r.c1; col <= r.c2; col++) {
        const v = ctx.host.value(r.sheet, row, col);
        if (v === null || v === '') n++;
      }
      return n;
    },
  },
  AVERAGE: {
    min: 1,
    max: MANY,
    fn: (a, ctx) => {
      const xs = list(a, ctx);
      if (!xs.length) throw DIV0;
      return sum(xs) / xs.length;
    },
  },
  MAX: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = list(a, ctx);
    return xs.length ? xs.reduce((m, x) => (x > m ? x : m), -Infinity) : 0;
  } },
  MIN: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = list(a, ctx);
    return xs.length ? xs.reduce((m, x) => (x < m ? x : m), Infinity) : 0;
  } },
  MEDIAN: { min: 1, max: MANY, fn: (a, ctx) => median(list(a, ctx)) },
  LARGE: { min: 2, max: 2, fn: (a, ctx) => nth(list([a[0]], ctx), num(scalarOf(a[1], ctx)), true) },
  SMALL: { min: 2, max: 2, fn: (a, ctx) => nth(list([a[0]], ctx), num(scalarOf(a[1], ctx)), false) },
  STDEV: { min: 1, max: MANY, fn: (a, ctx) => Math.sqrt(variance(list(a, ctx), true)) },
  'STDEV.S': { min: 1, max: MANY, fn: (a, ctx) => Math.sqrt(variance(list(a, ctx), true)) },
  STDEVP: { min: 1, max: MANY, fn: (a, ctx) => Math.sqrt(variance(list(a, ctx), false)) },
  'STDEV.P': { min: 1, max: MANY, fn: (a, ctx) => Math.sqrt(variance(list(a, ctx), false)) },
  VAR: { min: 1, max: MANY, fn: (a, ctx) => variance(list(a, ctx), true) },
  'VAR.S': { min: 1, max: MANY, fn: (a, ctx) => variance(list(a, ctx), true) },
  VARP: { min: 1, max: MANY, fn: (a, ctx) => variance(list(a, ctx), false) },
  'VAR.P': { min: 1, max: MANY, fn: (a, ctx) => variance(list(a, ctx), false) },
  RANK: { min: 2, max: 3, fn: (a, ctx) => rank(a, ctx) },
  'RANK.EQ': { min: 2, max: 3, fn: (a, ctx) => rank(a, ctx) },
  SUBTOTAL: {
    min: 2,
    max: MANY,
    fn: (a, ctx) => {
      const code = Math.trunc(num(scalarOf(a[0], ctx)));
      const f = AGGREGATES[code > 100 ? code - 100 : code];
      if (!f) throw VALUE;
      const rest = a.slice(1);
      return f(list(rest, ctx, { typed: 'coerce', skipSubtotals: true }), rest, ctx);
    },
  },
  COUNTIF: { min: 2, max: 2, fn: (a, ctx) => perCriterion(a, ctx, 1, (x) => matching([x[0], x[1]], ctx).hits.length) },
  COUNTIFS: {
    min: 2,
    max: MANY,
    fn: (a, ctx) => {
      if (a.length % 2) throw VALUE;
      return perCriterion(a, ctx, 1, (x) => matching(x, ctx).hits.length);
    },
  },
  SUMIF: { min: 2, max: 3, fn: (a, ctx) => perCriterion(a.slice(0, 2), ctx, 1, (x) => ifOne('sum', [...x, ...a.slice(2)], ctx)) },
  AVERAGEIF: { min: 2, max: 3, fn: (a, ctx) => perCriterion(a.slice(0, 2), ctx, 1, (x) => ifOne('avg', [...x, ...a.slice(2)], ctx)) },
  SUMIFS: { min: 3, max: MANY, fn: (a, ctx) => perCriterion(a, ctx, 2, (x) => ifs('sum', x[0], x.slice(1), ctx)) },
  AVERAGEIFS: { min: 3, max: MANY, fn: (a, ctx) => perCriterion(a, ctx, 2, (x) => ifs('avg', x[0], x.slice(1), ctx)) },
  MINIFS: { min: 3, max: MANY, fn: (a, ctx) => perCriterion(a, ctx, 2, (x) => ifs('min', x[0], x.slice(1), ctx)) },
  MAXIFS: { min: 3, max: MANY, fn: (a, ctx) => perCriterion(a, ctx, 2, (x) => ifs('max', x[0], x.slice(1), ctx)) },

  // Lookup and reference
  VLOOKUP: { min: 3, max: 4, fn: (a, ctx) => vlookup(a, ctx, false) },
  HLOOKUP: { min: 3, max: 4, fn: (a, ctx) => vlookup(a, ctx, true) },
  MATCH: { min: 2, max: 3, fn: (a, ctx) => match(a, ctx) },
  XMATCH: {
    min: 2,
    max: 4,
    fn: (a, ctx) => {
      const target = scalarOf(a[0], ctx);
      const { values } = flat(a[1], ctx);
      const mode = a[2] === undefined || a[2] === null ? 0 : num(scalarOf(a[2], ctx));
      const search = a[3] === undefined || a[3] === null ? 1 : num(scalarOf(a[3], ctx));
      const at = Math.abs(search) === 2 ? approximate(values, target, search < 0) : lookupIn(values, target, mode, search < 0);
      if (at < 0) throw NA;
      return at + 1;
    },
  },
  XLOOKUP: { min: 3, max: 6, fn: (a, ctx) => xlookup(a, ctx) },
  INDEX: { min: 2, max: 3, fn: (a, ctx) => index(a, ctx) },
  OFFSET: { min: 3, max: 5, volatile: true, fn: (a, ctx) => offset(a, ctx) },
  INDIRECT: { min: 1, max: 2, volatile: true, fn: (a, ctx) => indirect(a, ctx) },
  ROW: {
    min: 0,
    max: 1,
    fn: ([r], ctx) => {
      if (r === undefined) return ctx.row + 1;
      if (!(r instanceof Ref)) throw VALUE;
      return r.rows === 1 || ctx.legacy ? r.r1 + 1 : Matrix.build(r.rows, 1, (i) => r.r1 + i + 1);
    },
  },
  COLUMN: {
    min: 0,
    max: 1,
    fn: ([r], ctx) => {
      if (r === undefined) return ctx.col + 1;
      if (!(r instanceof Ref)) throw VALUE;
      return r.cols === 1 || ctx.legacy ? r.c1 + 1 : Matrix.build(1, r.cols, (_i, j) => r.c1 + j + 1);
    },
  },
  ROWS: { min: 1, max: 1, fn: ([r], ctx) => (r instanceof Ref ? r.rows : matrixOf(r, ctx).rows) },
  COLUMNS: { min: 1, max: 1, fn: ([r], ctx) => (r instanceof Ref ? r.cols : matrixOf(r, ctx).cols) },
  TRANSPOSE: {
    min: 1,
    max: 1,
    fn: ([a], ctx) => {
      const m = matrixOf(a, ctx);
      return Matrix.build(m.cols, m.rows, (r, c) => m.get(c, r));
    },
  },
  HYPERLINK: scalar(1, 2, ([link, friendly], ctx) => {
    const target = text(link);
    if (ctx.notes) ctx.notes.link = target;
    return friendly === undefined ? target : friendly;
  }),
  FILTER: {
    min: 2,
    max: 3,
    fn: (a, ctx) => {
      const m = matrixOf(a[0], ctx);
      const inc = matrixOf(a[1], ctx);
      const byRow = inc.cols === 1 && inc.rows === m.rows;
      const byCol = inc.rows === 1 && inc.cols === m.cols;
      if (!byRow && !byCol) throw VALUE;
      const keep: number[] = [];
      for (let i = 0; i < (byRow ? m.rows : m.cols); i++) {
        const v = byRow ? inc.get(i, 0) : inc.get(0, i);
        if (isError(v)) throw v;
        if (bool(v)) keep.push(i);
      }
      if (!keep.length) {
        if (a[2] !== undefined) return valuesOf(a[2], ctx);
        throw err('#CALC!');
      }
      return byRow ? Matrix.build(keep.length, m.cols, (r, c) => m.get(keep[r], c)) : Matrix.build(m.rows, keep.length, (r, c) => m.get(r, keep[c]));
    },
  },
  SORT: {
    min: 1,
    max: 4,
    fn: (a, ctx) => {
      const m = matrixOf(a[0], ctx);
      const by = a[3] === undefined || a[3] === null ? false : bool(scalarOf(a[3], ctx));
      const idx = a[1] === undefined || a[1] === null ? 1 : Math.trunc(num(scalarOf(a[1], ctx)));
      const order = a[2] === undefined || a[2] === null ? 1 : num(scalarOf(a[2], ctx));
      if (idx < 1 || idx > (by ? m.rows : m.cols) || (order !== 1 && order !== -1)) throw VALUE;
      return sortRows(m, [{ index: idx - 1, order }], by);
    },
  },
  SORTBY: {
    min: 2,
    max: MANY,
    fn: (a, ctx) => {
      const m = matrixOf(a[0], ctx);
      const keys: Matrix[] = [];
      const orders: number[] = [];
      for (let i = 1; i < a.length; i += 2) {
        keys.push(matrixOf(a[i], ctx));
        orders.push(a[i + 1] === undefined ? 1 : num(scalarOf(a[i + 1], ctx)));
      }
      const byCol = keys[0].rows === 1 && keys[0].cols === m.cols && m.rows !== keys[0].rows;
      const lines = byCol ? m.cols : m.rows;
      if (keys.some((k) => (byCol ? k.cols : k.rows) !== lines)) throw VALUE;
      const joined = byCol
        ? Matrix.build(keys.length, lines, (r, c) => keys[r].data[c])
        : Matrix.build(lines, keys.length, (r, c) => keys[c].data[r]);
      const order = sortRows(
        Matrix.build(byCol ? keys.length + 1 : lines, byCol ? lines : keys.length + 1, (r, c) => (byCol ? (r < keys.length ? joined.get(r, c) : c) : c < keys.length ? joined.get(r, c) : r)),
        keys.map((_k, i) => ({ index: i, order: orders[i] })),
        byCol,
      );
      const pos = (i: number) => (byCol ? (order.get(keys.length, i) as number) : (order.get(i, keys.length) as number));
      return byCol ? Matrix.build(m.rows, m.cols, (r, c) => m.get(r, pos(c))) : Matrix.build(m.rows, m.cols, (r, c) => m.get(pos(r), c));
    },
  },
  UNIQUE: {
    min: 1,
    max: 3,
    fn: (a, ctx) =>
      unique(matrixOf(a[0], ctx), a[1] === undefined || a[1] === null ? false : bool(scalarOf(a[1], ctx)), a[2] === undefined || a[2] === null ? false : bool(scalarOf(a[2], ctx))),
  },
  SEQUENCE: {
    min: 1,
    max: 4,
    fn: (a, ctx) => {
      const rows = Math.trunc(num(scalarOf(a[0], ctx)));
      const cols = a[1] === undefined || a[1] === null ? 1 : Math.trunc(num(scalarOf(a[1], ctx)));
      const start = a[2] === undefined || a[2] === null ? 1 : num(scalarOf(a[2], ctx));
      const step = a[3] === undefined || a[3] === null ? 1 : num(scalarOf(a[3], ctx));
      sized(rows, cols);
      return Matrix.build(rows, cols, (r, c) => start + (r * cols + c) * step);
    },
  },
  'WBE.SPILL': {
    min: 3,
    max: 4,
    fn: (a, ctx) => {
      const rows = num(scalarOf(a[0], ctx));
      const cols = num(scalarOf(a[1], ctx));
      const fixed = a[3] !== undefined && num(scalarOf(a[3], ctx)) === 1;
      const inner = valuesOf(a[2], ctx);
      const m = inner instanceof Matrix ? inner : new Matrix(1, 1, [inner]);
      if (ctx.notes) ctx.notes.measured = { rows: m.rows, cols: m.cols };
      if (rows === 1 && cols === 1) return m.get(0, 0);
      // Cut or padded to the room it was given, until it is given room for its new size; Ctrl+Shift+Enter arrays repeat a single row or column and show #N/A past their result.
      return Matrix.build(rows, cols, (r, c) => (fixed ? m.at(r, c) : r < m.rows && c < m.cols ? m.get(r, c) : null));
    },
  },

  // Text
  CONCATENATE: scalar(1, MANY, (a) => a.map((v) => text(v)).join('')),
  CONCAT: {
    min: 1,
    max: MANY,
    fn: (a, ctx) => {
      let out = '';
      for (const x of a) {
        if (x instanceof Ref || x instanceof Matrix) for (const v of matrixOf(x, ctx).data) out += text(v);
        else out += text(scalarOf(x, ctx));
      }
      return out;
    },
  },
  TEXTJOIN: {
    min: 3,
    max: MANY,
    fn: (a, ctx) => {
      const delim = text(scalarOf(a[0], ctx));
      const skip = bool(scalarOf(a[1], ctx));
      const parts: string[] = [];
      for (const x of a.slice(2)) {
        const vs = x instanceof Ref || x instanceof Matrix ? matrixOf(x, ctx).data : [scalarOf(x, ctx)];
        for (const v of vs) {
          const t = text(v);
          if (!skip || t !== '') parts.push(t);
        }
      }
      return parts.join(delim);
    },
  },
  LEFT: scalar(1, 2, ([t, n]) => {
    const k = optional(n, 1);
    if (k < 0) throw VALUE;
    return text(t).slice(0, Math.trunc(k));
  }),
  RIGHT: scalar(1, 2, ([t, n]) => {
    const k = Math.trunc(optional(n, 1));
    if (k < 0) throw VALUE;
    const s = text(t);
    return k === 0 ? '' : s.slice(-k);
  }),
  MID: scalar(3, 3, ([t, s, n]) => {
    const start = Math.trunc(num(s));
    const k = Math.trunc(num(n));
    if (start < 1 || k < 0) throw VALUE;
    return text(t).substr(start - 1, k);
  }),
  LEN: scalar(1, 1, ([t]) => text(t).length),
  UPPER: scalar(1, 1, ([t]) => text(t).toUpperCase()),
  LOWER: scalar(1, 1, ([t]) => text(t).toLowerCase()),
  PROPER: scalar(1, 1, ([t]) => proper(text(t))),
  TRIM: scalar(1, 1, ([t]) => text(t).replace(TRIM, '$1')),
  // The first 32 characters, which don't print.
  CLEAN: scalar(1, 1, ([t]) => [...text(t)].filter((ch) => ch.charCodeAt(0) >= 32).join('')),
  REPT: scalar(2, 2, ([t, n]) => {
    const k = Math.trunc(num(n));
    if (k < 0 || text(t).length * k > 32767) throw VALUE;
    return text(t).repeat(k);
  }),
  EXACT: scalar(2, 2, ([a, b]) => text(a) === text(b)),
  SUBSTITUTE: scalar(3, 4, ([t, from, to, nth]) => {
    const s = text(t);
    const f = text(from);
    const r = text(to);
    if (!f) return s;
    if (nth === undefined || nth === null) return s.split(f).join(r);
    const k = Math.trunc(num(nth));
    if (k < 1) throw VALUE;
    let at = -1;
    for (let i = 0; i < k; i++) {
      at = s.indexOf(f, at + 1);
      if (at < 0) return s;
    }
    return s.slice(0, at) + r + s.slice(at + f.length);
  }),
  REPLACE: scalar(4, 4, ([t, s, n, r]) => {
    const start = Math.trunc(num(s));
    const k = Math.trunc(num(n));
    if (start < 1 || k < 0) throw VALUE;
    const str = text(t);
    return str.slice(0, start - 1) + text(r) + str.slice(start - 1 + k);
  }),
  FIND: scalar(2, 3, ([f, w, s]) => {
    const start = Math.trunc(optional(s, 1));
    const within = text(w);
    if (start < 1 || start > within.length + 1) throw VALUE;
    const at = within.indexOf(text(f), start - 1);
    if (at < 0) throw VALUE;
    return at + 1;
  }),
  SEARCH: scalar(2, 3, ([f, w, s]) => {
    const start = Math.trunc(optional(s, 1));
    const within = text(w);
    if (start < 1 || start > within.length + 1) throw VALUE;
    const at = searchIn(parsePattern(text(f).toLowerCase()), within.toLowerCase().slice(start - 1));
    if (at < 0) throw VALUE;
    return at + start;
  }),
  TEXT: scalar(2, 2, ([v, f], ctx) => textFormat(v, text(f), ctx)),
  VALUE: scalar(1, 1, ([t], ctx) => (typeof t === 'number' ? t : dateValueOrNumber(text(t), ctx))),
  CHAR: scalar(1, 1, ([n]) => {
    const k = Math.trunc(num(n));
    if (k < 1 || k > 255) throw VALUE;
    return String.fromCharCode(k);
  }),
  CODE: scalar(1, 1, ([t]) => {
    const s = text(t);
    if (!s) throw VALUE;
    return s.charCodeAt(0);
  }),
  UNICHAR: scalar(1, 1, ([n]) => String.fromCodePoint(Math.trunc(num(n)))),
  UNICODE: scalar(1, 1, ([t]) => {
    const s = text(t);
    if (!s) throw VALUE;
    return s.codePointAt(0)!;
  }),
  T: scalar(1, 1, ([v]) => (typeof v === 'string' ? v : isError(v) ? v : '')),
  N: scalar(1, 1, ([v]) => (typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : isError(v) ? v : 0)),

  // Dates and times
  DATE: scalar(3, 3, ([y, m, d], ctx) => {
    let year = Math.trunc(num(y));
    if (year < 0 || year > 9999) throw NUM;
    if (year < 1900) year += 1900;
    const s = serialOf(year, Math.trunc(num(m)), Math.trunc(num(d)), ctx.host.date1904);
    if (s < 0) throw NUM;
    return s;
  }),
  TIME: scalar(3, 3, ([h, m, s]) => {
    const t = (Math.trunc(num(h)) * 3600 + Math.trunc(num(m)) * 60 + Math.trunc(num(s))) / 86400;
    if (t < 0) throw NUM;
    return t - Math.floor(t);
  }),
  TODAY: { min: 0, max: 0, volatile: true, fn: (_a, ctx) => {
    const now = new Date();
    return serialOf(now.getFullYear(), now.getMonth() + 1, now.getDate(), ctx.host.date1904);
  } },
  NOW: { min: 0, max: 0, volatile: true, fn: (_a, ctx) => {
    const now = new Date();
    return serialOf(now.getFullYear(), now.getMonth() + 1, now.getDate(), ctx.host.date1904) + (now.getHours() * 3600 + now.getMinutes() * 60 + now.getSeconds()) / 86400;
  } },
  YEAR: scalar(1, 1, ([v], ctx) => dateOf(dateArg(v, ctx), ctx.host.date1904).y),
  MONTH: scalar(1, 1, ([v], ctx) => dateOf(dateArg(v, ctx), ctx.host.date1904).m),
  DAY: scalar(1, 1, ([v], ctx) => dateOf(dateArg(v, ctx), ctx.host.date1904).d),
  HOUR: scalar(1, 1, ([v], ctx) => Math.floor(round15((dateArg(v, ctx) % 1) * 24))),
  MINUTE: scalar(1, 1, ([v], ctx) => Math.floor(round15((dateArg(v, ctx) * 1440) % 60))),
  SECOND: scalar(1, 1, ([v], ctx) => Math.round((dateArg(v, ctx) * 86400) % 60) % 60),
  WEEKDAY: scalar(1, 2, ([v, t], ctx) => weekday(dateArg(v, ctx), optional(t, 1), ctx)),
  EOMONTH: scalar(2, 2, ([v, m], ctx) => eomonth(dateArg(v, ctx), num(m), ctx, true)),
  EDATE: scalar(2, 2, ([v, m], ctx) => eomonth(dateArg(v, ctx), num(m), ctx, false)),
  DAYS: scalar(2, 2, ([e, s], ctx) => Math.floor(dateArg(e, ctx)) - Math.floor(dateArg(s, ctx))),
  DATEVALUE: scalar(1, 1, ([t], ctx) => Math.floor(dateValue(text(t), ctx))),
  DATEDIF: scalar(3, 3, ([s, e, u], ctx) => datedif(dateArg(s, ctx), dateArg(e, ctx), text(u).toUpperCase(), ctx)),
  NETWORKDAYS: { min: 2, max: 3, fn: (a, ctx) => workdays(dateArg(scalarOf(a[0], ctx), ctx), dateArg(scalarOf(a[1], ctx), ctx), holidaysOf(a[2], ctx), ctx) },
  WORKDAY: {
    min: 2,
    max: 3,
    fn: (a, ctx) => {
      let s = Math.floor(dateArg(scalarOf(a[0], ctx), ctx));
      let left = Math.trunc(num(scalarOf(a[1], ctx)));
      const holidays = holidaysOf(a[2], ctx);
      const step = left >= 0 ? 1 : -1;
      while (left !== 0) {
        s += step;
        if (weekday(s, 2, ctx) <= 5 && !holidays.has(s)) left -= step;
      }
      return s;
    },
  },

  // Information
  ISBLANK: scalar(1, 1, ([v]) => v === null),
  ISNUMBER: scalar(1, 1, ([v]) => typeof v === 'number'),
  ISTEXT: scalar(1, 1, ([v]) => typeof v === 'string'),
  ISNONTEXT: scalar(1, 1, ([v]) => typeof v !== 'string'),
  ISLOGICAL: scalar(1, 1, ([v]) => typeof v === 'boolean'),
  ISERROR: scalar(1, 1, ([v]) => isError(v)),
  ISERR: scalar(1, 1, ([v]) => isError(v) && v.error !== '#N/A'),
  ISNA: scalar(1, 1, ([v]) => isError(v) && v.error === '#N/A'),
  ISEVEN: scalar(1, 1, ([v]) => Math.trunc(num(v)) % 2 === 0),
  ISODD: scalar(1, 1, ([v]) => Math.abs(Math.trunc(num(v)) % 2) === 1),
  ISREF: { min: 1, max: 1, fn: ([r]) => r instanceof Ref },
  ISFORMULA: { min: 1, max: 1, fn: ([r], ctx) => (r instanceof Ref ? ctx.host.isFormula(r.sheet, r.r1, r.c1) : (() => { throw VALUE; })()) },
  NA: { min: 0, max: 0, fn: () => NA },

  // How files write A2# and @A1:A3.
  ANCHORARRAY: { min: 1, max: 1, fn: ([r], ctx) => (r instanceof Ref && r.single ? (ctx.host.spillOf(r.sheet, r.r1, r.c1) ?? REF) : REF) },
  SINGLE: { min: 1, max: 1, fn: ([v], ctx) => scalarOf(v, ctx) },
};

/** A text's value as a number, the way VALUE reads it: numbers, percentages, dates. */
function dateValueOrNumber(t: string, ctx: Ctx): number {
  try {
    return num(t);
  } catch {
    return dateValue(t, ctx);
  }
}

function datedif(start: number, end: number, unit: string, ctx: Ctx): number {
  if (start > end) throw NUM;
  const a = dateOf(start, ctx.host.date1904);
  const b = dateOf(end, ctx.host.date1904);
  const months = (b.y - a.y) * 12 + (b.m - a.m) - (b.d < a.d ? 1 : 0);
  switch (unit) {
    case 'D':
      return Math.floor(end) - Math.floor(start);
    case 'M':
      return months;
    case 'Y':
      return Math.floor(months / 12);
    case 'MD': {
      if (b.d >= a.d) return b.d - a.d;
      const prevEnd = dateOf(serialOf(b.y, b.m, 0, ctx.host.date1904), ctx.host.date1904).d;
      return prevEnd - a.d + b.d < 0 ? 0 : prevEnd - a.d + b.d;
    }
    case 'YM':
      return months % 12;
    case 'YD': {
      let s = serialOf(b.y, a.m, a.d, ctx.host.date1904);
      if (s > Math.floor(end)) s = serialOf(b.y - 1, a.m, a.d, ctx.host.date1904);
      return Math.floor(end) - s;
    }
    default:
      throw NUM;
  }
}

function logical(args: Operand[], ctx: Ctx, f: (xs: boolean[]) => boolean): boolean {
  const xs: boolean[] = [];
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof Ref || a instanceof Matrix) {
      for (const v of matrixOf(a, ctx).data) {
        if (isError(v)) throw v;
        if (typeof v === 'number' || typeof v === 'boolean') xs.push(Boolean(v));
      }
    } else xs.push(bool(scalarOf(a, ctx)));
  }
  if (!xs.length) throw VALUE;
  return f(xs);
}

function rank(a: Operand[], ctx: Ctx): number {
  const x = num(scalarOf(a[0], ctx));
  const xs = list([a[1]], ctx);
  const asc = a[2] !== undefined && a[2] !== null && num(scalarOf(a[2], ctx)) !== 0;
  if (!xs.some((v) => round15(v) === round15(x))) throw NA;
  return 1 + xs.filter((v) => (asc ? v < x : v > x)).length;
}
