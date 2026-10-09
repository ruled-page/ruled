import { format } from 'numfmt';
import { parseA1Ref, parseR1C1Ref } from '@borgar/fx';
import { BondError, Bonds, Calendar, ddb, validBasis, vdb, yearFrac } from '../core/bondMath';
import { excelRound } from '../core/excelNumbers';
import { eliminate } from '../core/matrixMath';
import { estimate, prepare, project, RegressionError } from '../core/regressionMath';
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

// -----------------------------------------------------------------------------
// Arrays reshaped, text split and joined, money over time

const given = (a: Operand[], i: number) => a[i] !== undefined && a[i] !== null;
const optNum = (a: Operand[], i: number, ctx: Ctx, d: number) => (given(a, i) ? num(scalarOf(a[i], ctx)) : d);
const optBool = (a: Operand[], i: number, ctx: Ctx, d: boolean) => (given(a, i) ? bool(scalarOf(a[i], ctx)) : d);

/** TAKE and DROP: rows and columns from the start (or, counted negative, the end). */
function takeOrDrop(a: Operand[], ctx: Ctx, take: boolean): Matrix {
  const m = matrixOf(a[0], ctx);
  const span = (n: number, i: number): [number, number] => {
    if (!given(a, i)) return [0, n];
    const k = Math.trunc(num(scalarOf(a[i], ctx)));
    if (take) return k >= 0 ? [0, Math.min(k, n)] : [Math.max(0, n + k), n];
    return k >= 0 ? [Math.min(k, n), n] : [0, Math.max(0, n + k)];
  };
  const [r1, r2] = span(m.rows, 1);
  const [c1, c2] = span(m.cols, 2);
  if (r2 <= r1 || c2 <= c1) throw err('#CALC!');
  return Matrix.build(r2 - r1, c2 - c1, (r, c) => m.get(r1 + r, c1 + c));
}

/** CHOOSEROWS and CHOOSECOLS: lines by position, negative from the end. */
function chooseLines(a: Operand[], ctx: Ctx, rows: boolean): Matrix {
  const m = matrixOf(a[0], ctx);
  const n = rows ? m.rows : m.cols;
  const picks: number[] = [];
  for (const x of a.slice(1)) {
    const v = valuesOf(x, ctx);
    for (const k of v instanceof Matrix ? v.data : [v]) {
      const i = Math.trunc(num(k));
      if (i === 0 || Math.abs(i) > n) throw VALUE;
      picks.push(i > 0 ? i - 1 : n + i);
    }
  }
  return rows ? Matrix.build(picks.length, m.cols, (r, c) => m.get(picks[r], c)) : Matrix.build(m.rows, picks.length, (r, c) => m.get(r, picks[c]));
}

/** A reference's blank cells as 0, as array results show them. */
const filled = (m: Matrix) => (m.data.includes(null) ? new Matrix(m.rows, m.cols, m.data.map((v) => v ?? 0)) : m);

/** VSTACK and HSTACK: arrays one after another, short ones padded with #N/A. */
function stack(a: Operand[], ctx: Ctx, vertical: boolean): Matrix {
  const ms = a.map((x) => filled(matrixOf(x, ctx)));
  const across = Math.max(...ms.map((m) => (vertical ? m.cols : m.rows)));
  const lines: Scalar[][] = [];
  for (const m of ms) {
    const count = vertical ? m.rows : m.cols;
    for (let i = 0; i < count; i++) lines.push(Array.from({ length: across }, (_, j) => (vertical ? (j < m.cols ? m.get(i, j) : NA) : j < m.rows ? m.get(j, i) : NA)));
  }
  return vertical ? Matrix.of(lines) : Matrix.build(across, lines.length, (r, c) => lines[c][r]);
}

/** TOCOL and TOROW: an array's values in one line, by row or by column, blanks or errors left out if asked. */
function toLine(a: Operand[], ctx: Ctx, column: boolean): Matrix {
  const m = matrixOf(a[0], ctx);
  const ignore = optNum(a, 1, ctx, 0);
  const byCol = optBool(a, 2, ctx, false);
  if (ignore < 0 || ignore > 3) throw VALUE;
  const out: Scalar[] = [];
  const visit = (v: Scalar) => {
    if (ignore & 1 && v === null) return;
    if (ignore & 2 && isError(v)) return;
    out.push(v ?? 0);
  };
  if (byCol) for (let c = 0; c < m.cols; c++) for (let r = 0; r < m.rows; r++) visit(m.get(r, c));
  else m.data.forEach(visit);
  if (!out.length) throw err('#CALC!');
  return column ? new Matrix(out.length, 1, out) : new Matrix(1, out.length, out);
}

/** WRAPROWS and WRAPCOLS: a row or column of values into lines of a given length. */
function wrapLine(a: Operand[], ctx: Ctx, rows: boolean): Matrix {
  const m = matrixOf(a[0], ctx);
  if (m.rows > 1 && m.cols > 1) throw VALUE;
  const n = Math.trunc(num(scalarOf(a[1], ctx)));
  if (n < 1) throw NUM;
  const pad = given(a, 2) ? scalarOf(a[2], ctx) : NA;
  const values = m.data.map((v) => v ?? 0);
  const lines = Math.ceil(values.length / n);
  const at = (line: number, i: number) => (line * n + i < values.length ? values[line * n + i] : pad);
  return rows ? Matrix.build(lines, n, (r, c) => at(r, c)) : Matrix.build(n, lines, (r, c) => at(c, r));
}

/** Where the nth delimiter (any of several) is, from the start or, for negative n, the end; the text's end counts as one with `matchEnd`. */
function findNth(t: string, delims: string[], n: number, insensitive: boolean, matchEnd: boolean): [number, number] | undefined {
  const s = insensitive ? t.toLowerCase() : t;
  const ds = delims.map((d) => (insensitive ? d.toLowerCase() : d));
  let found: [number, number] | undefined;
  if (n > 0) {
    let from = 0;
    for (let k = 0; k < n; k++) {
      let best = -1;
      let len = 0;
      for (const d of ds) {
        const i = s.indexOf(d, from);
        if (i >= 0 && (best < 0 || i < best)) [best, len] = [i, d.length];
      }
      if (best < 0) return matchEnd && k === n - 1 && from <= s.length ? [s.length, 0] : undefined;
      found = [best, len];
      from = best + Math.max(len, 1);
    }
    return found;
  }
  let end = s.length;
  for (let k = 0; k < -n; k++) {
    let best = -1;
    let len = 0;
    for (const d of ds) {
      const i = end - d.length < 0 ? -1 : s.lastIndexOf(d, end - d.length);
      if (i > best) [best, len] = [i, d.length];
    }
    if (best < 0) return matchEnd && k === -n - 1 ? [0, 0] : undefined;
    found = [best, len];
    end = best - (len ? 0 : 1);
  }
  return found;
}

const textsOf = (op: Operand, ctx: Ctx) => {
  const v = valuesOf(op, ctx);
  return (v instanceof Matrix ? v.data : [v]).map((x) => text(x));
};

/** TEXTBEFORE and TEXTAFTER. */
function textAround(a: Operand[], ctx: Ctx, after: boolean): Scalar {
  const t = text(scalarOf(a[0], ctx));
  const n = Math.trunc(optNum(a, 2, ctx, 1));
  if (n === 0 || Math.abs(n) > Math.max(t.length, 1)) throw VALUE;
  const at = findNth(t, textsOf(a[1], ctx), n, optNum(a, 3, ctx, 0) === 1, optNum(a, 4, ctx, 0) === 1);
  if (!at) return given(a, 5) ? scalarOf(a[5], ctx) : NA;
  return after ? t.slice(at[0] + at[1]) : t.slice(0, at[0]);
}

/** Text split at any of several delimiters. */
function splitOn(t: string, delims: string[], insensitive: boolean): string[] {
  const live = delims.filter((d) => d !== '');
  if (!live.length) return [t];
  const s = insensitive ? t.toLowerCase() : t;
  const ds = live.map((d) => (insensitive ? d.toLowerCase() : d));
  const out: string[] = [];
  let from = 0;
  for (;;) {
    let best = -1;
    let len = 0;
    for (const d of ds) {
      const i = s.indexOf(d, from);
      if (i >= 0 && (best < 0 || i < best || (i === best && d.length > len))) [best, len] = [i, d.length];
    }
    if (best < 0) break;
    out.push(t.slice(from, best));
    from = best + len;
  }
  out.push(t.slice(from));
  return out;
}

/** A value as VALUETOTEXT and ARRAYTOTEXT write it: as shown, or (strict) text in quotes. */
function valueText(v: Scalar, strict: boolean): string {
  if (isError(v)) return v.error;
  if (typeof v === 'string') return strict ? `"${v.replace(/"/g, '""')}"` : v;
  return text(v);
}

const groupThousands = (digits: string) => digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

/** Excel's present value relations: payment, future and present value, at rate r over n periods, paid at the start (type 1) or end. */
function payment(r: number, n: number, pv: number, fv: number, type: number): number {
  if (r === 0) return -(pv + fv) / n;
  const g = (1 + r) ** n;
  return -(r * (fv + pv * g)) / ((1 + r * type) * (g - 1));
}

function futureValue(r: number, n: number, pmt: number, pv: number, type: number): number {
  if (r === 0) return -(pv + pmt * n);
  const g = (1 + r) ** n;
  return -(pv * g + (pmt * (1 + r * type) * (g - 1)) / r);
}

function presentValue(r: number, n: number, pmt: number, fv: number, type: number): number {
  if (r === 0) return -(fv + pmt * n);
  const g = (1 + r) ** n;
  return -(fv + (pmt * (1 + r * type) * (g - 1)) / r) / g;
}

/** Newton's method from a guess, as Excel's RATE and IRR search; #NUM! when it doesn't settle. */
function solve(f: (x: number) => number, guess: number): number {
  let x = guess;
  for (let i = 0; i < 100; i++) {
    const y = f(x);
    if (Math.abs(y) < 1e-10) return x;
    const h = Math.max(1e-7, Math.abs(x) * 1e-7);
    const slope = (f(x + h) - f(x - h)) / (2 * h);
    if (!Number.isFinite(slope) || slope === 0) break;
    const next = x - y / slope;
    if (!Number.isFinite(next) || next <= -1) break;
    if (Math.abs(next - x) < 1e-12) return next;
    x = next;
  }
  throw NUM;
}

const ERROR_NUMBERS: Record<string, number> = {
  '#NULL!': 1,
  '#DIV/0!': 2,
  '#VALUE!': 3,
  '#REF!': 4,
  '#NAME?': 5,
  '#NUM!': 6,
  '#N/A': 7,
  '#GETTING_DATA': 8,
  '#SPILL!': 9,
  '#CONNECT!': 10,
  '#BLOCKED!': 11,
  '#UNKNOWN!': 12,
  '#FIELD!': 13,
  '#CALC!': 14,
};

// -----------------------------------------------------------------------------
// Statistics, distributions and matrices

/** PERCENTILE.INC (inclusive) or PERCENTILE.EXC of sorted values. */
function percentile(sorted: number[], k: number, inclusive: boolean): number {
  const n = sorted.length;
  if (!n || k < 0 || k > 1) throw NUM;
  const rank = inclusive ? k * (n - 1) + 1 : k * (n + 1);
  if (rank < 1 || rank > n) throw NUM;
  const lo = Math.floor(rank);
  const v = sorted[lo - 1];
  return lo === n ? v : v + (rank - lo) * (sorted[lo] - v);
}

const ascending = (xs: number[]) => [...xs].sort((a, b) => a - b);

/** The values that occur most often (at least twice), in the order they first appear. */
function modes(values: number[]): number[] {
  const counts = new Map<number, number>();
  for (const v of values) counts.set(v, (counts.get(v) ?? 0) + 1);
  const most = Math.max(0, ...counts.values());
  return most < 2 ? [] : [...counts].filter(([, n]) => n === most).map(([v]) => v);
}

/** Where x falls among sorted values, as a fraction: PERCENTRANK, inclusive or exclusive of the ends, cut (not rounded) to `digits`. */
function percentRank(a: Operand[], ctx: Ctx, inclusive: boolean): number {
  const sorted = ascending(list([a[0]], ctx));
  const x = num(scalarOf(a[1], ctx));
  const digits = Math.trunc(optNum(a, 2, ctx, 3));
  if (!sorted.length || digits < 1) throw NUM;
  const n = sorted.length;
  if (x < sorted[0] || x > sorted[n - 1]) throw NA;
  let below = 0;
  while (below < n && sorted[below] < x) below++;
  let r: number;
  if (sorted[below] === x) r = inclusive ? (n === 1 ? 1 : below / (n - 1)) : (below + 1) / (n + 1);
  else {
    const fraction = (x - sorted[below - 1]) / (sorted[below] - sorted[below - 1]);
    r = inclusive ? (below - 1 + fraction) / (n - 1) : (below + fraction) / (n + 1);
  }
  const scale = 10 ** digits;
  return Math.floor(round15(r * scale)) / scale;
}

/** The numbers AGGREGATE works on: errors and other subtotals left out as its option says. */
function aggregateValues(args: Operand[], ctx: Ctx, skipErrors: boolean, skipSubtotals: boolean): number[] {
  const out: number[] = [];
  const take = (v: Scalar) => {
    if (typeof v === 'number') out.push(v);
    else if (isError(v) && !skipErrors) throw v;
  };
  for (const x of args) {
    if (x instanceof Ref) ctx.host.each(x, take, skipSubtotals);
    else if (x instanceof Matrix) x.data.forEach(take);
    else if (x !== undefined) {
      const v = scalarOf(x, ctx);
      if (isError(v)) {
        if (!skipErrors) throw v;
      } else out.push(num(v));
    }
  }
  return out;
}

/** AGGREGATE's functions 1 to 13, on all their arguments. */
const AGGREGATE_LISTS: Record<number, (v: number[], args: Operand[], ctx: Ctx) => number> = {
  ...AGGREGATES,
  12: (v) => median(v),
  13: (v) => {
    const m = modes(v);
    if (!m.length) throw NA;
    return m[0];
  },
};

/** Functions 14 to 19, which take one array and a k. */
const AGGREGATE_RANKED: Record<number, (sorted: number[], k: number) => number> = {
  14: (s, k) => nth(s, k, true),
  15: (s, k) => nth(s, k, false),
  16: (s, k) => percentile(s, k, true),
  17: (s, q) => {
    if (Math.trunc(q) < 0 || Math.trunc(q) > 4) throw NUM;
    return percentile(s, Math.trunc(q) / 4, true);
  },
  18: (s, k) => percentile(s, k, false),
  19: (s, q) => {
    if (Math.trunc(q) < 1 || Math.trunc(q) > 3) throw NUM;
    return percentile(s, Math.trunc(q) / 4, false);
  },
};

/** Pairs of numbers from two equal-sized ranges or arrays, for the regression functions; a pair with text or a blank is left out. */
function pairs(ys: Operand, xs: Operand, ctx: Ctx): { y: number[]; x: number[] } {
  const a = matrixOf(ys, ctx).data;
  const b = matrixOf(xs, ctx).data;
  if (a.length !== b.length) throw NA;
  const y: number[] = [];
  const x: number[] = [];
  for (let i = 0; i < a.length; i++) {
    if (isError(a[i])) throw a[i];
    if (isError(b[i])) throw b[i];
    if (typeof a[i] === 'number' && typeof b[i] === 'number') {
      y.push(a[i] as number);
      x.push(b[i] as number);
    }
  }
  return { y, x };
}

/** The sums a straight line through pairs is worked out from. */
function line(ys: Operand, xs: Operand, ctx: Ctx) {
  const { y, x } = pairs(ys, xs, ctx);
  const n = y.length;
  if (!n) throw DIV0;
  const my = sum(y) / n;
  const mx = sum(x) / n;
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < n; i++) {
    sxy += (x[i] - mx) * (y[i] - my);
    sxx += (x[i] - mx) ** 2;
    syy += (y[i] - my) ** 2;
  }
  return { n, mx, my, sxy, sxx, syy };
}

function slopeOf(ys: Operand, xs: Operand, ctx: Ctx): { slope: number; intercept: number } {
  const l = line(ys, xs, ctx);
  if (l.sxx === 0) throw DIV0;
  const slope = l.sxy / l.sxx;
  return { slope, intercept: l.my - slope * l.mx };
}

function correl(ys: Operand, xs: Operand, ctx: Ctx): number {
  const l = line(ys, xs, ctx);
  if (l.sxx === 0 || l.syy === 0) throw DIV0;
  return l.sxy / Math.sqrt(l.sxx * l.syy);
}

/** The error function, to full double precision (W. J. Cody's rational approximations): erf, or with `complement`, erfc. */
function erf(x: number, complement = false): number {
  const y = Math.abs(x);
  if (y <= 0.46875) {
    const A = [3.1611237438705655, 113.86415415105016, 377.485237685302, 3209.3775891384694, 0.18577770618460315];
    const B = [23.601290952344122, 244.02463793444417, 1282.6165260773723, 2844.236833439171];
    const z = y > 1.11e-16 ? y * y : 0;
    let num = A[4] * z;
    let den = z;
    for (let i = 0; i < 3; i++) {
      num = (num + A[i]) * z;
      den = (den + B[i]) * z;
    }
    const r = (x * (num + A[3])) / (den + B[3]);
    return complement ? 1 - r : r;
  }
  let c: number;
  if (y <= 4) {
    const C = [0.5641884969886701, 8.883149794388375, 66.11919063714163, 298.6351381974001, 881.952221241769, 1712.0476126340707, 2051.0783778260716, 1230.3393547979972, 2.1531153547440383e-8];
    const D = [15.744926110709835, 117.6939508913125, 537.1811018620099, 1621.3895745666903, 3290.7992357334597, 4362.619090143247, 3439.3676741437216, 1230.3393548037495];
    let num = C[8] * y;
    let den = y;
    for (let i = 0; i < 7; i++) {
      num = (num + C[i]) * y;
      den = (den + D[i]) * y;
    }
    c = (num + C[7]) / (den + D[7]);
  } else {
    const P = [0.30532663496123236, 0.36034489994980445, 0.12578172611122926, 0.016083785148742275, 0.0006587491615298367, 0.016315387137302097];
    const Q = [2.5685201922898684, 1.8729528499234604, 0.5279051029514283, 0.06051834131244132, 0.0023352049762686918];
    const z = 1 / (y * y);
    let num = P[5] * z;
    let den = z;
    for (let i = 0; i < 4; i++) {
      num = (num + P[i]) * z;
      den = (den + Q[i]) * z;
    }
    c = (0.5641895835477563 - (z * (num + P[4])) / (den + Q[4])) / y;
  }
  const s = Math.trunc(y * 16) / 16;
  c *= Math.exp(-s * s) * Math.exp(-(y - s) * (y + s));
  if (complement) return x < 0 ? 2 - c : c;
  return x < 0 ? c - 1 : 1 - c;
}

const normalCdf = (z: number) => 0.5 * erf(-z / Math.SQRT2, true);
const normalPdf = (z: number) => Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);

/** The standard normal quantile: P. J. Acklam's approximation, then refined against normalCdf. */
function normalInverse(p: number): number {
  if (p <= 0 || p >= 1) throw NUM;
  const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
  const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
  const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
  let x: number;
  if (p < 0.02425) {
    const q = Math.sqrt(-2 * Math.log(p));
    x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p > 1 - 0.02425) {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else {
    const q = p - 0.5;
    const r = q * q;
    x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  }
  for (let i = 0; i < 2; i++) {
    const e = normalCdf(x) - p;
    const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2);
    x -= u / (1 + (x * u) / 2);
  }
  return x;
}

/** ln(n!) by summing, exact enough for the binomial sums here. */
function logFactorial(n: number): number {
  let s = 0;
  for (let i = 2; i <= n; i++) s += Math.log(i);
  return s;
}

/** P(lo <= X <= hi) for X binomial with n trials and success chance p. */
function binomialRange(n: number, p: number, lo: number, hi: number): number {
  if (p === 0 || p === 1) {
    const certain = p === 0 ? 0 : n;
    return certain >= lo && certain <= hi ? 1 : 0;
  }
  let lnChoose = logFactorial(n) - logFactorial(lo) - logFactorial(n - lo);
  let total = 0;
  for (let k = lo; k <= hi; k++) {
    total += Math.exp(lnChoose + k * Math.log(p) + (n - k) * Math.log1p(-p));
    lnChoose += Math.log(n - k) - Math.log(k + 1);
  }
  return total;
}

/** The largest matrix MINVERSE and MDETERM work on, as the work grows with the cube of its size. */
const MAX_MATRIX = 400;

/** A matrix of numbers; a blank or text in it is #VALUE!. */
function numberMatrix(op: Operand, ctx: Ctx): number[][] {
  const m = matrixOf(op, ctx);
  return m.rowsOf().map((row) =>
    row.map((v) => {
      if (isError(v)) throw v;
      if (typeof v !== 'number') throw VALUE;
      return v;
    }),
  );
}

function square(op: Operand, ctx: Ctx): number[][] {
  const m = numberMatrix(op, ctx);
  if (m.length !== m[0].length) throw VALUE;
  if (m.length > MAX_MATRIX) throw NUM;
  return m;
}

// -----------------------------------------------------------------------------
// Bonds, accrued interest, discounted securities and depreciation (shared with the HyperFormula add-on)

const calendarOf = (ctx: Ctx) => new Calendar(ctx.host.date1904 ? { year: 1904, month: 1, day: 1 } : { year: 1899, month: 12, day: 30 });

/**
 * A bond function: its date arguments (by position) read as dates, the rest
 * as numbers, missing ones from `defaults`; #NUM! for arguments Excel refuses.
 */
function bond(min: number, dates: number[], defaults: Record<number, number | boolean>, f: (b: Bonds, args: number[], flags: boolean[]) => number): FunctionSpec {
  const max = Math.max(min, ...Object.keys(defaults).map((k) => Number(k) + 1));
  return scalar(min, max, (a, ctx) => {
    const nums: number[] = [];
    const flags: boolean[] = [];
    for (let i = 0; i < max; i++) {
      const v = a[i];
      const missing = v === undefined || v === null;
      const d = defaults[i];
      if (typeof d === 'boolean') flags[i] = missing ? d : bool(v);
      else nums[i] = missing ? (d ?? 0) : dates.includes(i) ? dateArg(v, ctx) : num(v);
    }
    try {
      return f(new Bonds(calendarOf(ctx)), nums, flags);
    } catch (e) {
      if (e instanceof BondError) throw NUM;
      throw e;
    }
  });
}

const couponFn = (pick: (b: Bonds, s: number, m: number, f: number, basis: number) => number) => bond(3, [0, 1], { 3: 0 }, (b, [s, m, f, basis]) => pick(b, s, m, f, basis));
const discountedFn = (pick: (b: Bonds, s: number, m: number, x: number, y: number, basis: number) => number) => bond(4, [0, 1], { 4: 0 }, (b, [s, m, x, y, basis]) => pick(b, s, m, x, y, basis));

// -----------------------------------------------------------------------------
// Least-squares regression (shared with the HyperFormula add-on)

/** Runs shared regression code, with the error it refuses arguments with. */
function regression<T>(f: () => T): T {
  try {
    return f();
  } catch (e) {
    if (e instanceof RegressionError) throw err(e.code);
    throw e;
  }
}

/** TREND, or GROWTH with `log`. */
function projection(a: Operand[], ctx: Ctx, log: boolean): Matrix {
  const ys = numberMatrix(a[0], ctx);
  const xs = given(a, 1) ? numberMatrix(a[1], ctx) : undefined;
  const news = given(a, 2) ? numberMatrix(a[2], ctx) : undefined;
  const constant = optBool(a, 3, ctx, true);
  return Matrix.of(regression(() => project(prepare(ys, xs, log), constant, log, news)));
}

/** LINEST, or LOGEST with `log`. */
function estimation(a: Operand[], ctx: Ctx, log: boolean): Matrix {
  const ys = numberMatrix(a[0], ctx);
  const xs = given(a, 1) ? numberMatrix(a[1], ctx) : undefined;
  const rows = regression(() => estimate(prepare(ys, xs, log), optBool(a, 2, ctx, true), log, optBool(a, 3, ctx, false)));
  return Matrix.of(rows.map((row) => row.map((v) => (typeof v === 'number' ? v : err(v.error)))));
}

// -----------------------------------------------------------------------------
// Loans, depreciation and cash flows on dates; weeks and working days

/** IPMT: the interest in one period of a loan or annuity. */
function interestIn(r: number, per: number, n: number, pv: number, fv: number, type: number): number {
  if (per < 1 || per > n) throw NUM;
  if (type === 1 && per === 1) return 0;
  const pmt = payment(r, n, pv, fv, type);
  const interest = futureValue(r, per - 1, pmt, pv, type) * r;
  return type === 1 ? interest / (1 + r) : interest;
}

/** CUMIPMT and CUMPRINC: interest (or principal) paid from one period to another. */
function cumulative(a: Scalar[], principal: boolean): number {
  const [r, n, pv, s, e] = a.slice(0, 5).map(num);
  const type = num(a[5]);
  const start = Math.trunc(s);
  const end = Math.trunc(e);
  if (r <= 0 || n <= 0 || pv <= 0 || start < 1 || end < start || end > n || (type !== 0 && type !== 1)) throw NUM;
  const pmt = payment(r, n, pv, 0, type);
  let total = 0;
  for (let per = start; per <= end; per++) {
    const interest = interestIn(r, per, n, pv, 0, type);
    total += principal ? pmt - interest : interest;
  }
  return total;
}

/** Cash flows and their dates for XNPV and XIRR: numbers, as many dates as values, none before the first. */
function datedFlows(a: Operand[], ctx: Ctx): { values: number[]; days: number[] } {
  const values = numberMatrix(a[0], ctx).flat();
  const dates = numberMatrix(a[1], ctx).flat().map(Math.trunc);
  if (values.length !== dates.length || !values.length) throw NUM;
  if (dates.some((d) => d < dates[0])) throw NUM;
  return { values, days: dates.map((d) => d - dates[0]) };
}

const xnpv = (r: number, f: { values: number[]; days: number[] }) => f.values.reduce((s, v, i) => s + v / (1 + r) ** (f.days[i] / 365), 0);

/** A JavaScript date (UTC) for a serial number. */
function jsDate(serial: number, ctx: Ctx): Date {
  const { y, m, d } = dateOf(serial, ctx.host.date1904);
  return new Date(Date.UTC(y, m - 1, d));
}

/** The ISO 8601 week a date falls in: weeks start on Monday, and week 1 holds the year's first Thursday. */
function isoWeek(serial: number, ctx: Ctx): number {
  const d = jsDate(serial, ctx);
  const day = (d.getUTCDay() + 6) % 7;
  const thursday = new Date(d.getTime() + (3 - day) * DAY_MS);
  const first = new Date(Date.UTC(thursday.getUTCFullYear(), 0, 1));
  return 1 + Math.floor((thursday.getTime() - first.getTime()) / DAY_MS / 7);
}

/** The weekend for NETWORKDAYS.INTL and WORKDAY.INTL, Monday first: a code 1–7 or 11–17, or seven 0s and 1s. */
function weekendOf(op: Operand | undefined, ctx: Ctx): boolean[] {
  const v = op === undefined || op === null ? 1 : scalarOf(op, ctx);
  if (typeof v === 'string') {
    if (!/^[01]{7}$/.test(v)) throw VALUE;
    return [...v].map((c) => c === '1');
  }
  const code = Math.trunc(num(v));
  const days = Array<boolean>(7).fill(false);
  if (code >= 1 && code <= 7) {
    // 1 is Saturday and Sunday, 2 Sunday and Monday, …, 7 Friday and Saturday.
    days[(code + 4) % 7] = true;
    days[(code + 5) % 7] = true;
  } else if (code >= 11 && code <= 17) days[(code - 11 + 6) % 7] = true;
  else throw NUM;
  return days;
}

const workday = (serial: number, weekend: boolean[], holidays: Set<number>, ctx: Ctx) => !weekend[weekday(serial, 2, ctx) - 1] && !holidays.has(serial);

/** DAYS360: days between dates on a 360-day year, by the US (NASD) method or the European one. */
function days360(start: number, end: number, european: boolean, ctx: Ctx): number {
  let { y: y1, m: m1, d: d1 } = dateOf(start, ctx.host.date1904);
  let { y: y2, m: m2, d: d2 } = dateOf(end, ctx.host.date1904);
  const last = (y: number, m: number, d: number) => dateOf(serialOf(y, m + 1, 0, ctx.host.date1904), ctx.host.date1904).d === d;
  if (european) {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  } else {
    const startLast = last(y1, m1, d1);
    if (startLast) d1 = 30;
    if (d2 === 31) {
      if (d1 < 30) {
        d2 = 1;
        m2 += 1;
      } else d2 = 30;
    }
  }
  if (m2 > 12) {
    m2 -= 12;
    y2 += 1;
  }
  return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
}

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
      // Blank cells a result reads from show as 0 in it, as in Excel.
      const m = inner instanceof Matrix ? filled(inner) : new Matrix(1, 1, [inner]);
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

  // Lookups, and arrays reshaped
  LOOKUP: {
    min: 2,
    max: 3,
    fn: (a, ctx) => {
      const target = scalarOf(a[0], ctx);
      const look = matrixOf(a[1], ctx);
      let keys: Scalar[];
      let results: Scalar[];
      if (a[2] === undefined) {
        // The array form: the first row or column searched, the last one returned.
        const across = look.cols > look.rows;
        keys = across ? look.data.slice(0, look.cols) : Array.from({ length: look.rows }, (_, r) => look.get(r, 0));
        results = across ? look.data.slice((look.rows - 1) * look.cols) : Array.from({ length: look.rows }, (_, r) => look.get(r, look.cols - 1));
      } else {
        keys = look.data;
        results = matrixOf(a[2], ctx).data;
      }
      const at = approximate(keys, target);
      if (at < 0 || at >= results.length) throw NA;
      return results[at];
    },
  },
  TAKE: { min: 2, max: 3, fn: (a, ctx) => takeOrDrop(a, ctx, true) },
  DROP: { min: 2, max: 3, fn: (a, ctx) => takeOrDrop(a, ctx, false) },
  CHOOSEROWS: { min: 2, max: MANY, fn: (a, ctx) => chooseLines(a, ctx, true) },
  CHOOSECOLS: { min: 2, max: MANY, fn: (a, ctx) => chooseLines(a, ctx, false) },
  VSTACK: { min: 1, max: MANY, fn: (a, ctx) => stack(a, ctx, true) },
  HSTACK: { min: 1, max: MANY, fn: (a, ctx) => stack(a, ctx, false) },
  TOCOL: { min: 1, max: 3, fn: (a, ctx) => toLine(a, ctx, true) },
  TOROW: { min: 1, max: 3, fn: (a, ctx) => toLine(a, ctx, false) },
  WRAPROWS: { min: 2, max: 3, fn: (a, ctx) => wrapLine(a, ctx, true) },
  WRAPCOLS: { min: 2, max: 3, fn: (a, ctx) => wrapLine(a, ctx, false) },
  EXPAND: {
    min: 2,
    max: 4,
    fn: (a, ctx) => {
      const m = filled(matrixOf(a[0], ctx));
      const rows = Math.trunc(optNum(a, 1, ctx, m.rows));
      const cols = Math.trunc(optNum(a, 2, ctx, m.cols));
      if (rows < m.rows || cols < m.cols) throw VALUE;
      sized(rows, cols);
      const pad = given(a, 3) ? scalarOf(a[3], ctx) : NA;
      return Matrix.build(rows, cols, (r, c) => (r < m.rows && c < m.cols ? m.get(r, c) : pad));
    },
  },

  // Text, split and joined
  TEXTBEFORE: { min: 2, max: 6, fn: (a, ctx) => textAround(a, ctx, false) },
  TEXTAFTER: { min: 2, max: 6, fn: (a, ctx) => textAround(a, ctx, true) },
  TEXTSPLIT: {
    min: 2,
    max: 6,
    fn: (a, ctx) => {
      const t = text(scalarOf(a[0], ctx));
      const insensitive = optNum(a, 4, ctx, 0) === 1;
      const skip = optBool(a, 3, ctx, false);
      const pad = given(a, 5) ? scalarOf(a[5], ctx) : NA;
      const keep = (parts: string[]) => (skip ? parts.filter((p) => p !== '') : parts);
      const lines = keep(given(a, 2) ? splitOn(t, textsOf(a[2], ctx), insensitive) : [t]).map((line) => keep(given(a, 1) ? splitOn(line, textsOf(a[1], ctx), insensitive) : [line]));
      if (!lines.length) throw err('#CALC!');
      const cols = Math.max(1, ...lines.map((l) => l.length));
      return Matrix.build(lines.length, cols, (r, c) => (c < lines[r].length ? lines[r][c] : pad));
    },
  },
  VALUETOTEXT: { min: 1, max: 2, fn: (a, ctx) => valueText(scalarOf(a[0], ctx), optNum(a, 1, ctx, 0) === 1) },
  ARRAYTOTEXT: {
    min: 1,
    max: 2,
    fn: (a, ctx) => {
      const m = matrixOf(a[0], ctx);
      const strict = optNum(a, 1, ctx, 0) === 1;
      if (!strict) return m.data.map((v) => valueText(v, false)).join(', ');
      return `{${m.rowsOf().map((row) => row.map((v) => valueText(v, true)).join(',')).join(';')}}`;
    },
  },
  FIXED: scalar(1, 3, ([x, d, plain]) => {
    const digits = Math.trunc(optional(d, 2));
    if (digits > 127) throw VALUE;
    const n = excelRound(num(x), digits, 'half');
    const [whole, frac] = Math.abs(n).toFixed(Math.max(0, digits)).split('.');
    const body = (plain !== undefined && plain !== null && bool(plain) ? whole : groupThousands(whole)) + (frac ? `.${frac}` : '');
    return n < 0 ? `-${body}` : body;
  }),
  NUMBERVALUE: scalar(1, 3, ([t, dec, group]) => {
    let s = text(t).replace(/\s/g, '');
    if (!s) return 0;
    const d = dec === undefined || dec === null ? '.' : text(dec).charAt(0);
    const g = group === undefined || group === null ? ',' : text(group).charAt(0);
    if (!d || d === g) throw VALUE;
    let percent = 0;
    while (s.endsWith('%')) {
      percent++;
      s = s.slice(0, -1);
    }
    const at = s.indexOf(d);
    const whole = (at < 0 ? s : s.slice(0, at)).split(g).join('');
    const frac = at < 0 ? '' : s.slice(at + 1);
    if (frac.includes(g) || frac.includes(d)) throw VALUE;
    const n = Number(`${whole}${frac ? `.${frac}` : ''}`);
    if (!Number.isFinite(n) || whole + frac === '') throw VALUE;
    return n / 100 ** percent;
  }),
  // Bytes are characters, as in Excel set to a language that writes one byte a character.
  LENB: scalar(1, 1, ([t]) => text(t).length),
  MIDB: scalar(3, 3, ([t, s, n]) => {
    const start = Math.trunc(num(s));
    const k = Math.trunc(num(n));
    if (start < 1 || k < 0) throw VALUE;
    return text(t).substr(start - 1, k);
  }),
  TYPE: {
    min: 1,
    max: 1,
    fn: ([v], ctx) => {
      if ((v instanceof Ref && !v.single) || v instanceof Matrix) return 64;
      if (v instanceof Lambda) return 128;
      const x = scalarOf(v, ctx);
      return isError(x) ? 16 : typeof x === 'string' ? 2 : typeof x === 'boolean' ? 4 : 1;
    },
  },
  'ERROR.TYPE': scalar(1, 1, ([v]) => (isError(v) && ERROR_NUMBERS[v.error] ? ERROR_NUMBERS[v.error] : NA)),

  // Money over time
  PMT: scalar(3, 5, ([r, n, pv, fv, t]) => payment(num(r), num(n), num(pv), optional(fv, 0), optional(t, 0) ? 1 : 0)),
  FV: scalar(3, 5, ([r, n, p, pv, t]) => futureValue(num(r), num(n), num(p), optional(pv, 0), optional(t, 0) ? 1 : 0)),
  PV: scalar(3, 5, ([r, n, p, fv, t]) => presentValue(num(r), num(n), num(p), optional(fv, 0), optional(t, 0) ? 1 : 0)),
  NPER: scalar(3, 5, ([r, p, pv, fv, t]) => {
    const rate = num(r);
    const pay = num(p);
    const f = optional(fv, 0);
    const type = optional(t, 0) ? 1 : 0;
    if (rate === 0) {
      if (pay === 0) throw NUM;
      return -(num(pv) + f) / pay;
    }
    const x = (pay * (1 + rate * type) - f * rate) / (pay * (1 + rate * type) + num(pv) * rate);
    if (x <= 0) throw NUM;
    return Math.log(x) / Math.log(1 + rate);
  }),
  RATE: scalar(3, 6, ([n, p, pv, fv, t, g]) => {
    const periods = num(n);
    const pay = num(p);
    const present = num(pv);
    const future = optional(fv, 0);
    const type = optional(t, 0) ? 1 : 0;
    return solve((r) => (r === 0 ? present + pay * periods + future : present * (1 + r) ** periods + (pay * (1 + r * type) * ((1 + r) ** periods - 1)) / r + future), optional(g, 0.1));
  }),
  NPV: {
    min: 2,
    max: MANY,
    fn: (a, ctx) => {
      const r = num(scalarOf(a[0], ctx));
      let total = 0;
      let i = 1;
      numbers(a.slice(1), ctx, (v) => (total += v / (1 + r) ** i++));
      return total;
    },
  },
  IRR: {
    min: 1,
    max: 2,
    fn: (a, ctx) => {
      const flows = list([a[0]], ctx);
      if (!flows.some((v) => v > 0) || !flows.some((v) => v < 0)) throw NUM;
      return solve((r) => flows.reduce((s, v, i) => s + v / (1 + r) ** i, 0), optNum(a, 1, ctx, 0.1));
    },
  },
  YEARFRAC: scalar(2, 3, ([s, e, b], ctx) => {
    const basis = Math.trunc(optional(b, 0));
    if (!validBasis(basis)) throw NUM;
    const start = Math.floor(dateArg(s, ctx));
    const end = Math.floor(dateArg(e, ctx));
    return yearFrac(calendarOf(ctx), start, end, basis);
  }),
  PERCENTOF: { min: 2, max: 2, fn: (a, ctx) => {
    const all = sum(list([a[1]], ctx));
    if (all === 0) throw DIV0;
    return sum(list([a[0]], ctx)) / all;
  } },
  GESTEP: scalar(1, 2, ([x, s]) => (num(x) >= optional(s, 0) ? 1 : 0)),
  PERMUT: scalar(2, 2, ([n, k]) => {
    const a = Math.trunc(num(n));
    const b = Math.trunc(num(k));
    if (a <= 0 || b < 0 || a < b) throw NUM;
    let p = 1;
    for (let i = 0; i < b; i++) p *= a - i;
    return p;
  }),
  PERMUTATIONA: scalar(2, 2, ([n, k]) => {
    const a = Math.trunc(num(n));
    const b = Math.trunc(num(k));
    if (a < 0 || b < 0) throw NUM;
    return a ** b;
  }),

  // Statistics
  AGGREGATE: {
    min: 3,
    max: MANY,
    fn: (a, ctx) => {
      const code = Math.trunc(num(scalarOf(a[0], ctx)));
      const option = Math.trunc(num(scalarOf(a[1], ctx)));
      if (option < 0 || option > 7) throw VALUE;
      const skipErrors = option === 2 || option === 3 || option === 6 || option === 7;
      const skipSubtotals = option <= 3;
      if (AGGREGATE_RANKED[code]) {
        if (a.length < 4) throw VALUE;
        return AGGREGATE_RANKED[code](ascending(aggregateValues([a[2]], ctx, skipErrors, skipSubtotals)), num(scalarOf(a[3], ctx)));
      }
      const f = code === 3 ? undefined : AGGREGATE_LISTS[code];
      const rest = a.slice(2);
      if (code === 3) return counta(rest, ctx, skipSubtotals);
      if (!f) throw VALUE;
      return f(aggregateValues(rest, ctx, skipErrors, skipSubtotals), rest, ctx);
    },
  },
  MODE: { min: 1, max: MANY, fn: (a, ctx) => {
    const m = modes(list(a, ctx));
    if (!m.length) throw NA;
    return m[0];
  } },
  'MODE.SNGL': { min: 1, max: MANY, fn: (a, ctx) => {
    const m = modes(list(a, ctx));
    if (!m.length) throw NA;
    return m[0];
  } },
  'MODE.MULT': { min: 1, max: MANY, fn: (a, ctx) => {
    const m = modes(list(a, ctx));
    if (!m.length) throw NA;
    return new Matrix(m.length, 1, m);
  } },
  PERCENTRANK: { min: 2, max: 3, fn: (a, ctx) => percentRank(a, ctx, true) },
  'PERCENTRANK.INC': { min: 2, max: 3, fn: (a, ctx) => percentRank(a, ctx, true) },
  'PERCENTRANK.EXC': { min: 2, max: 3, fn: (a, ctx) => percentRank(a, ctx, false) },
  PERCENTILE: { min: 2, max: 2, fn: (a, ctx) => percentile(ascending(list([a[0]], ctx)), num(scalarOf(a[1], ctx)), true) },
  'PERCENTILE.INC': { min: 2, max: 2, fn: (a, ctx) => percentile(ascending(list([a[0]], ctx)), num(scalarOf(a[1], ctx)), true) },
  'PERCENTILE.EXC': { min: 2, max: 2, fn: (a, ctx) => percentile(ascending(list([a[0]], ctx)), num(scalarOf(a[1], ctx)), false) },
  QUARTILE: { min: 2, max: 2, fn: (a, ctx) => AGGREGATE_RANKED[17](ascending(list([a[0]], ctx)), num(scalarOf(a[1], ctx))) },
  'QUARTILE.INC': { min: 2, max: 2, fn: (a, ctx) => AGGREGATE_RANKED[17](ascending(list([a[0]], ctx)), num(scalarOf(a[1], ctx))) },
  'QUARTILE.EXC': { min: 2, max: 2, fn: (a, ctx) => AGGREGATE_RANKED[19](ascending(list([a[0]], ctx)), num(scalarOf(a[1], ctx))) },
  'RANK.AVG': {
    min: 2,
    max: 3,
    fn: (a, ctx) => {
      const x = num(scalarOf(a[0], ctx));
      const xs = list([a[1]], ctx);
      const asc = given(a, 2) && num(scalarOf(a[2], ctx)) !== 0;
      const ties = xs.filter((v) => round15(v) === round15(x)).length;
      if (!ties) throw NA;
      const before = xs.filter((v) => (asc ? v < x : v > x) && round15(v) !== round15(x)).length;
      return before + (ties + 1) / 2;
    },
  },
  TRIMMEAN: {
    min: 2,
    max: 2,
    fn: (a, ctx) => {
      const xs = ascending(list([a[0]], ctx));
      const p = num(scalarOf(a[1], ctx));
      if (p < 0 || p >= 1 || !xs.length) throw NUM;
      const cut = Math.floor((xs.length * p) / 2);
      const kept = xs.slice(cut, xs.length - cut);
      return sum(kept) / kept.length;
    },
  },
  KURT: {
    min: 1,
    max: MANY,
    fn: (a, ctx) => {
      const xs = list(a, ctx);
      const n = xs.length;
      if (n < 4) throw DIV0;
      const s = Math.sqrt(variance(xs, true));
      if (s === 0) throw DIV0;
      const mean = sum(xs) / n;
      const fourth = sum(xs.map((x) => ((x - mean) / s) ** 4));
      return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * fourth - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
    },
  },
  FREQUENCY: {
    min: 2,
    max: 2,
    fn: (a, ctx) => {
      const data = list([a[0]], ctx);
      const bins = list([a[1]], ctx);
      const order = bins.map((b, i) => [b, i]).sort((p, q) => p[0] - q[0]);
      const counts = new Array<number>(bins.length + 1).fill(0);
      for (const v of data) {
        const at = order.find(([b]) => v <= b);
        counts[at ? at[1] : bins.length]++;
      }
      return new Matrix(counts.length, 1, counts);
    },
  },
  PROB: {
    min: 3,
    max: 4,
    fn: (a, ctx) => {
      const xs = matrixOf(a[0], ctx).data;
      const ps = matrixOf(a[1], ctx).data;
      if (xs.length !== ps.length) throw NA;
      const lo = num(scalarOf(a[2], ctx));
      const hi = given(a, 3) ? num(scalarOf(a[3], ctx)) : lo;
      let total = 0;
      let all = 0;
      xs.forEach((x, i) => {
        const p = ps[i];
        if (typeof p !== 'number' || p <= 0 || p > 1) throw NUM;
        all += p;
        if (typeof x === 'number' && x >= lo && x <= hi) total += p;
      });
      if (Math.abs(all - 1) > 1e-6) throw NUM;
      return total;
    },
  },
  'BINOM.DIST.RANGE': scalar(3, 4, ([n, p, s, s2]) => {
    const trials = Math.trunc(num(n));
    const chance = num(p);
    const lo = Math.trunc(num(s));
    const hi = s2 === undefined || s2 === null ? lo : Math.trunc(num(s2));
    if (trials < 0 || chance < 0 || chance > 1 || lo < 0 || lo > trials || hi < lo || hi > trials) throw NUM;
    return binomialRange(trials, chance, lo, hi);
  }),
  SLOPE: { min: 2, max: 2, fn: (a, ctx) => slopeOf(a[0], a[1], ctx).slope },
  INTERCEPT: { min: 2, max: 2, fn: (a, ctx) => slopeOf(a[0], a[1], ctx).intercept },
  FORECAST: { min: 3, max: 3, fn: (a, ctx) => {
    const { slope, intercept } = slopeOf(a[1], a[2], ctx);
    return intercept + slope * num(scalarOf(a[0], ctx));
  } },
  'FORECAST.LINEAR': { min: 3, max: 3, fn: (a, ctx) => {
    const { slope, intercept } = slopeOf(a[1], a[2], ctx);
    return intercept + slope * num(scalarOf(a[0], ctx));
  } },
  CORREL: { min: 2, max: 2, fn: (a, ctx) => correl(a[0], a[1], ctx) },
  PEARSON: { min: 2, max: 2, fn: (a, ctx) => correl(a[0], a[1], ctx) },
  RSQ: { min: 2, max: 2, fn: (a, ctx) => correl(a[0], a[1], ctx) ** 2 },
  STEYX: { min: 2, max: 2, fn: (a, ctx) => {
    const l = line(a[0], a[1], ctx);
    if (l.n < 3 || l.sxx === 0) throw DIV0;
    return Math.sqrt((l.syy - l.sxy ** 2 / l.sxx) / (l.n - 2));
  } },
  'COVARIANCE.P': { min: 2, max: 2, fn: (a, ctx) => {
    const l = line(a[0], a[1], ctx);
    return l.sxy / l.n;
  } },
  COVAR: { min: 2, max: 2, fn: (a, ctx) => {
    const l = line(a[0], a[1], ctx);
    return l.sxy / l.n;
  } },
  'COVARIANCE.S': { min: 2, max: 2, fn: (a, ctx) => {
    const l = line(a[0], a[1], ctx);
    if (l.n < 2) throw DIV0;
    return l.sxy / (l.n - 1);
  } },

  // The error function and the normal distribution
  ERF: scalar(1, 2, ([lo, hi]) => (hi === undefined || hi === null ? erf(num(lo)) : erf(num(hi)) - erf(num(lo)))),
  'ERF.PRECISE': scalar(1, 1, ([x]) => erf(num(x))),
  ERFC: scalar(1, 1, ([x]) => erf(num(x), true)),
  'ERFC.PRECISE': scalar(1, 1, ([x]) => erf(num(x), true)),
  'NORM.S.DIST': scalar(2, 2, ([z, c]) => (bool(c) ? normalCdf(num(z)) : normalPdf(num(z)))),
  NORMSDIST: scalar(1, 1, ([z]) => normalCdf(num(z))),
  'NORM.DIST': scalar(4, 4, ([x, m, sd, c]) => {
    const s = num(sd);
    if (s <= 0) throw NUM;
    const z = (num(x) - num(m)) / s;
    return bool(c) ? normalCdf(z) : normalPdf(z) / s;
  }),
  NORMDIST: scalar(4, 4, ([x, m, sd, c]) => {
    const s = num(sd);
    if (s <= 0) throw NUM;
    const z = (num(x) - num(m)) / s;
    return bool(c) ? normalCdf(z) : normalPdf(z) / s;
  }),
  'NORM.S.INV': scalar(1, 1, ([p]) => normalInverse(num(p))),
  NORMSINV: scalar(1, 1, ([p]) => normalInverse(num(p))),
  'NORM.INV': scalar(3, 3, ([p, m, sd]) => {
    if (num(sd) <= 0) throw NUM;
    return num(m) + num(sd) * normalInverse(num(p));
  }),
  NORMINV: scalar(3, 3, ([p, m, sd]) => {
    if (num(sd) <= 0) throw NUM;
    return num(m) + num(sd) * normalInverse(num(p));
  }),
  STANDARDIZE: scalar(3, 3, ([x, m, sd]) => {
    if (num(sd) <= 0) throw NUM;
    return (num(x) - num(m)) / num(sd);
  }),

  // Matrices
  MUNIT: scalar(1, 1, ([n]) => {
    const k = Math.trunc(num(n));
    if (k < 1) throw VALUE;
    sized(k, k);
    return Matrix.build(k, k, (r, c) => (r === c ? 1 : 0));
  }),
  MDETERM: { min: 1, max: 1, fn: ([m], ctx) => eliminate(square(m, ctx)).det },
  MINVERSE: {
    min: 1,
    max: 1,
    fn: ([m], ctx) => {
      const inverse = eliminate(square(m, ctx)).inverse;
      if (!inverse) throw NUM;
      return Matrix.of(inverse);
    },
  },
  MMULT: {
    min: 2,
    max: 2,
    fn: ([x, y], ctx) => {
      const a = numberMatrix(x, ctx);
      const b = numberMatrix(y, ctx);
      if (a[0].length !== b.length) throw VALUE;
      sized(a.length, b[0].length);
      return Matrix.build(a.length, b[0].length, (r, c) => a[r].reduce((s, v, k) => s + v * b[k][c], 0));
    },
  },

  // Bonds, accrued interest, discounted securities and depreciation
  COUPDAYBS: couponFn((b, ...a) => b.coupdaybs(...a)),
  COUPDAYS: couponFn((b, ...a) => b.coupdays(...a)),
  COUPDAYSNC: couponFn((b, ...a) => b.coupdaysnc(...a)),
  COUPNCD: couponFn((b, ...a) => b.coupncd(...a)),
  COUPNUM: couponFn((b, ...a) => b.coupnum(...a)),
  COUPPCD: couponFn((b, ...a) => b.couppcd(...a)),
  PRICE: bond(6, [0, 1], { 6: 0 }, (b, [s, m, rate, yld, red, f, basis]) => b.price(s, m, rate, yld, red, f, basis)),
  YIELD: bond(6, [0, 1], { 6: 0 }, (b, [s, m, rate, pr, red, f, basis]) => b.yield(s, m, rate, pr, red, f, basis)),
  DURATION: bond(5, [0, 1], { 5: 0 }, (b, [s, m, c, yld, f, basis]) => b.duration(s, m, c, yld, f, basis)),
  MDURATION: bond(5, [0, 1], { 5: 0 }, (b, [s, m, c, yld, f, basis]) => b.duration(s, m, c, yld, f, basis, true)),
  ACCRINT: bond(6, [0, 1, 2], { 4: 1000, 6: 0, 7: true }, (b, [i, first, s, rate, par, f, basis], flags) => b.accrint(i, first, s, rate, par, f, basis, flags[7])),
  ACCRINTM: bond(3, [0, 1], { 3: 1000, 4: 0 }, (b, [i, s, rate, par, basis]) => b.accrintm(i, s, rate, par, basis)),
  DISC: discountedFn((b, ...a) => b.disc(...a)),
  INTRATE: discountedFn((b, ...a) => b.intrate(...a)),
  RECEIVED: discountedFn((b, ...a) => b.received(...a)),
  PRICEDISC: discountedFn((b, ...a) => b.pricedisc(...a)),
  YIELDDISC: discountedFn((b, ...a) => b.yielddisc(...a)),
  PRICEMAT: bond(5, [0, 1, 2], { 5: 0 }, (b, [s, m, i, rate, yld, basis]) => b.pricemat(s, m, i, rate, yld, basis)),
  YIELDMAT: bond(5, [0, 1, 2], { 5: 0 }, (b, [s, m, i, rate, pr, basis]) => b.yieldmat(s, m, i, rate, pr, basis)),
  VDB: bond(5, [], { 5: 2, 6: false }, (_b, [cost, salvage, life, start, end, factor], flags) => vdb(cost, salvage, life, start, end, factor, flags[6])),

  // Least-squares regression
  TREND: { min: 1, max: 4, fn: (a, ctx) => projection(a, ctx, false) },
  GROWTH: { min: 1, max: 4, fn: (a, ctx) => projection(a, ctx, true) },
  LINEST: { min: 1, max: 4, fn: (a, ctx) => estimation(a, ctx, false) },
  LOGEST: { min: 1, max: 4, fn: (a, ctx) => estimation(a, ctx, true) },

  // Loans and depreciation
  IPMT: scalar(4, 6, ([r, per, n, pv, fv, t]) => interestIn(num(r), num(per), num(n), num(pv), optional(fv, 0), optional(t, 0) ? 1 : 0)),
  PPMT: scalar(4, 6, ([r, per, n, pv, fv, t]) => {
    const type = optional(t, 0) ? 1 : 0;
    return payment(num(r), num(n), num(pv), optional(fv, 0), type) - interestIn(num(r), num(per), num(n), num(pv), optional(fv, 0), type);
  }),
  CUMIPMT: scalar(6, 6, (a) => cumulative(a, false)),
  CUMPRINC: scalar(6, 6, (a) => cumulative(a, true)),
  ISPMT: scalar(4, 4, ([r, per, n, pv]) => {
    if (num(n) === 0) throw DIV0;
    return num(pv) * num(r) * (num(per) / num(n) - 1);
  }),
  SLN: scalar(3, 3, ([c, s, l]) => {
    if (num(l) === 0) throw DIV0;
    return (num(c) - num(s)) / num(l);
  }),
  SYD: scalar(4, 4, ([c, s, l, p]) => {
    const life = num(l);
    const per = num(p);
    if (life <= 0 || per <= 0 || per > life) throw NUM;
    return ((num(c) - num(s)) * (life - per + 1) * 2) / (life * (life + 1));
  }),
  DDB: scalar(4, 5, ([c, s, l, p, f]) => {
    const [cost, salvage, life, period, factor] = [num(c), num(s), num(l), num(p), optional(f, 2)];
    if (cost < 0 || salvage < 0 || life <= 0 || period <= 0 || period > life || factor <= 0) throw NUM;
    return ddb(cost, salvage, life, period, factor);
  }),
  DB: scalar(4, 5, ([c, s, l, p, mo]) => {
    const [cost, salvage, life, period] = [num(c), num(s), num(l), Math.trunc(num(p))];
    const months = Math.trunc(optional(mo, 12));
    if (cost < 0 || salvage < 0 || life <= 0 || months < 1 || months > 12 || period < 1 || period > life + (months < 12 ? 1 : 0)) throw NUM;
    if (cost === 0) return 0;
    const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
    let total = (cost * rate * months) / 12;
    if (period === 1) return total;
    for (let k = 2; k < period; k++) total += (cost - total) * rate;
    return period === life + 1 ? ((cost - total) * rate * (12 - months)) / 12 : (cost - total) * rate;
  }),
  // Cash flows and rates
  XNPV: {
    min: 3,
    max: 3,
    fn: (a, ctx) => {
      const r = num(scalarOf(a[0], ctx));
      if (r <= -1) throw NUM;
      return xnpv(r, datedFlows(a.slice(1), ctx));
    },
  },
  XIRR: {
    min: 2,
    max: 3,
    fn: (a, ctx) => {
      const flows = datedFlows(a, ctx);
      if (!flows.values.some((v) => v > 0) || !flows.values.some((v) => v < 0)) throw NUM;
      return solve((r) => xnpv(r, flows), optNum(a, 2, ctx, 0.1));
    },
  },
  MIRR: {
    min: 3,
    max: 3,
    fn: (a, ctx) => {
      const flows = list([a[0]], ctx);
      const finance = num(scalarOf(a[1], ctx));
      const reinvest = num(scalarOf(a[2], ctx));
      const n = flows.length;
      let negative = 0;
      let positive = 0;
      flows.forEach((v, i) => {
        if (v < 0) negative += v / (1 + finance) ** i;
        else positive += v * (1 + reinvest) ** (n - 1 - i);
      });
      if (negative === 0 || positive === 0) throw DIV0;
      return (positive / -negative) ** (1 / (n - 1)) - 1;
    },
  },
  EFFECT: scalar(2, 2, ([r, n]) => {
    const rate = num(r);
    const periods = Math.trunc(num(n));
    if (rate <= 0 || periods < 1) throw NUM;
    return (1 + rate / periods) ** periods - 1;
  }),
  NOMINAL: scalar(2, 2, ([r, n]) => {
    const rate = num(r);
    const periods = Math.trunc(num(n));
    if (rate <= 0 || periods < 1) throw NUM;
    return periods * ((1 + rate) ** (1 / periods) - 1);
  }),
  RRI: scalar(3, 3, ([n, pv, fv]) => {
    if (num(n) <= 0 || num(pv) === 0) throw NUM;
    const r = (num(fv) / num(pv)) ** (1 / num(n)) - 1;
    if (!Number.isFinite(r)) throw NUM;
    return r;
  }),
  PDURATION: scalar(3, 3, ([r, pv, fv]) => {
    if (num(r) <= 0 || num(pv) <= 0 || num(fv) <= 0) throw NUM;
    return (Math.log(num(fv)) - Math.log(num(pv))) / Math.log(1 + num(r));
  }),
  FVSCHEDULE: {
    min: 2,
    max: 2,
    fn: (a, ctx) => {
      let v = num(scalarOf(a[0], ctx));
      for (const r of matrixOf(a[1], ctx).data) {
        if (isError(r)) throw r;
        v *= 1 + (r === null ? 0 : num(r));
      }
      return v;
    },
  },
  DOLLARDE: scalar(2, 2, ([d, f]) => {
    const fraction = Math.trunc(num(f));
    if (fraction < 0) throw NUM;
    if (fraction === 0) throw DIV0;
    const x = num(d);
    const whole = Math.trunc(x);
    return whole + ((x - whole) * 10 ** Math.ceil(Math.log10(fraction))) / fraction;
  }),
  DOLLARFR: scalar(2, 2, ([d, f]) => {
    const fraction = Math.trunc(num(f));
    if (fraction < 0) throw NUM;
    if (fraction === 0) throw DIV0;
    const x = num(d);
    const whole = Math.trunc(x);
    return whole + ((x - whole) * fraction) / 10 ** Math.ceil(Math.log10(fraction));
  }),

  // Weeks and working days
  WEEKNUM: scalar(1, 2, ([v, t], ctx) => {
    const serial = dateArg(v, ctx);
    const type = Math.trunc(optional(t, 1));
    if (type === 21) return isoWeek(serial, ctx);
    const starts: Record<number, number> = { 1: 0, 17: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6 };
    const start = starts[type];
    if (start === undefined) throw NUM;
    const d = jsDate(serial, ctx);
    const jan1 = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    const offset = (jan1.getUTCDay() - start + 7) % 7;
    const dayOfYear = Math.round((d.getTime() - jan1.getTime()) / DAY_MS);
    return Math.floor((dayOfYear + offset) / 7) + 1;
  }),
  ISOWEEKNUM: scalar(1, 1, ([v], ctx) => isoWeek(dateArg(v, ctx), ctx)),
  DAYS360: scalar(2, 3, ([s, e, m], ctx) => days360(Math.floor(dateArg(s, ctx)), Math.floor(dateArg(e, ctx)), m !== undefined && m !== null && bool(m), ctx)),
  'NETWORKDAYS.INTL': {
    min: 2,
    max: 4,
    fn: (a, ctx) => {
      const start = Math.floor(dateArg(scalarOf(a[0], ctx), ctx));
      const end = Math.floor(dateArg(scalarOf(a[1], ctx), ctx));
      const weekend = weekendOf(a[2], ctx);
      const holidays = holidaysOf(a[3], ctx);
      const step = end >= start ? 1 : -1;
      let n = 0;
      for (let d = start; step > 0 ? d <= end : d >= end; d += step) if (workday(d, weekend, holidays, ctx)) n += step;
      return n;
    },
  },
  'WORKDAY.INTL': {
    min: 2,
    max: 4,
    fn: (a, ctx) => {
      let d = Math.floor(dateArg(scalarOf(a[0], ctx), ctx));
      let left = Math.trunc(num(scalarOf(a[1], ctx)));
      const weekend = weekendOf(a[2], ctx);
      if (weekend.every(Boolean)) throw VALUE;
      const holidays = holidaysOf(a[3], ctx);
      const step = left >= 0 ? 1 : -1;
      while (left !== 0) {
        d += step;
        if (workday(d, weekend, holidays, ctx)) left -= step;
      }
      return d;
    },
  },
  TIMEVALUE: scalar(1, 1, ([t]) => {
    const m = /^\s*(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([ap]m)?\s*$/i.exec(text(t));
    if (!m) throw VALUE;
    let h = Number(m[1]);
    const ampm = m[4]?.toLowerCase();
    if (ampm) {
      if (h < 1 || h > 12) throw VALUE;
      h = (h % 12) + (ampm === 'pm' ? 12 : 0);
    }
    const v = (h * 3600 + Number(m[2]) * 60 + Number(m[3] ?? 0)) / 86400;
    return v - Math.floor(v);
  }),

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
