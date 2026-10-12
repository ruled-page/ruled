import { format } from 'numfmt';
import { parseA1Ref, parseR1C1Ref } from '@borgar/fx';
import { BondError, Bonds, Calendar, ddb, validBasis, vdb, yearFrac } from '../core/bondMath';
import { excelRound, numberText } from '../core/excelNumbers';
import { eliminate } from '../core/matrixMath';
import { estimate, prepare, project, RegressionError } from '../core/regressionMath';
import { colToLetters, MAX_COLS, MAX_ROWS, quoteSheetName } from '../core/refs';
import { MAX_ARRAY_CELLS } from '../core/sizeLimits';
import { hasWildcards, matchesWhole, parsePattern, searchIn } from '../core/wildcard';
import { scalarOf, valuesOf, type Ctx } from './operands';
import { Areas, bool, CALC, compare, DIV0, err, ERROR_CODES, isError, Lambda, Matrix, NA, NUM, num, REF, Ref, round15, text, VALUE, type Operand, type Scalar } from './values';

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
    else if (a instanceof Areas) for (const r of a.refs) ctx.host.each(r, fromCells, how.skipSubtotals);
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
  4: (xs) => (xs.length ? xs.reduce((m, x) => (x > m ? x : m), -Infinity) : 0),
  5: (xs) => (xs.length ? xs.reduce((m, x) => (x < m ? x : m), Infinity) : 0),
  6: (xs) => (xs.length ? xs.reduce((p, x) => p * x, 1) : 0),
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
    if (a instanceof Ref || a instanceof Areas) for (const r of a instanceof Areas ? a.refs : [a]) ctx.host.each(r, (v) => void (v !== null && n++), skipSubtotals);
    else if (a instanceof Matrix) for (const v of a.data) n += v === null ? 0 : 1;
    else n++;
  }
  return n;
}

// -----------------------------------------------------------------------------
// Criteria, as COUNTIF and SUMIFS take them

/**
 * Whether a value meets a criterion such as 5, ">=10", "<>", "North" or "N*",
 * by Excel's rules: a criterion that is an empty cell is 0; "=5" (or 5)
 * matches the text "5" as well as the number, though "<>5" passes over only
 * the number; "<>" alone matches anything not empty, a formula's "" too;
 * and a comparison with nothing to compare with matches nothing.
 */
function criterion(c: Scalar): (v: Scalar) => boolean {
  if (c === null) return criterion(0);
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
    if (asNumber) return typeof v === 'number' ? round15(v) === round15(n!) : typeof v === 'string' && v.trim() !== '' && Number(v) === n;
    if (asBool !== undefined) return v === asBool;
    if (typeof v !== 'string') return false;
    return hasWildcards(rest) ? matchesWhole(parsePattern(rest.toLowerCase()), v.toLowerCase()) : v.toLowerCase() === rest.toLowerCase();
  };
  // Text that spells an error means the error.
  const code = rest.toUpperCase();
  if (ERROR_CODES.has(code) && (op === '' || op === '=' || op === '<>')) {
    const same = (v: Scalar) => isError(v) && v.error === code;
    return op === '<>' ? (v) => !same(v) : same;
  }
  if (op === '' || op === '=') return equal;
  if (op === '<>') {
    if (rest === '') return (v) => v !== null;
    if (asNumber) return (v) => !(typeof v === 'number' && round15(v) === round15(n!));
    return (v) => !equal(v);
  }
  if (rest === '') return () => false;
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
  if (op instanceof Areas) throw VALUE;
  const v = scalarOf(op, ctx);
  return { rows: 1, cols: 1, at: () => v };
}

/**
 * The offsets of the cells meeting every (range, criterion) pair, for
 * COUNTIFS and its kind; ranges must be the same size. Whole rows and
 * columns are read only as far as the sheet is used: `blanks` counts the
 * empty cells past that, when empty cells meet every criterion.
 */
function matching(pairs: Operand[], ctx: Ctx): { hits: [number, number][]; rows: number; cols: number; blanks: number } {
  const ranges = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const op = pairs[i];
    const g = grid(op, ctx);
    ranges.push({ g, full: op instanceof Ref ? { rows: op.rows, cols: op.cols } : { rows: g.rows, cols: g.cols }, test: criterion(scalarOf(pairs[i + 1], ctx)) });
  }
  const { rows, cols } = ranges[0].full;
  if (ranges.some((r) => r.full.rows !== rows || r.full.cols !== cols)) throw VALUE;
  const usedRows = Math.min(rows, Math.max(...ranges.map((r) => r.g.rows)));
  const usedCols = Math.min(cols, Math.max(...ranges.map((r) => r.g.cols)));
  const hits: [number, number][] = [];
  for (let r = 0; r < usedRows; r++) {
    for (let c = 0; c < usedCols; c++) if (ranges.every((x) => x.test(x.g.at(r, c)))) hits.push([r, c]);
  }
  const past = rows * cols - usedRows * usedCols;
  return { hits, rows, cols, blanks: past > 0 && ranges.every((x) => x.test(null)) ? past : 0 };
}

const count = (m: { hits: unknown[]; blanks: number }) => m.hits.length + m.blanks;

function ifs(kind: 'sum' | 'avg' | 'min' | 'max', target: Operand, pairs: Operand[], ctx: Ctx): number {
  const t = grid(target, ctx);
  const { hits, rows, cols } = matching(pairs, ctx);
  const size = target instanceof Ref ? target : t;
  if (size.rows !== rows || size.cols !== cols) throw VALUE;
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
  return xs.reduce((m, x) => (kind === 'min' ? (x < m ? x : m) : x > m ? x : m));
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
  if (op instanceof Areas) throw VALUE;
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
  // XLOOKUP's -1 and 1: an exact match, else the next smaller or larger; searching from the last, the last of equals.
  let best = -1;
  for (let k = 0; k < values.length; k++) {
    const i = last ? values.length - 1 - k : k;
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

/**
 * XMATCH's and XLOOKUP's binary search (search mode 2, or -2 for a list
 * sorted descending): an exact match (mode 0), else the next smaller (-1)
 * or larger (1).
 */
function binaryMatch(values: Scalar[], target: Scalar, mode: number, descending: boolean): number {
  // Ascending, the last value not above it; descending, the last not below.
  const at = approximate(values, target, descending);
  if (at >= 0 && compare(values[at], target) === 0) return at;
  if (mode === 0 || mode === 2) return -1;
  if (descending ? mode > 0 : mode < 0) return at;
  return at + 1 < values.length ? at + 1 : -1;
}

/** A key for a value, for finding exact matches by it: numbers to 15 digits, text without regard to case. */
function matchKey(v: Scalar): string | undefined {
  if (typeof v === 'number') return `n${round15(v)}`;
  if (typeof v === 'string') return `s${v.toLowerCase()}`;
  if (typeof v === 'boolean') return `b${v}`;
  return undefined;
}

/**
 * How to find a value among others by a match mode and search mode: for
 * many values looked up at once, exact matches go through an index built
 * once, so the list is read once rather than once per value.
 */
function finder(values: Scalar[], mode: number, search: number, many: boolean): (target: Scalar) => number {
  if (Math.abs(search) === 2) return (target) => binaryMatch(values, target, mode, search < 0);
  if (!many || (mode !== 0 && mode !== 2)) return (target) => lookupIn(values, target, mode, search < 0);
  const index = new Map<string, number>();
  for (let k = 0; k < values.length; k++) {
    const i = search < 0 ? values.length - 1 - k : k;
    const key = matchKey(values[i]);
    if (key !== undefined && !index.has(key)) index.set(key, i);
  }
  return (target) => {
    if (mode === 2 && typeof target === 'string' && hasWildcards(target)) return lookupIn(values, target, mode, search < 0);
    const key = matchKey(target);
    return key === undefined ? lookupIn(values, target, mode, search < 0) : (index.get(key) ?? -1);
  };
}

/**
 * Runs `f` on the value given or, given a range or array, on each of its
 * values, as Excel lifts a function over an argument that takes one value.
 * A formula from before dynamic arrays meets a range at its row or column.
 */
function eachValue(op: Operand, ctx: Ctx, f: (v: Scalar) => Scalar): Operand {
  const many = (op instanceof Ref && !op.single && !ctx.legacy) || (op instanceof Matrix && op.data.length > 1);
  if (!many) return f(scalarOf(op, ctx));
  const m = matrixOf(op, ctx);
  return Matrix.build(m.rows, m.cols, (r, c) => {
    try {
      return f(m.get(r, c));
    } catch (e) {
      if (isError(e)) return e;
      throw e;
    }
  });
}

/** A function lifted over its first argument (see eachValue); a result of more than one value gives its first. */
const liftFirst = (spec: FunctionSpec): FunctionSpec => ({
  ...spec,
  fn: (a, ctx) =>
    eachValue(a[0], ctx, (v) => {
      const out = spec.fn!([v, ...a.slice(1)], ctx);
      if (isError(out)) throw out;
      return out instanceof Matrix ? out.get(0, 0) : out instanceof Ref ? scalarOf(out, ctx) : (out as Scalar);
    }),
});

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
  // Of a union, the area the fourth argument picks (the first by default).
  const which = given(args, 3) ? Math.trunc(num(scalarOf(args[3], ctx))) : 1;
  const areas = args[0] instanceof Areas ? args[0].refs : [args[0]];
  if (which < 1 || which > areas.length) throw REF;
  const area = areas[which - 1];
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
  if (Math.abs(search) === 2) at = binaryMatch(look.values, target, mode, search < 0);
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
      // Numbers, text, TRUE and FALSE, then errors (reversed when descending); blanks last either way.
      const xb = x === null;
      const yb = y === null;
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
  // Excel's dates run to the end of 31 December 9999.
  if (n < 0 || n >= serialOf(9999, 12, 31, ctx.host.date1904) + 1) throw NUM;
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

/**
 * The date `left` working days from `start`, a day at a time: #NUM! once it
 * passes the first or last date there is (31 December 9999), as Excel
 * gives, and at once for more working days than that span holds.
 */
function stepWorkdays(start: number, left: number, isWorkday: (d: number) => boolean, ctx: Ctx): number {
  const first = ctx.host.date1904 ? 0 : 1;
  const last = serialOf(9999, 12, 31, ctx.host.date1904);
  if (Math.abs(left) > last) throw NUM;
  const step = left >= 0 ? 1 : -1;
  let d = start;
  while (left !== 0) {
    d += step;
    if (d < first || d > last) throw NUM;
    if (isWorkday(d)) left -= step;
  }
  return d;
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
  if (n < 1 || n > (rows ? MAX_COLS : MAX_ROWS)) throw NUM;
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
  const at = findNth(t, textsOf(a[1], ctx), n, flag(a, 3, ctx), flag(a, 4, ctx));
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

/** A number at or above 0 to 15 significant digits, written out without an exponent: 1E+21 as 1000000000000000000000. */
function writtenOut(n: number): string {
  if (n === 0) return '0';
  const [mantissa, e] = n.toExponential(14).split('e');
  const digits = mantissa.replace('.', '').replace(/0+$/, '');
  const exp = Number(e);
  if (exp >= digits.length - 1) return digits + '0'.repeat(exp - digits.length + 1);
  if (exp >= 0) return `${digits.slice(0, exp + 1)}.${digits.slice(exp + 1)}`;
  return `0.${'0'.repeat(-exp - 1)}${digits}`;
}

/** An argument that is 0 or 1 (FALSE or TRUE), as TEXTSPLIT's and VALUETOTEXT's switches are; anything else is #VALUE!. */
function flag(a: Operand[], i: number, ctx: Ctx): boolean {
  if (!given(a, i)) return false;
  const n = num(scalarOf(a[i], ctx));
  if (n !== 0 && n !== 1) throw VALUE;
  return n === 1;
}

/** The longest text a cell holds; longer is #VALUE!. */
const MAX_TEXT = 32767;

function cellText(t: string): string {
  if (t.length > MAX_TEXT) throw VALUE;
  return t;
}

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

/** ROWS or COLUMNS: of a range or array; a union is #REF!, an error passes on. */
function lineCount(r: Operand, ctx: Ctx, rows: boolean): number {
  if (r instanceof Areas) throw REF;
  if (isError(r)) throw r;
  const m = r instanceof Ref ? r : matrixOf(r, ctx);
  return rows ? m.rows : m.cols;
}

/** FORECAST: the line through the pairs, at x, or at each x of an array. */
function forecast(a: Operand[], ctx: Ctx): Operand {
  const { slope, intercept } = slopeOf(a[1], a[2], ctx);
  return eachValue(a[0], ctx, (x) => intercept + slope * num(x));
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

/**
 * The most steps a function takes one at a time (periods, terms of a sum):
 * beyond it, #NUM!, rather than a calculation that doesn't finish.
 */
const MAX_STEPS = 10_000_000;

/** Up to this many trials the binomial terms are summed exactly; beyond, they come from the gamma and beta functions. */
const SUMMED_TRIALS = 10_000;

/** P(X <= k) for X binomial with n trials and success chance p, from the incomplete beta function. */
const binomialCdf = (n: number, p: number, k: number) => (k < 0 ? 0 : k >= n ? 1 : betaI(1 - p, p, n - k, k + 1));

/** P(lo <= X <= hi) for X binomial with n trials and success chance p. */
function binomialRange(n: number, p: number, lo: number, hi: number): number {
  // Past whole numbers a double holds exactly, as Excel's #NUM! for 1E+30 trials.
  if (n > 2 ** 53) throw NUM;
  if (p === 0 || p === 1) {
    const certain = p === 0 ? 0 : n;
    return certain >= lo && certain <= hi ? 1 : 0;
  }
  if (n > SUMMED_TRIALS) {
    if (lo === hi) return Math.exp(logChoose(n, lo) + lo * Math.log(p) + (n - lo) * Math.log1p(-p));
    return Math.max(0, binomialCdf(n, p, hi) - binomialCdf(n, p, lo - 1));
  }
  let lnChoose = logFactorial(n) - logFactorial(lo) - logFactorial(n - lo);
  let total = 0;
  for (let k = lo; k <= hi; k++) {
    total += Math.exp(lnChoose + k * Math.log(p) + (n - k) * Math.log1p(-p));
    lnChoose += Math.log(n - k) - Math.log(k + 1);
  }
  return total;
}

// -----------------------------------------------------------------------------
// Probability distributions, from the gamma and beta functions

/**
 * Values as AVERAGEA and its kind count them: in ranges TRUE is 1, FALSE and
 * text 0; typed values read as numbers. Excel's AVERAGEA counts an array's
 * values as a range's, the others only an array's numbers.
 */
function listA(args: Operand[], ctx: Ctx, arraysAsRanges = false): number[] {
  const out: number[] = [];
  const fromCells = (v: Scalar) => {
    if (typeof v === 'number') out.push(v);
    else if (typeof v === 'boolean') out.push(v ? 1 : 0);
    else if (typeof v === 'string') out.push(0);
    else if (isError(v)) throw v;
  };
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof Ref) ctx.host.each(a, fromCells);
    else if (a instanceof Areas) for (const r of a.refs) ctx.host.each(r, fromCells);
    else if (a instanceof Matrix) a.data.forEach(arraysAsRanges ? fromCells : (v) => (isError(v) ? fromCells(v) : typeof v === 'number' && out.push(v)));
    else if (a instanceof Lambda) throw VALUE;
    else if (isError(a)) throw a;
    else out.push(num(a));
  }
  return out;
}

/** The third moment about the mean, over `s` cubed. */
function cubes(xs: number[], s: number): number {
  if (s === 0) throw DIV0;
  const mean = sum(xs) / xs.length;
  return sum(xs.map((x) => ((x - mean) / s) ** 3));
}

/** A distribution's result: an overflow or a result that isn't a number is #NUM!. */
const dist = (min: number, max: number, f: (a: Scalar[]) => number): FunctionSpec =>
  scalar(min, max, (a) => {
    const r = f(a);
    if (!Number.isFinite(r)) throw NUM;
    return r;
  });

/** Degrees of freedom: cut to a whole number, at least 1 and below 10^10. */
function degrees(v: Scalar): number {
  const d = Math.trunc(num(v));
  if (d < 1 || d >= 1e10) throw NUM;
  return d;
}

/** A probability from 0 to 1, or with `open`, strictly between them. */
function probability(v: Scalar, open = false): number {
  const p = num(v);
  if (open ? p <= 0 || p >= 1 : p < 0 || p > 1) throw NUM;
  return p;
}

function positive(v: Scalar): number {
  const x = num(v);
  if (x <= 0) throw NUM;
  return x;
}

/** Godfrey's coefficients for Lanczos' approximation with g = 607/128. */
const LANCZOS = [
  0.9999999999999971, 57.15623566586292, -59.59796035547549, 14.136097974741746, -0.4919138160976202, 0.00003399464998481189,
  0.00004652362892704858, -0.00009837447530487956, 0.0001580887032249125, -0.00021026444172410488, 0.00021743961811521265,
  -0.0001643181065367639, 0.00008441822398385275, -0.000026190838401581408, 0.0000036899182659531625,
];

/** ln Γ(x) for x > 0, to about 15 digits. */
function logGamma(x: number): number {
  if (x < 0.5) return Math.log(Math.PI / Math.sin(Math.PI * x)) - logGamma(1 - x);
  const z = x - 1;
  let a = LANCZOS[0];
  for (let i = 1; i < LANCZOS.length; i++) a += LANCZOS[i] / (z + i);
  const t = z + 607 / 128 + 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(a);
}

/** Γ(x): exact for whole numbers, #NUM! at 0 and the negative whole numbers. */
function gamma(x: number): number {
  if (x <= 0 && Number.isInteger(x)) throw NUM;
  if (Number.isInteger(x) && x <= 171) {
    let f = 1;
    for (let i = 2; i < x; i++) f *= i;
    return f;
  }
  return x < 0.5 ? Math.PI / (Math.sin(Math.PI * x) * gamma(1 - x)) : Math.exp(logGamma(x));
}

const logBeta = (a: number, b: number) => logGamma(a) + logGamma(b) - logGamma(a + b);
const logChoose = (n: number, k: number) => logGamma(n + 1) - logGamma(k + 1) - logGamma(n - k + 1);

const TINY = 1e-300;

/** The regularised incomplete gamma function P(a, x), or with `upper`, Q(a, x) = 1 − P(a, x), each worked out directly. */
function gammaP(a: number, x: number, upper = false): number {
  if (x <= 0) return upper ? 1 : 0;
  const front = Math.exp(a * Math.log(x) - x - logGamma(a));
  if (x < a + 1) {
    let term = 1 / a;
    let total = term;
    for (let n = 1; n < 100000 && Math.abs(term) > Math.abs(total) * 1e-17; n++) {
      term *= x / (a + n);
      total += term;
    }
    const p = Math.min(1, total * front);
    return upper ? 1 - p : p;
  }
  // Lentz's continued fraction for Q.
  let b = x + 1 - a;
  let c = 1 / TINY;
  let d = 1 / b;
  let h = d;
  for (let i = 1; i < 100000; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < TINY) d = TINY;
    c = b + an / c;
    if (Math.abs(c) < TINY) c = TINY;
    d = 1 / d;
    const step = d * c;
    h *= step;
    if (Math.abs(step - 1) < 1e-16) break;
  }
  const q = front * h;
  return upper ? q : 1 - q;
}

/** The continued fraction for the incomplete beta function, by Lentz's method. */
function betaFraction(x: number, a: number, b: number): number {
  let c = 1;
  let d = 1 - ((a + b) * x) / (a + 1);
  if (Math.abs(d) < TINY) d = TINY;
  d = 1 / d;
  let h = d;
  for (let m = 1; m < 100000; m++) {
    const m2 = 2 * m;
    for (const aa of [(m * (b - m) * x) / ((a - 1 + m2) * (a + m2)), (-(a + m) * (a + b + m) * x) / ((a + m2) * (a + 1 + m2))]) {
      d = 1 + aa * d;
      if (Math.abs(d) < TINY) d = TINY;
      c = 1 + aa / c;
      if (Math.abs(c) < TINY) c = TINY;
      d = 1 / d;
      h *= d * c;
    }
    if (Math.abs(d * c - 1) < 1e-15) break;
  }
  return h;
}

/**
 * The regularised incomplete beta function I_x(a, b), or with `upper`,
 * 1 − I_x(a, b); `y` is 1 − x, passed in so a tail keeps its precision.
 */
function betaI(x: number, y: number, a: number, b: number, upper = false): number {
  if (x <= 0) return upper ? 1 : 0;
  if (y <= 0) return upper ? 0 : 1;
  const front = Math.exp(a * Math.log(x) + b * Math.log(y) - logBeta(a, b));
  if (x < (a + 1) / (a + b + 2)) {
    const lower = (front * betaFraction(x, a, b)) / a;
    return upper ? 1 - lower : lower;
  }
  const high = (front * betaFraction(y, b, a)) / b;
  return upper ? high : 1 - high;
}

/** Where a monotonic `f` reaches `target`, from x = 0 up: the bracket doubles until it holds the answer, then halves to the last representable step. */
function bisect(f: (x: number) => number, target: number, rising: boolean, hi = 1): number {
  const short = (x: number) => (rising ? f(x) < target : f(x) > target);
  let lo = 0;
  while (short(hi)) {
    lo = hi;
    hi *= 2;
    if (hi > 1e300) throw NUM;
  }
  for (;;) {
    const mid = lo + (hi - lo) / 2;
    if (mid <= lo || mid >= hi) return hi;
    if (short(mid)) lo = mid;
    else hi = mid;
  }
}

/** P(T > t) for Student's t distribution with v degrees of freedom. */
function tUpper(t: number, v: number): number {
  if (t < 0) return 1 - tUpper(-t, v);
  const t2 = t * t;
  if (t2 === Infinity) return 0;
  return 0.5 * betaI(v / (v + t2), t2 / (v + t2), v / 2, 0.5);
}

const tDensity = (t: number, v: number) => Math.exp(logGamma((v + 1) / 2) - logGamma(v / 2) - ((v + 1) / 2) * Math.log1p((t * t) / v)) / Math.sqrt(v * Math.PI);

/** The t with P(T > t) = p. */
const tUpperInverse = (p: number, v: number) => bisect((t) => tUpper(t, v), p, false);

function tInverse(p: number, v: number): number {
  if (p === 0.5) return 0;
  return p < 0.5 ? -tUpperInverse(p, v) : tUpperInverse(1 - p, v);
}

/** P(F ≤ f), or with `upper`, P(F > f), for the F distribution. */
function fDist(f: number, d1: number, d2: number, upper = false): number {
  if (f <= 0) return upper ? 1 : 0;
  const s = d1 * f + d2;
  return betaI((d1 * f) / s, d2 / s, d1 / 2, d2 / 2, upper);
}

function fDensity(f: number, d1: number, d2: number): number {
  if (f === 0) return d1 === 2 ? 1 : d1 > 2 ? 0 : Infinity;
  return Math.exp(0.5 * d1 * Math.log(d1 * f) + 0.5 * d2 * Math.log(d2) - 0.5 * (d1 + d2) * Math.log(d1 * f + d2) - logBeta(d1 / 2, d2 / 2)) / f;
}

function gammaDensity(x: number, a: number, b: number): number {
  if (x === 0) return a === 1 ? 1 / b : a > 1 ? 0 : Infinity;
  return Math.exp((a - 1) * Math.log(x) - x / b - a * Math.log(b) - logGamma(a));
}

/** The inverse of a cumulative distribution on 0 to infinity: 0 at p = 0. */
const inverseFrom0 = (cdf: (x: number) => number, p: number) => (p === 0 ? 0 : bisect(cdf, p, true));

/** The inverse of an upper tail on 0 to infinity: 0 at p = 1. */
const inverseUpper = (tail: (x: number) => number, p: number) => (p === 1 ? 0 : bisect(tail, p, false));

/** BETA.DIST's x, A and B, as a fraction of the way from A to B. */
function betaPlace(x: Scalar, lo: Scalar | undefined, hi: Scalar | undefined): { x: number; width: number } {
  const a = optional(lo, 0);
  const b = optional(hi, 1);
  const v = num(x);
  if (v < a || v > b || a === b) throw NUM;
  return { x: (v - a) / (b - a), width: b - a };
}

function betaShape(alpha: Scalar, beta: Scalar): [number, number] {
  return [positive(alpha), positive(beta)];
}

function betaInverse(p: Scalar, alpha: Scalar, beta: Scalar, lo: Scalar | undefined, hi: Scalar | undefined): number {
  const q = num(p);
  const [a, b] = betaShape(alpha, beta);
  const from = optional(lo, 0);
  const to = optional(hi, 1);
  if (q <= 0 || q > 1 || from >= to) throw NUM;
  return from + (to - from) * bisect((x) => betaI(x, 1 - x, a, b), q, true);
}

/** The smallest number of successes whose cumulative binomial probability reaches alpha. */
function binomialInverse(n: number, p: number, alpha: number): number {
  if (p === 0) return 0;
  if (p === 1) return n;
  if (n > 2 ** 53) throw NUM;
  if (n > SUMMED_TRIALS) {
    // The cumulative probability only rises with k: halve the range it could be in.
    let lo = 0;
    let hi = n;
    while (lo < hi) {
      const mid = Math.floor((lo + hi) / 2);
      if (binomialCdf(n, p, mid) >= alpha * (1 - 1e-14)) hi = mid;
      else lo = mid + 1;
    }
    return lo;
  }
  let lnChoose = 0;
  let total = 0;
  for (let k = 0; k < n; k++) {
    total += Math.exp(lnChoose + k * Math.log(p) + (n - k) * Math.log1p(-p));
    if (total >= alpha * (1 - 1e-14)) return k;
    lnChoose += Math.log(n - k) - Math.log(k + 1);
  }
  return n;
}

function negativeBinomial(f: Scalar, s: Scalar, p: Scalar, cumulative: boolean): number {
  const failures = Math.trunc(num(f));
  const successes = Math.trunc(num(s));
  const chance = probability(p);
  if (failures < 0 || successes < 1) throw NUM;
  if (cumulative) return betaI(chance, 1 - chance, successes, failures + 1);
  return Math.exp(logChoose(failures + successes - 1, failures) + successes * Math.log(chance) + failures * Math.log1p(-chance));
}

function hypergeometric(args: Scalar[], cumulative: boolean): number {
  const [k, n, big, total] = args.slice(0, 4).map((v) => Math.trunc(num(v)));
  if (total <= 0 || n <= 0 || n > total || big <= 0 || big > total) throw NUM;
  const least = Math.max(0, n - total + big);
  if (k < least || (k > Math.min(n, big) && !cumulative)) throw NUM;
  if (k > Math.min(n, big)) return 1;
  const term = (i: number) => Math.exp(logChoose(big, i) + logChoose(total - big, n - i) - logChoose(total, n));
  if (!cumulative) return term(k);
  if (k - least > MAX_STEPS) throw NUM;
  let s = 0;
  for (let i = least; i <= k; i++) s += term(i);
  return Math.min(1, s);
}

/** The two-tailed chance that a ratio of variances this far from 1 happens by chance. */
function fTest(a: Operand[], ctx: Ctx): number {
  const xs = list([a[0]], ctx);
  const ys = list([a[1]], ctx);
  if (xs.length < 2 || ys.length < 2) throw DIV0;
  const vx = variance(xs, true);
  const vy = variance(ys, true);
  if (vx === 0 || vy === 0) throw DIV0;
  const f = vx / vy;
  const d1 = xs.length - 1;
  const d2 = ys.length - 1;
  return 2 * Math.min(fDist(f, d1, d2), fDist(f, d1, d2, true));
}

/** Student's t-test: paired (1), equal variances (2) or unequal (3); one or two tails. */
function tTest(a: Operand[], ctx: Ctx): number {
  const tails = Math.trunc(num(scalarOf(a[2], ctx)));
  const type = Math.trunc(num(scalarOf(a[3], ctx)));
  if ((tails !== 1 && tails !== 2) || type < 1 || type > 3) throw NUM;
  let t: number;
  let df: number;
  if (type === 1) {
    const { y, x } = pairs(a[0], a[1], ctx);
    const d = y.map((v, i) => v - x[i]);
    if (d.length < 2) throw DIV0;
    const s = variance(d, true);
    if (s === 0) throw DIV0;
    t = sum(d) / d.length / Math.sqrt(s / d.length);
    df = d.length - 1;
  } else {
    const xs = list([a[0]], ctx);
    const ys = list([a[1]], ctx);
    const n1 = xs.length;
    const n2 = ys.length;
    if (n1 < 2 || n2 < 2) throw DIV0;
    const v1 = variance(xs, true);
    const v2 = variance(ys, true);
    const diff = sum(xs) / n1 - sum(ys) / n2;
    if (type === 2) {
      df = n1 + n2 - 2;
      const pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / df;
      t = diff / Math.sqrt(pooled * (1 / n1 + 1 / n2));
    } else {
      const e1 = v1 / n1;
      const e2 = v2 / n2;
      t = diff / Math.sqrt(e1 + e2);
      df = (e1 + e2) ** 2 / (e1 ** 2 / (n1 - 1) + e2 ** 2 / (n2 - 1));
    }
    if (!Number.isFinite(t)) throw DIV0;
  }
  return tails * tUpper(Math.abs(t), df);
}

function chiTest(a: Operand[], ctx: Ctx): number {
  const actual = matrixOf(a[0], ctx);
  const expected = matrixOf(a[1], ctx);
  if (actual.rows !== expected.rows || actual.cols !== expected.cols) throw NA;
  let chi = 0;
  for (let i = 0; i < actual.data.length; i++) {
    const o = actual.data[i];
    const e = expected.data[i];
    if (isError(o)) throw o;
    if (isError(e)) throw e;
    if (typeof o !== 'number' || typeof e !== 'number') continue;
    if (e === 0) throw DIV0;
    chi += (o - e) ** 2 / e;
  }
  const df = actual.rows > 1 && actual.cols > 1 ? (actual.rows - 1) * (actual.cols - 1) : actual.data.length - 1;
  if (df < 1) throw NA;
  return gammaP(df / 2, chi / 2, true);
}

function zTest(a: Operand[], ctx: Ctx): number {
  const xs = list([a[0]], ctx);
  if (!xs.length) throw NA;
  const x = num(scalarOf(a[1], ctx));
  const sigma = given(a, 2) ? num(scalarOf(a[2], ctx)) : Math.sqrt(variance(xs, true));
  if (sigma === 0) throw DIV0;
  return normalCdf(-((sum(xs) / xs.length - x) / (sigma / Math.sqrt(xs.length))));
}

function nonNegative(v: Scalar): number {
  const x = num(v);
  if (x < 0) throw NUM;
  return x;
}

function confidence(alpha: Scalar, sd: Scalar, size: Scalar, student: boolean): number {
  const a = probability(alpha, true);
  const s = positive(sd);
  const n = Math.trunc(num(size));
  if (n < 1) throw NUM;
  if (student && n === 1) throw DIV0;
  return ((student ? tUpperInverse(a / 2, n - 1) : -normalInverse(a / 2)) * s) / Math.sqrt(n);
}

function logNormal(x: Scalar, mean: Scalar, sd: Scalar, cumulative: boolean): number {
  const v = positive(x);
  const s = positive(sd);
  const z = (Math.log(v) - num(mean)) / s;
  return cumulative ? normalCdf(z) : normalPdf(z) / (v * s);
}

function twoTailedT(p: Scalar, v: Scalar): number {
  const q = num(p);
  if (q <= 0 || q > 1) throw NUM;
  const df = degrees(v);
  return q === 1 ? 0 : tUpperInverse(q / 2, df);
}

function chiUpperInverse(p: Scalar, v: Scalar): number {
  const k = degrees(v);
  return inverseUpper((x) => gammaP(k / 2, x / 2, true), probability(p));
}

function fUpperInverse(p: Scalar, v1: Scalar, v2: Scalar): number {
  const d1 = degrees(v1);
  const d2 = degrees(v2);
  return inverseUpper((f) => fDist(f, d1, d2, true), probability(p));
}

function gammaDist(x: Scalar, a: Scalar, b: Scalar, cumulative: boolean): number {
  const v = nonNegative(x);
  const alpha = positive(a);
  const beta = positive(b);
  return cumulative ? gammaP(alpha, v / beta) : gammaDensity(v, alpha, beta);
}

function gammaInverse(p: Scalar, a: Scalar, b: Scalar): number {
  const q = probability(p);
  const alpha = positive(a);
  const beta = positive(b);
  return inverseFrom0((x) => gammaP(alpha, x / beta), q);
}

function binomial(s: Scalar, n: Scalar, p: Scalar, cumulative: boolean): number {
  const k = Math.trunc(num(s));
  const trials = Math.trunc(num(n));
  const chance = probability(p);
  if (k < 0 || k > trials) throw NUM;
  return binomialRange(trials, chance, cumulative ? 0 : k, k);
}

function criticalBinomial(n: Scalar, p: Scalar, alpha: Scalar): number {
  const trials = Math.trunc(num(n));
  const chance = probability(p);
  const a = probability(alpha);
  if (trials < 0) throw NUM;
  return binomialInverse(trials, chance, a);
}

function poisson(x: Scalar, m: Scalar, cumulative: boolean): number {
  const k = Math.trunc(num(x));
  const mean = num(m);
  if (k < 0 || mean < 0) throw NUM;
  if (cumulative) return gammaP(k + 1, mean, true);
  if (mean === 0) return k === 0 ? 1 : 0;
  return Math.exp(k * Math.log(mean) - mean - logGamma(k + 1));
}

function exponential(x: Scalar, l: Scalar, cumulative: boolean): number {
  const v = nonNegative(x);
  const rate = positive(l);
  return cumulative ? -Math.expm1(-rate * v) : rate * Math.exp(-rate * v);
}

function weibull(x: Scalar, a: Scalar, b: Scalar, cumulative: boolean): number {
  const v = nonNegative(x);
  const alpha = positive(a);
  const beta = positive(b);
  const u = (v / beta) ** alpha;
  return cumulative ? -Math.expm1(-u) : (alpha / beta) * (v / beta) ** (alpha - 1) * Math.exp(-u);
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
// Bonds, accrued interest, discounted securities and depreciation (src/core/bondMath.ts)

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
// Least-squares regression (src/core/regressionMath.ts)

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
  if (r <= 0 || n <= 0 || pv <= 0 || start < 1 || end < start || end > n || (type !== 0 && type !== 1) || end - start > MAX_STEPS) throw NUM;
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

// -----------------------------------------------------------------------------
// Counting, number bases, Roman numerals and trigonometry

function factorial(n: number): number {
  let f = 1;
  for (let i = 2; i <= n; i++) f *= i;
  return f;
}

/** n choose k, exact for the whole numbers Excel gives; #NUM! once too large for a number. */
function choose(n: number, k: number): number {
  let r = 1;
  for (let i = 1; i <= Math.min(k, n - k); i++) {
    r = (r * (n - k + i)) / i;
    if (!Number.isFinite(r)) throw NUM;
  }
  return Math.round(r);
}

const gcd2 = (a: number, b: number): number => (b === 0 ? a : gcd2(b, a % b));

/** Whole, non-negative numbers from the arguments, for GCD, LCM and MULTINOMIAL. */
function wholes(a: Operand[], ctx: Ctx): number[] {
  const out = list(a, ctx).map(Math.trunc);
  if (out.some((n) => n < 0)) throw NUM;
  return out;
}

/** SUMX2MY2 and its kind: a sum over pairs of numbers from two equal-sized arrays. */
function pairSum(a: Operand[], ctx: Ctx, f: (x: number, y: number) => number): number {
  const xs = matrixOf(a[0], ctx).data;
  const ys = matrixOf(a[1], ctx).data;
  if (xs.length !== ys.length) throw NA;
  let total = 0;
  xs.forEach((x, i) => {
    const y = ys[i];
    if (isError(x)) throw x;
    if (isError(y)) throw y;
    if (typeof x === 'number' && typeof y === 'number') total += f(x, y);
  });
  return total;
}

const ROMAN_CHARS = 'MDCLXVI';
const ROMAN_VALUES = [1000, 500, 100, 50, 10, 5, 1];

/** A number in Roman numerals: form 0 classic, 1 to 4 more and more concise, as Excel writes them. */
function roman(value: number, form: number): string {
  let n = value;
  let out = '';
  const last = ROMAN_VALUES.length - 1;
  for (let i = 0; i <= last / 2; i++) {
    let index = 2 * i;
    const digit = Math.floor(n / ROMAN_VALUES[index]);
    if (digit % 5 === 4) {
      const index2 = digit === 4 ? index - 1 : index - 2;
      let steps = 0;
      while (steps < form && index < last) {
        steps++;
        if (ROMAN_VALUES[index2] - ROMAN_VALUES[index + 1] <= n) index++;
        else steps = form;
      }
      out += ROMAN_CHARS[index] + ROMAN_CHARS[index2];
      n = n + ROMAN_VALUES[index] - ROMAN_VALUES[index2];
    } else {
      if (digit > 4) out += ROMAN_CHARS[index - 1];
      out += ROMAN_CHARS[index].repeat(digit % 5);
      n %= ROMAN_VALUES[index];
    }
  }
  return out;
}

/** Roman numerals read as a number: a numeral before a larger one is taken away. */
function arabic(t: string): number {
  let s = t.trim().toUpperCase();
  const negative = s.startsWith('-');
  if (negative) s = s.slice(1);
  if (s.length > 255) throw VALUE;
  let total = 0;
  let largest = 0;
  for (let i = s.length - 1; i >= 0; i--) {
    const at = ROMAN_CHARS.indexOf(s[i]);
    if (at < 0) throw VALUE;
    const v = ROMAN_VALUES[at];
    if (v < largest) total -= v;
    else {
      total += v;
      largest = v;
    }
  }
  return negative ? -total : total;
}

const DIGITS = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';

/** A trigonometric function of one number, #NUM! outside its domain, #DIV/0! where it is infinite. */
const trig = (f: (x: number) => number, divides = false): FunctionSpec =>
  scalar(1, 1, ([x]) => {
    const v = num(x);
    if (Math.abs(v) >= 2 ** 27) throw NUM;
    const r = f(v);
    if (!Number.isFinite(r)) throw divides ? DIV0 : NUM;
    return r;
  });

// -----------------------------------------------------------------------------
// Database functions and Treasury bills

/**
 * A database criterion: as COUNTIF's, except that a blank matches anything
 * and plain text matches text beginning with it, as Excel's filters do.
 */
function dbCriterion(c: Scalar): (v: Scalar) => boolean {
  if (c === null || c === '') return () => true;
  if (typeof c === 'string' && !/^(<=|>=|<>|<|>|=)/.test(c) && !Number.isFinite(Number(c)) && !/^(true|false)$/i.test(c)) return criterion(`${c}*`);
  return criterion(c);
}

/** Which column of a database a header names, ignoring case; -1 if none. */
function columnNamed(name: Scalar, headers: Scalar[]): number {
  if (name === null) return -1;
  const wanted = text(name).toLowerCase();
  return headers.findIndex((h) => h !== null && !isError(h) && text(h).toLowerCase() === wanted);
}

/**
 * The records of a database (a range whose first row names its columns)
 * that meet the criteria (a range of headers above conditions): those
 * meeting every condition in any one criteria row. Gives each such record's
 * value in the field, or with no field, null for each.
 */
function records(a: Operand[], ctx: Ctx, fieldNeeded = true): Scalar[] {
  const db = grid(a[0], ctx);
  const criteria = grid(a[2], ctx);
  if (criteria.rows < 2) throw VALUE;
  const headers = Array.from({ length: db.cols }, (_, c) => db.at(0, c));
  let field: number | undefined;
  if (given(a, 1)) {
    const f = scalarOf(a[1], ctx);
    if (typeof f === 'number') field = Math.trunc(f) - 1;
    else if (typeof f === 'string') field = columnNamed(f, headers);
    else if (isError(f)) throw f;
    else throw VALUE;
    if (field < 0 || field >= db.cols) throw VALUE;
  } else if (fieldNeeded) throw VALUE;
  const rows: { col: number; test: (v: Scalar) => boolean }[][] = [];
  for (let r = 1; r < criteria.rows; r++) {
    const row = [];
    for (let c = 0; c < criteria.cols; c++) {
      const condition = criteria.at(r, c);
      if (isError(condition)) throw condition;
      if (condition === null || condition === '') continue;
      const col = columnNamed(criteria.at(0, c), headers);
      // A condition under a header the database lacks is a worked-out one: TRUE lets every record through.
      row.push(col < 0 ? { col, test: () => condition === true } : { col, test: dbCriterion(condition) });
    }
    rows.push(row);
  }
  const out: Scalar[] = [];
  for (let r = 1; r < db.rows; r++) {
    if (rows.some((row) => row.every((t) => t.test(t.col < 0 ? null : db.at(r, t.col))))) out.push(field === undefined ? null : db.at(r, field));
  }
  return out;
}

/** The numbers among the matching records' values in the field. */
function dbNumbers(a: Operand[], ctx: Ctx): number[] {
  const out: number[] = [];
  for (const v of records(a, ctx)) {
    if (isError(v)) throw v;
    if (typeof v === 'number') out.push(v);
  }
  return out;
}

/** Days from a Treasury bill's settlement to its maturity, which must come after it and within a year. */
function billDays(s: Scalar, m: Scalar, ctx: Ctx): number {
  const settle = Math.trunc(dateArg(s, ctx));
  const mature = Math.trunc(dateArg(m, ctx));
  if (settle >= mature || mature > eomonth(settle, 12, ctx, false)) throw NUM;
  return mature - settle;
}

// -----------------------------------------------------------------------------
// LAMBDA's helpers

/** A LAMBDA argument taking `count` values; anything else is #VALUE!. */
function lambdaOf(op: Operand, count: number): Lambda {
  if (!(op instanceof Lambda) || op.params.length !== count) throw VALUE;
  return op;
}

/** A LAMBDA's result where one value is wanted: more than one is #CALC!, as arrays can't hold arrays. */
function single(v: Operand, ctx: Ctx): Scalar {
  if (v instanceof Lambda) return CALC;
  if (v instanceof Areas) return VALUE;
  if (v instanceof Ref) return v.single ? ctx.host.value(v.sheet, v.r1, v.c1) : CALC;
  if (v instanceof Matrix) return v.rows === 1 && v.cols === 1 ? v.get(0, 0) : CALC;
  return v;
}

/** Each row (or column) of a range or array, a range's as a reference, as BYROW and BYCOL give them to their LAMBDA. */
function lines(op: Operand, ctx: Ctx, byCol: boolean): Operand[] {
  if (op instanceof Ref) {
    const r = ctx.host.bound(op);
    return byCol
      ? Array.from({ length: r.cols }, (_, c) => new Ref(r.sheet, r.r1, r.c1 + c, r.r2, r.c1 + c))
      : Array.from({ length: r.rows }, (_, i) => new Ref(r.sheet, r.r1 + i, r.c1, r.r1 + i, r.c2));
  }
  const m = matrixOf(op, ctx);
  return byCol
    ? Array.from({ length: m.cols }, (_, c) => Matrix.build(m.rows, 1, (r) => m.get(r, c)))
    : Array.from({ length: m.rows }, (_, r) => Matrix.build(1, m.cols, (_r, c) => m.get(r, c)));
}

// -----------------------------------------------------------------------------
// Engineering: binary, octal and hexadecimal, bits, Bessel functions, complex numbers

/** The bits ten binary, octal or hexadecimal digits hold; a negative number is written as their two's complement. */
const RADIX_BITS: Record<number, number> = { 2: 10, 8: 30, 16: 40 };
const RADIX_DIGITS: Record<number, RegExp> = { 2: /^[01]{0,10}$/, 8: /^[0-7]{0,10}$/, 16: /^[0-9a-f]{0,10}$/i };

/** A binary, octal or hexadecimal number read from its digits: #NUM! with more than ten, or one the base lacks. */
function fromRadix(v: Scalar, radix: number): number {
  if (typeof v === 'boolean') throw VALUE;
  const t = text(v);
  if (!RADIX_DIGITS[radix].test(t)) throw NUM;
  const n = t ? parseInt(t, radix) : 0;
  const bits = RADIX_BITS[radix];
  return n >= 2 ** (bits - 1) ? n - 2 ** bits : n;
}

/** A number in a base, padded to `places` digits; #NUM! beyond what ten digits hold. A negative number is ten digits, whatever `places`. */
function toRadix(n: number, radix: number, places: Scalar | undefined): string {
  const bits = RADIX_BITS[radix];
  if (n < -(2 ** (bits - 1)) || n >= 2 ** (bits - 1)) throw NUM;
  if (n < 0) return (n + 2 ** bits).toString(radix).toUpperCase();
  const digits = n.toString(radix).toUpperCase();
  if (places === undefined || places === null) return digits;
  if (typeof places === 'boolean') throw VALUE;
  const width = Math.trunc(num(places));
  if (width < digits.length || width > 10) throw NUM;
  return digits.padStart(width, '0');
}

/** BIN2DEC and its kind: from a base (or decimal, 10) to another. */
const convert = (from: number, to: number): FunctionSpec =>
  scalar(1, to === 10 ? 1 : 2, ([x, places]) => {
    if (typeof x === 'boolean') throw VALUE;
    const n = from === 10 ? Math.trunc(num(x)) : fromRadix(x, from);
    return to === 10 ? n : toRadix(n, to, places);
  });

const BIT_LIMIT = 2 ** 48;
const HALF = 2 ** 24;

/** A whole number from 0 to 2^48 − 1, as the bit functions take. */
function bitNumber(v: Scalar): number {
  const n = num(v);
  if (n < 0 || n >= BIT_LIMIT || !Number.isInteger(n)) throw NUM;
  return n;
}

/** A bitwise operation on 48-bit numbers, a 24-bit half at a time. */
const bitwise = (op: (a: number, b: number) => number): FunctionSpec =>
  scalar(2, 2, ([x, y]) => {
    const a = bitNumber(x);
    const b = bitNumber(y);
    return (op(Math.floor(a / HALF), Math.floor(b / HALF)) >>> 0) * HALF + (op(a % HALF, b % HALF) >>> 0);
  });

const bitShift = (left: boolean): FunctionSpec =>
  scalar(2, 2, ([x, by]) => {
    const n = bitNumber(x);
    const s = Math.trunc(num(by));
    if (Math.abs(s) > 53) throw NUM;
    const up = left ? s : -s;
    const r = up >= 0 ? n * 2 ** up : Math.floor(n / 2 ** -up);
    if (r >= BIT_LIMIT) throw NUM;
    return r;
  });

/*
 * The Bessel functions as Excel works them out: rational fits (Hart) for
 * J₀, J₁, Y₀ and Y₁ below 8 and asymptotic forms above it, the polynomial
 * fits of Abramowitz and Stegun 9.8 for I₀, I₁, K₀ and K₁, and recurrences
 * from those for higher orders. They are good to about eight digits, and
 * reproduce Excel's own digits, such as BESSELJ(0,0) = 1.00000000283141.
 */

const TWO_OVER_PI = 0.636619772;

/** Horner's rule: c[0] + c[1]·t + c[2]·t² + … */
const poly = (c: number[], t: number) => c.reduceRight((s, k) => s * t + k, 0);

/** J₀ or J₁ (and the matching Y) beyond 8: amplitude and phase corrections. */
function besselFar(x: number, one: boolean, y: boolean): number {
  const z = 8 / x;
  const t = z * z;
  const xx = x - (one ? 2.356194491 : 0.785398164);
  const p = one
    ? poly([1, 0.183105e-2, -0.3516396496e-4, 0.2457520174e-5, -0.240337019e-6], t)
    : poly([1, -0.1098628627e-2, 0.2734510407e-4, -0.2073370639e-5, 0.2093887211e-6], t);
  const q = one
    ? poly([0.04687499995, -0.2002690873e-3, 0.8449199096e-5, -0.88228987e-6, 0.105787412e-6], t)
    : poly([-0.1562499995e-1, 0.1430488765e-3, -0.6911147651e-5, 0.7621095161e-6, y ? -0.934945152e-7 : -0.934935152e-7], t);
  const amplitude = Math.sqrt(TWO_OVER_PI / x);
  return y ? amplitude * (Math.sin(xx) * p + z * Math.cos(xx) * q) : amplitude * (Math.cos(xx) * p - z * Math.sin(xx) * q);
}

function besselJ0(x: number): number {
  const ax = Math.abs(x);
  if (ax >= 8) return besselFar(ax, false, false);
  const t = x * x;
  return (
    poly([57568490574, -13362590354, 651619640.7, -11214424.18, 77392.33017, -184.9052456], t) /
    poly([57568490411, 1029532985, 9494680.718, 59272.64853, 267.8532712, 1], t)
  );
}

function besselJ1(x: number): number {
  const ax = Math.abs(x);
  if (ax >= 8) return Math.sign(x) * besselFar(ax, true, false);
  const t = x * x;
  return (
    (x * poly([72362614232, -7895059235, 242396853.1, -2972611.439, 15704.4826, -30.16036606], t)) /
    poly([144725228442, 2300535178, 18583304.74, 99447.43394, 376.9991397, 1], t)
  );
}

/**
 * Jₙ: upward from J₀ and J₁ when x is past n, else Miller's downward
 * recurrence from 2⌊(n + ⌊√(40n)⌋)/2⌋, scaled so J₀ + 2J₂ + 2J₄ + … = 1.
 */
function besselJ(x: number, n: number): number {
  if (n === 0) return besselJ0(x);
  if (n === 1) return besselJ1(x);
  const ax = Math.abs(x);
  if (ax === 0) return 0;
  const tox = 2 / ax;
  let result: number;
  if (ax > n) {
    let below = besselJ0(ax);
    let here = besselJ1(ax);
    for (let j = 1; j < n; j++) [below, here] = [here, j * tox * here - below];
    result = here;
  } else {
    let even = false;
    let above = 0;
    let here = 1;
    let total = 0;
    result = 0;
    for (let j = 2 * Math.floor((n + Math.floor(Math.sqrt(40 * n))) / 2); j > 0; j--) {
      [above, here] = [here, j * tox * here - above];
      if (Math.abs(here) > 1e10) {
        here *= 1e-10;
        above *= 1e-10;
        result *= 1e-10;
        total *= 1e-10;
      }
      if (even) total += here;
      even = !even;
      if (j === n) result = above;
    }
    result /= 2 * total - here;
  }
  return x < 0 && n % 2 ? -result : result;
}

function besselY0(x: number): number {
  if (x >= 8) return besselFar(x, false, true);
  const t = x * x;
  return (
    poly([-2957821389, 7062834065, -512359803.6, 10879881.29, -86327.92757, 228.4622733], t) /
      poly([40076544269, 745249964.8, 7189466.438, 47447.2647, 226.1030244, 1], t) +
    TWO_OVER_PI * besselJ0(x) * Math.log(x)
  );
}

function besselY1(x: number): number {
  if (x >= 8) return besselFar(x, true, true);
  const t = x * x;
  return (
    (x * poly([-0.4900604943e13, 0.127527439e13, -0.5153438139e11, 0.7349264551e9, -0.4237922726e7, 0.8511937935e4], t)) /
      poly([0.249958057e14, 0.4244419664e12, 0.3733650367e10, 0.2245904002e8, 0.102042605e6, 0.3549632885e3, 1], t) +
    TWO_OVER_PI * (besselJ1(x) * Math.log(x) - 1 / x)
  );
}

/** Yₙ, upward from Y₀ and Y₁. */
function besselY(x: number, n: number): number {
  if (n === 0) return besselY0(x);
  let below = besselY0(x);
  let here = besselY1(x);
  for (let j = 1; j < n; j++) [below, here] = [here, ((2 * j) / x) * here - below];
  return here;
}

function besselI0(x: number): number {
  const ax = Math.abs(x);
  if (ax < 3.75) return poly([1, 3.5156229, 3.0899424, 1.2067492, 0.2659732, 0.360768e-1, 0.45813e-2], (x / 3.75) ** 2);
  return (Math.exp(ax) / Math.sqrt(ax)) * poly([0.39894228, 0.1328592e-1, 0.225319e-2, -0.157565e-2, 0.916281e-2, -0.2057706e-1, 0.2635537e-1, -0.1647633e-1, 0.392377e-2], 3.75 / ax);
}

function besselI1(x: number): number {
  const ax = Math.abs(x);
  const r =
    ax < 3.75
      ? ax * poly([0.5, 0.87890594, 0.51498869, 0.15084934, 0.2658733e-1, 0.301532e-2, 0.32411e-3], (x / 3.75) ** 2)
      : (Math.exp(ax) / Math.sqrt(ax)) * poly([0.39894228, -0.3988024e-1, -0.362018e-2, 0.163801e-2, -0.1031555e-1, 0.2282967e-1, -0.2895312e-1, 0.1787654e-1, -0.420059e-2], 3.75 / ax);
  return x < 0 ? -r : r;
}

/** Iₙ: Miller's downward recurrence from 2(n + ⌊√(40n)⌋), scaled to I₀. */
function besselI(x: number, n: number): number {
  if (n === 0) return besselI0(x);
  if (n === 1) return besselI1(x);
  if (x === 0) return 0;
  const tox = 2 / Math.abs(x);
  let above = 0;
  let here = 1;
  let result = 0;
  for (let j = 2 * (n + Math.floor(Math.sqrt(40 * n))); j > 0; j--) {
    [above, here] = [here, above + j * tox * here];
    if (Math.abs(here) > 1e10) {
      result *= 1e-10;
      here *= 1e-10;
      above *= 1e-10;
    }
    if (j === n) result = above;
  }
  result *= besselI0(x) / here;
  return x < 0 && n % 2 ? -result : result;
}

function besselK0(x: number): number {
  if (x <= 2) return -Math.log(x / 2) * besselI0(x) + poly([-0.57721566, 0.4227842, 0.23069756, 0.348859e-1, 0.262698e-2, 0.1075e-3, 0.74e-5], (x * x) / 4);
  return (Math.exp(-x) / Math.sqrt(x)) * poly([1.25331414, -0.7832358e-1, 0.2189568e-1, -0.1062446e-1, 0.587872e-2, -0.25154e-2, 0.53208e-3], 2 / x);
}

function besselK1(x: number): number {
  if (x <= 2) return Math.log(x / 2) * besselI1(x) + (1 / x) * poly([1, 0.15443144, -0.67278579, -0.18156897, -0.1919402e-1, -0.110404e-2, -0.4686e-4], (x * x) / 4);
  return (Math.exp(-x) / Math.sqrt(x)) * poly([1.25331414, 0.23498619, -0.365562e-1, 0.1504268e-1, -0.780353e-2, 0.325614e-2, -0.68245e-3], 2 / x);
}

/** Kₙ, upward from K₀ and K₁. */
function besselK(x: number, n: number): number {
  if (n === 0) return besselK0(x);
  let below = besselK0(x);
  let here = besselK1(x);
  for (let j = 1; j < n; j++) [below, here] = [here, below + ((2 * j) / x) * here];
  return here;
}

/** The Bessel functions' x and order: the order cut to a whole number, at least 0. */
function besselArgs(x: Scalar, n: Scalar, positiveX: boolean): [number, number] {
  if (typeof x === 'boolean' || typeof n === 'boolean') throw VALUE;
  const v = num(x);
  const order = Math.trunc(num(n));
  // The recurrences take a step per order: beyond a million, #NUM! rather than a wait.
  if (order < 0 || order > 1_000_000 || (positiveX && v <= 0)) throw NUM;
  return [v, order];
}

interface Complex {
  re: number;
  im: number;
  /** The letter its text used for √−1, when it had one. */
  unit?: string;
}

const COMPLEX_PART = '(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][+-]?\\d+)?';
/** "3", "3+4i", "3-i", "4j", "-i": a real part, an imaginary part, or both. */
const COMPLEX_TEXT = new RegExp(`^(?:([+-]?${COMPLEX_PART})(?:([+-])(${COMPLEX_PART})?([ij]))?|([+-]?)(${COMPLEX_PART})?([ij]))$`);

/** A complex number from its text, or a number as its real part. */
function complexOf(v: Scalar): Complex {
  if (typeof v === 'number') return { re: v, im: 0 };
  if (typeof v === 'boolean') throw VALUE;
  if (v === null) return { re: 0, im: 0 };
  if (isError(v)) throw v;
  const m = COMPLEX_TEXT.exec(v);
  if (!m) throw NUM;
  if (m[7]) return { re: 0, im: (m[5] === '-' ? -1 : 1) * (m[6] ? Number(m[6]) : 1), unit: m[7] };
  if (!m[4]) return { re: Number(m[1]), im: 0 };
  return { re: Number(m[1]), im: (m[2] === '-' ? -1 : 1) * (m[3] ? Number(m[3]) : 1), unit: m[4] };
}

/** The letter for √−1 in a result: j if the numbers used it; mixing i and j is #VALUE!. */
function unitOf(zs: Complex[]): string {
  const units = new Set(zs.map((z) => z.unit).filter(Boolean));
  if (units.size > 1) throw VALUE;
  return units.has('j') ? 'j' : 'i';
}

/** A complex number as Excel writes it: "3+4i", "3-i", "4i", "3", parts to 15 significant digits. */
function complexText(z: Complex, unit: string): string {
  if (!Number.isFinite(z.re) || !Number.isFinite(z.im)) throw NUM;
  const re = round15(z.re);
  const im = round15(z.im);
  if (im === 0) return numberText(re);
  const imaginary = im === 1 ? '' : im === -1 ? '-' : numberText(im);
  if (re === 0) return imaginary + unit;
  return `${numberText(re)}${im > 0 ? '+' : ''}${imaginary}${unit}`;
}

const ONE: Complex = { re: 1, im: 0 };
const times = (a: Complex, b: Complex): Complex => ({ re: a.re * b.re - a.im * b.im, im: a.re * b.im + a.im * b.re });

function over(a: Complex, b: Complex): Complex {
  const d = b.re * b.re + b.im * b.im;
  if (d === 0) throw NUM;
  return { re: (a.re * b.re + a.im * b.im) / d, im: (a.im * b.re - a.re * b.im) / d };
}

const sinC = (z: Complex): Complex => ({ re: Math.sin(z.re) * Math.cosh(z.im), im: Math.cos(z.re) * Math.sinh(z.im) });
const cosC = (z: Complex): Complex => ({ re: Math.cos(z.re) * Math.cosh(z.im), im: -Math.sin(z.re) * Math.sinh(z.im) });
const sinhC = (z: Complex): Complex => ({ re: Math.sinh(z.re) * Math.cos(z.im), im: Math.cosh(z.re) * Math.sin(z.im) });
const coshC = (z: Complex): Complex => ({ re: Math.cosh(z.re) * Math.cos(z.im), im: Math.sinh(z.re) * Math.sin(z.im) });

/**
 * A quotient over a real denominator, as the double-angle forms of tan, cot,
 * sec and the rest give them (tan z = (sin 2a + i sinh 2b) / (cos 2a + cosh 2b)),
 * which keep the digits Excel shows; #NUM! where the function is infinite.
 */
function halves(re: number, im: number, denominator: number): Complex {
  if (denominator === 0) throw NUM;
  return { re: re / denominator, im: im / denominator };
}

function lnC(z: Complex): Complex {
  if (z.re === 0 && z.im === 0) throw NUM;
  return { re: Math.log(Math.hypot(z.re, z.im)), im: Math.atan2(z.im, z.re) };
}

/** z to a power, by the polar form, as Excel does even for whole powers: (2+3i)³ is −46+9.00000000000001i. */
function powerC(z: Complex, p: number): Complex {
  if (z.re === 0 && z.im === 0) {
    if (p <= 0) throw NUM;
    return { re: 0, im: 0 };
  }
  const r = Math.hypot(z.re, z.im) ** p;
  const angle = Math.atan2(z.im, z.re) * p;
  return { re: r * Math.cos(angle), im: r * Math.sin(angle) };
}

/** Complex numbers among the arguments, from ranges and arrays too; blank cells are left out. */
function complexes(args: Operand[], ctx: Ctx): Complex[] {
  const out: Complex[] = [];
  const take = (v: Scalar) => {
    if (v !== null) out.push(complexOf(v));
  };
  for (const a of args) {
    if (a === undefined) continue;
    if (a instanceof Ref) ctx.host.each(a, take);
    else if (a instanceof Areas) for (const r of a.refs) ctx.host.each(r, take);
    else if (a instanceof Matrix) a.data.forEach(take);
    else if (a instanceof Lambda) throw VALUE;
    else take(a);
  }
  return out;
}

/** A function of one complex number, giving one. */
const complex1 = (f: (z: Complex) => Complex): FunctionSpec =>
  scalar(1, 1, ([x]) => {
    const z = complexOf(x);
    return complexText(f(z), unitOf([z]));
  });

/** A function of two complex numbers, giving one. */
const complex2 = (f: (a: Complex, b: Complex) => Complex): FunctionSpec =>
  scalar(2, 2, ([x, y]) => {
    const a = complexOf(x);
    const b = complexOf(y);
    return complexText(f(a, b), unitOf([a, b]));
  });

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
  // Of no numbers, 0.
  PRODUCT: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = list(a, ctx);
    return xs.length ? xs.reduce((p, x) => p * x, 1) : 0;
  } },
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
  RANDARRAY: {
    min: 0,
    max: 5,
    volatile: true,
    fn: (a, ctx) => {
      const arg = (i: number, d: number) => (given(a, i) ? num(scalarOf(a[i], ctx)) : d);
      const rows = Math.trunc(arg(0, 1));
      const cols = Math.trunc(arg(1, 1));
      const lo = arg(2, 0);
      const hi = arg(3, 1);
      const whole = given(a, 4) && bool(scalarOf(a[4], ctx));
      // No rows or columns is an empty array, which Excel shows as #CALC!.
      if ((rows === 0 || cols === 0) && rows >= 0 && cols >= 0) throw CALC;
      sized(rows, cols);
      if (lo > hi || (whole && (!Number.isInteger(lo) || !Number.isInteger(hi)))) throw VALUE;
      return Matrix.build(rows, cols, () => (whole ? lo + Math.floor(Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo)));
    },
  },
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
  COUNTIF: { min: 2, max: 2, fn: (a, ctx) => perCriterion(a, ctx, 1, (x) => count(matching([x[0], x[1]], ctx))) },
  COUNTIFS: {
    min: 2,
    max: MANY,
    fn: (a, ctx) => {
      if (a.length % 2) throw VALUE;
      return perCriterion(a, ctx, 1, (x) => count(matching(x, ctx)));
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
      const { values, rows, cols } = flat(a[1], ctx);
      if (rows > 1 && cols > 1) throw VALUE;
      const mode = given(a, 2) ? Math.trunc(num(scalarOf(a[2], ctx))) : 0;
      const search = given(a, 3) ? Math.trunc(num(scalarOf(a[3], ctx))) : 1;
      if (![-1, 0, 1, 2].includes(mode) || ![-2, -1, 1, 2].includes(search)) throw VALUE;
      const find = finder(values, mode, search, a[0] instanceof Matrix || (a[0] instanceof Ref && !a[0].single));
      return eachValue(a[0], ctx, (target) => {
        if (isError(target)) throw target;
        const at = find(target);
        if (at < 0) throw NA;
        return at + 1;
      });
    },
  },
  XLOOKUP: { min: 3, max: 6, fn: (a, ctx) => xlookup(a, ctx) },
  INDEX: { min: 2, max: 4, fn: (a, ctx) => index(a, ctx) },
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
  ROWS: { min: 1, max: 1, fn: ([r], ctx) => lineCount(r, ctx, true) },
  COLUMNS: { min: 1, max: 1, fn: ([r], ctx) => lineCount(r, ctx, false) },
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
        const key = matrixOf(a[i], ctx);
        // Each key is one row or column; its order 1 or -1.
        if (key.rows > 1 && key.cols > 1) throw VALUE;
        keys.push(key);
        const order = given(a, i + 1) ? Math.trunc(num(scalarOf(a[i + 1], ctx))) : 1;
        if (order !== 1 && order !== -1) throw VALUE;
        orders.push(order);
      }
      // Keys that are rows sort the columns.
      const byCol = keys[0].rows === 1 && keys[0].cols > 1;
      if (keys.some((k) => (k.rows === 1 && k.cols > 1) !== byCol)) throw VALUE;
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
  CONCATENATE: scalar(1, MANY, (a) => cellText(a.map((v) => text(v)).join(''))),
  CONCAT: {
    min: 1,
    max: MANY,
    fn: (a, ctx) => {
      let out = '';
      for (const x of a) {
        if (x instanceof Ref || x instanceof Matrix) for (const v of matrixOf(x, ctx).data) out += text(v);
        else out += text(scalarOf(x, ctx));
      }
      return cellText(out);
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
      const s = Math.floor(dateArg(scalarOf(a[0], ctx), ctx));
      const left = Math.trunc(num(scalarOf(a[1], ctx)));
      const holidays = holidaysOf(a[2], ctx);
      return stepWorkdays(s, left, (d) => weekday(d, 2, ctx) <= 5 && !holidays.has(d), ctx);
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
      const look = matrixOf(a[1], ctx);
      let keys: Scalar[];
      let results: Scalar[];
      if (a[2] === undefined) {
        // The array form: the first row or column searched, the last one returned.
        const across = look.cols > look.rows;
        keys = across ? look.data.slice(0, look.cols) : Array.from({ length: look.rows }, (_, r) => look.get(r, 0));
        results = across ? look.data.slice((look.rows - 1) * look.cols) : Array.from({ length: look.rows }, (_, r) => look.get(r, look.cols - 1));
      } else {
        const result = matrixOf(a[2], ctx);
        // The vector form returns from a single row or column.
        if (result.rows > 1 && result.cols > 1) throw NA;
        keys = look.data;
        results = result.data;
      }
      return eachValue(a[0], ctx, (target) => {
        const at = approximate(keys, target);
        if (at < 0 || at >= results.length) throw NA;
        return results[at];
      });
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
  TEXTBEFORE: liftFirst({ min: 2, max: 6, fn: (a, ctx) => textAround(a, ctx, false) }),
  TEXTAFTER: liftFirst({ min: 2, max: 6, fn: (a, ctx) => textAround(a, ctx, true) }),
  TEXTSPLIT: {
    min: 2,
    max: 6,
    fn: (a, ctx) => {
      // Given several texts, each gives its first piece.
      if (a[0] instanceof Matrix || (a[0] instanceof Ref && !a[0].single)) return liftFirst(FUNCTIONS.TEXTSPLIT).fn!(a, ctx);
      const t = text(scalarOf(a[0], ctx));
      const insensitive = flag(a, 4, ctx);
      if (given(a, 1) && textsOf(a[1], ctx).every((d) => d === '')) throw VALUE;
      const skip = optBool(a, 3, ctx, false);
      const pad = given(a, 5) ? scalarOf(a[5], ctx) : NA;
      const keep = (parts: string[]) => (skip ? parts.filter((p) => p !== '') : parts);
      const lines = keep(given(a, 2) ? splitOn(t, textsOf(a[2], ctx), insensitive) : [t]).map((line) => keep(given(a, 1) ? splitOn(line, textsOf(a[1], ctx), insensitive) : [line]));
      // Nothing left once empty pieces are left out.
      if (lines.every((l) => !l.length)) throw err('#CALC!');
      const cols = Math.max(1, ...lines.map((l) => l.length));
      return Matrix.build(lines.length, cols, (r, c) => (c < lines[r].length ? lines[r][c] : pad));
    },
  },
  VALUETOTEXT: liftFirst({ min: 1, max: 2, fn: (a, ctx) => valueText(scalarOf(a[0], ctx), flag(a, 1, ctx)) }),
  ARRAYTOTEXT: {
    min: 1,
    max: 2,
    fn: (a, ctx) => {
      const m = matrixOf(a[0], ctx);
      const strict = flag(a, 1, ctx);
      if (!strict) return m.data.map((v) => valueText(v, false)).join(', ');
      return `{${m.rowsOf().map((row) => row.map((v) => valueText(v, true)).join(',')).join(';')}}`;
    },
  },
  FIXED: scalar(1, 3, ([x, d, plain]) => {
    const digits = Math.trunc(optional(d, 2));
    if (digits > 127) throw VALUE;
    // Rounded, then written out in full from its 15 significant digits, as Excel writes it.
    const n = excelRound(num(x), digits, 'half');
    const [whole, frac = ''] = writtenOut(Math.abs(n)).split('.');
    const decimals = Math.max(0, digits);
    const body = (plain !== undefined && plain !== null && bool(plain) ? whole : groupThousands(whole)) + (decimals ? `.${frac.padEnd(decimals, '0').slice(0, decimals)}` : '');
    return n < 0 ? `-${body}` : body;
  }),
  NUMBERVALUE: scalar(1, 3, ([t, dec, group]) => {
    let s = text(t).replace(/\s/g, '');
    if (!s) return 0;
    const d = dec === undefined || dec === null ? '.' : text(dec).charAt(0);
    // Without a group separator, the one the decimal separator isn't.
    const g = group === undefined || group === null ? (d === ',' ? '.' : ',') : text(group).charAt(0);
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
    for (let i = 0; i < b; i++) {
      p *= a - i;
      if (!Number.isFinite(p)) throw NUM;
    }
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
      if (code === 3) {
        // COUNTA, leaving out errors when asked.
        let n = 0;
        const count = (v: Scalar) => void (v !== null && !(skipErrors && isError(v)) && n++);
        for (const x of rest) {
          if (x instanceof Ref) ctx.host.each(x, count, skipSubtotals);
          else if (x instanceof Matrix) x.data.forEach(count);
          else if (x !== undefined) count(scalarOf(x, ctx));
        }
        return n;
      }
      if (!f) throw VALUE;
      // COUNT passes over errors whatever the option.
      return f(aggregateValues(rest, ctx, skipErrors || code === 2, skipSubtotals), rest, ctx);
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
        if (isError(x)) throw x;
        if (isError(p)) throw p;
        // A pair without a number for its chance is left out.
        if (typeof p !== 'number') return;
        if (p <= 0 || p > 1) throw NUM;
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
  FORECAST: { min: 3, max: 3, fn: (a, ctx) => forecast(a, ctx) },
  'FORECAST.LINEAR': { min: 3, max: 3, fn: (a, ctx) => forecast(a, ctx) },
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

  // Descriptive statistics
  AVEDEV: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = list(a, ctx);
    if (!xs.length) throw NUM;
    const mean = sum(xs) / xs.length;
    return sum(xs.map((x) => Math.abs(x - mean))) / xs.length;
  } },
  DEVSQ: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = list(a, ctx);
    if (!xs.length) throw NUM;
    const mean = sum(xs) / xs.length;
    return sum(xs.map((x) => (x - mean) ** 2));
  } },
  GEOMEAN: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = list(a, ctx);
    if (!xs.length || xs.some((x) => x <= 0)) throw NUM;
    return Math.exp(sum(xs.map(Math.log)) / xs.length);
  } },
  HARMEAN: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = list(a, ctx);
    if (!xs.length || xs.some((x) => x <= 0)) throw NUM;
    return xs.length / sum(xs.map((x) => 1 / x));
  } },
  SKEW: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = list(a, ctx);
    const n = xs.length;
    if (n < 3) throw DIV0;
    return (n / ((n - 1) * (n - 2))) * cubes(xs, Math.sqrt(variance(xs, true)));
  } },
  'SKEW.P': { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = list(a, ctx);
    return cubes(xs, Math.sqrt(variance(xs, false))) / xs.length;
  } },
  AVERAGEA: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = listA(a, ctx, true);
    if (!xs.length) throw DIV0;
    return sum(xs) / xs.length;
  } },
  MAXA: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = listA(a, ctx);
    return xs.length ? xs.reduce((m, x) => (x > m ? x : m), -Infinity) : 0;
  } },
  MINA: { min: 1, max: MANY, fn: (a, ctx) => {
    const xs = listA(a, ctx);
    return xs.length ? xs.reduce((m, x) => (x < m ? x : m), Infinity) : 0;
  } },
  STDEVA: { min: 1, max: MANY, fn: (a, ctx) => Math.sqrt(variance(listA(a, ctx), true)) },
  STDEVPA: { min: 1, max: MANY, fn: (a, ctx) => Math.sqrt(variance(listA(a, ctx), false)) },
  VARA: { min: 1, max: MANY, fn: (a, ctx) => variance(listA(a, ctx), true) },
  VARPA: { min: 1, max: MANY, fn: (a, ctx) => variance(listA(a, ctx), false) },

  // More of the normal distribution
  GAUSS: dist(1, 1, ([z]) => 0.5 * erf(num(z) / Math.SQRT2)),
  PHI: dist(1, 1, ([x]) => normalPdf(num(x))),
  FISHER: dist(1, 1, ([x]) => {
    const v = num(x);
    if (v <= -1 || v >= 1) throw NUM;
    return Math.atanh(v);
  }),
  FISHERINV: dist(1, 1, ([y]) => Math.tanh(num(y))),
  CONFIDENCE: dist(3, 3, ([alpha, sd, n]) => confidence(alpha, sd, n, false)),
  'CONFIDENCE.NORM': dist(3, 3, ([alpha, sd, n]) => confidence(alpha, sd, n, false)),
  'CONFIDENCE.T': dist(3, 3, ([alpha, sd, n]) => confidence(alpha, sd, n, true)),
  'Z.TEST': { min: 2, max: 3, fn: (a, ctx) => zTest(a, ctx) },
  ZTEST: { min: 2, max: 3, fn: (a, ctx) => zTest(a, ctx) },
  'LOGNORM.DIST': dist(4, 4, ([x, m, sd, c]) => logNormal(x, m, sd, bool(c))),
  LOGNORMDIST: dist(3, 3, ([x, m, sd]) => logNormal(x, m, sd, true)),
  'LOGNORM.INV': dist(3, 3, ([p, m, sd]) => Math.exp(num(m) + positive(sd) * normalInverse(num(p)))),
  LOGINV: dist(3, 3, ([p, m, sd]) => Math.exp(num(m) + positive(sd) * normalInverse(num(p)))),

  // Student's t
  'T.DIST': dist(3, 3, ([x, v, c]) => {
    const t = num(x);
    const df = degrees(v);
    return bool(c) ? 1 - tUpper(t, df) : tDensity(t, df);
  }),
  'T.DIST.RT': dist(2, 2, ([x, v]) => tUpper(num(x), degrees(v))),
  'T.DIST.2T': dist(2, 2, ([x, v]) => {
    const t = num(x);
    if (t < 0) throw NUM;
    return 2 * tUpper(t, degrees(v));
  }),
  TDIST: dist(3, 3, ([x, v, tails]) => {
    const t = num(x);
    const n = Math.trunc(num(tails));
    if (t < 0 || (n !== 1 && n !== 2)) throw NUM;
    return n * tUpper(t, degrees(v));
  }),
  'T.INV': dist(2, 2, ([p, v]) => tInverse(probability(p, true), degrees(v))),
  'T.INV.2T': dist(2, 2, ([p, v]) => twoTailedT(p, v)),
  TINV: dist(2, 2, ([p, v]) => twoTailedT(p, v)),
  'T.TEST': { min: 4, max: 4, fn: (a, ctx) => tTest(a, ctx) },
  TTEST: { min: 4, max: 4, fn: (a, ctx) => tTest(a, ctx) },

  // Chi-squared
  'CHISQ.DIST': dist(3, 3, ([x, v, c]) => {
    const k = degrees(v);
    const y = nonNegative(x);
    return bool(c) ? gammaP(k / 2, y / 2) : gammaDensity(y, k / 2, 2);
  }),
  'CHISQ.DIST.RT': dist(2, 2, ([x, v]) => gammaP(degrees(v) / 2, nonNegative(x) / 2, true)),
  CHIDIST: dist(2, 2, ([x, v]) => gammaP(degrees(v) / 2, nonNegative(x) / 2, true)),
  'CHISQ.INV': dist(2, 2, ([p, v]) => {
    const k = degrees(v);
    return inverseFrom0((x) => gammaP(k / 2, x / 2), probability(p));
  }),
  'CHISQ.INV.RT': dist(2, 2, ([p, v]) => chiUpperInverse(p, v)),
  CHIINV: dist(2, 2, ([p, v]) => chiUpperInverse(p, v)),
  'CHISQ.TEST': { min: 2, max: 2, fn: (a, ctx) => chiTest(a, ctx) },
  CHITEST: { min: 2, max: 2, fn: (a, ctx) => chiTest(a, ctx) },

  // F
  'F.DIST': dist(4, 4, ([x, v1, v2, c]) => {
    const f = nonNegative(x);
    const d1 = degrees(v1);
    const d2 = degrees(v2);
    return bool(c) ? fDist(f, d1, d2) : fDensity(f, d1, d2);
  }),
  'F.DIST.RT': dist(3, 3, ([x, v1, v2]) => fDist(nonNegative(x), degrees(v1), degrees(v2), true)),
  FDIST: dist(3, 3, ([x, v1, v2]) => fDist(nonNegative(x), degrees(v1), degrees(v2), true)),
  'F.INV': dist(3, 3, ([p, v1, v2]) => {
    const d1 = degrees(v1);
    const d2 = degrees(v2);
    return inverseFrom0((f) => fDist(f, d1, d2), probability(p));
  }),
  'F.INV.RT': dist(3, 3, ([p, v1, v2]) => fUpperInverse(p, v1, v2)),
  FINV: dist(3, 3, ([p, v1, v2]) => fUpperInverse(p, v1, v2)),
  'F.TEST': { min: 2, max: 2, fn: (a, ctx) => fTest(a, ctx) },
  FTEST: { min: 2, max: 2, fn: (a, ctx) => fTest(a, ctx) },

  // Gamma and beta
  GAMMA: dist(1, 1, ([x]) => gamma(num(x))),
  GAMMALN: dist(1, 1, ([x]) => logGamma(positive(x))),
  'GAMMALN.PRECISE': dist(1, 1, ([x]) => logGamma(positive(x))),
  'GAMMA.DIST': dist(4, 4, ([x, a, b, c]) => gammaDist(x, a, b, bool(c))),
  GAMMADIST: dist(4, 4, ([x, a, b, c]) => gammaDist(x, a, b, bool(c))),
  'GAMMA.INV': dist(3, 3, ([p, a, b]) => gammaInverse(p, a, b)),
  GAMMAINV: dist(3, 3, ([p, a, b]) => gammaInverse(p, a, b)),
  'BETA.DIST': dist(4, 6, ([x, alpha, beta, c, lo, hi]) => {
    const [a, b] = betaShape(alpha, beta);
    const at = betaPlace(x, lo, hi);
    if (bool(c)) return betaI(at.x, 1 - at.x, a, b);
    return Math.exp((a - 1) * Math.log(at.x) + (b - 1) * Math.log1p(-at.x) - logBeta(a, b)) / at.width;
  }),
  BETADIST: dist(3, 5, ([x, alpha, beta, lo, hi]) => {
    const [a, b] = betaShape(alpha, beta);
    const at = betaPlace(x, lo, hi);
    return betaI(at.x, 1 - at.x, a, b);
  }),
  'BETA.INV': dist(3, 5, ([p, a, b, lo, hi]) => betaInverse(p, a, b, lo, hi)),
  BETAINV: dist(3, 5, ([p, a, b, lo, hi]) => betaInverse(p, a, b, lo, hi)),

  // Counts of events
  'BINOM.DIST': dist(4, 4, ([s, n, p, c]) => binomial(s, n, p, bool(c))),
  BINOMDIST: dist(4, 4, ([s, n, p, c]) => binomial(s, n, p, bool(c))),
  'BINOM.INV': dist(3, 3, ([n, p, alpha]) => criticalBinomial(n, p, alpha)),
  CRITBINOM: dist(3, 3, ([n, p, alpha]) => criticalBinomial(n, p, alpha)),
  'NEGBINOM.DIST': dist(4, 4, ([f, s, p, c]) => negativeBinomial(f, s, p, bool(c))),
  NEGBINOMDIST: dist(3, 3, ([f, s, p]) => negativeBinomial(f, s, p, false)),
  'HYPGEOM.DIST': dist(5, 5, (a) => hypergeometric(a, bool(a[4]))),
  HYPGEOMDIST: dist(4, 4, (a) => hypergeometric(a, false)),
  'POISSON.DIST': dist(3, 3, ([x, m, c]) => poisson(x, m, bool(c))),
  POISSON: dist(3, 3, ([x, m, c]) => poisson(x, m, bool(c))),

  // Times between events
  'EXPON.DIST': dist(3, 3, ([x, l, c]) => exponential(x, l, bool(c))),
  EXPONDIST: dist(3, 3, ([x, l, c]) => exponential(x, l, bool(c))),
  'WEIBULL.DIST': dist(4, 4, ([x, a, b, c]) => weibull(x, a, b, bool(c))),
  WEIBULL: dist(4, 4, ([x, a, b, c]) => weibull(x, a, b, bool(c))),

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
    if (cost < 0 || salvage < 0 || life <= 0 || months < 1 || months > 12 || period < 1 || period > life + (months < 12 ? 1 : 0) || period > MAX_STEPS) throw NUM;
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
      const d = Math.floor(dateArg(scalarOf(a[0], ctx), ctx));
      const left = Math.trunc(num(scalarOf(a[1], ctx)));
      const weekend = weekendOf(a[2], ctx);
      if (weekend.every(Boolean)) throw VALUE;
      const holidays = holidaysOf(a[3], ctx);
      return stepWorkdays(d, left, (day) => workday(day, weekend, holidays, ctx), ctx);
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

  // Counting and sums of pairs
  FACT: scalar(1, 1, ([x]) => {
    const n = Math.trunc(num(x));
    if (n < 0 || n > 170) throw NUM;
    return factorial(n);
  }),
  FACTDOUBLE: scalar(1, 1, ([x]) => {
    const n = Math.trunc(num(x));
    if (n < -1 || n > 300) throw NUM;
    let f = 1;
    for (let i = n; i > 1; i -= 2) f *= i;
    return f;
  }),
  COMBIN: scalar(2, 2, ([n, k]) => {
    const a = Math.trunc(num(n));
    const b = Math.trunc(num(k));
    if (a < 0 || b < 0 || a < b) throw NUM;
    return choose(a, b);
  }),
  COMBINA: scalar(2, 2, ([n, k]) => {
    const a = Math.trunc(num(n));
    const b = Math.trunc(num(k));
    if (a < 0 || b < 0 || (a === 0 && b > 0)) throw NUM;
    return b === 0 ? 1 : choose(a + b - 1, b);
  }),
  GCD: { min: 1, max: MANY, fn: (a, ctx) => wholes(a, ctx).reduce(gcd2, 0) },
  LCM: { min: 1, max: MANY, fn: (a, ctx) => wholes(a, ctx).reduce((l, n) => (l === 0 || n === 0 ? 0 : (l / gcd2(l, n)) * n), 1) },
  MULTINOMIAL: {
    min: 1,
    max: MANY,
    fn: (a, ctx) => {
      const ns = wholes(a, ctx);
      const total = sum(ns);
      if (total > 170) throw NUM;
      return Math.round(ns.reduce((r, n) => r / factorial(n), factorial(total)));
    },
  },
  SERIESSUM: {
    min: 4,
    max: 4,
    fn: (a, ctx) => {
      const x = num(scalarOf(a[0], ctx));
      const n = num(scalarOf(a[1], ctx));
      const m = num(scalarOf(a[2], ctx));
      return numberMatrix(a[3], ctx)
        .flat()
        .reduce((s, c, i) => s + c * x ** (n + i * m), 0);
    },
  },
  SUMX2MY2: { min: 2, max: 2, fn: (a, ctx) => pairSum(a, ctx, (x, y) => x * x - y * y) },
  SUMX2PY2: { min: 2, max: 2, fn: (a, ctx) => pairSum(a, ctx, (x, y) => x * x + y * y) },
  SUMXMY2: { min: 2, max: 2, fn: (a, ctx) => pairSum(a, ctx, (x, y) => (x - y) ** 2) },
  SQRTPI: math1((x) => (x < 0 ? NUM : Math.sqrt(x * Math.PI))),
  'CEILING.PRECISE': scalar(1, 2, ([x, s]) => {
    const k = Math.abs(optional(s, 1));
    return k === 0 ? 0 : excelRound(Math.ceil(num(x) / k) * k, 10, 'half');
  }),
  'ISO.CEILING': scalar(1, 2, ([x, s]) => {
    const k = Math.abs(optional(s, 1));
    return k === 0 ? 0 : excelRound(Math.ceil(num(x) / k) * k, 10, 'half');
  }),
  'FLOOR.PRECISE': scalar(1, 2, ([x, s]) => {
    const k = Math.abs(optional(s, 1));
    return k === 0 ? 0 : excelRound(Math.floor(num(x) / k) * k, 10, 'half');
  }),

  // Number bases and Roman numerals
  ROMAN: scalar(1, 2, ([x, f]) => {
    const n = Math.trunc(num(x));
    const form = f === undefined || f === null ? 0 : typeof f === 'boolean' ? (f ? 0 : 4) : Math.trunc(num(f));
    if (n < 0 || n > 3999 || form < 0 || form > 4) throw VALUE;
    return roman(n, form);
  }),
  ARABIC: scalar(1, 1, ([t]) => arabic(text(t))),
  BASE: scalar(2, 3, ([x, r, len]) => {
    const n = Math.trunc(num(x));
    const radix = Math.trunc(num(r));
    const min = Math.trunc(optional(len, 0));
    if (n < 0 || n >= 2 ** 53 || radix < 2 || radix > 36 || min < 0 || min > 255) throw NUM;
    return n.toString(radix).toUpperCase().padStart(min, '0');
  }),
  DECIMAL: scalar(2, 2, ([t, r]) => {
    const s = text(t).trim().toUpperCase();
    const radix = Math.trunc(num(r));
    if (radix < 2 || radix > 36) throw NUM;
    if (s.length > 255) throw VALUE;
    let n = 0;
    for (const ch of s) {
      const d = DIGITS.indexOf(ch);
      if (d < 0 || d >= radix) throw NUM;
      n = n * radix + d;
    }
    return n;
  }),

  // Trigonometry
  SIN: trig(Math.sin),
  COS: trig(Math.cos),
  TAN: trig(Math.tan),
  ASIN: math1((x) => (x < -1 || x > 1 ? NUM : Math.asin(x))),
  ACOS: math1((x) => (x < -1 || x > 1 ? NUM : Math.acos(x))),
  ATAN: math1(Math.atan),
  ATAN2: scalar(2, 2, ([x, y]) => {
    if (num(x) === 0 && num(y) === 0) throw DIV0;
    return Math.atan2(num(y), num(x));
  }),
  SINH: math1(Math.sinh),
  COSH: math1(Math.cosh),
  TANH: math1(Math.tanh),
  ASINH: math1(Math.asinh),
  ACOSH: math1((x) => (x < 1 ? NUM : Math.acosh(x))),
  ATANH: math1((x) => (x <= -1 || x >= 1 ? NUM : Math.atanh(x))),
  COT: trig((x) => 1 / Math.tan(x), true),
  CSC: trig((x) => 1 / Math.sin(x), true),
  SEC: trig((x) => 1 / Math.cos(x), true),
  COTH: math1((x) => (x === 0 ? DIV0 : 1 / Math.tanh(x))),
  CSCH: math1((x) => (x === 0 ? DIV0 : 1 / Math.sinh(x))),
  SECH: math1((x) => 1 / Math.cosh(x)),
  ACOT: math1((x) => Math.PI / 2 - Math.atan(x)),
  ACOTH: math1((x) => (Math.abs(x) <= 1 ? NUM : 0.5 * Math.log((x + 1) / (x - 1)))),
  RADIANS: math1((x) => (x * Math.PI) / 180),
  DEGREES: math1((x) => (x * 180) / Math.PI),

  // Database functions
  DSUM: { min: 3, max: 3, fn: (a, ctx) => sum(dbNumbers(a, ctx)) },
  DAVERAGE: { min: 3, max: 3, fn: (a, ctx) => {
    const xs = dbNumbers(a, ctx);
    if (!xs.length) throw DIV0;
    return sum(xs) / xs.length;
  } },
  DCOUNT: { min: 3, max: 3, fn: (a, ctx) => (given(a, 1) ? dbNumbers(a, ctx).length : records(a, ctx, false).length) },
  DCOUNTA: { min: 3, max: 3, fn: (a, ctx) => records(a, ctx, false).filter((v) => !given(a, 1) || (v !== null && v !== '')).length },
  DMAX: { min: 3, max: 3, fn: (a, ctx) => {
    const xs = dbNumbers(a, ctx);
    return xs.length ? xs.reduce((m, x) => (x > m ? x : m), -Infinity) : 0;
  } },
  DMIN: { min: 3, max: 3, fn: (a, ctx) => {
    const xs = dbNumbers(a, ctx);
    return xs.length ? xs.reduce((m, x) => (x < m ? x : m), Infinity) : 0;
  } },
  DPRODUCT: { min: 3, max: 3, fn: (a, ctx) => {
    const xs = dbNumbers(a, ctx);
    return xs.length ? xs.reduce((p, x) => p * x, 1) : 0;
  } },
  DSTDEV: { min: 3, max: 3, fn: (a, ctx) => Math.sqrt(variance(dbNumbers(a, ctx), true)) },
  DSTDEVP: { min: 3, max: 3, fn: (a, ctx) => Math.sqrt(variance(dbNumbers(a, ctx), false)) },
  DVAR: { min: 3, max: 3, fn: (a, ctx) => variance(dbNumbers(a, ctx), true) },
  DVARP: { min: 3, max: 3, fn: (a, ctx) => variance(dbNumbers(a, ctx), false) },
  DGET: { min: 3, max: 3, fn: (a, ctx) => {
    const found = records(a, ctx);
    if (!found.length) throw VALUE;
    if (found.length > 1) throw NUM;
    return found[0] ?? 0;
  } },

  // Treasury bills
  TBILLPRICE: scalar(3, 3, ([s, m, d], ctx) => {
    const days = billDays(s, m, ctx);
    const discount = num(d);
    if (discount <= 0) throw NUM;
    const price = 100 * (1 - (discount * days) / 360);
    if (price <= 0) throw NUM;
    return price;
  }),
  TBILLYIELD: scalar(3, 3, ([s, m, p], ctx) => {
    const days = billDays(s, m, ctx);
    const price = num(p);
    if (price <= 0) throw NUM;
    return ((100 - price) / price) * (360 / days);
  }),
  TBILLEQ: scalar(3, 3, ([s, m, d], ctx) => {
    const days = billDays(s, m, ctx);
    const discount = num(d);
    if (discount <= 0) throw NUM;
    if (days <= 182) return (365 * discount) / (360 - discount * days);
    const price = 100 * (1 - (discount * days) / 360);
    if (price <= 0) throw NUM;
    const t = days / 365;
    return (-t + Math.sqrt(t * t - (2 * t - 1) * (1 - 100 / price))) / (t - 0.5);
  }),

  // LAMBDA's helpers
  MAP: {
    min: 2,
    max: MANY,
    fn: (a, ctx) => {
      const f = lambdaOf(a[a.length - 1], a.length - 1);
      const arrays = a.slice(0, -1).map((x) => matrixOf(x, ctx));
      const rows = Math.max(...arrays.map((m) => m.rows));
      const cols = Math.max(...arrays.map((m) => m.cols));
      // Arrays of other sizes aren't stretched: where one has no value, it gives #N/A.
      return Matrix.build(rows, cols, (r, c) => single(f.call(arrays.map((m) => (r < m.rows && c < m.cols ? m.get(r, c) : NA))), ctx));
    },
  },
  REDUCE: {
    min: 3,
    max: 3,
    fn: (a, ctx) => {
      const f = lambdaOf(a[2], 2);
      let total: Operand = a[0] ?? null;
      for (const v of matrixOf(a[1], ctx).data) total = f.call([total, v]);
      return total;
    },
  },
  SCAN: {
    min: 3,
    max: 3,
    fn: (a, ctx) => {
      const f = lambdaOf(a[2], 2);
      const m = matrixOf(a[1], ctx);
      let total: Operand = a[0] ?? null;
      return new Matrix(m.rows, m.cols, m.data.map((v) => single((total = f.call([total, v])), ctx)));
    },
  },
  BYROW: {
    min: 2,
    max: 2,
    fn: (a, ctx) => {
      const f = lambdaOf(a[1], 1);
      const rows = lines(a[0], ctx, false);
      return new Matrix(rows.length, 1, rows.map((row) => single(f.call([row]), ctx)));
    },
  },
  BYCOL: {
    min: 2,
    max: 2,
    fn: (a, ctx) => {
      const f = lambdaOf(a[1], 1);
      const cols = lines(a[0], ctx, true);
      return new Matrix(1, cols.length, cols.map((col) => single(f.call([col]), ctx)));
    },
  },
  MAKEARRAY: {
    min: 3,
    max: 3,
    fn: (a, ctx) => {
      const rows = Math.trunc(num(scalarOf(a[0], ctx)));
      const cols = Math.trunc(num(scalarOf(a[1], ctx)));
      const f = lambdaOf(a[2], 2);
      sized(rows, cols);
      return Matrix.build(rows, cols, (r, c) => single(f.call([r + 1, c + 1]), ctx));
    },
  },

  // Binary, octal and hexadecimal
  BIN2DEC: convert(2, 10),
  BIN2HEX: convert(2, 16),
  BIN2OCT: convert(2, 8),
  DEC2BIN: convert(10, 2),
  DEC2HEX: convert(10, 16),
  DEC2OCT: convert(10, 8),
  HEX2BIN: convert(16, 2),
  HEX2DEC: convert(16, 10),
  HEX2OCT: convert(16, 8),
  OCT2BIN: convert(8, 2),
  OCT2DEC: convert(8, 10),
  OCT2HEX: convert(8, 16),

  // Bits
  BITAND: bitwise((a, b) => a & b),
  BITOR: bitwise((a, b) => a | b),
  BITXOR: bitwise((a, b) => a ^ b),
  BITLSHIFT: bitShift(true),
  BITRSHIFT: bitShift(false),
  DELTA: scalar(1, 2, ([a, b]) => (num(a) === optional(b, 0) ? 1 : 0)),

  // Bessel functions
  BESSELJ: dist(2, 2, ([x, n]) => besselJ(...besselArgs(x, n, false))),
  BESSELI: dist(2, 2, ([x, n]) => besselI(...besselArgs(x, n, false))),
  BESSELY: dist(2, 2, ([x, n]) => besselY(...besselArgs(x, n, true))),
  BESSELK: dist(2, 2, ([x, n]) => besselK(...besselArgs(x, n, true))),

  // Complex numbers
  COMPLEX: scalar(2, 3, ([re, im, suffix]) => {
    if (typeof re === 'boolean' || typeof im === 'boolean') throw VALUE;
    const unit = suffix === undefined || suffix === null ? 'i' : text(suffix);
    if (unit !== 'i' && unit !== 'j') throw VALUE;
    return complexText({ re: num(re), im: num(im) }, unit);
  }),
  IMREAL: scalar(1, 1, ([z]) => complexOf(z).re),
  IMAGINARY: scalar(1, 1, ([z]) => complexOf(z).im),
  IMABS: scalar(1, 1, ([z]) => {
    const c = complexOf(z);
    return Math.hypot(c.re, c.im);
  }),
  IMARGUMENT: scalar(1, 1, ([z]) => {
    const c = complexOf(z);
    if (c.re === 0 && c.im === 0) throw DIV0;
    return Math.atan2(c.im, c.re);
  }),
  IMCONJUGATE: complex1((z) => ({ re: z.re, im: -z.im })),
  IMSUM: { min: 1, max: MANY, fn: (a, ctx) => {
    const zs = complexes(a, ctx);
    return complexText(zs.reduce((s, z) => ({ re: s.re + z.re, im: s.im + z.im }), { re: 0, im: 0 }), unitOf(zs));
  } },
  IMPRODUCT: { min: 1, max: MANY, fn: (a, ctx) => {
    const zs = complexes(a, ctx);
    return complexText(zs.reduce(times, ONE), unitOf(zs));
  } },
  IMSUB: complex2((a, b) => ({ re: a.re - b.re, im: a.im - b.im })),
  IMDIV: complex2(over),
  IMPOWER: scalar(2, 2, ([x, p]) => {
    const z = complexOf(x);
    if (typeof p === 'boolean') throw VALUE;
    return complexText(powerC(z, num(p)), unitOf([z]));
  }),
  IMSQRT: complex1((z) => powerC(z, 0.5)),
  IMEXP: complex1((z) => ({ re: Math.exp(z.re) * Math.cos(z.im), im: Math.exp(z.re) * Math.sin(z.im) })),
  IMLN: complex1(lnC),
  IMLOG10: complex1((z) => {
    const l = lnC(z);
    return { re: l.re / Math.LN10, im: l.im / Math.LN10 };
  }),
  IMLOG2: complex1((z) => {
    const l = lnC(z);
    return { re: l.re / Math.LN2, im: l.im / Math.LN2 };
  }),
  IMSIN: complex1(sinC),
  IMCOS: complex1(cosC),
  IMTAN: complex1(({ re: a, im: b }) => halves(Math.sin(2 * a), Math.sinh(2 * b), Math.cos(2 * a) + Math.cosh(2 * b))),
  IMCOT: complex1(({ re: a, im: b }) => halves(Math.sin(2 * a), -Math.sinh(2 * b), Math.cosh(2 * b) - Math.cos(2 * a))),
  IMSEC: complex1(({ re: a, im: b }) => halves(2 * Math.cos(a) * Math.cosh(b), 2 * Math.sin(a) * Math.sinh(b), Math.cos(2 * a) + Math.cosh(2 * b))),
  IMCSC: complex1(({ re: a, im: b }) => halves(2 * Math.sin(a) * Math.cosh(b), -2 * Math.cos(a) * Math.sinh(b), Math.cosh(2 * b) - Math.cos(2 * a))),
  IMSINH: complex1(sinhC),
  IMCOSH: complex1(coshC),
  IMSECH: complex1(({ re: a, im: b }) => halves(2 * Math.cosh(a) * Math.cos(b), -2 * Math.sinh(a) * Math.sin(b), Math.cosh(2 * a) + Math.cos(2 * b))),
  IMCSCH: complex1(({ re: a, im: b }) => halves(2 * Math.sinh(a) * Math.cos(b), -2 * Math.cosh(a) * Math.sin(b), Math.cosh(2 * a) - Math.cos(2 * b))),

  // References and sheets
  ADDRESS: scalar(2, 5, ([r, c, absolute, a1, sheet]) => {
    const row = Math.trunc(num(r));
    const col = Math.trunc(num(c));
    const abs = Math.trunc(optional(absolute, 1));
    if (row < 1 || row > MAX_ROWS || col < 1 || col > MAX_COLS || abs < 1 || abs > 4) throw VALUE;
    const absRow = abs === 1 || abs === 2;
    const absCol = abs === 1 || abs === 3;
    const ref =
      a1 === undefined || a1 === null || bool(a1)
        ? `${absCol ? '$' : ''}${colToLetters(col - 1)}${absRow ? '$' : ''}${row}`
        : `R${absRow ? row : `[${row}]`}C${absCol ? col : `[${col}]`}`;
    if (sheet === undefined || sheet === null) return ref;
    const name = text(sheet);
    if (!name) return `!${ref}`;
    // Excel quotes a [Book]Sheet name only when the sheet part needs it.
    const book = /^\[[^\]']+\]/.exec(name)?.[0] ?? '';
    const quoted = quoteSheetName(name.slice(book.length));
    return `${quoted.startsWith("'") ? `'${book}${quoted.slice(1)}` : book + quoted}!${ref}`;
  }),
  FORMULATEXT: {
    min: 1,
    max: 1,
    fn: ([r], ctx) => {
      if (!(r instanceof Ref)) throw r instanceof Areas ? VALUE : NA;
      const t = ctx.host.formulaText(r.sheet, r.r1, r.c1);
      if (t === undefined) throw NA;
      return t;
    },
  },
  // Only what the cell itself gives: formats, widths and protection aren't known here (see CELL_INFO in native.ts).
  CELL: {
    min: 2,
    max: 2,
    fn: ([info, r], ctx) => {
      if (!(r instanceof Ref)) throw VALUE;
      const v = ctx.host.value(r.sheet, r.r1, r.c1);
      switch (text(scalarOf(info, ctx)).toLowerCase()) {
        case 'row':
          return r.r1 + 1;
        case 'col':
          return r.c1 + 1;
        case 'contents':
          return v;
        case 'type':
          return v === null ? 'b' : typeof v === 'string' ? 'l' : 'v';
        case 'address':
          return `$${colToLetters(r.c1)}$${r.r1 + 1}`;
        default:
          throw VALUE;
      }
    },
  },
  AREAS: { min: 1, max: 1, fn: ([r]) => (r instanceof Areas ? r.refs.length : r instanceof Ref ? 1 : (() => { throw VALUE; })()) },
  SHEET: {
    min: 0,
    max: 1,
    fn: ([v], ctx) => {
      if (v === undefined) return ctx.sheet + 1;
      if (v instanceof Ref) return v.sheet + 1;
      const at = ctx.host.sheetIndex(text(scalarOf(v, ctx)));
      if (at < 0) throw NA;
      return at + 1;
    },
  },
  SHEETS: { min: 0, max: 1, fn: ([v], ctx) => (v === undefined ? ctx.host.sheetCount() : v instanceof Ref ? 1 : (() => { throw VALUE; })()) },

  // How files write A2# and @A1:A3.
  ANCHORARRAY: { min: 1, max: 1, fn: ([r], ctx) => (r instanceof Ref && r.single ? (ctx.host.spillOf(r.sheet, r.r1, r.c1) ?? REF) : REF) },
  SINGLE: { min: 1, max: 1, fn: ([v], ctx) => scalarOf(v, ctx) },
};

// Bytes are characters, as in Excel set to a language that writes one byte a character.
for (const name of ['LEFT', 'RIGHT', 'FIND', 'SEARCH', 'REPLACE']) FUNCTIONS[`${name}B`] = FUNCTIONS[name];

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
