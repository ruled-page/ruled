import { numberText } from './excelNumbers';

/**
 * Values as formulas see them, and Excel's rules for turning one kind into
 * another and for comparing them.
 */

/** A single value: a number, text, TRUE/FALSE, an error, or a blank cell (null). */
export type Scalar = number | string | boolean | null | ErrorValue;

export interface ErrorValue {
  readonly error: string;
}

const errors = new Map<string, ErrorValue>();

/** The error with this code; one object per code, so errors compare by identity. */
export function err(code: string): ErrorValue {
  let e = errors.get(code);
  if (!e) errors.set(code, (e = Object.freeze({ error: code })));
  return e;
}

export const DIV0 = err('#DIV/0!');
export const NA = err('#N/A');
export const VALUE = err('#VALUE!');
export const REF = err('#REF!');
export const NAME = err('#NAME?');
export const NUM = err('#NUM!');
export const NULL = err('#NULL!');
export const SPILL = err('#SPILL!');
export const CALC = err('#CALC!');

/** Error codes a cell can hold, as Excel writes them. */
export const ERROR_CODES = new Set(['#DIV/0!', '#N/A', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!', '#GETTING_DATA', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!']);

export function isError(v: unknown): v is ErrorValue {
  return typeof v === 'object' && v !== null && 'error' in v;
}

/** A rectangle of values: an array a formula works out, or the cells of a range, read. */
export class Matrix {
  constructor(
    readonly rows: number,
    readonly cols: number,
    readonly data: Scalar[],
  ) {}

  get(r: number, c: number): Scalar {
    return this.data[r * this.cols + c];
  }

  /** The value at a row and column, a single row or column repeated across the other way, as Excel lines arrays up. */
  at(r: number, c: number): Scalar {
    const rr = this.rows === 1 ? 0 : r;
    const cc = this.cols === 1 ? 0 : c;
    return rr < this.rows && cc < this.cols ? this.data[rr * this.cols + cc] : NA;
  }

  static of(rows: Scalar[][]): Matrix {
    const cols = Math.max(1, ...rows.map((r) => r.length));
    const data: Scalar[] = [];
    for (const r of rows) for (let c = 0; c < cols; c++) data.push(c < r.length ? r[c] : NA);
    return new Matrix(rows.length, cols, data);
  }

  static build(rows: number, cols: number, f: (r: number, c: number) => Scalar): Matrix {
    const data: Scalar[] = new Array(rows * cols);
    for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) data[r * cols + c] = f(r, c);
    return new Matrix(rows, cols, data);
  }

  rowsOf(): Scalar[][] {
    return Array.from({ length: this.rows }, (_, r) => this.data.slice(r * this.cols, (r + 1) * this.cols));
  }
}

/** A reference to cells: a sheet and a rectangle, 0-based and inclusive. */
export class Ref {
  constructor(
    readonly sheet: number,
    readonly r1: number,
    readonly c1: number,
    readonly r2: number,
    readonly c2: number,
  ) {}

  get rows(): number {
    return this.r2 - this.r1 + 1;
  }

  get cols(): number {
    return this.c2 - this.c1 + 1;
  }

  get single(): boolean {
    return this.r1 === this.r2 && this.c1 === this.c2;
  }
}

/**
 * Several references at once: a union, (A1:B2,D4), or the same cells on a
 * run of sheets, Jan:Mar!B2. SUM and its kind read every area; most other
 * functions take one reference, and give #VALUE! for these.
 */
export class Areas {
  constructor(readonly refs: Ref[]) {}
}

/** A LAMBDA's parameters and body, with the names it can see where it was written. */
export class Lambda {
  constructor(
    readonly params: string[],
    readonly call: (args: Operand[]) => Operand,
  ) {}
}

/** What evaluating part of a formula gives: a value, an array, a reference (or several), or a LAMBDA. */
export type Operand = Scalar | Matrix | Ref | Areas | Lambda;

/** Numbers as Excel types them: 1,000.50, 12%, $5, 1.5E3. */
// A currency sign may come before the number ($100, €100) or, for the euro, pound and yen, after it (100€).
const NUMBER_TEXT = /^[+-]?(?:[$€£¥]\s*)?(?:\d{1,3}(?:,\d{3})+|\d*)(?:\.\d*)?(?:e[+-]?\d+)?(?:\s*[€£¥])?%?$/i;

/** Text read as a number, as Excel reads "12" or "1,000" in arithmetic; undefined when it doesn't read as one. */
/**
 * x to the power y as Excel works it out: 0 to the power 0 is #NUM!, 0 to a
 * negative power #DIV/0!, a negative number has its odd roots (-8 to the
 * 1/3 is -2) and no others, and a result too large is #NUM!.
 */
export function power(x: number, y: number): Scalar {
  if (x === 0 && y < 0) return DIV0;
  if (x === 0 && y === 0) return NUM;
  if (x < 0 && !Number.isInteger(y)) {
    const root = Math.round(1 / y);
    if (Math.abs(1 / y - root) > 1e-10 || root % 2 === 0) return NUM;
    return -Math.exp(Math.log(-x) * y);
  }
  const r = Math.pow(x, y);
  return Number.isFinite(r) ? r : NUM;
}

export function parseNumber(text: string): number | undefined {
  const s = text.trim();
  if (!s || !/\d/.test(s) || !NUMBER_TEXT.test(s)) return undefined;
  const percent = s.endsWith('%');
  const n = Number(s.replace(/[$€£¥,%\s]/g, ''));
  if (!Number.isFinite(n)) return undefined;
  return percent ? n / 100 : n;
}

const DATE_TEXT = /^(?:(\d{4})[-/](\d{1,2})[-/](\d{1,2})|(\d{1,2})\/(\d{1,2})\/(\d{4}))(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?(?:\s*([AaPp])[Mm])?)?$/;
const NAMED_DATE = /^(\d{1,2})[- ]([A-Za-z]{3,9})[- ](\d{4})$/;
const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
// 14:30, 4:5 pm and 4 pm.
const TIME_TEXT = /^(\d{1,2})(?::(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?|(?=\s*[AaPp][Mm]))(?:\s*([AaPp])[Mm])?$/;

/**
 * Text read as a date or time, as Excel reads "2024-01-10", "1/10/2024" or
 * "14:30" in arithmetic: a serial number in the 1900 date system (a
 * workbook in the 1904 system reads such text 1,462 days late here).
 * Undefined when it isn't one, or names a day that doesn't exist.
 */
export function parseDateText(written: string): number | undefined {
  // Spaces around a separator don't count: 2026 - 01 - 01, 4 : 35.
  const text = written.replace(/\s*([-/:])\s*/g, '$1');
  const t = TIME_TEXT.exec(text);
  if (t) {
    let h = Number(t[1]);
    if (t[4]) {
      if (h < 1 || h > 12) return undefined;
      h = (h % 12) + (/p/i.test(t[4]) ? 12 : 0);
    }
    const m = Number(t[2] ?? 0);
    const sec = Number(t[3] ?? 0);
    // A sixtieth second, 23:59:60, is the next minute, as Excel reads it.
    return m < 60 && sec <= 60 ? (h * 3600 + m * 60 + sec) / 86400 : undefined;
  }
  // 2-Feb-2026 and 2 February 2026.
  const named = NAMED_DATE.exec(text);
  if (named) {
    const name = named[2].toLowerCase();
    const month = MONTHS.findIndex((m) => name === m || name === m.slice(0, 3)) + 1;
    return month ? parseDateText(`${named[3]}-${month}-${named[1]}`) : undefined;
  }
  const d = DATE_TEXT.exec(text);
  if (!d) return undefined;
  const [y, mo, day] = d[1] ? [Number(d[1]), Number(d[2]), Number(d[3])] : [Number(d[6]), Number(d[4]), Number(d[5])];
  let hours = d[7] === undefined ? 0 : Number(d[7]);
  if (d[10]) {
    if (hours < 1 || hours > 12) return undefined;
    hours = (hours % 12) + (/p/i.test(d[10]) ? 12 : 0);
  }
  const time = d[7] === undefined ? 0 : (hours * 3600 + Number(d[8]) * 60 + Number(d[9] ?? 0)) / 86400;
  // The 1900 system counts a 29 February 1900 that never was, as day 60.
  if (y === 1900 && mo === 2 && day === 29) return 60 + time;
  const at = new Date(Date.UTC(y, mo - 1, day));
  if (y < 1900 || at.getUTCFullYear() !== y || at.getUTCMonth() !== mo - 1 || at.getUTCDate() !== day) return undefined;
  let serial = at.getTime() / 86400000 + 25569;
  if (serial < 61) serial -= 1;
  return serial + time;
}

/** A value in arithmetic: blank 0, TRUE 1, text that reads as a number; anything else is #VALUE!. Errors are thrown. */
export function num(v: Scalar): number {
  if (typeof v === 'number') return v;
  if (v === null) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string') {
    const n = parseNumber(v) ?? parseDateText(v);
    if (n === undefined) throw VALUE;
    return n;
  }
  throw v;
}

/** A value as text: numbers to 15 significant digits, TRUE and FALSE in capitals, blank as "". Errors are thrown. */
export function text(v: Scalar): string {
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numberText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v === null) return '';
  throw v;
}

/** A value as TRUE or FALSE: numbers by whether they are 0, the text "TRUE" or "FALSE"; other text is #VALUE!. Errors are thrown. */
export function bool(v: Scalar): boolean {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null) return false;
  if (typeof v === 'string') {
    const u = v.toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    throw VALUE;
  }
  throw v;
}

/** A number to Excel's 15 significant digits, so 0.1+0.2 equals 0.3. */
export function round15(n: number): number {
  return n === 0 || !Number.isFinite(n) ? n : Number(n.toPrecision(15));
}

/**
 * A sum or difference as Excel gives it: a result that is only the binary
 * noise of numbers that cancel (0.3-0.1-0.2) is 0.
 */
export function cancel(result: number, a: number, b: number): number {
  return result !== 0 && Math.abs(result) < Math.max(Math.abs(a), Math.abs(b)) * 1e-15 ? 0 : result;
}

const RANK = (v: Exclude<Scalar, null>) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);

/**
 * Compares as Excel's = < > do: numbers before text before TRUE and FALSE,
 * text without regard to case, numbers to 15 significant digits, a blank
 * as 0, "" or FALSE to match the other side. Errors sort last; what they
 * mean is the caller's to decide.
 */
export function compare(a: Scalar, b: Scalar): number {
  const blank = (other: Scalar) => (typeof other === 'string' ? '' : typeof other === 'boolean' ? false : 0);
  const x = a === null ? blank(b) : a;
  const y = b === null ? blank(a) : b;
  const rx = RANK(x);
  const ry = RANK(y);
  if (rx !== ry) return rx - ry;
  if (typeof x === 'number') {
    const d = round15(x) - round15(y as number);
    return d < 0 ? -1 : d > 0 ? 1 : 0;
  }
  if (typeof x === 'string') {
    const p = x.toLowerCase();
    const q = (y as string).toLowerCase();
    return p < q ? -1 : p > q ? 1 : 0;
  }
  if (typeof x === 'boolean') return x === y ? 0 : x ? 1 : -1;
  return 0;
}
