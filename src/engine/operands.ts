import type { FunctionSpec } from './functions';
import { Areas, err, isError, Lambda, Matrix, Ref, VALUE, type Operand, type Scalar } from './values';

/** How formulas see the workbook and their own cell, and how operands become values. */

/** What a formula reads from the workbook it is in. */
export interface Host {
  readonly date1904: boolean;
  /** Our index of a sheet by its name, or -1. */
  sheetIndex(name: string): number;
  value(sheet: number, row: number, col: number): Scalar;
  /** The values of a reference; whole rows and columns stop at the last used. */
  read(ref: Ref): Matrix;
  /** A reference of whole rows or columns cut to the sheet's used rows and columns. */
  bound(ref: Ref): Ref;
  /** Each value of a reference, blanks included up to the last used row and column; `skipSubtotals` leaves out cells that are SUBTOTAL formulas. */
  each(ref: Ref, f: (v: Scalar, row: number, col: number) => void, skipSubtotals?: boolean): void;
  /** A defined name as the formula at `ctx` sees it, its sheet's own first. */
  name(name: string, ctx: Ctx): Operand;
  /** The range a formula in a cell spills into (for A2#), if it spills. */
  spillOf(sheet: number, row: number, col: number): Ref | undefined;
  isFormula(sheet: number, row: number, col: number): boolean;
  /** How many sheets the workbook has, for SHEETS. */
  sheetCount(): number;
  /** A cell's formula as the formula bar shows it, for FORMULATEXT; undefined when it holds none. */
  formulaText(sheet: number, row: number, col: number): string | undefined;
}

/** What a formula's cell notes as it is worked out. */
export interface CellNotes {
  /** The size of a spilling formula's whole result. */
  measured?: { rows: number; cols: number };
  /** Where a HYPERLINK formula links to. */
  link?: string;
}

export interface Ctx {
  host: Host;
  sheet: number;
  row: number;
  col: number;
  /**
   * A formula from before dynamic arrays: a range it gives as its result
   * meets the formula's row or column (implicit intersection), as Excel
   * shows such formulas, rather than spilling.
   */
  legacy: boolean;
  notes?: CellNotes;
  /** The values of the LAMBDA parameters in scope, by name in capitals. */
  vars?: Record<string, Operand>;
}

// -----------------------------------------------------------------------------
// Operands

/** A reference where one value is wanted, in a formula from before dynamic arrays: the cell in the formula's row or column. */
export function intersect(ref: Ref, ctx: Ctx): Scalar {
  if (ref.single) return ctx.host.value(ref.sheet, ref.r1, ref.c1);
  if (ref.c1 === ref.c2 && ctx.row >= ref.r1 && ctx.row <= ref.r2) return ctx.host.value(ref.sheet, ctx.row, ref.c1);
  if (ref.r1 === ref.r2 && ctx.col >= ref.c1 && ctx.col <= ref.c2) return ctx.host.value(ref.sheet, ref.r1, ctx.col);
  return VALUE;
}

/**
 * An operand as values: a reference read, arrays as they are. Ranges are
 * worked on whole, in older formulas too (whose final result alone meets
 * the formula's row or column), so array formulas entered with
 * Ctrl+Shift+Enter in one cell, =SUM(A1:A9*B1:B9), work as Excel's do.
 */
export function valuesOf(op: Operand, ctx: Ctx): Scalar | Matrix {
  if (op instanceof Ref) return op.single ? ctx.host.value(op.sheet, op.r1, op.c1) : ctx.host.read(op);
  if (op instanceof Lambda) return err('#CALC!');
  if (op instanceof Areas) return VALUE;
  return op;
}

/** An operand as one value: a range met by the formula's row or column, an array's first value. */
export function scalarOf(op: Operand, ctx: Ctx): Scalar {
  if (op instanceof Ref) return op.single ? ctx.host.value(op.sheet, op.r1, op.c1) : intersect(op, ctx);
  if (op instanceof Matrix) return op.get(0, 0);
  if (op instanceof Lambda) return err('#CALC!');
  if (op instanceof Areas) return VALUE;
  return op;
}

/** A value worked out from each pair of values, arrays lined up as Excel lines them up. */
export function pairwise(a: Scalar | Matrix, b: Scalar | Matrix, f: (x: Scalar, y: Scalar) => Scalar): Scalar | Matrix {
  if (!(a instanceof Matrix) && !(b instanceof Matrix)) return f(a, b);
  const ra = a instanceof Matrix ? a.rows : 1;
  const ca = a instanceof Matrix ? a.cols : 1;
  const rb = b instanceof Matrix ? b.rows : 1;
  const cb = b instanceof Matrix ? b.cols : 1;
  const rows = ra === 1 ? rb : rb === 1 ? ra : Math.max(ra, rb);
  const cols = ca === 1 ? cb : cb === 1 ? ca : Math.max(ca, cb);
  return Matrix.build(rows, cols, (r, c) => {
    const x = a instanceof Matrix ? a.at(r, c) : a;
    const y = b instanceof Matrix ? b.at(r, c) : b;
    return f(x, y);
  });
}

/** Runs f on a value, or on each value of an array; an error thrown is its result. */
export function each(v: Scalar | Matrix, f: (x: Scalar) => Scalar): Scalar | Matrix {
  const safe = (x: Scalar) => {
    try {
      return f(x);
    } catch (e) {
      if (isError(e)) return e;
      throw e;
    }
  };
  return v instanceof Matrix ? Matrix.build(v.rows, v.cols, (r, c) => safe(v.get(r, c))) : safe(v);
}

export function attempt(f: () => Operand): Operand {
  try {
    return f();
  } catch (e) {
    if (isError(e)) return e;
    throw e;
  }
}

/**
 * Calls a function. Parameters it takes one value at a time get one; given
 * an array (or a range, in a dynamic-array formula), the function runs on
 * each value and gives an array of results, as Excel lifts it.
 */
export function apply(spec: FunctionSpec, ops: Operand[], ctx: Ctx): Operand {
  return attempt(() => {
    const scalars = spec.scalar;
    if (!scalars) return spec.fn!(ops, ctx);
    const vals: (Scalar | Matrix | Operand)[] = ops.map((op, i) => (scalars === true || scalars.includes(i) ? valuesOf(op, ctx) : op));
    let rows = 0;
    let cols = 0;
    vals.forEach((v, i) => {
      if (v instanceof Matrix && (scalars === true || scalars.includes(i))) {
        rows = Math.max(rows, v.rows);
        cols = Math.max(cols, v.cols);
      }
    });
    if (!rows) return spec.fn!(vals as Operand[], ctx);
    return Matrix.build(rows, cols, (r, c) => {
      const one = vals.map((v, i) => (v instanceof Matrix && (scalars === true || scalars.includes(i)) ? v.at(r, c) : v));
      const out = attempt(() => spec.fn!(one as Operand[], ctx));
      return out instanceof Matrix ? out.get(0, 0) : out instanceof Ref || out instanceof Areas ? scalarOf(out, ctx) : out instanceof Lambda ? err('#CALC!') : out;
    });
  });
}

