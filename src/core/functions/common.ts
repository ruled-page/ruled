import { ArraySize, CellError, EmptyValue, ErrorType, FunctionPlugin, SimpleRangeValue } from 'hyperformula';

/** A node of the engine's syntax tree, as far as these functions read one. */
export interface Node {
  type: string;
  value?: unknown;
  args?: Node[];
  procedureName?: string;
  expression?: Node;
  left?: Node;
  right?: Node;
}

/** A function call as the engine hands it to a plugin method. */
export interface Call extends Node {
  args: Node[];
}

export type State = Parameters<FunctionPlugin['runFunction']>[1];

/** A cell value, with dates, percentages and currencies as the plain numbers they hold. */
export type Value = number | string | boolean | CellError | typeof EmptyValue;

export type Grid = Value[][];

/** Excel's #CALC! error, which the engine has no type for; registerExcelFunctions gives it a name. */
export const CALC = 'CALC' as ErrorType;

export const error = (type: ErrorType = ErrorType.VALUE) => new CellError(type);
export const calcError = () => new CellError(CALC);

/** The engine wraps dates, percentages and currencies in objects holding the number as `val`. */
export function plain(v: unknown): Value {
  if (v !== null && typeof v === 'object' && !(v instanceof CellError) && 'val' in v) return (v as { val: number }).val;
  return v as Value;
}

/** A value as rows of cells: a range or array as it is, anything else as a single cell. */
export function grid(v: unknown): Grid {
  return v instanceof SimpleRangeValue ? v.data.map((row) => row.map(plain)) : [[plain(v)]];
}

/** The numbers in a range or array, passing over text, logicals and blanks; or the first error in it. */
export function rangeNumbers(v: unknown): number[] | CellError {
  const out: number[] = [];
  for (const row of grid(v)) {
    for (const x of row) {
      if (x instanceof CellError) return x;
      if (typeof x === 'number') out.push(x);
    }
  }
  return out;
}

export function array(rows: Grid): SimpleRangeValue {
  return SimpleRangeValue.onlyValues(rows as never);
}

/** Rounds the 15-digit decimal value, as Excel does, rather than the binary one. */
export function excelRound(x: number, digits: number, mode: 'half' | 'up' | 'down'): number {
  if (!Number.isFinite(x) || x === 0) return x;
  const d = Math.trunc(digits);
  const sign = x < 0 ? -1 : 1;
  const [mantissa, exp] = Math.abs(x).toExponential(14).split('e');
  // All 15 digits are already left of the rounding point.
  if (Number(exp) + d >= 14) return x;
  const shifted = Number(`${mantissa}e${Number(exp) + d}`);
  const r = mode === 'half' ? Math.round(shifted) : mode === 'up' ? Math.ceil(shifted) : Math.floor(shifted);
  return sign * Number(`${r}e${-d}`);
}

/**
 * A number as Excel writes it when it becomes text ("a"&0.1+0.2 is "a0.3"):
 * 15 significant digits, in scientific notation when very large or small.
 */
export function numberText(n: number): string {
  if (n === 0) return '0';
  const [mantissa, e] = n.toExponential(14).split('e');
  const exp = Number(e);
  if (exp >= -4 && exp < 15) return String(Number(n.toPrecision(15)));
  const m = mantissa.replace(/\.?0+$/, '');
  return `${m}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
}

/** A value as text, the way Excel converts one: TRUE, 0.3, or the text itself. */
export function valueText(v: Exclude<Value, CellError>): string {
  if (typeof v === 'number') return numberText(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return v === EmptyValue ? '' : v;
}

const collator = new Intl.Collator('en', { sensitivity: 'accent' });

/** Text compared as Excel does in lookups and sorting: ignoring case, not accents. */
export function compareText(a: string, b: string): number {
  return collator.compare(a, b);
}

/** Where a value sorts by type: numbers, text, logicals, then errors. */
function typeRank(v: Value): number {
  if (typeof v === 'number') return 0;
  if (typeof v === 'string') return 1;
  return typeof v === 'boolean' ? 2 : 3;
}

/** Excel's ascending order across types, for anything but empty cells; errors are all equal. */
export function compareValues(a: Value, b: Value): number {
  const ta = typeRank(a);
  const tb = typeRank(b);
  if (ta !== tb) return ta - tb;
  if (typeof a === 'string') return compareText(a, b as string);
  return ta < 3 ? Number(a) - Number(b) : 0;
}

/** Whether a lookup finds b when looking for a: same type, text ignoring case. */
export function sameValue(a: Value, b: Value): boolean {
  if (typeof a === 'string' && typeof b === 'string') return compareText(a, b) === 0;
  return a === b;
}

/**
 * The number an argument is written as, when it is a literal: 3, -2, "4",
 * TRUE. NaN for literal text that isn't a number, which will be an error.
 */
export function literalNumber(node: Node): number | undefined {
  switch (node.type) {
    case 'NUMBER':
      return node.value as number;
    case 'STRING': {
      const n = Number(node.value as string);
      return Number.isFinite(n) ? n : NaN;
    }
    case 'MINUS_UNARY_OP':
    case 'PLUS_UNARY_OP': {
      const n = literalNumber(node.value as Node);
      return n === undefined ? undefined : node.type === 'MINUS_UNARY_OP' ? -n : n;
    }
    case 'PARENTHESES':
      return literalNumber(node.expression!);
    case 'FUNCTION_CALL':
      if (!node.args?.length && (node.procedureName === 'TRUE' || node.procedureName === 'FALSE')) return node.procedureName === 'TRUE' ? 1 : 0;
      return undefined;
    default:
      return undefined;
  }
}

/** The logical an argument is written as, when it is a literal: TRUE, 0, "FALSE". */
export function literalBoolean(node: Node): boolean | undefined {
  while (node.type === 'PARENTHESES') node = node.expression!;
  if (node.type === 'STRING') return ({ TRUE: true, FALSE: false } as Record<string, boolean>)[(node.value as string).toUpperCase()];
  const n = literalNumber(node);
  return n === undefined ? undefined : n !== 0;
}

/** Whether an argument was left out or left empty: TAKE(A1:B9,,1). */
export const omitted = (node: Node | undefined) => !node || node.type === 'EMPTY';

const REFERENCES = new Set(['CELL_REFERENCE', 'CELL_RANGE', 'COLUMN_RANGE', 'ROW_RANGE']);

/** Arrays a function may build before it gives up, so a formula can't exhaust memory. */
export const MAX_RESULT_CELLS = 1_000_000;

/** Shared plumbing for the plugins that add Excel's functions. */
export abstract class ExcelPlugin extends FunctionPlugin {
  protected run(ast: Call, state: State, name: string, fn: (...args: never[]) => unknown): unknown {
    return this.runFunction(ast.args as never, state, this.metadata(name), fn as never);
  }

  protected evaluate(node: Node, state: State): unknown {
    return this.evaluateAst(node as never, state);
  }

  /** An error as it is shown in a cell: #N/A. */
  protected errorText(e: CellError): string {
    return (this.config as unknown as { translationPackage: { getErrorTranslation(type: string): string } }).translationPackage.getErrorTranslation(e.type);
  }

  /** What another function gives for the same arguments: one Excel keeps under a second name. */
  protected delegate(ast: Call, state: State, name: string): unknown {
    return this.evaluate({ type: 'FUNCTION_CALL', procedureName: name, args: ast.args }, state);
  }

  /** A value converted to a number the way arithmetic does: "12" is 12, "a" an error. */
  protected toNumber(v: Value): number | CellError {
    return plain(this.coerceScalarToNumberOrError(v as never)) as number | CellError;
  }

  /**
   * The numbers a list of arguments holds, following Excel's rule for
   * functions such as AVERAGE: numbers, logicals and numeric text typed as
   * arguments count, while in references and arrays only numbers do. Returns
   * the first error met instead, if there is one.
   */
  protected numbers(args: Node[], state: State): number[] | CellError {
    const out: number[] = [];
    for (const node of args) {
      const v = node.type === 'EMPTY' ? 0 : this.evaluate(node, state);
      const found = v instanceof SimpleRangeValue || REFERENCES.has(node.type) || node.type === 'NAMED_EXPRESSION' ? rangeNumbers(v) : this.toNumber(plain(v));
      if (found instanceof CellError) return found;
      out.push(...(typeof found === 'number' ? [found] : found));
    }
    return out;
  }

  /**
   * The size an argument's value will have, when it can be known before the
   * formula runs. Unknown when the argument, or anything inside it, has a
   * size that depends on values (SEQUENCE(A1)), since a result predicted
   * smaller than it turns out would break the engine.
   */
  protected sizeOf(node: Node | undefined, state: State): { width: number; height: number } | undefined {
    if (omitted(node)) return undefined;
    while (node!.type === 'PARENTHESES') node = node!.expression!;
    if (REFERENCES.has(node!.type)) return this.arraySizeForAst(node as never, state);
    if (this.unknownSize(node!, state)) return undefined;
    const s = this.arraySizeForAst(node as never, state);
    return { width: s.width, height: s.height };
  }

  private unknownSize(node: Node, state: State): boolean {
    while (node.type === 'PARENTHESES') node = node.expression!;
    if (REFERENCES.has(node.type) || node.type === 'EMPTY' || node.type === 'ERROR') return false;
    if (this.arraySizeForAst(node as never, state).isRef) return true;
    const children = [
      ...(node.type === 'ARRAY' ? (node.args as unknown as Node[][]).flat() : (node.args ?? [])),
      ...[node.expression, node.left, node.right, node.type.endsWith('_OP') ? (node.value as Node) : undefined].filter((n): n is Node => !!n),
    ];
    return children.some((c) => this.unknownSize(c, state));
  }
}

/** A size to predict for an array result, or "unknown" when it is empty or too large to build. */
export function arraySize(width: number, height: number): ArraySize {
  if (!(width >= 1 && height >= 1) || width * height > MAX_RESULT_CELLS || !Number.isInteger(width) || !Number.isInteger(height)) return ArraySize.error();
  return new ArraySize(width, height);
}
