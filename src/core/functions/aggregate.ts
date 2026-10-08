import { CellError, EmptyValue, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, error, grid, largest, plain, smallest, type Call, type Node, type State, type Value } from './common';

const sum = (v: number[]) => v.reduce((s, x) => s + x, 0);

/** Sum of squared deviations from the mean, in two passes for accuracy. */
function squares(v: number[]): number {
  const mean = sum(v) / v.length;
  return v.reduce((s, x) => s + (x - mean) ** 2, 0);
}

/** PERCENTILE.INC (inclusive) or PERCENTILE.EXC of sorted values. */
function percentile(sorted: number[], k: number, inclusive: boolean): number | CellError {
  const n = sorted.length;
  if (!n || k < 0 || k > 1) return error(ErrorType.NUM);
  const rank = inclusive ? k * (n - 1) + 1 : k * (n + 1);
  if (rank < 1 || rank > n) return error(ErrorType.NUM);
  const lo = Math.floor(rank);
  const v = sorted[lo - 1];
  return lo === n ? v : v + (rank - lo) * (sorted[lo] - v);
}

/** The values of AGGREGATE's functions 1–13, which work on all their references. */
const LISTS: Record<number, (v: number[]) => number | CellError> = {
  1: (v) => (v.length ? sum(v) / v.length : error(ErrorType.DIV_BY_ZERO)),
  2: (v) => v.length,
  4: (v) => (v.length ? largest(v) : 0),
  5: (v) => (v.length ? smallest(v) : 0),
  6: (v) => (v.length ? v.reduce((p, x) => p * x, 1) : 0),
  7: (v) => (v.length > 1 ? Math.sqrt(squares(v) / (v.length - 1)) : error(ErrorType.DIV_BY_ZERO)),
  8: (v) => (v.length ? Math.sqrt(squares(v) / v.length) : error(ErrorType.DIV_BY_ZERO)),
  9: sum,
  10: (v) => (v.length > 1 ? squares(v) / (v.length - 1) : error(ErrorType.DIV_BY_ZERO)),
  11: (v) => (v.length ? squares(v) / v.length : error(ErrorType.DIV_BY_ZERO)),
  12: (v) => {
    if (!v.length) return error(ErrorType.NUM);
    const s = [...v].sort((a, b) => a - b);
    const mid = s.length >> 1;
    return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  },
  13: (v) => {
    const counts = new Map<number, number>();
    for (const x of v) counts.set(x, (counts.get(x) ?? 0) + 1);
    let best: number | undefined;
    for (const [x, n] of counts) if (n > 1 && (best === undefined || n > counts.get(best)!)) best = x;
    return best ?? error(ErrorType.NA);
  },
};

/** Functions 14–19, which take one array and a k. */
const RANKED: Record<number, (sorted: number[], k: number) => number | CellError> = {
  14: (s, k) => {
    const i = Math.ceil(k);
    return i < 1 || i > s.length ? error(ErrorType.NUM) : s[s.length - i];
  },
  15: (s, k) => {
    const i = Math.ceil(k);
    return i < 1 || i > s.length ? error(ErrorType.NUM) : s[i - 1];
  },
  16: (s, k) => percentile(s, k, true),
  17: (s, q) => (Math.trunc(q) < 0 || Math.trunc(q) > 4 ? error(ErrorType.NUM) : percentile(s, Math.trunc(q) / 4, true)),
  18: (s, k) => percentile(s, k, false),
  19: (s, q) => (Math.trunc(q) < 1 || Math.trunc(q) > 3 ? error(ErrorType.NUM) : percentile(s, Math.trunc(q) / 4, false)),
};

/** Whether a formula calls SUBTOTAL or AGGREGATE anywhere in it. */
function hasSubtotal(node: Node | undefined): boolean {
  if (!node || typeof node !== 'object') return false;
  if (node.type === 'FUNCTION_CALL' && (node.procedureName === 'SUBTOTAL' || node.procedureName === 'AGGREGATE')) return true;
  const children = node.type === 'ARRAY' ? (node.args as unknown as Node[][]).flat() : (node.args ?? []);
  return [...children, node.expression, node.left, node.right, node.value as Node].some(hasSubtotal);
}

type Graph = { getCell(address: { sheet: number; row: number; col: number }): { formula?: Node } | undefined };
type Range = { start: { sheet: number; row: number; col: number } };

export class AggregateFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    AGGREGATE: { method: 'aggregate', repeatLastArgs: 1, parameters: [{ argumentType: FunctionArgumentType.ANY }] },
  };

  /**
   * AGGREGATE(function, options, ref1, …) or AGGREGATE(function, options,
   * array, k). Options 0–3 leave out cells holding SUBTOTAL or AGGREGATE,
   * and 2, 3, 6 and 7 leave out errors. Hidden rows (options 1, 3, 5 and 7)
   * can't be told apart here and count as shown.
   */
  aggregate(ast: Call, state: State) {
    if (ast.args.length < 3) return error();
    const head = ast.args.slice(0, 2).map((a) => (a.type === 'EMPTY' ? 0 : this.toNumber(plain(this.evaluate(a, state)))));
    for (const h of head) if (h instanceof CellError) return h;
    const [fn, options] = (head as number[]).map(Math.trunc);
    if (!(fn >= 1 && fn <= 19) || !(options >= 0 && options <= 7)) return error();
    const skipErrors = [2, 3, 6, 7].includes(options);
    const skipNested = options <= 3;
    const ranked = fn >= 14;
    if (ranked && ast.args.length !== 4) return error();
    const values: Value[] = [];
    for (const node of ast.args.slice(2, ranked ? 3 : undefined)) {
      const v = this.evaluate(node, state);
      if (v instanceof SimpleRangeValue) for (const x of this.cells(v, skipNested)) values.push(x);
      else values.push(plain(v));
    }
    // COUNT and COUNTA never fail on errors; COUNTA counts them unless they are left out.
    if (fn === 3) return values.filter((v) => v !== EmptyValue && !(skipErrors && v instanceof CellError)).length;
    const numbers: number[] = [];
    for (const v of values) {
      if (v instanceof CellError) {
        if (skipErrors || fn === 2) continue;
        return v;
      }
      if (typeof v === 'number') numbers.push(v);
    }
    if (!ranked) return LISTS[fn](numbers);
    const k = this.toNumber(plain(this.evaluate(ast.args[3], state)));
    return k instanceof CellError ? k : RANKED[fn](numbers.sort((a, b) => a - b), k);
  }

  /** A range's values, without the cells whose formulas are subtotals themselves. */
  private cells(v: SimpleRangeValue, skipNested: boolean): Value[] {
    const rows = grid(v);
    const range = v.range as unknown as Range | undefined;
    if (!skipNested || !range) return rows.flat();
    const graph = this.dependencyGraph as unknown as Graph;
    return rows.flatMap((row, r) =>
      row.filter((_, c) => !hasSubtotal(graph.getCell({ sheet: range.start.sheet, row: range.start.row + r, col: range.start.col + c })?.formula)),
    );
  }
}
