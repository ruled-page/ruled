import { ArraySize, CellError, EmptyValue, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, MAX_RESULT_CELLS, array, arraySize, calcError, compareValues, error, grid, literalNumber, omitted, plain, type Call, type Grid, type Node, type State, type Value } from './common';

/** What a count argument left empty stands for: all of them. */
const ALL = Number.POSITIVE_INFINITY;

const count = { argumentType: FunctionArgumentType.NUMBER, defaultValue: ALL, emptyAsDefault: true };
const optionalCount = { ...count, optionalArg: true };

const transpose = (rows: Grid): Grid => rows[0].map((_, c) => rows.map((r) => r[c]));

/** The first or (counting from the end) last n of a list. */
const take = <T>(list: T[], n: number) => (n >= 0 ? list.slice(0, n) : list.slice(Math.max(0, list.length + n)));
const drop = <T>(list: T[], n: number) => (n >= 0 ? list.slice(n) : list.slice(0, Math.max(0, list.length + n)));

/** A count as written, ALL when left out, or undefined when only known once calculated. */
function literalCount(node: Node | undefined): number | undefined {
  if (omitted(node)) return ALL;
  const n = literalNumber(node!);
  return n === undefined ? undefined : Math.trunc(n);
}

/** Splits a list into rows of n, the last one padded. */
function wrap(values: Value[], n: number, pad: Value): Grid {
  const rows: Grid = [];
  for (let i = 0; i < values.length; i += n) {
    const row = values.slice(i, i + n);
    rows.push(row.length < n ? [...row, ...Array<Value>(n - row.length).fill(pad)] : row);
  }
  return rows;
}

const padding = (pad: Value | undefined) => (pad === undefined ? error(ErrorType.NA) : pad);

export class ArrayFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    TAKE: {
      method: 'take',
      sizeOfResultArrayMethod: 'takeSize',
      vectorizationForbidden: true,
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, count, optionalCount],
    },
    DROP: {
      method: 'drop',
      sizeOfResultArrayMethod: 'dropSize',
      vectorizationForbidden: true,
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, count, optionalCount],
    },
    CHOOSEROWS: {
      method: 'chooserows',
      sizeOfResultArrayMethod: 'chooserowsSize',
      vectorizationForbidden: true,
      repeatLastArgs: 1,
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }],
    },
    CHOOSECOLS: {
      method: 'choosecols',
      sizeOfResultArrayMethod: 'choosecolsSize',
      vectorizationForbidden: true,
      repeatLastArgs: 1,
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY }],
    },
    TOCOL: { method: 'tocol', sizeOfResultArrayMethod: 'tocolSize', vectorizationForbidden: true, parameters: flattenParameters() },
    TOROW: { method: 'torow', sizeOfResultArrayMethod: 'torowSize', vectorizationForbidden: true, parameters: flattenParameters() },
    WRAPROWS: { method: 'wraprows', sizeOfResultArrayMethod: 'wraprowsSize', vectorizationForbidden: true, parameters: wrapParameters() },
    WRAPCOLS: { method: 'wrapcols', sizeOfResultArrayMethod: 'wrapcolsSize', vectorizationForbidden: true, parameters: wrapParameters() },
    EXPAND: {
      method: 'expand',
      sizeOfResultArrayMethod: 'expandSize',
      vectorizationForbidden: true,
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, count, optionalCount, { argumentType: FunctionArgumentType.SCALAR, optionalArg: true }],
    },
    SORTBY: {
      method: 'sortby',
      sizeOfResultArrayMethod: 'sortbySize',
      vectorizationForbidden: true,
      repeatLastArgs: 2,
      parameters: [
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 1, emptyAsDefault: true },
      ],
    },
    RANDARRAY: {
      method: 'randarray',
      sizeOfResultArrayMethod: 'randarraySize',
      vectorizationForbidden: true,
      isVolatile: true,
      parameters: [
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 1, emptyAsDefault: true },
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 1, emptyAsDefault: true },
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0, emptyAsDefault: true },
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 1, emptyAsDefault: true },
        { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: false, emptyAsDefault: true },
      ],
    },
  };

  take(ast: Call, state: State) {
    return this.run(ast, state, 'TAKE', (range: SimpleRangeValue, rows: number, cols: number) => {
      const r = Math.trunc(rows);
      const c = Math.trunc(cols);
      if (r === 0 || c === 0) return calcError();
      return array(take(grid(range), r).map((row) => take(row, c)));
    });
  }

  takeSize(ast: Call, state: State) {
    const size = this.sizeOf(ast.args[0], state);
    const r = literalCount(ast.args[1]);
    const c = literalCount(ast.args[2]);
    if (!size || r === 0 || c === 0) return ArraySize.error();
    // A count only known once calculated can't take more than there is.
    return arraySize(Math.min(Math.abs(c ?? ALL), size.width), Math.min(Math.abs(r ?? ALL), size.height));
  }

  drop(ast: Call, state: State) {
    return this.run(ast, state, 'DROP', (range: SimpleRangeValue, rows: number, cols: number) => {
      const r = rows === ALL ? 0 : Math.trunc(rows);
      const c = cols === ALL ? 0 : Math.trunc(cols);
      const out = drop(grid(range), r).map((row) => drop(row, c));
      return out.length && out[0].length ? array(out) : calcError();
    });
  }

  dropSize(ast: Call, state: State) {
    const size = this.sizeOf(ast.args[0], state);
    if (!size) return ArraySize.error();
    const left = (n: number | undefined, of: number) => (n === undefined || n === ALL ? of : of - Math.min(Math.abs(n), of));
    return arraySize(left(literalCount(ast.args[2]), size.width), left(literalCount(ast.args[1]), size.height));
  }

  chooserows(ast: Call, state: State) {
    return this.run(ast, state, 'CHOOSEROWS', (range: SimpleRangeValue, ...picks: unknown[]) => {
      const rows = grid(range);
      const which = this.indices(picks, rows.length);
      return which instanceof CellError ? which : array(which.map((i) => rows[i]));
    });
  }

  chooserowsSize(ast: Call, state: State) {
    return this.chooseSize(ast, state, false);
  }

  choosecols(ast: Call, state: State) {
    return this.run(ast, state, 'CHOOSECOLS', (range: SimpleRangeValue, ...picks: unknown[]) => {
      const rows = grid(range);
      const which = this.indices(picks, rows[0].length);
      return which instanceof CellError ? which : array(rows.map((row) => which.map((i) => row[i])));
    });
  }

  choosecolsSize(ast: Call, state: State) {
    return this.chooseSize(ast, state, true);
  }

  /** The positions CHOOSEROWS or CHOOSECOLS picks, from 1 or (negative) from the end. */
  private indices(picks: unknown[], length: number): number[] | CellError {
    const out: number[] = [];
    for (const v of picks.flatMap((p) => grid(p).flat())) {
      const n = v instanceof CellError ? v : plain(this.coerceScalarToNumberOrError(v as never));
      if (n instanceof CellError) return n;
      const i = Math.trunc(n as number);
      if (i === 0 || Math.abs(i) > length) return error();
      out.push(i > 0 ? i - 1 : length + i);
    }
    return out;
  }

  /** As many rows (or columns) as the picks list numbers. */
  private chooseSize(ast: Call, state: State, columns: boolean) {
    const [source, ...picks] = ast.args;
    const size = this.sizeOf(source, state);
    const sizes = picks.map((p) => this.sizeOf(p, state));
    if (!size || !picks.length || sizes.some((s) => !s)) return ArraySize.error();
    const n = sizes.reduce((sum, s) => sum + s!.width * s!.height, 0);
    return columns ? arraySize(n, size.height) : arraySize(size.width, n);
  }

  tocol(ast: Call, state: State) {
    return this.run(ast, state, 'TOCOL', (range: SimpleRangeValue, ignore: number, byColumn: boolean) => {
      const values = this.flatten(range, ignore, byColumn);
      return values instanceof CellError ? values : array(values.map((v) => [v]));
    });
  }

  tocolSize(ast: Call, state: State) {
    const n = this.flattenSize(ast, state);
    return n === undefined ? ArraySize.error() : arraySize(1, n);
  }

  torow(ast: Call, state: State) {
    return this.run(ast, state, 'TOROW', (range: SimpleRangeValue, ignore: number, byColumn: boolean) => {
      const values = this.flatten(range, ignore, byColumn);
      return values instanceof CellError ? values : array([values]);
    });
  }

  torowSize(ast: Call, state: State) {
    const n = this.flattenSize(ast, state);
    return n === undefined ? ArraySize.error() : arraySize(n, 1);
  }

  /** TOCOL and TOROW's values in order, leaving out blanks (ignore 1), errors (2) or both (3). */
  private flatten(range: SimpleRangeValue, ignore: number, byColumn: boolean): Value[] | CellError {
    const mode = Math.trunc(ignore);
    if (mode < 0 || mode > 3) return error();
    const rows = grid(range);
    const values = (byColumn ? transpose(rows) : rows).flat().filter((v) => !((mode & 1 && v === EmptyValue) || (mode & 2 && v instanceof CellError)));
    return values.length ? values : calcError();
  }

  /** All the values; when some are left out the rest are padded with empty cells. */
  private flattenSize(ast: Call, state: State): number | undefined {
    const size = this.sizeOf(ast.args[0], state);
    return size ? size.width * size.height : undefined;
  }

  wraprows(ast: Call, state: State) {
    return this.run(ast, state, 'WRAPROWS', (range: SimpleRangeValue, n: number, pad?: Value) => {
      const values = this.vector(range, n);
      return values instanceof CellError ? values : array(wrap(values, Math.trunc(n), padding(pad)));
    });
  }

  wraprowsSize(ast: Call, state: State) {
    const fit = this.wrapSize(ast, state);
    return fit ? arraySize(fit[0], fit[1]) : ArraySize.error();
  }

  wrapcols(ast: Call, state: State) {
    return this.run(ast, state, 'WRAPCOLS', (range: SimpleRangeValue, n: number, pad?: Value) => {
      const values = this.vector(range, n);
      return values instanceof CellError ? values : array(transpose(wrap(values, Math.trunc(n), padding(pad))));
    });
  }

  wrapcolsSize(ast: Call, state: State) {
    const fit = this.wrapSize(ast, state);
    return fit ? arraySize(fit[1], fit[0]) : ArraySize.error();
  }

  /** The values of a single row or column to wrap. */
  private vector(range: SimpleRangeValue, n: number): Value[] | CellError {
    const rows = grid(range);
    if (rows.length > 1 && rows[0].length > 1) return error();
    return Math.trunc(n) < 1 ? error(ErrorType.NUM) : rows.flat();
  }

  /** [values per row, rows] of a wrapped vector. */
  private wrapSize(ast: Call, state: State): [number, number] | undefined {
    const size = this.sizeOf(ast.args[0], state);
    const n = literalCount(ast.args[1]);
    if (!size || n === undefined || n < 1 || n === ALL || (size.width > 1 && size.height > 1)) return undefined;
    return [n, Math.ceil((size.width * size.height) / n)];
  }

  expand(ast: Call, state: State) {
    return this.run(ast, state, 'EXPAND', (range: SimpleRangeValue, rows: number, cols: number, pad?: Value) => {
      const data = grid(range);
      const h = rows === ALL ? data.length : Math.trunc(rows);
      const w = cols === ALL ? data[0].length : Math.trunc(cols);
      if (h < data.length || w < data[0].length) return error();
      if (h * w > MAX_RESULT_CELLS) return error(ErrorType.NUM);
      const filler = padding(pad);
      const out = data.map((row) => [...row, ...Array<Value>(w - row.length).fill(filler)]);
      while (out.length < h) out.push(Array<Value>(w).fill(filler));
      return array(out);
    });
  }

  expandSize(ast: Call, state: State) {
    const size = this.sizeOf(ast.args[0], state);
    const r = literalCount(ast.args[1]);
    const c = literalCount(ast.args[2]);
    if (!size || r === undefined || c === undefined) return ArraySize.error();
    const h = r === ALL ? size.height : r;
    const w = c === ALL ? size.width : c;
    return h < size.height || w < size.width ? ArraySize.error() : arraySize(w, h);
  }

  sortby(ast: Call, state: State) {
    return this.run(ast, state, 'SORTBY', (range: SimpleRangeValue, ...pairs: unknown[]) => {
      const rows = grid(range);
      const keys: { values: Value[]; order: number; columns: boolean }[] = [];
      for (let i = 0; i < pairs.length; i += 2) {
        const by = grid(pairs[i]);
        const order = pairs[i + 1] as number;
        if (order !== 1 && order !== -1) return error();
        // A column of keys sorts the rows; a row of keys sorts the columns.
        const columns = by.length === 1 && by[0].length > 1;
        if (by.length > 1 && by[0].length > 1) return error();
        if (columns ? by[0].length !== rows[0].length : by.length !== rows.length) return error();
        keys.push({ values: by.flat(), order, columns });
      }
      const columns = keys[0].columns;
      if (keys.some((k) => k.columns !== columns)) return error();
      const lines = columns ? transpose(rows) : rows;
      const order = lines.map((_, i) => i);
      order.sort((a, b) => {
        for (const k of keys) {
          const c = compareKeys(k.values[a], k.values[b], k.order);
          if (c) return c;
        }
        return 0;
      });
      const sorted = order.map((i) => lines[i]);
      return array(columns ? transpose(sorted) : sorted);
    });
  }

  sortbySize(ast: Call, state: State) {
    const size = this.sizeOf(ast.args[0], state);
    return size ? arraySize(size.width, size.height) : ArraySize.error();
  }

  randarray(ast: Call, state: State) {
    return this.run(ast, state, 'RANDARRAY', (rows: number, cols: number, min: number, max: number, whole: boolean) => {
      const h = Math.trunc(rows);
      const w = Math.trunc(cols);
      if (h < 0 || w < 0 || min > max || (whole && (!Number.isInteger(min) || !Number.isInteger(max)))) return error();
      if (h === 0 || w === 0) return calcError();
      if (h * w > MAX_RESULT_CELLS) return error(ErrorType.NUM);
      const next = whole ? () => min + Math.floor(Math.random() * (max - min + 1)) : () => min + Math.random() * (max - min);
      return array(Array.from({ length: h }, () => Array.from({ length: w }, next)));
    });
  }

  randarraySize(ast: Call) {
    const [rows, cols] = ast.args;
    const h = omitted(rows) ? 1 : literalNumber(rows!);
    const w = omitted(cols) ? 1 : literalNumber(cols!);
    return h === undefined || w === undefined ? ArraySize.error() : arraySize(Math.trunc(w), Math.trunc(h));
  }
}

/** Sort order for SORTBY's keys: Excel's order across types, reversed for -1, with empty cells always last. */
function compareKeys(a: Value, b: Value, order: number): number {
  if (a === EmptyValue || b === EmptyValue) return (a === EmptyValue ? 1 : 0) - (b === EmptyValue ? 1 : 0);
  return compareValues(a, b) * order;
}

function flattenParameters() {
  return [
    { argumentType: FunctionArgumentType.RANGE },
    { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0, emptyAsDefault: true },
    { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: false, emptyAsDefault: true },
  ];
}

function wrapParameters() {
  return [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.SCALAR, optionalArg: true }];
}
