import { ArraySize, CellError, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, array, arraySize, error, grid, literalBoolean, omitted, plain, type Call, type Grid, type Node, type State, type Value } from './common';
import { estimate, mean, prepare, project, RegressionError, variables, type Data } from '../regressionMath';

/** Observations of y with the variables they were seen at: x[i] holds observation i's values. */
/** Every value as a number, or the error the first non-number gives. */
function numbersOf(rows: Grid): number[][] | CellError {
  for (const row of rows) {
    for (const v of row) {
      if (v instanceof CellError) return v;
      if (typeof v !== 'number') return error();
    }
  }
  return rows as number[][];
}

const codeError = (code: string) => error(code === '#NUM!' ? ErrorType.NUM : code === '#N/A' ? ErrorType.NA : code === '#REF!' ? ErrorType.REF : ErrorType.VALUE);

/** Runs shared regression code, with the error it refuses arguments with. */
function shared<T>(f: () => T): T | CellError {
  try {
    return f();
  } catch (e) {
    if (e instanceof RegressionError) return codeError(e.code);
    throw e;
  }
}

export class RegressionFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    FORECAST: { method: 'forecast', parameters: forecastParameters() },
    'FORECAST.LINEAR': { method: 'forecast', parameters: forecastParameters() },
    INTERCEPT: { method: 'intercept', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }] },
    TREND: { method: 'trend', sizeOfResultArrayMethod: 'trendSize', vectorizationForbidden: true, parameters: trendParameters() },
    GROWTH: { method: 'growth', sizeOfResultArrayMethod: 'trendSize', vectorizationForbidden: true, parameters: trendParameters() },
    LINEST: { method: 'linest', sizeOfResultArrayMethod: 'linestSize', vectorizationForbidden: true, parameters: linestParameters() },
    LOGEST: { method: 'logest', sizeOfResultArrayMethod: 'linestSize', vectorizationForbidden: true, parameters: linestParameters() },
  };

  forecast(ast: Call, state: State) {
    return this.run(ast, state, 'FORECAST', (x: number, ys: SimpleRangeValue, xs: SimpleRangeValue) => {
      const line = this.line(ys, xs);
      return line instanceof CellError ? line : line.b + line.slope * x;
    });
  }

  intercept(ast: Call, state: State) {
    return this.run(ast, state, 'INTERCEPT', (ys: SimpleRangeValue, xs: SimpleRangeValue) => {
      const line = this.line(ys, xs);
      return line instanceof CellError ? line : line.b;
    });
  }

  /** The straight line through pairs where both values are numbers. */
  private line(ys: SimpleRangeValue, xs: SimpleRangeValue): { slope: number; b: number } | CellError {
    const y = ys.valuesFromTopLeftCorner().map(plain);
    const x = xs.valuesFromTopLeftCorner().map(plain);
    if (y.length !== x.length) return error(ErrorType.NA);
    const pairs: [number, number][] = [];
    for (let i = 0; i < y.length; i++) {
      if (y[i] instanceof CellError) return y[i] as CellError;
      if (x[i] instanceof CellError) return x[i] as CellError;
      if (typeof y[i] === 'number' && typeof x[i] === 'number') pairs.push([x[i] as number, y[i] as number]);
    }
    if (!pairs.length) return error(ErrorType.DIV_BY_ZERO);
    const xBar = mean(pairs.map((p) => p[0]));
    const yBar = mean(pairs.map((p) => p[1]));
    const sxx = pairs.reduce((s, [xi]) => s + (xi - xBar) ** 2, 0);
    if (sxx === 0) return error(ErrorType.DIV_BY_ZERO);
    const slope = pairs.reduce((s, [xi, yi]) => s + (xi - xBar) * (yi - yBar), 0) / sxx;
    return { slope, b: yBar - slope * xBar };
  }

  /** known_y's and known_x's checked and lined up; known_x's left out are 1, 2, 3…. */
  private data(ys: SimpleRangeValue, xs: SimpleRangeValue | undefined, log: boolean): Data | CellError {
    const yRows = numbersOf(grid(ys));
    if (yRows instanceof CellError) return yRows;
    const xRows = xs ? numbersOf(grid(xs)) : undefined;
    if (xRows instanceof CellError) return xRows;
    return shared(() => prepare(yRows, xRows, log));
  }

  trend(ast: Call, state: State) {
    return this.run(ast, state, 'TREND', this.projection(ast, false));
  }

  growth(ast: Call, state: State) {
    return this.run(ast, state, 'GROWTH', this.projection(ast, true));
  }

  /** TREND (or GROWTH, on logarithms) at new_x's, which default to known_x's. */
  private projection(ast: Call, log: boolean) {
    const [, hasX, hasNew] = ast.args.map((a) => !omitted(a));
    return (ys: SimpleRangeValue, xs: SimpleRangeValue, newXs: SimpleRangeValue, constant: boolean) => {
      const d = this.data(ys, hasX ? xs : undefined, log);
      if (d instanceof CellError) return d;
      const news = hasNew ? numbersOf(grid(newXs)) : undefined;
      if (news instanceof CellError) return news;
      const rows = shared(() => project(d, constant, log, news));
      return rows instanceof CellError ? rows : array(rows);
    };
  }

  trendSize(ast: Call, state: State) {
    const [ys, xs, news] = ast.args;
    const y = this.sizeOf(ys, state);
    const x = omitted(xs) ? undefined : this.sizeOf(xs, state);
    const k = !omitted(xs) && x && y ? variables(y, x) : 1;
    if (!y || k === undefined || (!omitted(xs) && !x)) return ArraySize.error();
    if (omitted(news)) return arraySize(y.width, y.height);
    const n = this.sizeOf(news, state);
    if (!n) return ArraySize.error();
    if (k === 1) return arraySize(n.width, n.height);
    return y.height === 1 ? arraySize(n.width, 1) : arraySize(1, n.height);
  }

  linest(ast: Call, state: State) {
    return this.run(ast, state, 'LINEST', this.estimate(ast, false));
  }

  logest(ast: Call, state: State) {
    return this.run(ast, state, 'LOGEST', this.estimate(ast, true));
  }

  /** LINEST's coefficients, last variable first and the constant last, and with stats four rows more. */
  private estimate(ast: Call, log: boolean) {
    const hasX = !omitted(ast.args[1]);
    return (ys: SimpleRangeValue, xs: SimpleRangeValue, constant: boolean, stats: boolean) => {
      const d = this.data(ys, hasX ? xs : undefined, log);
      if (d instanceof CellError) return d;
      return array(estimate(d, constant, log, stats).map((row) => row.map((v): Value => (typeof v === 'number' ? v : codeError(v.error)))));
    };
  }

  linestSize(ast: Call, state: State) {
    const [ys, xs, , stats] = ast.args;
    const y = this.sizeOf(ys, state);
    const x = omitted(xs) ? { width: 1, height: 1 } : this.sizeOf(xs, state);
    const k = omitted(xs) ? 1 : y && x ? variables(y, x) : undefined;
    if (k === undefined) return ArraySize.error();
    // Statistics only known once calculated may leave four empty rows.
    const wanted = omitted(stats) ? false : literalBoolean(stats as Node);
    return arraySize(k + 1, wanted === false ? 1 : 5);
  }
}

function forecastParameters() {
  return [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }];
}

function trendParameters() {
  return [
    { argumentType: FunctionArgumentType.RANGE },
    { argumentType: FunctionArgumentType.RANGE, optionalArg: true },
    { argumentType: FunctionArgumentType.RANGE, optionalArg: true },
    { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true, emptyAsDefault: true },
  ];
}

function linestParameters() {
  return [
    { argumentType: FunctionArgumentType.RANGE },
    { argumentType: FunctionArgumentType.RANGE, optionalArg: true },
    { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: true, emptyAsDefault: true },
    { argumentType: FunctionArgumentType.BOOLEAN, defaultValue: false, emptyAsDefault: true },
  ];
}
