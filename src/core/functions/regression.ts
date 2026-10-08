import { ArraySize, CellError, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, array, arraySize, error, grid, literalBoolean, omitted, plain, type Call, type Grid, type Node, type State, type Value } from './common';

/** Observations of y with the variables they were seen at: x[i] holds observation i's values. */
interface Data {
  y: number[];
  x: number[][];
  /** y was given as a row, so each row of x is a variable. */
  byRow: boolean;
  shape: { width: number; height: number };
}

/** A least-squares fit and the statistics LINEST reports for it. */
interface Fit {
  /** One per variable; 0 for a variable dropped as a combination of the others. */
  coef: number[];
  b: number;
  se: (number | CellError)[];
  seB: number | CellError;
  r2: number | CellError;
  sey: number | CellError;
  f: number | CellError;
  df: number;
  ssreg: number;
  ssresid: number;
}

/** The most steps a fit may take (observations × variables²), so that a huge one can't hang the editor. */
const MAX_WORK = 100_000_000;

const dot = (a: number[], b: number[]) => a.reduce((s, v, i) => s + v * b[i], 0);
const mean = (v: number[]) => v.reduce((s, x) => s + x, 0) / v.length;

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

/**
 * Least squares by Gram-Schmidt on the centred columns (when there is a
 * constant), dropping any column that is a combination of earlier ones, as
 * Excel's LINEST does.
 */
function leastSquares(y: number[], x: number[][], constant: boolean): Fit {
  const n = y.length;
  const k = x[0].length;
  const xBar = Array.from({ length: k }, (_, j) => (constant ? mean(x.map((r) => r[j])) : 0));
  const yBar = constant ? mean(y) : 0;
  const q: number[][] = [];
  const kept: number[] = [];
  const r: number[][] = [];
  for (let j = 0; j < k; j++) {
    const column = x.map((row) => row[j] - xBar[j]);
    const v = column.slice();
    const rj = Array<number>(kept.length).fill(0);
    // Twice over, which keeps the columns orthogonal to working precision.
    for (let pass = 0; pass < 2; pass++) {
      q.forEach((qt, t) => {
        const p = dot(qt, v);
        rj[t] += p;
        for (let i = 0; i < n; i++) v[i] -= p * qt[i];
      });
    }
    const norm = Math.sqrt(dot(v, v));
    if (norm <= 1e-12 * Math.sqrt(dot(column, column)) || norm === 0) continue;
    kept.push(j);
    q.push(v.map((e) => e / norm));
    r.push(Array<number>(k).fill(0));
    rj.forEach((p, t) => (r[t][j] = p));
    r[kept.length - 1][j] = norm;
  }
  const m = kept.length;
  // Solve R c = Qᵀy, and invert R for the standard errors.
  const qy = q.map((qt) => dot(qt, y.map((v) => v - yBar)));
  const c = Array<number>(m).fill(0);
  for (let t = m - 1; t >= 0; t--) {
    let s = qy[t];
    for (let u = t + 1; u < m; u++) s -= r[t][kept[u]] * c[u];
    c[t] = s / r[t][kept[t]];
  }
  const inv = Array.from({ length: m }, () => Array<number>(m).fill(0));
  for (let col = 0; col < m; col++) {
    for (let t = col; t >= 0; t--) {
      let s = t === col ? 1 : 0;
      for (let u = t + 1; u <= col; u++) s -= r[t][kept[u]] * inv[u][col];
      inv[t][col] = s / r[t][kept[t]];
    }
  }
  const coef = Array<number>(k).fill(0);
  kept.forEach((j, t) => (coef[j] = c[t]));
  const b = constant ? yBar - dot(coef, xBar) : 0;
  const fitted = x.map((row) => b + dot(coef, row));
  const ssresid = y.reduce((s, v, i) => s + (v - fitted[i]) ** 2, 0);
  const ssreg = fitted.reduce((s, v) => s + (v - yBar) ** 2, 0);
  const sstot = y.reduce((s, v) => s + (v - yBar) ** 2, 0);
  const df = n - m - (constant ? 1 : 0);
  const num = (v: number) => (Number.isFinite(v) ? v : error(ErrorType.NUM));
  const sey = df > 0 ? Math.sqrt(ssresid / df) : NaN;
  const se: (number | CellError)[] = Array<number>(k).fill(0);
  kept.forEach((j, t) => (se[j] = num(sey * Math.sqrt(inv[t].reduce((s, v) => s + v * v, 0)))));
  // The constant's variance: sey²(1/n + |R⁻ᵀ x̄|²).
  const w = kept.map((_, u) => inv.reduce((s, row, t) => s + row[u] * xBar[kept[t]], 0));
  return {
    coef,
    b,
    se,
    seB: constant ? num(sey * Math.sqrt(1 / n + dot(w, w))) : error(ErrorType.NA),
    r2: num(ssreg / sstot),
    sey: num(sey),
    f: m && ssresid > 0 ? num(ssreg / m / (ssresid / df)) : error(ErrorType.NUM),
    df,
    ssreg,
    ssresid,
  };
}

/** How many variables known_x's holds for known_y's of the given size, or undefined if they can't go together. */
function variables(y: { width: number; height: number }, x: { width: number; height: number }): number | undefined {
  if (y.width === 1) return x.height === y.height ? x.width : undefined;
  if (y.height === 1) return x.width === y.width ? x.height : undefined;
  return x.width === y.width && x.height === y.height ? 1 : undefined;
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
    const shape = { width: yRows[0].length, height: yRows.length };
    let y = yRows.flat();
    if (log) {
      if (y.some((v) => v <= 0)) return error(ErrorType.NUM);
      y = y.map(Math.log);
    }
    const byRow = shape.height === 1 && shape.width > 1;
    if (!xs) return { y, x: y.map((_, i) => [i + 1]), byRow, shape };
    const xRows = numbersOf(grid(xs));
    if (xRows instanceof CellError) return xRows;
    const k = variables(shape, { width: xRows[0].length, height: xRows.length });
    if (k === undefined) return error(ErrorType.REF);
    const x = shape.width === 1 ? xRows : shape.height === 1 ? y.map((_, i) => xRows.map((row) => row[i])) : xRows.flat().map((v) => [v]);
    return y.length * x[0].length ** 2 > MAX_WORK ? error(ErrorType.NUM) : { y, x, byRow, shape };
  }

  trend(ast: Call, state: State) {
    return this.run(ast, state, 'TREND', this.projection(ast, false));
  }

  growth(ast: Call, state: State) {
    return this.run(ast, state, 'GROWTH', this.projection(ast, true));
  }

  /** TREND (or GROWTH, on logarithms) at new_x's, which default to known_x's. */
  private projection(ast: Call, log: boolean) {
    // An argument left empty counts as left out.
    const [, hasX, hasNew] = ast.args.map((a) => !omitted(a));
    return (ys: SimpleRangeValue, xs: SimpleRangeValue, newXs: SimpleRangeValue, constant: boolean) => {
      const d = this.data(ys, hasX ? xs : undefined, log);
      const news = hasNew ? newXs : undefined;
      if (d instanceof CellError) return d;
      const fit = leastSquares(d.y, d.x, constant);
      const at = (row: number[]) => (log ? Math.exp(fit.b + dot(fit.coef, row)) : fit.b + dot(fit.coef, row));
      const k = d.x[0].length;
      if (!news) {
        const out = d.x.map(at);
        return array(d.byRow ? [out] : Array.from({ length: d.shape.height }, (_, r) => out.slice(r * d.shape.width, (r + 1) * d.shape.width)));
      }
      const rows = numbersOf(grid(news));
      if (rows instanceof CellError) return rows;
      // One variable: a result for each new x, in its place. Several: one per row (or column) of new values.
      if (k === 1) return array(rows.map((row) => row.map((v) => at([v]))));
      if (d.byRow) return rows.length === k ? array([rows[0].map((_, i) => at(rows.map((row) => row[i])))]) : error(ErrorType.REF);
      return rows[0].length === k ? array(rows.map((row) => [at(row)])) : error(ErrorType.REF);
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
      const fit = leastSquares(d.y, d.x, constant);
      const k = fit.coef.length;
      const values = [...fit.coef].reverse().concat(fit.b);
      const top: Value[] = log ? values.map(Math.exp) : values;
      if (!stats) return array([top]);
      const na = (n: number) => Array<Value>(Math.max(0, k + 1 - n)).fill(error(ErrorType.NA));
      return array([
        top,
        [...[...fit.se].reverse(), fit.seB],
        [fit.r2, fit.sey, ...na(2)],
        [fit.f, fit.df, ...na(2)],
        [fit.ssreg, fit.ssresid, ...na(2)],
      ]);
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
