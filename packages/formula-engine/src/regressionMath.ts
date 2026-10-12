/** Least-squares regression for LINEST, LOGEST, TREND and GROWTH, shared by the calculation engines. */

/** An error a result holds, as Excel writes it. */
export interface Code {
  error: string;
}

const NUM_CODE: Code = { error: '#NUM!' };
const NA_CODE: Code = { error: '#N/A' };

/** Arguments Excel refuses, with this error. */
export class RegressionError extends Error {
  constructor(readonly code: '#NUM!' | '#REF!') {
    super(code);
  }
}

export interface Data {
  y: number[];
  x: number[][];
  /** y was given as a row, so each row of x is a variable. */
  byRow: boolean;
  shape: { width: number; height: number };
}

/** A least-squares fit and the statistics LINEST reports for it. */
export interface Fit {
  /** One per variable; 0 for a variable dropped as a combination of the others. */
  coef: number[];
  b: number;
  se: (number | Code)[];
  seB: number | Code;
  r2: number | Code;
  sey: number | Code;
  f: number | Code;
  df: number;
  ssreg: number;
  ssresid: number;
}

/** The most steps a fit may take (observations × variables²), so that a huge one can't hang the editor. */
export const MAX_WORK = 100_000_000;

export const dot = (a: number[], b: number[]) => a.reduce((s, v, i) => s + v * b[i], 0);
export const mean = (v: number[]) => v.reduce((s, x) => s + x, 0) / v.length;


/**
 * Least squares by Gram-Schmidt on the centred columns (when there is a
 * constant), dropping any column that is a combination of earlier ones, as
 * Excel's LINEST does.
 */
export function leastSquares(y: number[], x: number[][], constant: boolean): Fit {
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
  const num = (v: number) => (Number.isFinite(v) ? v : NUM_CODE);
  const sey = df > 0 ? Math.sqrt(ssresid / df) : NaN;
  const se: (number | Code)[] = Array<number>(k).fill(0);
  kept.forEach((j, t) => (se[j] = num(sey * Math.sqrt(inv[t].reduce((s, v) => s + v * v, 0)))));
  // The constant's variance: sey²(1/n + |R⁻ᵀ x̄|²).
  const w = kept.map((_, u) => inv.reduce((s, row, t) => s + row[u] * xBar[kept[t]], 0));
  return {
    coef,
    b,
    se,
    seB: constant ? num(sey * Math.sqrt(1 / n + dot(w, w))) : NA_CODE,
    r2: num(ssreg / sstot),
    sey: num(sey),
    f: m && ssresid > 0 ? num(ssreg / m / (ssresid / df)) : NUM_CODE,
    df,
    ssreg,
    ssresid,
  };
}

/** How many variables known_x's holds for known_y's of the given size, or undefined if they can't go together. */
export function variables(y: { width: number; height: number }, x: { width: number; height: number }): number | undefined {
  if (y.width === 1) return x.height === y.height ? x.width : undefined;
  if (y.height === 1) return x.width === y.width ? x.height : undefined;
  return x.width === y.width && x.height === y.height ? 1 : undefined;
}

/** known_y's and known_x's lined up, as numbers already checked; known_x's left out are 1, 2, 3…. */
export function prepare(yRows: number[][], xRows: number[][] | undefined, log: boolean): Data {
  const shape = { width: yRows[0].length, height: yRows.length };
  let y = yRows.flat();
  if (log) {
    if (y.some((v) => v <= 0)) throw new RegressionError('#NUM!');
    y = y.map(Math.log);
  }
  const byRow = shape.height === 1 && shape.width > 1;
  if (!xRows) return { y, x: y.map((_, i) => [i + 1]), byRow, shape };
  const k = variables(shape, { width: xRows[0].length, height: xRows.length });
  if (k === undefined) throw new RegressionError('#REF!');
  const x = shape.width === 1 ? xRows : shape.height === 1 ? y.map((_, i) => xRows.map((row) => row[i])) : xRows.flat().map((v) => [v]);
  if (y.length * x[0].length ** 2 > MAX_WORK) throw new RegressionError('#NUM!');
  return { y, x, byRow, shape };
}

/** TREND (or GROWTH, on logarithms) at new_x's, which default to known_x's, as rows. */
export function project(d: Data, constant: boolean, log: boolean, news: number[][] | undefined): number[][] {
  const fit = leastSquares(d.y, d.x, constant);
  const at = (row: number[]) => (log ? Math.exp(fit.b + dot(fit.coef, row)) : fit.b + dot(fit.coef, row));
  const k = d.x[0].length;
  if (!news) {
    const out = d.x.map(at);
    return d.byRow ? [out] : Array.from({ length: d.shape.height }, (_, r) => out.slice(r * d.shape.width, (r + 1) * d.shape.width));
  }
  if (k === 1) return news.map((row) => row.map((v) => at([v])));
  if (d.byRow) {
    if (news.length !== k) throw new RegressionError('#REF!');
    return [news[0].map((_, i) => at(news.map((row) => row[i])))];
  }
  if (news[0].length !== k) throw new RegressionError('#REF!');
  return news.map((row) => [at(row)]);
}

/** LINEST's (or LOGEST's) rows: the coefficients, last variable first and the constant last, and with stats four rows more. */
export function estimate(d: Data, constant: boolean, log: boolean, stats: boolean): (number | Code)[][] {
  const fit = leastSquares(d.y, d.x, constant);
  const k = fit.coef.length;
  const values = [...fit.coef].reverse().concat(fit.b);
  const top = log ? values.map(Math.exp) : values;
  if (!stats) return [top];
  const na = (n: number) => Array<Code>(Math.max(0, k + 1 - n)).fill(NA_CODE);
  return [top, [...[...fit.se].reverse(), fit.seB], [fit.r2, fit.sey, ...na(2)], [fit.f, fit.df, ...na(2)], [fit.ssreg, fit.ssresid, ...na(2)]];
}
