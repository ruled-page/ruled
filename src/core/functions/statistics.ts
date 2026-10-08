import { ArraySize, CellError, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, array, arraySize, error, plain, rangeNumbers, type Call, type State, type Value } from './common';

/** A criterion parsed by the engine, as SUMIFS and COUNTIFS use them. */
type Criteria = { fromCellValue(raw: unknown, helper: unknown): { lambda: (v: unknown) => boolean } | undefined };

/** Rounds away float noise before a floor, so 0.7 × 1000 is 700 and not 699.999…. */
const floor15 = (x: number) => Math.floor(Number(x.toPrecision(15)));

/** The values that occur most often (at least twice), in the order they first appear. */
function modes(values: number[]): number[] {
  const counts = new Map<number, number>();
  for (const v of values) counts.set(v, (counts.get(v) ?? 0) + 1);
  const most = Math.max(0, ...counts.values());
  return most < 2 ? [] : [...counts].filter(([, n]) => n === most).map(([v]) => v);
}

/** Where x falls among sorted values, as a fraction: PERCENTRANK, inclusive or exclusive of the ends. */
function percentRank(sorted: number[], x: number, inclusive: boolean): number {
  const n = sorted.length;
  // How many values are smaller than x, and the first one that isn't.
  let below = 0;
  while (below < n && sorted[below] < x) below++;
  if (sorted[below] === x) {
    if (inclusive) return n === 1 ? 1 : below / (n - 1);
    return (below + 1) / (n + 1);
  }
  const fraction = (x - sorted[below - 1]) / (sorted[below] - sorted[below - 1]);
  return inclusive ? (below - 1 + fraction) / (n - 1) : (below + fraction) / (n + 1);
}

/** ln(n!) by summing, exact enough for the binomial sums here. */
function logFactorial(n: number): number {
  let s = 0;
  for (let i = 2; i <= n; i++) s += Math.log(i);
  return s;
}

/** P(lo ≤ X ≤ hi) for X binomial with n trials and success chance p. */
function binomialRange(n: number, p: number, lo: number, hi: number): number {
  if (p === 0 || p === 1) {
    const certain = p === 0 ? 0 : n;
    return certain >= lo && certain <= hi ? 1 : 0;
  }
  // Each term from logarithms, as their products can underflow long before the sum does.
  let lnChoose = logFactorial(n) - logFactorial(lo) - logFactorial(n - lo);
  let sum = 0;
  for (let k = lo; k <= hi; k++) {
    sum += Math.exp(lnChoose + k * Math.log(p) + (n - k) * Math.log1p(-p));
    lnChoose += Math.log(n - k) - Math.log(k + 1);
  }
  return sum;
}

export class StatisticalFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    AVERAGEIFS: {
      method: 'averageifs',
      repeatLastArgs: 2,
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NOERROR }],
    },
    MODE: { method: 'mode', repeatLastArgs: 1, parameters: [{ argumentType: FunctionArgumentType.ANY }] },
    'MODE.SNGL': { method: 'mode', repeatLastArgs: 1, parameters: [{ argumentType: FunctionArgumentType.ANY }] },
    'MODE.MULT': { method: 'modemult', sizeOfResultArrayMethod: 'modemultSize', repeatLastArgs: 1, parameters: [{ argumentType: FunctionArgumentType.ANY }] },
    FREQUENCY: {
      method: 'frequency',
      sizeOfResultArrayMethod: 'frequencySize',
      vectorizationForbidden: true,
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }],
    },
    TRIMMEAN: { method: 'trimmean', parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER }] },
    KURT: { method: 'kurt', repeatLastArgs: 1, parameters: [{ argumentType: FunctionArgumentType.ANY }] },
    PERCENTRANK: { method: 'percentrankinc', parameters: percentRankParameters() },
    'PERCENTRANK.INC': { method: 'percentrankinc', parameters: percentRankParameters() },
    'PERCENTRANK.EXC': { method: 'percentrankexc', parameters: percentRankParameters() },
    PERMUT: { method: 'permut', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }] },
    PERMUTATIONA: { method: 'permutationa', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }] },
    'RANK.AVG': {
      method: 'rankavg',
      parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0 }],
    },
    PROB: {
      method: 'prob',
      parameters: [
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.NUMBER },
        { argumentType: FunctionArgumentType.NUMBER, optionalArg: true },
      ],
    },
    GESTEP: { method: 'gestep', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0 }] },
    'ERF.PRECISE': { method: 'erfprecise', parameters: [{ argumentType: FunctionArgumentType.NUMBER }] },
    'ERFC.PRECISE': { method: 'erfcprecise', parameters: [{ argumentType: FunctionArgumentType.NUMBER }] },
    PERCENTOF: { method: 'percentof', repeatLastArgs: 1, parameters: [{ argumentType: FunctionArgumentType.ANY }] },
    'BINOM.DIST.RANGE': {
      method: 'binomdistrange',
      parameters: [
        { argumentType: FunctionArgumentType.NUMBER },
        { argumentType: FunctionArgumentType.NUMBER },
        { argumentType: FunctionArgumentType.NUMBER },
        { argumentType: FunctionArgumentType.NUMBER, optionalArg: true },
      ],
    },
  };

  averageifs(ast: Call, state: State) {
    return this.run(ast, state, 'AVERAGEIFS', (values: SimpleRangeValue, ...conditions: unknown[]) => {
      const tests: { cells: Value[]; test: (v: unknown) => boolean }[] = [];
      for (let i = 0; i < conditions.length; i += 2) {
        const range = conditions[i] as SimpleRangeValue;
        // Each criteria range must be the same shape as the values averaged.
        if (range.width() !== values.width() || range.height() !== values.height()) return error();
        const criterion = (this.interpreter as unknown as { criterionBuilder: Criteria }).criterionBuilder.fromCellValue(conditions[i + 1], this.arithmeticHelper);
        if (!criterion) return error();
        tests.push({ cells: range.valuesFromTopLeftCorner().map(plain), test: criterion.lambda });
      }
      let sum = 0;
      let count = 0;
      const cells = values.valuesFromTopLeftCorner().map(plain);
      for (let i = 0; i < cells.length; i++) {
        if (!tests.every((t) => t.test(t.cells[i]))) continue;
        const v = cells[i];
        if (v instanceof CellError) return v;
        if (typeof v === 'number') {
          sum += v;
          count++;
        }
      }
      return count ? sum / count : error(ErrorType.DIV_BY_ZERO);
    });
  }

  mode(ast: Call, state: State) {
    const values = this.numbers(ast.args, state);
    if (values instanceof CellError) return values;
    return modes(values)[0] ?? error(ErrorType.NA);
  }

  modemult(ast: Call, state: State) {
    const values = this.numbers(ast.args, state);
    if (values instanceof CellError) return values;
    const found = modes(values);
    return found.length ? array(found.map((v) => [v])) : error(ErrorType.NA);
  }

  /**
   * How many modes there are is only known once calculated: one fits in the
   * cell, more give #VALUE!. Most often there is one, and room kept for more
   * would collide with the cells below.
   */
  modemultSize() {
    return ArraySize.error();
  }

  frequency(ast: Call, state: State) {
    return this.run(ast, state, 'FREQUENCY', (data: SimpleRangeValue, binsRange: SimpleRangeValue) => {
      const values = rangeNumbers(data);
      if (values instanceof CellError) return values;
      const bins = rangeNumbers(binsRange);
      if (bins instanceof CellError) return bins;
      // Bins count up to their value from the next lower bin, whatever order they are given in.
      const order = bins.map((b, i) => [b, i]).sort((a, b) => a[0] - b[0]);
      const counts = Array<number>(bins.length + 1).fill(0);
      for (const v of values) counts[order.find(([b]) => v <= b)?.[1] ?? bins.length]++;
      return array(counts.map((n) => [n]));
    });
  }

  frequencySize(ast: Call, state: State) {
    const bins = this.sizeOf(ast.args[1], state);
    // Text and blanks among the bins only make the result shorter.
    return bins ? arraySize(1, bins.width * bins.height + 1) : ArraySize.error();
  }

  trimmean(ast: Call, state: State) {
    return this.run(ast, state, 'TRIMMEAN', (range: SimpleRangeValue, percent: number) => {
      const values = rangeNumbers(range);
      if (values instanceof CellError) return values;
      if (percent < 0 || percent > 1 || !values.length) return error(ErrorType.NUM);
      // The same number of values goes from each end, rounded down.
      const cut = floor15((values.length * percent) / 2);
      if (2 * cut >= values.length) return error(ErrorType.NUM);
      const kept = values.sort((a, b) => a - b).slice(cut, values.length - cut);
      return kept.reduce((s, v) => s + v, 0) / kept.length;
    });
  }

  kurt(ast: Call, state: State) {
    const values = this.numbers(ast.args, state);
    if (values instanceof CellError) return values;
    const n = values.length;
    if (n < 4) return error(ErrorType.DIV_BY_ZERO);
    const mean = values.reduce((s, v) => s + v, 0) / n;
    const sd = Math.sqrt(values.reduce((s, v) => s + (v - mean) ** 2, 0) / (n - 1));
    if (sd === 0) return error(ErrorType.DIV_BY_ZERO);
    const sum4 = values.reduce((s, v) => s + ((v - mean) / sd) ** 4, 0);
    return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * sum4 - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
  }

  percentrankinc(ast: Call, state: State) {
    return this.run(ast, state, 'PERCENTRANK.INC', this.percentRanker(true));
  }

  percentrankexc(ast: Call, state: State) {
    return this.run(ast, state, 'PERCENTRANK.EXC', this.percentRanker(false));
  }

  private percentRanker(inclusive: boolean) {
    return (range: SimpleRangeValue, x: number, significance: number) => {
      const values = rangeNumbers(range);
      if (values instanceof CellError) return values;
      const digits = Math.trunc(significance);
      if (digits < 1 || !values.length) return error(ErrorType.NUM);
      const sorted = values.sort((a, b) => a - b);
      if (x < sorted[0] || x > sorted[sorted.length - 1]) return error(ErrorType.NA);
      // Excel cuts the result down to the significant digits rather than rounding it.
      const scale = 10 ** digits;
      return floor15(percentRank(sorted, x, inclusive) * scale) / scale;
    };
  }

  permut(ast: Call, state: State) {
    return this.run(ast, state, 'PERMUT', (number: number, chosen: number) => {
      const n = Math.trunc(number);
      const k = Math.trunc(chosen);
      if (n < 0 || k < 0 || n < k) return error(ErrorType.NUM);
      let p = 1;
      for (let i = 0; i < k && Number.isFinite(p); i++) p *= n - i;
      return p;
    });
  }

  permutationa(ast: Call, state: State) {
    return this.run(ast, state, 'PERMUTATIONA', (number: number, chosen: number) => {
      const n = Math.trunc(number);
      const k = Math.trunc(chosen);
      return n < 0 || k < 0 ? error(ErrorType.NUM) : n ** k;
    });
  }

  rankavg(ast: Call, state: State) {
    return this.run(ast, state, 'RANK.AVG', (x: number, range: SimpleRangeValue, order: number) => {
      const values = rangeNumbers(range);
      if (values instanceof CellError) return values;
      const ties = values.filter((v) => v === x).length;
      if (!ties) return error(ErrorType.NA);
      const ahead = values.filter((v) => (order ? v < x : v > x)).length;
      return ahead + (ties + 1) / 2;
    });
  }

  prob(ast: Call, state: State) {
    return this.run(ast, state, 'PROB', (xs: SimpleRangeValue, ps: SimpleRangeValue, lower: number, upper?: number) => {
      const x = xs.valuesFromTopLeftCorner().map(plain);
      const p = ps.valuesFromTopLeftCorner().map(plain);
      if (x.length !== p.length) return error(ErrorType.NA);
      let total = 0;
      let inside = 0;
      const hi = upper ?? lower;
      for (let i = 0; i < x.length; i++) {
        const [xi, pi] = [x[i], p[i]];
        if (xi instanceof CellError) return xi;
        if (pi instanceof CellError) return pi;
        if (typeof xi !== 'number' || typeof pi !== 'number') continue;
        if (pi <= 0 || pi > 1) return error(ErrorType.NUM);
        total += pi;
        if (xi >= lower && xi <= hi) inside += pi;
      }
      return Math.abs(total - 1) > 1e-10 ? error(ErrorType.NUM) : inside;
    });
  }

  gestep(ast: Call, state: State) {
    return this.run(ast, state, 'GESTEP', (n: number, step: number) => (n >= step ? 1 : 0));
  }

  // The engine's ERF and ERFC, which also take two arguments where these take one.
  erfprecise(ast: Call, state: State) {
    return ast.args.length === 1 ? this.delegate(ast, state, 'ERF') : error(ErrorType.NA);
  }

  erfcprecise(ast: Call, state: State) {
    return ast.args.length === 1 ? this.delegate(ast, state, 'ERFC') : error(ErrorType.NA);
  }

  percentof(ast: Call, state: State) {
    if (ast.args.length !== 2) return error(ErrorType.NA);
    const [part, all] = ast.args.map((a) => this.numbers([a], state));
    if (part instanceof CellError) return part;
    if (all instanceof CellError) return all;
    const total = all.reduce((s, v) => s + v, 0);
    return total ? part.reduce((s, v) => s + v, 0) / total : error(ErrorType.DIV_BY_ZERO);
  }

  binomdistrange(ast: Call, state: State) {
    return this.run(ast, state, 'BINOM.DIST.RANGE', (trials: number, p: number, s: number, s2?: number) => {
      const n = Math.trunc(trials);
      const lo = Math.trunc(s);
      const hi = s2 === undefined ? lo : Math.trunc(s2);
      if (n < 0 || p < 0 || p > 1 || lo < 0 || lo > n || hi < lo || hi > n) return error(ErrorType.NUM);
      return binomialRange(n, p, lo, hi);
    });
  }
}

function percentRankParameters() {
  return [
    { argumentType: FunctionArgumentType.RANGE },
    { argumentType: FunctionArgumentType.NUMBER },
    { argumentType: FunctionArgumentType.NUMBER, defaultValue: 3, emptyAsDefault: true },
  ];
}
