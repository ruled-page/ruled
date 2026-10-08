import { CellError, EmptyValue, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, compareValues, error, grid, sameValue, type Call, type State, type Value } from './common';

/** Whether a lookup may compare the two: only values of the same type are ordered against each other. */
const comparable = (a: Value, b: Value) => typeof a === typeof b && !(a instanceof CellError) && a !== EmptyValue;

/** Where the values of the key's type are; approximate lookups pass over the rest. */
const ofKeyType = (values: Value[], key: Value) => values.flatMap((v, i) => (comparable(v, key) ? [i] : []));

/**
 * Which of those positions holds the last value not past the key, found by
 * binary search as Excel's approximate lookups do, so values out of order
 * give Excel's answers. Searching past every value is what makes
 * LOOKUP(2,1/(A:A="x"),B:B) find the last match.
 */
function lastAtMost(values: Value[], idx: number[], key: Value, order = 1): number {
  let lo = 0;
  let hi = idx.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (order * compareValues(values[idx[mid]], key) <= 0) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return found;
}

export class LookupFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    LOOKUP: {
      method: 'lookup',
      parameters: [{ argumentType: FunctionArgumentType.NOERROR }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE, optionalArg: true }],
    },
    XMATCH: {
      method: 'xmatch',
      parameters: [
        { argumentType: FunctionArgumentType.NOERROR },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0, emptyAsDefault: true },
        { argumentType: FunctionArgumentType.NUMBER, defaultValue: 1, emptyAsDefault: true },
      ],
    },
  };

  lookup(ast: Call, state: State) {
    return this.run(ast, state, 'LOOKUP', (key: Value, where: SimpleRangeValue, result?: SimpleRangeValue) => {
      const data = grid(where);
      const h = data.length;
      const w = data[0].length;
      // The array form searches the first row or column, whichever is longer, and answers from the last.
      const wide = w > h;
      const keys = h === 1 || w === 1 ? data.flat() : wide ? data[0] : data.map((r) => r[0]);
      let answers: Value[];
      if (result) {
        const r = grid(result);
        if (r.length > 1 && r[0].length > 1) return error(ErrorType.NA);
        answers = r.flat();
      } else {
        answers = wide ? data[h - 1] : data.map((r) => r[w - 1]);
      }
      const k = key === EmptyValue ? 0 : key;
      const idx = ofKeyType(keys, k);
      const found = lastAtMost(keys, idx, k);
      if (found < 0 || idx[found] >= answers.length) return error(ErrorType.NA);
      const at = idx[found];
      const v = answers[at];
      return v === EmptyValue ? 0 : v;
    });
  }

  xmatch(ast: Call, state: State) {
    return this.run(ast, state, 'XMATCH', (key: Value, where: SimpleRangeValue, matchMode: number, searchMode: number) => {
      const match = Math.trunc(matchMode);
      const search = Math.trunc(searchMode);
      if (![0, -1, 1, 2].includes(match) || ![1, -1, 2, -2].includes(search)) return error();
      const data = grid(where);
      if (data.length > 1 && data[0].length > 1) return error();
      const values = data.flat();
      const k = key === EmptyValue ? 0 : key;
      let found: number;
      if (match === 2 && typeof k === 'string') {
        const test = this.arithmeticHelper.eqMatcherFunction(k);
        found = this.linear(values, (v) => test(v as never), search === -1);
      } else if (Math.abs(search) === 2) {
        found = this.binary(values, k, match === 2 ? 0 : match, search === 2 ? 1 : -1);
      } else {
        found = this.nearest(values, k, match === 2 ? 0 : match, search === -1);
      }
      return found < 0 ? error(ErrorType.NA) : found + 1;
    });
  }

  private linear(values: Value[], test: (v: Value) => boolean, backwards: boolean): number {
    for (let n = 0; n < values.length; n++) {
      const i = backwards ? values.length - 1 - n : n;
      if (test(values[i])) return i;
    }
    return -1;
  }

  /** The exact match, or with mode -1 the largest smaller value and with 1 the smallest larger one; ties go to the first found. */
  private nearest(values: Value[], key: Value, mode: number, backwards: boolean): number {
    let best = -1;
    for (let n = 0; n < values.length; n++) {
      const i = backwards ? values.length - 1 - n : n;
      const v = values[i];
      if (sameValue(v, key)) return i;
      if (mode === 0 || !comparable(v, key)) continue;
      const side = compareValues(v, key) * mode;
      if (side > 0 && (best < 0 || compareValues(v, values[best]) * mode < 0)) best = i;
    }
    return best;
  }

  /** Binary search over values sorted ascending (order 1) or descending (-1). */
  private binary(values: Value[], key: Value, mode: number, order: number): number {
    const idx = ofKeyType(values, key);
    const p = lastAtMost(values, idx, key, order);
    if (p >= 0 && sameValue(values[idx[p]], key)) return idx[p];
    if (mode === 0) return -1;
    const smaller = order === 1 ? p : p + 1;
    const larger = order === 1 ? p + 1 : p;
    const q = mode === -1 ? smaller : larger;
    return q >= 0 && q < idx.length ? idx[q] : -1;
  }
}
