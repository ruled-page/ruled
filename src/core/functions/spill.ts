import { ArraySize, EmptyValue, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { array, arraySize, ExcelPlugin, grid, literalNumber, plain, type Call, type Grid, type State } from './common';

/**
 * The function the engine is given a spilling formula in, never written to a
 * file: WBE.SPILL(rows, cols, formula) gives the formula's result in exactly
 * that many rows and columns. The engine needs an array's size before it
 * works the array out, and guesses it from the formula (FILTER as many rows
 * as it filters, SEQUENCE(A1) not at all); this gives it the size the result
 * really has, and notes the size each time the formula is worked out, so it
 * can be given room again when its result grows or shrinks.
 */
export const SPILL_FUNCTION = 'WBE.SPILL';

export interface ArrayDims {
  rows: number;
  cols: number;
}

/** Per engine (by its dependency graph): the size each spilling formula's result had when last worked out, by "sheet:row:col". */
const measured = new WeakMap<object, Map<string, ArrayDims>>();

/** The sizes noted for one engine's spilling formulas. */
export function measuredSizes(graph: object): Map<string, ArrayDims> {
  let m = measured.get(graph);
  if (!m) measured.set(graph, (m = new Map()));
  return m;
}

export class SpillFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    [SPILL_FUNCTION]: {
      method: 'spill',
      sizeOfResultArrayMethod: 'spillSize',
      vectorizationForbidden: true,
      parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.ANY }],
    },
  };

  spill(ast: Call, state: State) {
    const rows = literalNumber(ast.args[0]) ?? 1;
    const cols = literalNumber(ast.args[1]) ?? 1;
    const v = ast.args[2] ? this.evaluate(ast.args[2], state) : EmptyValue;
    const values: Grid = v instanceof SimpleRangeValue ? grid(v) : [[plain(v)]];
    const at = state.formulaAddress;
    measuredSizes(this.dependencyGraph).set(`${at.sheet}:${at.row}:${at.col}`, { rows: values.length, cols: values[0].length });
    if (rows === 1 && cols === 1) return values[0][0];
    // Cut or padded to the room it has, until it is given room for its new size.
    const out: Grid = [];
    for (let r = 0; r < rows; r++) {
      const line = values[r] ?? [];
      out.push(Array.from({ length: cols }, (_, c) => line[c] ?? EmptyValue));
    }
    return array(out);
  }

  spillSize(ast: Call) {
    const rows = literalNumber(ast.args[0]);
    const cols = literalNumber(ast.args[1]);
    return rows === undefined || cols === undefined ? ArraySize.error() : arraySize(cols, rows);
  }
}
