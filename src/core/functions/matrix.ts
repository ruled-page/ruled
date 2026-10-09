import { ArraySize, CellError, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, MAX_RESULT_CELLS, array, arraySize, error, grid, literalNumber, omitted, type Call, type State } from './common';
import { eliminate } from '../matrixMath';

/** The largest matrix to invert, as the work grows with the cube of its size. */
const MAX_SIZE = 400;

/** A square matrix of numbers, or the error its cells give: a blank or text is #VALUE!. */
function squareMatrix(range: SimpleRangeValue): number[][] | CellError {
  const rows = grid(range);
  if (rows.length !== rows[0].length) return error();
  if (rows.length > MAX_SIZE) return error(ErrorType.NUM);
  for (const row of rows) {
    for (const v of row) {
      if (v instanceof CellError) return v;
      if (typeof v !== 'number') return error();
    }
  }
  return rows.map((row) => row.slice() as number[]);
}

export class MatrixFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    MDETERM: { method: 'mdeterm', vectorizationForbidden: true, parameters: [{ argumentType: FunctionArgumentType.RANGE }] },
    MINVERSE: { method: 'minverse', sizeOfResultArrayMethod: 'minverseSize', vectorizationForbidden: true, parameters: [{ argumentType: FunctionArgumentType.RANGE }] },
    MUNIT: { method: 'munit', sizeOfResultArrayMethod: 'munitSize', vectorizationForbidden: true, parameters: [{ argumentType: FunctionArgumentType.NUMBER }] },
  };

  mdeterm(ast: Call, state: State) {
    return this.run(ast, state, 'MDETERM', (range: SimpleRangeValue) => {
      const m = squareMatrix(range);
      return m instanceof CellError ? m : eliminate(m).det;
    });
  }

  minverse(ast: Call, state: State) {
    return this.run(ast, state, 'MINVERSE', (range: SimpleRangeValue) => {
      const m = squareMatrix(range);
      if (m instanceof CellError) return m;
      const { inverse } = eliminate(m);
      return inverse ? array(inverse) : error(ErrorType.NUM);
    });
  }

  minverseSize(ast: Call, state: State) {
    const size = this.sizeOf(ast.args[0], state);
    return size && size.width === size.height ? arraySize(size.width, size.height) : ArraySize.error();
  }

  munit(ast: Call, state: State) {
    return this.run(ast, state, 'MUNIT', (dimension: number) => {
      const n = Math.trunc(dimension);
      if (n < 1) return error();
      if (n * n > MAX_RESULT_CELLS) return error(ErrorType.NUM);
      return array(Array.from({ length: n }, (_, i) => Array.from({ length: n }, (__, j) => (i === j ? 1 : 0))));
    });
  }

  munitSize(ast: Call) {
    const n = omitted(ast.args[0]) ? undefined : literalNumber(ast.args[0]);
    return n === undefined ? ArraySize.error() : arraySize(Math.trunc(n), Math.trunc(n));
  }
}
