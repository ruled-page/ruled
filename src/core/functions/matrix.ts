import { ArraySize, CellError, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { ExcelPlugin, MAX_RESULT_CELLS, array, arraySize, error, grid, literalNumber, omitted, type Call, type State } from './common';

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

/**
 * Gauss-Jordan elimination with partial pivoting, on the matrix and an
 * identity beside it. Returns the determinant and the inverse (undefined
 * when the matrix is singular).
 */
function eliminate(a: number[][]): { det: number; inverse?: number[][] } {
  const n = a.length;
  const inv = a.map((_, i) => a.map((__, j) => (i === j ? 1 : 0)));
  let det = 1;
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
    if (a[p][c] === 0) return { det: 0 };
    if (p !== c) {
      [a[p], a[c]] = [a[c], a[p]];
      [inv[p], inv[c]] = [inv[c], inv[p]];
      det = -det;
    }
    const pivot = a[c][c];
    det *= pivot;
    for (let j = 0; j < n; j++) {
      a[c][j] /= pivot;
      inv[c][j] /= pivot;
    }
    for (let r = 0; r < n; r++) {
      if (r === c || a[r][c] === 0) continue;
      const f = a[r][c];
      for (let j = 0; j < n; j++) {
        a[r][j] -= f * a[c][j];
        inv[r][j] -= f * inv[c][j];
      }
    }
  }
  return { det, inverse: inv };
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
