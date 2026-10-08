import { CellError, EmptyValue, ErrorType, FunctionArgumentType, SimpleRangeValue, type ImplementedFunctions } from 'hyperformula';
import { CALC, ExcelPlugin, error, plain, type Call, type State, type Value } from './common';

/** ERROR.TYPE's numbers for the errors the engine can produce. */
const ERROR_NUMBERS: Partial<Record<string, number>> = {
  [ErrorType.DIV_BY_ZERO]: 2,
  [ErrorType.VALUE]: 3,
  [ErrorType.REF]: 4,
  [ErrorType.NAME]: 5,
  [ErrorType.NUM]: 6,
  [ErrorType.NA]: 7,
  [ErrorType.SPILL]: 9,
  [CALC]: 14,
};

export class InfoFunctions extends ExcelPlugin {
  static implementedFunctions: ImplementedFunctions = {
    TYPE: { method: 'type', parameters: [{ argumentType: FunctionArgumentType.ANY }] },
    'ERROR.TYPE': { method: 'errortype', parameters: [{ argumentType: FunctionArgumentType.SCALAR }] },
  };

  type(ast: Call, state: State) {
    if (ast.args.length !== 1) return error(ErrorType.NA);
    // Read without the usual argument handling, which would pass on an error or split up an array.
    const v = this.evaluate(ast.args[0], state);
    if (v instanceof SimpleRangeValue) return 64;
    const x = plain(v);
    if (typeof x === 'number' || x === EmptyValue) return 1;
    if (typeof x === 'string') return 2;
    if (typeof x === 'boolean') return 4;
    return 16;
  }

  errortype(ast: Call, state: State) {
    return this.run(ast, state, 'ERROR.TYPE', (v: Value) => {
      if (!(v instanceof CellError)) return error(ErrorType.NA);
      // The engine's own errors (#CYCLE!, #ERROR!) have no number in Excel; they pass through.
      return ERROR_NUMBERS[v.type] ?? v;
    });
  }
}
