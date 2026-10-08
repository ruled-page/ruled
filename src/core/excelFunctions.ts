import { CellError, ErrorType, FunctionArgumentType, FunctionPlugin, HyperFormula, SimpleRangeValue, type FunctionPluginDefinition } from 'hyperformula';
import { format } from 'numfmt';
import { ArrayFunctions } from './functions/arrays';
import { excelRound } from './functions/common';
import { InfoFunctions } from './functions/info';
import { LookupFunctions } from './functions/lookup';
import { TextFunctions } from './functions/text';

/**
 * Functions where HyperFormula's results differ from Excel's, replaced with
 * versions that follow Excel (each was caught by the calculation-accuracy
 * tests): INT rounds down rather than towards zero, MOD takes the divisor's
 * sign, ROUND works on the decimal value Excel shows (1.005 → 1.01), TEXT
 * understands Excel number formats, DATEDIF and RANK.EQ exist. Functions it
 * lacks altogether are added from ./functions.
 */

/** A function call as the engine hands it to a plugin method. */
type Ast = { args: Parameters<FunctionPlugin['runFunction']>[0] };
type State = Parameters<FunctionPlugin['runFunction']>[1];

const DAY_MS = 86_400_000;

class ExcelCompatPlugin extends FunctionPlugin {
  /** Excel's serial number for a JS date, in this engine's date system. */
  private dateOf(serial: number): Date {
    const nd = (this.config as unknown as { nullDate: { year: number; month: number; day: number } }).nullDate;
    return new Date(Date.UTC(nd.year, nd.month - 1, nd.day) + Math.floor(serial) * DAY_MS);
  }

  int(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('INT'), (x: number) => Math.floor(x));
  }

  mod(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('MOD'), (n: number, d: number) => {
      if (d === 0) return new CellError(ErrorType.DIV_BY_ZERO);
      const r = n - d * Math.floor(n / d);
      // Keep results like 0.30000000000000004 to Excel's 15 digits.
      return Number(r.toPrecision(15));
    });
  }

  round(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('ROUND'), (x: number, d: number) => excelRound(x, d, 'half'));
  }

  roundup(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('ROUNDUP'), (x: number, d: number) => excelRound(x, d, 'up'));
  }

  rounddown(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('ROUNDDOWN'), (x: number, d: number) => excelRound(x, d, 'down'));
  }

  text(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('TEXT'), (value: number | string | boolean, pattern: string) => {
      const nd = (this.config as unknown as { nullDate: { year: number } }).nullDate;
      // Formats work in the 1900 system; 1904-system serials are moved into it.
      const v = typeof value === 'number' && nd.year === 1904 ? value + 1462 : value;
      try {
        return format(pattern, v, { leap1900: true, throws: true, nbsp: false });
      } catch {
        return new CellError(ErrorType.VALUE);
      }
    });
  }

  datedif(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('DATEDIF'), (start: number, end: number, unit: string) => {
      if (start > end) return new CellError(ErrorType.NUM);
      const a = this.dateOf(start);
      const b = this.dateOf(end);
      const [ay, am, ad] = [a.getUTCFullYear(), a.getUTCMonth(), a.getUTCDate()];
      const [by, bm, bd] = [b.getUTCFullYear(), b.getUTCMonth(), b.getUTCDate()];
      const months = (by - ay) * 12 + (bm - am) - (bd < ad ? 1 : 0);
      switch (unit.toUpperCase()) {
        case 'Y':
          return Math.floor(months / 12);
        case 'M':
          return months;
        case 'D':
          return Math.floor(end) - Math.floor(start);
        case 'YM':
          return months % 12;
        case 'MD': {
          if (bd >= ad) return bd - ad;
          // Days from the same day in the month before the end date.
          const prev = Date.UTC(by, bm - 1, ad);
          return Math.round((b.getTime() - prev) / DAY_MS);
        }
        case 'YD': {
          let anniversary = Date.UTC(by, am, ad);
          if (anniversary > b.getTime()) anniversary = Date.UTC(by - 1, am, ad);
          return Math.round((b.getTime() - anniversary) / DAY_MS);
        }
        default:
          return new CellError(ErrorType.NUM);
      }
    });
  }

  rankeq(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('RANK.EQ'), (n: number, range: SimpleRangeValue, order: number) => {
      const values = range.valuesFromTopLeftCorner().filter((v): v is number => typeof v === 'number');
      if (!values.includes(n)) return new CellError(ErrorType.NA);
      return 1 + values.filter((v) => (order ? v < n : v > n)).length;
    });
  }
}

const rank = {
  method: 'rankeq',
  parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0 }],
};

(ExcelCompatPlugin as unknown as { implementedFunctions: Record<string, unknown> }).implementedFunctions = {
  INT: { method: 'int', parameters: [{ argumentType: FunctionArgumentType.NUMBER }] },
  MOD: { method: 'mod', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }] },
  ROUND: { method: 'round', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0 }] },
  ROUNDUP: { method: 'roundup', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0 }] },
  ROUNDDOWN: { method: 'rounddown', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER, defaultValue: 0 }] },
  TEXT: { method: 'text', parameters: [{ argumentType: FunctionArgumentType.SCALAR }, { argumentType: FunctionArgumentType.STRING }] },
  DATEDIF: { method: 'datedif', parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.STRING }] },
  'RANK.EQ': rank,
  RANK: rank,
};

/** Functions HyperFormula lacks, each plugin with the functions it implements. */
const ADDED: FunctionPluginDefinition[] = [ArrayFunctions, InfoFunctions, LookupFunctions, TextFunctions] as never[];

let registered = false;

/** Puts the Excel versions in place of HyperFormula's, and adds the ones it lacks; engines built afterwards use them. */
export function registerExcelFunctions(): void {
  if (registered) return;
  registered = true;
  // Excel's #CALC! (an empty array, for one) has no name in the engine.
  const errors = (HyperFormula.getLanguage('enGB') as unknown as { errors: Record<string, string> }).errors;
  errors.CALC ??= '#CALC!';
  const plugins: FunctionPluginDefinition[] = [ExcelCompatPlugin as never, ...ADDED];
  for (const plugin of plugins) {
    for (const id of Object.keys(plugin.implementedFunctions)) {
      try {
        HyperFormula.unregisterFunction(id);
      } catch {
        // Not registered before (RANK.EQ, for one).
      }
      HyperFormula.registerFunction(id, plugin, { enGB: { [id]: id } });
    }
  }
}
