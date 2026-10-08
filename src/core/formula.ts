import {
  parse,
  stringifyTokens,
  tokenize,
  translateTokensToA1,
  translateTokensToR1C1,
  type Token,
} from '@borgar/fx';
import { msg, type Text } from '../i18n/message';
import { toA1 } from './refs';

/**
 * Functions added after Excel 2007 are stored in the file with an `_xlfn.`
 * prefix (and `_xlfn._xlws.` for a few). Without it Excel shows #NAME? until
 * the cell is re-entered, so formulas typed in the app must gain the prefix.
 */
export const FUTURE_FUNCTIONS = new Set(
  (
    'ACOT ACOTH AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV ' +
    'BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT ' +
    'CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT ' +
    'COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS DECIMAL DROP ECMA.CEILING ERF.PRECISE ERFC.PRECISE ' +
    'ERROR.TYPE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FIELDVALUE FILTERXML FLOOR.MATH ' +
    'FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR ' +
    'FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HSTACK HYPGEOM.DIST IFNA IFS IMCOSH IMCOT ' +
    'IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV ' +
    'MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV ' +
    'NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTOF PERCENTRANK.EXC PERCENTRANK.INC ' +
    'PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE RRI ' +
    'SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV ' +
    'T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR UNICODE UNIQUE VALUETOTEXT ' +
    'VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST'
  ).split(' '),
);
const WORKSHEET_FUNCTIONS = new Set(['FILTER', 'SORT']);

/**
 * Functions that can return an array. Excel saves a formula using one as a
 * dynamic array formula; saved as a plain one, Excel would show it as
 * =@SORT(…), a single value. The same list as XlsxWriter's.
 */
const DYNAMIC_ARRAY_FUNCTIONS = new Set(
  (
    'ANCHORARRAY BYCOL BYROW CHOOSECOLS CHOOSEROWS DROP EXPAND FILTER HSTACK LAMBDA MAKEARRAY MAP RANDARRAY ' +
    'REDUCE SCAN SEQUENCE SINGLE SORT SORTBY SWITCH TAKE TEXTSPLIT TOCOL TOROW UNIQUE VSTACK WRAPCOLS WRAPROWS XLOOKUP'
  ).split(' '),
);

/** Function prefixes, and the one Excel puts on LET and LAMBDA parameter names. */
const PREFIX_RE = /_xlfn\.(?:_xlws\.)?|_xlws\.|_xlpm\./gi;

export function stripFunctionPrefixes(formula: string): string {
  return formula.indexOf('_xl') === -1 ? formula : formula.replace(PREFIX_RE, '');
}

/** How a file formula is shown to the user. */
export function displayFormula(fileFormula: string): string {
  return '=' + stripFunctionPrefixes(fileFormula);
}

/**
 * How a file formula is given to the calculation engine, which reads a few
 * things differently from Excel: it has no TRUE/FALSE literals, only the
 * TRUE() and FALSE() functions; it reads exponents only with a lower-case
 * "e" (1e+3, not 1E+3); and it can't read a quote doubled inside a string,
 * so "say ""hi""" becomes ("say "&CHAR(34)&"hi"&CHAR(34)&""), except in an
 * array constant, where & isn't allowed.
 */
export function engineFormula(fileFormula: string): string {
  let f = stripFunctionPrefixes(fileFormula);
  if (/TRUE|FALSE|\dE|""/i.test(f)) {
    try {
      let depth = 0;
      f = stringifyTokens(
        tokens(f).map((t) => {
          if (t.type === 'operator' && (t.value === '{' || t.value === '}')) depth += t.value === '{' ? 1 : -1;
          if (t.type === 'bool') return { ...t, value: `${t.value}()` };
          if (t.type === 'number') return { ...t, value: t.value.toLowerCase() };
          if (t.type === 'string' && !depth && t.value.slice(1, -1).includes('""')) {
            const parts = t.value.slice(1, -1).split('""');
            return { ...t, value: `(${parts.map((part) => `"${part}"`).join('&CHAR(34)&')})` };
          }
          return t;
        }),
      );
    } catch {
      // Leave the formula as it is; the engine will report it if it cannot parse it.
    }
  }
  return '=' + f;
}

function tokens(formula: string): Token[] {
  return tokenize(formula, { mergeRefs: true, allowTernary: false });
}

/** Converts a formula typed by the user (without "=") into file form. */
export function toFileFormula(userFormula: string): string {
  const list = tokens(userFormula);
  const params = lambdaParameters(list);
  const out = list.map((t, i) => {
    if (params.has(i) && !/^_xlpm\./i.test(t.value)) return { ...t, value: `_xlpm.${t.value}` };
    if (t.type !== 'func') return t;
    const upper = t.value.toUpperCase();
    if (upper.startsWith('_XL')) return t;
    if (WORKSHEET_FUNCTIONS.has(upper)) return { ...t, value: `_xlfn._xlws.${t.value}` };
    if (FUTURE_FUNCTIONS.has(upper)) return { ...t, value: `_xlfn.${t.value}` };
    return t;
  });
  return stringifyTokens(out);
}

/**
 * Tokens that name a LET or LAMBDA parameter, where it is declared or used.
 * Excel stores these names with an `_xlpm.` prefix, without which it reads
 * them as defined names and shows #NAME?. LET(x, 1, y, 2, x + y) declares x
 * and y; LAMBDA(a, b, a + b) declares a and b; either's names count anywhere
 * inside it.
 */
function lambdaParameters(list: Token[]): Set<number> {
  const out = new Set<number>();
  for (let i = 0; i < list.length; i++) {
    const name = list[i].type === 'func' ? stripFunctionPrefixes(list[i].value).toUpperCase() : '';
    if ((name !== 'LET' && name !== 'LAMBDA') || list[i + 1]?.value !== '(') continue;
    // The call's arguments, as token ranges, up to its closing parenthesis.
    const args: [number, number][] = [];
    let depth = 0;
    let start = i + 2;
    let end = list.length;
    for (let k = i + 1; k < list.length; k++) {
      const v = list[k].type === 'operator' ? list[k].value : '';
      if (v === '(' || v === '{') depth++;
      else if (v === ')' || v === '}') {
        if (--depth === 0) {
          args.push([start, k]);
          end = k;
          break;
        }
      } else if (v === ',' && depth === 1) {
        args.push([start, k]);
        start = k + 1;
      }
    }
    const declared = new Set<string>();
    args.slice(0, -1).forEach(([a, b], n) => {
      const solid = list.slice(a, b).filter((t) => t.type !== 'whitespace');
      if ((name === 'LAMBDA' || n % 2 === 0) && solid.length === 1 && solid[0].type === 'range_named') declared.add(solid[0].value.toUpperCase());
    });
    for (let k = i + 2; k < end; k++) {
      const t = list[k];
      if (t.type === 'range_named' && declared.has(t.value.toUpperCase())) out.add(k);
    }
  }
  return out;
}

/** Whether Excel would save this formula as a dynamic array formula. */
export function isDynamicArrayFormula(fileFormula: string): boolean {
  return functionNames(fileFormula).some((n) => DYNAMIC_ARRAY_FUNCTIONS.has(n));
}

/**
 * Returns what is wrong when the formula is not valid Excel syntax: a
 * message, or the parser's own (English) description.
 */
export function validateFormula(userFormula: string): Text | undefined {
  if (!userFormula.trim()) return msg('core.formula.empty');
  try {
    parse(tokens(userFormula));
    return undefined;
  } catch (e) {
    return e instanceof Error ? e.message : String(e);
  }
}

export function functionNames(fileFormula: string): string[] {
  try {
    return tokens(fileFormula)
      .filter((t) => t.type === 'func')
      .map((t) => stripFunctionPrefixes(t.value).toUpperCase());
  } catch {
    return [];
  }
}

const VOLATILE = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY']);

export function isVolatile(fileFormula: string): boolean {
  return functionNames(fileFormula).some((n) => VOLATILE.has(n));
}

/**
 * Returns a function that rewrites the formula, written at (row, col), as if
 * it had been copied to another cell: relative references move, absolute
 * ones stay, and references pushed off the sheet become #REF!.
 */
export function makeShifter(fileFormula: string, row: number, col: number): (toRow: number, toCol: number) => string {
  const rel = translateTokensToR1C1(tokens(fileFormula), toA1(row, col));
  return (toRow, toCol) =>
    toRow === row && toCol === col
      ? fileFormula
      : stringifyTokens(translateTokensToA1(rel, toA1(toRow, toCol), { wrapEdges: false }));
}

/** Like makeShifter for a single move; returns undefined if the formula cannot be translated. */
export function shiftFormula(fileFormula: string, fromRow: number, fromCol: number, toRow: number, toCol: number): string | undefined {
  try {
    return makeShifter(fileFormula, fromRow, fromCol)(toRow, toCol);
  } catch {
    return undefined;
  }
}
