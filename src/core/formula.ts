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
const FUTURE_FUNCTIONS = new Set(
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
    'NORM.S.DIST NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC ' +
    'PERMUTATIONA PHI POISSON.DIST QUARTILE.EXC QUARTILE.INC QUERYSTRING RANDARRAY RANK.AVG RANK.EQ REDUCE RRI ' +
    'SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV ' +
    'T.INV.2T T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR UNICODE UNIQUE VALUETOTEXT ' +
    'VAR.P VAR.S VSTACK WEBSERVICE WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST'
  ).split(' '),
);
const WORKSHEET_FUNCTIONS = new Set(['FILTER', 'SORT']);

const PREFIX_RE = /_xlfn\.(?:_xlws\.)?|_xlws\./gi;

export function stripFunctionPrefixes(formula: string): string {
  return formula.indexOf('_xl') === -1 ? formula : formula.replace(PREFIX_RE, '');
}

/** How a file formula is shown to the user. */
export function displayFormula(fileFormula: string): string {
  return '=' + stripFunctionPrefixes(fileFormula);
}

/**
 * How a file formula is given to the calculation engine. HyperFormula has
 * no TRUE/FALSE literals, only the TRUE() and FALSE() functions.
 */
export function engineFormula(fileFormula: string): string {
  let f = stripFunctionPrefixes(fileFormula);
  if (/TRUE|FALSE/i.test(f)) {
    try {
      f = stringifyTokens(tokens(f).map((t) => (t.type === 'bool' ? { ...t, value: `${t.value}()` } : t)));
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
  const out = tokens(userFormula).map((t) => {
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
