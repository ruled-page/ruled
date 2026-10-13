import {
  stringifyTokens,
  tokenize,
  translateTokensToA1,
  translateTokensToR1C1,
  type Token,
} from '@borgar/fx';
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
 * array constant, where & isn't allowed. It reads a reference whose cells
 * were deleted, which Excel writes with its sheet (Prices!#REF!), only as
 * the error alone. It has no LET either, which is written out (expandLets).
 */
export function engineFormula(fileFormula: string): string {
  let f = stripFunctionPrefixes(expandLets(fileFormula) ?? fileFormula);
  if (/TRUE|FALSE|\dE|""|!#REF!/i.test(f)) {
    try {
      let depth = 0;
      const list = tokens(f);
      const deleted = (i: number) => list[i]?.type === 'context' && list[i + 1]?.value === '!' && list[i + 2]?.type === 'error';
      f = stringifyTokens(
        list.flatMap((t, i) => {
          if (deleted(i) || deleted(i - 1)) return [];
          if (t.type === 'operator' && (t.value === '{' || t.value === '}')) depth += t.value === '{' ? 1 : -1;
          if (t.type === 'bool') return [{ ...t, value: `${t.value}()` }];
          if (t.type === 'number') return [{ ...t, value: t.value.toLowerCase() }];
          if (t.type === 'string' && !depth && t.value.slice(1, -1).includes('""')) {
            const parts = t.value.slice(1, -1).split('""');
            return [{ ...t, value: `(${parts.map((part) => `"${part}"`).join('&CHAR(34)&')})` }];
          }
          return [t];
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

/**
 * The arguments of the function call whose name is token `at`, as token
 * ranges, and where its closing bracket is; undefined if it isn't closed.
 */
function callArgs(list: Token[], at: number): { args: [number, number][]; close: number } | undefined {
  const args: [number, number][] = [];
  let depth = 0;
  let start = at + 2;
  for (let k = at + 1; k < list.length; k++) {
    const v = list[k].type === 'operator' ? list[k].value : '';
    if (v === '(' || v === '{') depth++;
    else if (v === ')' || v === '}') {
      if (--depth === 0) {
        args.push([start, k]);
        return { args, close: k };
      }
    } else if (v === ',' && depth === 1) {
      args.push([start, k]);
      start = k + 1;
    }
  }
  return undefined;
}

/** A LET or LAMBDA parameter's name as the formula uses it, with or without Excel's prefix. */
const parameterName = (t: Token) => t.value.replace(/^_xlpm\./i, '').toUpperCase();

/** Longer than this, a formula with its LETs written out is left for Excel to calculate. */
const MAX_EXPANDED = 8192;

const bracket = (inner: Token[]): Token[] => [{ type: 'operator', value: '(' } as Token, ...inner, { type: 'operator', value: ')' } as Token];

/**
 * The formula with every LET call written out, each name replaced by its
 * value in brackets: LET(x, A1*2, x + x) is ((A1*2) + (A1*2)). That works
 * out the same as long as nothing in a value is volatile. The innermost LET
 * goes first, so a name a nested LET reuses means its own value there.
 * Undefined when a LET can't be read, has a volatile value, or written out
 * would be too long.
 */
export function expandLets(fileFormula: string): string | undefined {
  if (!/LET\s*\(/i.test(fileFormula)) return fileFormula;
  let list: Token[];
  try {
    list = tokens(fileFormula);
  } catch {
    return undefined;
  }
  // From the end: the last LET has no other inside it.
  for (let at = list.length - 1; at >= 0; at--) {
    if (list[at].type !== 'func' || stripFunctionPrefixes(list[at].value).toUpperCase() !== 'LET' || list[at + 1]?.value !== '(') continue;
    const call = callArgs(list, at);
    if (!call || call.args.length < 3 || call.args.length % 2 === 0) return undefined;
    const values = new Map<string, Token[]>();
    const written = ([a, b]: [number, number]) => {
      // Without the spaces around the argument.
      while (a < b && list[a].type === 'whitespace') a++;
      while (b > a && list[b - 1].type === 'whitespace') b--;
      // A name's uses, and its calls when it holds a LAMBDA, f(5).
      return list.slice(a, b).flatMap((t) => (t.type === 'range_named' || t.type === 'func' ? (values.get(parameterName(t)) ?? [t]) : [t]));
    };
    for (let n = 0; n < call.args.length - 1; n += 2) {
      const name = list.slice(...call.args[n]).filter((t) => t.type !== 'whitespace');
      if (name.length !== 1 || name[0].type !== 'range_named') return undefined;
      const value = written(call.args[n + 1]);
      if (value.some((t) => t.type === 'func' && VOLATILE.has(stripFunctionPrefixes(t.value).toUpperCase()))) return undefined;
      values.set(parameterName(name[0]), bracket(value));
    }
    list = [...list.slice(0, at), ...bracket(written(call.args[call.args.length - 1])), ...list.slice(call.close + 1)];
    if (list.reduce((n, t) => n + t.value.length, 0) > MAX_EXPANDED) return undefined;
  }
  return stringifyTokens(list);
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
    const call = callArgs(list, i);
    if (!call) continue;
    const { args, close: end } = call;
    const declared = new Set<string>();
    args.slice(0, -1).forEach(([a, b], n) => {
      const solid = list.slice(a, b).filter((t) => t.type !== 'whitespace');
      if ((name === 'LAMBDA' || n % 2 === 0) && solid.length === 1 && solid[0].type === 'range_named') declared.add(solid[0].value.toUpperCase());
    });
    for (let k = i + 2; k < end; k++) {
      const t = list[k];
      // A parameter, or a call of one that holds a LAMBDA, f(5).
      if ((t.type === 'range_named' || (t.type === 'func' && list[k + 1]?.value === '(')) && declared.has(t.value.toUpperCase())) out.add(k);
    }
  }
  return out;
}

/** Whether Excel would save this formula as a dynamic array formula. */
export function isDynamicArrayFormula(fileFormula: string): boolean {
  return functionNames(fileFormula).some((n) => DYNAMIC_ARRAY_FUNCTIONS.has(n));
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
