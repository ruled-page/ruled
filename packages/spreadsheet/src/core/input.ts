import { isDateFormat, isTextFormat, parseValue } from 'numfmt';
import { msg, type Text } from '../i18n/message';
import { parse, tokenize } from '@borgar/fx';
import { fifteenDigits } from '@ruled/formula/engine';
import { toFileFormula } from '@ruled/formula/formula';
import { localeInfo, normalizeForLocale } from './locale';
import { EMPTY, type CellContent } from './types';

export const ERROR_LITERALS = new Set([
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA',
  '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!',
]);

export interface ParsedInput {
  content: CellContent;
  /** Number format implied by the input (e.g. "12%" → "0%"), for General cells. */
  numFmt?: string;
  /** Why the input was rejected, e.g. a malformed formula. */
  error?: Text;
}

/** Days between the 1900 and 1904 date systems' day zero. */
export const DAYS_1900_TO_1904 = 1462;

/**
 * Interprets text the way Excel does when it is typed or pasted into a cell
 * with the given number format, reading dates and decimals in `locale`, and
 * dates as serial numbers in the workbook's date system.
 */
export function parseInput(text: string, cellNumFmt: string, locale = 'en-US', date1904 = false): ParsedInput {
  if (text === '') return { content: EMPTY };
  if (text[0] === "'") return { content: { kind: 'string', value: text.slice(1) } };

  let textFormat = false;
  try {
    textFormat = cellNumFmt !== 'General' && isTextFormat(cellNumFmt);
  } catch {
    textFormat = false;
  }
  if (textFormat) return { content: { kind: 'string', value: text } };

  if (text[0] === '=' && text.length > 1) {
    const body = text.slice(1);
    const problem = validateFormula(body);
    if (problem) return { content: EMPTY, error: msg('core.edit.badFormula', { problem }) };
    return { content: { kind: 'formula', formula: toFileFormula(body) } };
  }

  const trimmed = text.trim();
  if (ERROR_LITERALS.has(trimmed.toUpperCase())) return { content: { kind: 'error', value: trimmed.toUpperCase() } };

  const info = localeInfo(locale);
  const local = normalizeForLocale(text, info);
  // Where a point is the decimal separator, a comma after it isn't a thousands
  // separator: Excel keeps "1.234,5" as text rather than reading 1.2345.
  if (!info.decimalComma && /\d\.\d*,/.test(local.text)) return { content: { kind: 'string', value: text } };
  let parsed: ReturnType<typeof parseValue> = null;
  try {
    parsed = parseValue(local.text);
  } catch {
    parsed = null;
  }
  if (parsed) {
    if (typeof parsed.v === 'boolean') return { content: { kind: 'bool', value: parsed.v } };
    if (typeof parsed.v === 'number' && Number.isFinite(parsed.v)) {
      const z = local.numFmt ?? parsed.z;
      const numFmt = z && cellNumFmt === 'General' ? z : undefined;
      // The parser counts days from 1900; a time on its own has no date to move.
      // A plain number keeps fifteen significant digits, the rest becoming zeros, as Excel keeps them.
      const plain = local.text.trim().replace(/,/g, '');
      const v = /^[+-]?\d*\.?\d*(e[+-]?\d+)?$/i.test(plain) ? fifteenDigits(plain) : parsed.v;
      const value = date1904 && z && isDateFormat(z) && v >= 1 ? v - DAYS_1900_TO_1904 : v;
      return { content: { kind: 'number', value }, numFmt };
    }
  }
  return { content: { kind: 'string', value: text } };
}

/**
 * Returns what is wrong when the formula is not valid Excel syntax: a
 * message, or the parser's own (English) description.
 */
export function validateFormula(userFormula: string): Text | undefined {
  if (!userFormula.trim()) return msg('core.formula.empty');
  try {
    parse(tokenize(userFormula, { mergeRefs: true, allowTernary: false }));
    return undefined;
  } catch (e) {
    return e instanceof Error ? e.message : String(e);
  }
}
