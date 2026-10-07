import { isTextFormat, parseValue } from 'numfmt';
import { toFileFormula, validateFormula } from './formula';
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
  error?: string;
}

/**
 * Interprets text the way Excel does when it is typed or pasted into a cell
 * with the given number format, reading dates and decimals in `locale`.
 */
export function parseInput(text: string, cellNumFmt: string, locale = 'en-US'): ParsedInput {
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
    if (problem) return { content: EMPTY, error: `There's a problem with this formula: ${problem}` };
    return { content: { kind: 'formula', formula: toFileFormula(body) } };
  }

  const trimmed = text.trim();
  if (ERROR_LITERALS.has(trimmed.toUpperCase())) return { content: { kind: 'error', value: trimmed.toUpperCase() } };

  const local = normalizeForLocale(text, localeInfo(locale));
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
      return { content: { kind: 'number', value: parsed.v }, numFmt };
    }
  }
  return { content: { kind: 'string', value: text } };
}
