import { hasWildcards, matchesWhole, parsePattern } from '@ruled/formula/wildcard';
import { isDateFormat } from 'numfmt';
import { parseInput } from '../core/input';
import type { FilterCondition, FilterOperator } from '../core/types';

/** A cell as a filter sees it: the text it shows, and its number when it holds one. */
export interface FilterCell {
  text: string;
  n?: number;
}

/** The number a rule's value stands for, written as Excel stores it ("5", "45292.5"). */
const asNumber = (value: string) => (value.trim() !== '' && Number.isFinite(Number(value)) ? Number(value) : undefined);

/**
 * Whether a cell meets one rule, as Excel's custom filter tests it: numbers
 * against a number, otherwise the text the cell shows, ignoring case and
 * with Excel's wildcards. Only numbers meet a comparison with a number, and
 * an empty cell meets only "does not equal".
 */
function meets(operator: FilterOperator, value: string, cell: FilterCell): boolean {
  if (operator === 'notEqual') return !meets('equal', value, cell);
  const n = asNumber(value);
  if (operator === 'equal') {
    if (n !== undefined && cell.n !== undefined) return cell.n === n;
    const want = value.toLowerCase();
    const text = cell.text.toLowerCase();
    return hasWildcards(want) ? matchesWhole(parsePattern(want), text) : text === want;
  }
  if (cell.text === '' && cell.n === undefined) return false;
  let d: number;
  if (n !== undefined) {
    if (cell.n === undefined) return false;
    d = cell.n - n;
  } else {
    if (cell.n !== undefined) return false;
    d = cell.text.localeCompare(value, undefined, { sensitivity: 'base' });
  }
  return operator === 'greaterThan' ? d > 0 : operator === 'greaterThanOrEqual' ? d >= 0 : operator === 'lessThan' ? d < 0 : d <= 0;
}

/** The numbers among a column's cells. */
const numbersOf = (cells: readonly FilterCell[]) => cells.flatMap((c) => (c.n === undefined ? [] : [c.n]));

/**
 * The cut-off for Excel's Top 10: the smallest of the largest `count`
 * numbers (or the largest of the smallest), or of so many percent of them,
 * at least one. Ties with it pass too.
 */
export function top10Threshold(cond: Extract<FilterCondition, { kind: 'top10' }>, cells: readonly FilterCell[]): number | undefined {
  const numbers = numbersOf(cells).sort((a, b) => (cond.top ? b - a : a - b));
  if (!numbers.length) return undefined;
  const take = cond.percent ? Math.max(1, Math.floor((numbers.length * cond.count) / 100)) : Math.max(1, cond.count);
  return numbers[Math.min(take, numbers.length) - 1];
}

/** The average of a column's numbers, which Above and Below Average compare with. */
export function averageOf(cells: readonly FilterCell[]): number | undefined {
  const numbers = numbersOf(cells);
  return numbers.length ? numbers.reduce((a, b) => a + b, 0) / numbers.length : undefined;
}

/** Which of a column's cells, in order, a condition lets show. */
export function conditionPasses(cond: FilterCondition, cells: readonly FilterCell[]): boolean[] {
  switch (cond.kind) {
    case 'custom':
      return cells.map((cell) => {
        const results = cond.rules.map((r) => meets(r.operator, r.value, cell));
        return cond.and ? results.every(Boolean) : results.some(Boolean);
      });
    case 'top10': {
      const t = cond.threshold ?? top10Threshold(cond, cells);
      return cells.map((c) => t !== undefined && c.n !== undefined && (cond.top ? c.n >= t : c.n <= t));
    }
    case 'average': {
      const avg = cond.average ?? averageOf(cells);
      return cells.map((c) => avg !== undefined && c.n !== undefined && (cond.above ? c.n > avg : c.n < avg));
    }
  }
}

/** A rule as Excel's Custom AutoFilter dialog offers it: its operators, and the text ones it writes with wildcards. */
export type RuleOp = FilterOperator | 'beginsWith' | 'notBeginsWith' | 'endsWith' | 'notEndsWith' | 'contains' | 'notContains';

export const RULE_OPS: readonly RuleOp[] = [
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
  'beginsWith',
  'notBeginsWith',
  'endsWith',
  'notEndsWith',
  'contains',
  'notContains',
];

const WILDCARD_OPS: Partial<Record<RuleOp, [FilterOperator, (v: string) => string]>> = {
  beginsWith: ['equal', (v) => `${v}*`],
  notBeginsWith: ['notEqual', (v) => `${v}*`],
  endsWith: ['equal', (v) => `*${v}`],
  notEndsWith: ['notEqual', (v) => `*${v}`],
  contains: ['equal', (v) => `*${v}*`],
  notContains: ['notEqual', (v) => `*${v}*`],
};

/** A rule as the file keeps it, from the dialog's operator and value: "begins with abc" is "equals abc*". */
export function ruleFor(op: RuleOp, value: string): { operator: FilterOperator; value: string } {
  const wild = WILDCARD_OPS[op];
  return wild ? { operator: wild[0], value: wild[1](value) } : { operator: op as FilterOperator, value };
}

/** The dialog's operator and value for a rule from the file, as Excel shows it again: "equals abc*" is "begins with abc". */
export function ruleOp(rule: { operator: FilterOperator; value: string }): { op: RuleOp; value: string } {
  const plain = (v: string) => v !== '' && !/[*?~]/.test(v);
  if (rule.operator === 'equal' || rule.operator === 'notEqual') {
    const not = rule.operator === 'notEqual';
    const v = rule.value;
    if (v.length > 2 && v.startsWith('*') && v.endsWith('*') && plain(v.slice(1, -1))) return { op: not ? 'notContains' : 'contains', value: v.slice(1, -1) };
    if (v.endsWith('*') && plain(v.slice(0, -1))) return { op: not ? 'notBeginsWith' : 'beginsWith', value: v.slice(0, -1) };
    if (v.startsWith('*') && plain(v.slice(1))) return { op: not ? 'notEndsWith' : 'endsWith', value: v.slice(1) };
  }
  return { op: rule.operator, value: rule.value };
}

/**
 * A value typed into a rule, as the file keeps it: in a column of numbers or
 * dates, the number (a date's serial number) the text means in the locale;
 * text otherwise. Undefined where a number or date column is compared with
 * something that isn't one.
 */
export function typedRuleValue(text: string, op: RuleOp, kind: 'text' | 'number' | 'date', locale: string, date1904: boolean): string | undefined {
  const t = text.trim();
  const comparison = op === 'greaterThan' || op === 'greaterThanOrEqual' || op === 'lessThan' || op === 'lessThanOrEqual';
  if (kind === 'text' || WILDCARD_OPS[op] || t === '') return comparison && kind !== 'text' ? undefined : text;
  const parsed = parseInput(t, 'General', locale, date1904);
  const n = !parsed.error && parsed.content.kind === 'number' ? parsed.content.value : undefined;
  const date = !!parsed.numFmt && isDateFormat(parsed.numFmt);
  if (n !== undefined && (kind === 'number' || date)) return String(n);
  return comparison ? undefined : text;
}
