import { hasWildcards, matchesWhole, parsePattern } from '@ruled/formula/wildcard';
import { isDateFormat } from 'numfmt';
import { parseInput } from '../core/input';
import type { DatePeriod, FilterCondition, FilterOperator } from '../core/types';

/** A cell as a filter sees it: the text it shows, its number when it holds one, and its colours and icon as drawn. */
export interface FilterCell {
  text: string;
  n?: number;
  /** Its fill and text colours, conditional formatting's included; no fill, or the automatic text colour, missing. */
  fill?: string;
  font?: string;
  icon?: { set: string; index: number };
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
export function conditionPasses(cond: FilterCondition, cells: readonly FilterCell[], date1904 = false): boolean[] {
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
    case 'color': {
      const want = cond.color?.toUpperCase();
      return cells.map((c) => (cond.font ? c.font : c.fill)?.toUpperCase() === want);
    }
    case 'icon':
      return cells.map((c) => (cond.index === undefined ? !c.icon : c.icon?.set === cond.set && c.icon.index === cond.index));
    case 'period': {
      const range = cond.from !== undefined && cond.to !== undefined ? { from: cond.from, to: cond.to } : undefined;
      return cells.map((c) => c.n !== undefined && (range ? c.n >= range.from && c.n < range.to : inPeriod(cond.period, c.n, date1904)));
    }
  }
}

/** The periods counted from today, as the Date Filters menu lists them, then the quarters and months of any year. */
export const DATE_PERIODS: readonly DatePeriod[] = ['tomorrow', 'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth', 'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter', 'nextYear', 'thisYear', 'lastYear', 'yearToDate'];
export const ANY_YEAR_PERIODS: readonly DatePeriod[] = ['Q1', 'Q2', 'Q3', 'Q4', 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11', 'M12'];

const DAY = 86_400_000;
/** A day's serial number in the 1900 date system, from its UTC date. */
const serial = (y: number, m: number, d: number) => (Date.UTC(y, m, d) - Date.UTC(1899, 11, 30)) / DAY;

/** Whether a date (a serial number) is in a quarter or month of any year. */
function inPeriod(period: DatePeriod, n: number, date1904: boolean): boolean {
  const month = new Date(Date.UTC(1899, 11, 30) + (Math.floor(n) + (date1904 ? 1462 : 0)) * DAY).getUTCMonth();
  const q = /^Q(\d)$/.exec(period);
  if (q) return Math.floor(month / 3) === Number(q[1]) - 1;
  const m = /^M(\d+)$/.exec(period);
  return !!m && month === Number(m[1]) - 1;
}

/**
 * The dates a period counted from `today` covers, as serial numbers: from
 * the first up to but not including the last, weeks starting on Sunday as
 * in Excel. Undefined for the quarters and months of any year.
 */
export function periodRange(period: DatePeriod, today: Date, date1904: boolean): { from: number; to: number } | undefined {
  const y = today.getFullYear();
  const m = today.getMonth();
  const d = today.getDate();
  const shift = date1904 ? 1462 : 0;
  const span = (from: number, to: number) => ({ from: from - shift, to: to - shift });
  const day = (k: number) => span(serial(y, m, d + k), serial(y, m, d + k + 1));
  const week = (k: number) => span(serial(y, m, d - today.getDay() + 7 * k), serial(y, m, d - today.getDay() + 7 * k + 7));
  const month = (k: number) => span(serial(y, m + k, 1), serial(y, m + k + 1, 1));
  const quarter = (k: number) => span(serial(y, m - (m % 3) + 3 * k, 1), serial(y, m - (m % 3) + 3 * k + 3, 1));
  const year = (k: number) => span(serial(y + k, 0, 1), serial(y + k + 1, 0, 1));
  const ranges: Partial<Record<DatePeriod, () => { from: number; to: number }>> = {
    tomorrow: () => day(1),
    today: () => day(0),
    yesterday: () => day(-1),
    nextWeek: () => week(1),
    thisWeek: () => week(0),
    lastWeek: () => week(-1),
    nextMonth: () => month(1),
    thisMonth: () => month(0),
    lastMonth: () => month(-1),
    nextQuarter: () => quarter(1),
    thisQuarter: () => quarter(0),
    lastQuarter: () => quarter(-1),
    nextYear: () => year(1),
    thisYear: () => year(0),
    lastYear: () => year(-1),
    yearToDate: () => span(serial(y, 0, 1), serial(y, m, d + 1)),
  };
  return ranges[period]?.();
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
