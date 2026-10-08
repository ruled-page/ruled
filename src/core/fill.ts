import type { CellContent, SeriesOptions } from './types';

export interface FillSource {
  content: CellContent;
  /** The source cell holds a number shown as a date. */
  isDate: boolean;
}

/** For each target: copy a source cell (formulas get shifted) or use a computed value. */
export type FillPlan = { kind: 'copy'; source: number } | { kind: 'value'; source: number; content: CellContent };

/** Excel's built-in lists, in English: what filling continues when it is given no lists of its own. */
export const ENGLISH_LISTS: readonly (readonly string[])[] = [
  ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
  ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
  ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
];

const mod = (a: number, n: number) => ((a % n) + n) % n;

/** Quarter names that run 1 to 4 and start again: Q1, Qtr2, Quarter 3. */
const QUARTER = /^(Q|Qtr|Quarter)(\s?)([1-4])$/i;

const DAY_MS = 86_400_000;

/** Day zero of the workbook's date system, as a time in ms. */
const dayZero = (date1904: boolean) => (date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30));

/**
 * Dates a whole number of months apart, on the same day of the month (or
 * each on the last day of its month), as a step in months; otherwise null.
 */
function monthStep(serials: number[], date1904: boolean): { start: Date; step: number; monthEnd: boolean } | null {
  if (serials.some((s) => !Number.isInteger(s) || s < 61)) return null;
  const dates = serials.map((s) => new Date(dayZero(date1904) + s * DAY_MS));
  const lastDay = (d: Date) => new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)).getUTCDate();
  const monthEnd = dates.every((d) => d.getUTCDate() === lastDay(d));
  if (!monthEnd && dates.some((d) => d.getUTCDate() !== dates[0].getUTCDate())) return null;
  const months = dates.map((d) => d.getUTCFullYear() * 12 + d.getUTCMonth());
  const step = months[1] - months[0];
  if (!step || months.some((m, i) => i > 0 && m - months[i - 1] !== step)) return null;
  return { start: dates[0], step, monthEnd };
}

/** Rounds away binary noise such as 0.30000000000000004, like Excel's 15 digits. */
const tidy = (n: number) => parseFloat(n.toPrecision(15));

/** The least-squares straight line through the values at x = 0, 1, 2…, unrounded. */
function fit(ys: readonly number[]): (x: number) => number {
  const k = ys.length;
  const mx = (k - 1) / 2;
  const my = ys.reduce((a, b) => a + b, 0) / k;
  let num = 0;
  let den = 0;
  ys.forEach((y, x) => {
    num += (x - mx) * (y - my);
    den += (x - mx) * (x - mx);
  });
  const slope = den ? num / den : 0;
  return (x) => my + slope * (x - mx);
}

function linearFit(ys: readonly number[]): (x: number) => number {
  const f = fit(ys);
  return (x) => tidy(f(x));
}

/**
 * A word from a list in the case of the cell it continues: all capitals, all
 * small letters or a capital first letter, as Excel keeps it; otherwise as
 * the list writes it.
 */
function matchCase(word: string, like: string): string {
  if (like === like.toUpperCase()) return word.toUpperCase();
  if (like === like.toLowerCase()) return word.toLowerCase();
  const first = like.charAt(0);
  const rest = like.slice(1);
  if (first !== first.toLowerCase() && rest === rest.toLowerCase()) return word.charAt(0).toUpperCase() + word.slice(1);
  return word;
}

/**
 * Plans an Excel-style drag fill along one row or column. Source cells sit
 * at positions 0..k-1; `positions` are the target positions, beyond k-1 when
 * filling down/right and below 0 when filling up/left. Text in one of the
 * `lists` (matched whatever its case) continues through that list; the
 * first list that has every source word in it is the one used.
 */
export function planFill(sources: FillSource[], positions: number[], date1904 = false, lists: readonly (readonly string[])[] = ENGLISH_LISTS): FillPlan[] {
  const k = sources.length;
  const copy = () => positions.map((x): FillPlan => ({ kind: 'copy', source: mod(x, k) }));
  if (k === 0) return [];

  const kinds = new Set(sources.map((s) => s.content.kind));
  if (kinds.size !== 1) return copy();
  const kind = sources[0].content.kind;

  if (kind === 'number') {
    const ys = sources.map((s) => (s.content as { value: number }).value);
    if (k === 1 && !sources[0].isDate) return copy();
    // 1 Jan, 1 Feb runs on by months, as Excel does, rather than by 31 days.
    const months = k > 1 && sources.every((s) => s.isDate) ? monthStep(ys, date1904) : null;
    if (months) {
      const { start, step, monthEnd } = months;
      return positions.map((x) => {
        const y = start.getUTCFullYear();
        const m = start.getUTCMonth() + step * x;
        const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
        const day = monthEnd ? last : Math.min(start.getUTCDate(), last);
        const value = Math.round((Date.UTC(y, m, day) - dayZero(date1904)) / DAY_MS);
        return { kind: 'value', source: mod(x, k), content: { kind: 'number', value } };
      });
    }
    const f = k === 1 ? (x: number) => ys[0] + x : linearFit(ys);
    return positions.map((x) => ({ kind: 'value', source: mod(x, k), content: { kind: 'number', value: f(x) } }));
  }

  if (kind === 'string') {
    const texts = sources.map((s) => (s.content as { value: string }).value);

    const lower = texts.map((t) => t.toLowerCase());
    for (const list of lists) {
      if (list.length < 2) continue;
      const idx = lower.map((t) => list.findIndex((w) => w.toLowerCase() === t));
      if (idx.some((i) => i < 0)) continue;
      const step = k === 1 ? 1 : idx[1] - idx[0];
      if (idx.some((v, i) => i > 0 && mod(v - idx[i - 1], list.length) !== mod(step, list.length))) return copy();
      return positions.map((x) => ({
        kind: 'value',
        source: mod(x, k),
        content: { kind: 'string', value: matchCase(list[mod(idx[0] + step * x, list.length)], texts[0]) },
      }));
    }

    const quarters = texts.map((t) => QUARTER.exec(t.trim()));
    if (quarters.every((q) => q && q[1].toLowerCase() === quarters[0]![1].toLowerCase() && q[2] === quarters[0]![2])) {
      const ns = quarters.map((q) => Number(q![3]) - 1);
      const step = k === 1 ? 1 : ns[1] - ns[0];
      if (ns.every((v, i) => i === 0 || mod(v - ns[i - 1], 4) === mod(step, 4))) {
        const [, word, space] = quarters[0]!;
        return positions.map((x) => ({ kind: 'value', source: mod(x, k), content: { kind: 'string', value: `${word}${space}${mod(ns[0] + step * x, 4) + 1}` } }));
      }
    }

    const parts = texts.map((t) => /^(.*?)(\d+)(\D*)$/.exec(t));
    if (parts.every(Boolean) && parts.every((m) => m![1] === parts[0]![1] && m![3] === parts[0]![3])) {
      const ns = parts.map((m) => parseInt(m![2], 10));
      const width = parts[0]![2].length;
      const f = k === 1 ? (x: number) => ns[0] + x : linearFit(ns);
      return positions.map((x) => {
        const n = Math.round(f(x));
        const digits = String(Math.abs(n)).padStart(width, '0');
        return {
          kind: 'value',
          source: mod(x, k),
          content: { kind: 'string', value: `${parts[0]![1]}${n < 0 ? '-' : ''}${digits}${parts[0]![3]}` },
        };
      });
    }
  }

  return copy();
}

/** The last day Excel shows as a date, 31 December 9999. */
const LAST_DATE = 2_958_465;

/** A serial date's weekday, 0 being Sunday, as Excel's WEEKDAY counts (1 January 1900 being a Sunday). */
const weekday = (day: number, date1904: boolean) => mod(day + (date1904 ? 5 : 6), 7);

/** A serial date's calendar day. Excel's 1900 system counts a 29 February 1900 that never was, so earlier days are one out. */
function calendarDay(day: number, date1904: boolean): Date {
  return new Date((!date1904 && day < 60 ? Date.UTC(1899, 11, 31) : dayZero(date1904)) + day * DAY_MS);
}

function serialDay(time: number, date1904: boolean): number {
  const days = Math.round((time - dayZero(date1904)) / DAY_MS);
  return !date1904 && days < 61 ? days - 1 : days;
}

/** The date `months` months on, on the same day of the month or the last day of a shorter month, at the same time of day. */
function addMonths(serial: number, months: number, date1904: boolean): number {
  const day = Math.floor(serial);
  const d = calendarDay(day, date1904);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + months;
  const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
  return serialDay(Date.UTC(y, m, Math.min(d.getUTCDate(), last)), date1904) + (serial - day);
}

/** The date `n` weekdays on (back, below zero), skipping Saturdays and Sundays, as Excel's WORKDAY counts. */
function addWeekdays(serial: number, n: number, date1904: boolean): number {
  let day = Math.floor(serial);
  const dir = Math.sign(n);
  let left = Math.abs(n);
  // Any seven days in a row have five weekdays in them.
  const weeks = Math.max(0, Math.floor((left - 1) / 5));
  day += dir * 7 * weeks;
  left -= 5 * weeks;
  while (left > 0) {
    day += dir;
    const w = weekday(day, date1904);
    if (w !== 0 && w !== 6) left--;
  }
  return day + (serial - Math.floor(serial));
}

export interface SeriesLine {
  /** The position the values start at: 1 after the first cell, or 0 when a trend replaces the numbers the line starts with. */
  from: number;
  values: number[];
  /** Whether the series goes up (1), down (-1) or stays where it is (0). */
  direction: number;
}

/**
 * One row or column of Excel's Fill Series (Linear, Growth or Date) from the
 * numbers it starts with, for a line of `length` cells: the value for each
 * cell from `from` on, ending before the first that would pass the stop
 * value, as Excel leaves the cells after it as they are. Without a trend only
 * the first number counts, and the step is added to it (Linear), multiplies
 * it (Growth) or moves the date on in its unit (Date). Weekdays skip
 * Saturdays and Sundays; months and years keep the day of the month, or take
 * the last day of a shorter month. With a trend, every number the line
 * starts with is used, and replaced, by a least-squares straight line
 * (Linear) or exponential curve (Growth), ignoring the step. Undefined when
 * there is no series: a growth trend through numbers not all above zero.
 */
export function seriesLine(starts: readonly number[], length: number, options: Omit<SeriesOptions, 'by'>, date1904 = false): SeriesLine | undefined {
  const { type, step, stop, dateUnit } = options;
  const trend = options.trend && (type === 'linear' || type === 'growth');
  const start = starts[0];
  const whole = Math.trunc(step);
  let at: (i: number, prev: number) => number;
  if (trend && type === 'growth') {
    if (starts.some((y) => !(y > 0))) return undefined;
    const f = fit(starts.map(Math.log));
    at = (i) => tidy(Math.exp(f(i)));
  } else if (trend) {
    at = linearFit(starts);
  } else if (type === 'growth') {
    at = (i) => tidy(start * step ** i);
  } else if (type === 'date' && dateUnit === 'weekday') {
    at = (_i, prev) => addWeekdays(prev, whole, date1904);
  } else if (type === 'date' && dateUnit !== 'day') {
    at = (i) => tidy(addMonths(start, whole * i * (dateUnit === 'year' ? 12 : 1), date1904));
  } else {
    at = (i) => tidy(start + step * i);
  }
  const first = trend ? at(0, start) : start;
  const direction = Math.sign(at(1, first) - first);
  const passes = (v: number) => stop !== undefined && (direction > 0 ? v > stop : direction < 0 && v < stop);
  // A trend replaces the numbers it is fitted to, whatever the stop value.
  const kept = trend ? starts.length : 1;
  const from = trend ? 0 : 1;
  const values: number[] = [];
  let prev = start;
  for (let i = from; i < length; i++) {
    const v = at(i, prev);
    if (!Number.isFinite(v) || (type === 'date' && (v < 0 || v > LAST_DATE - (date1904 ? 1462 : 0)))) break;
    if (i >= kept && passes(v)) break;
    values.push(v);
    prev = v;
  }
  return { from, values, direction };
}
