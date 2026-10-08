import type { CellContent } from './types';

export interface FillSource {
  content: CellContent;
  /** The source cell holds a number shown as a date. */
  isDate: boolean;
}

/** For each target: copy a source cell (formulas get shifted) or use a computed value. */
export type FillPlan = { kind: 'copy'; source: number } | { kind: 'value'; source: number; content: CellContent };

const LISTS = [
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

function linearFit(ys: number[]): (x: number) => number {
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
  return (x) => tidy(my + slope * (x - mx));
}

function matchCase(word: string, like: string): string {
  if (like === like.toUpperCase()) return word.toUpperCase();
  if (like === like.toLowerCase()) return word.toLowerCase();
  return word;
}

/**
 * Plans an Excel-style drag fill along one row or column. Source cells sit
 * at positions 0..k-1; `positions` are the target positions, beyond k-1 when
 * filling down/right and below 0 when filling up/left.
 */
export function planFill(sources: FillSource[], positions: number[], date1904 = false): FillPlan[] {
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

    for (const list of LISTS) {
      const idx = texts.map((t) => list.findIndex((w) => w.toLowerCase() === t.toLowerCase()));
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
