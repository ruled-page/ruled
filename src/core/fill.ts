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
export function planFill(sources: FillSource[], positions: number[]): FillPlan[] {
  const k = sources.length;
  const copy = () => positions.map((x): FillPlan => ({ kind: 'copy', source: mod(x, k) }));
  if (k === 0) return [];

  const kinds = new Set(sources.map((s) => s.content.kind));
  if (kinds.size !== 1) return copy();
  const kind = sources[0].content.kind;

  if (kind === 'number') {
    const ys = sources.map((s) => (s.content as { value: number }).value);
    if (k === 1 && !sources[0].isDate) return copy();
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
