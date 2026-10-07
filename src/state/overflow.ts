import { VT, type DisplayStyle } from '../core/types';
import { alignment, displayText, type CellRecord } from './format';
import type { SheetView } from './sheetView';

/** Horizontal padding Glide leaves inside a cell, in pixels. */
export const CELL_PADDING = 4.5;

/** Text that spills out of its cell into empty neighbours, as Excel draws it. */
export interface Overflow {
  /** The cell the text belongs to. */
  col: number;
  /** First and last columns the text covers, including its own. */
  from: number;
  to: number;
  text: string;
  style: DisplayStyle;
  color?: string;
  align: 'left' | 'right' | 'center';
  /** Width of the text in pixels. */
  width: number;
}

const isEmpty = (rec: CellRecord | null) => !rec || rec.t === VT.None;

/**
 * Text cells in a row that are wider than their column, with the empty
 * neighbours they spill into. Left-aligned text spills right, right-aligned
 * text spills left and centred text both ways; numbers, wrapped text and
 * merged cells never spill.
 */
export function rowOverflows(
  view: SheetView,
  row: number,
  styleOf: (col: number, rec: CellRecord) => DisplayStyle,
  date1904: boolean,
  measure: (text: string, style: DisplayStyle) => number,
): Overflow[] {
  const cells: [number, CellRecord][] = [];
  view.forEach({ r1: row, r2: row, c1: 0, c2: view.visibleCols[view.colCount - 1] ?? 0 }, (_r, c, rec) => {
    if (rec.t === VT.String) cells.push([c, rec]);
  });
  const out: Overflow[] = [];
  const free = (c: number) => c >= 0 && isEmpty(view.get(row, c)) && !view.mergeAt(row, c);
  for (const [col, rec] of cells) {
    const style = styleOf(col, rec);
    if (style.wrap || style.hAlign === 'fill' || style.hAlign === 'justify' || view.mergeAt(row, col)) continue;
    const gc = view.gridCol(col);
    if (gc < 0) continue;
    const { text, color } = displayText(rec, style, date1904);
    if (!text) continue;
    const width = measure(text, style);
    const own = view.colWidth(col);
    const extra = width + CELL_PADDING * 2 - own;
    if (extra <= 0) continue;
    const align = alignment(rec, style);
    // Walk visible columns outwards until the text fits or a cell is in the way.
    const spill = (dir: 1 | -1, need: number) => {
      let g = gc;
      while (need > 0) {
        const next = view.visibleCols[g + dir];
        if (next === undefined || !free(next)) break;
        g += dir;
        need -= view.colWidth(next);
      }
      return view.visibleCols[g];
    };
    let from = col;
    let to = col;
    if (align === 'left') to = spill(1, extra);
    else if (align === 'right') from = spill(-1, extra);
    else {
      from = spill(-1, extra / 2);
      to = spill(1, extra / 2);
    }
    if (from !== col || to !== col) out.push({ col, from, to, text, style, color, align, width });
  }
  return out;
}

/** Left edge of the text, relative to the left edge of its own cell. */
export function overflowTextX(o: Overflow, ownWidth: number): number {
  if (o.align === 'left') return CELL_PADDING;
  if (o.align === 'right') return ownWidth - CELL_PADDING - o.width;
  return ownWidth / 2 - o.width / 2;
}

/** Excel's scientific notation: 1.23E+14, with at least two exponent digits. */
function scientific(n: number, digits: number): string {
  const [mantissa, exp] = n.toExponential(digits).split('e');
  const e = Number(exp);
  return `${mantissa}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
}

/**
 * A number as Excel shows it when it is wider than its column: General
 * numbers drop decimals, then switch to scientific notation; numbers with a
 * format, and anything still too wide, become a row of #.
 */
export function fitNumber(text: string, n: number, general: boolean, available: number, width: (s: string) => number): string {
  if (width(text) <= available) return text;
  if (general && Number.isFinite(n)) {
    const abs = Math.abs(n);
    if (abs >= 1e-4 && abs < 1e11) {
      const decimals = (text.split('.')[1] ?? '').length;
      for (let d = decimals - 1; d >= 0; d--) {
        const t = String(Number(n.toFixed(d)));
        if (width(t) <= available) return t;
      }
    }
    for (let d = 5; d >= 0; d--) {
      const t = scientific(n, d);
      if (width(t) <= available) return t;
    }
  }
  return '#'.repeat(Math.max(1, Math.floor(available / Math.max(1, width('#')))));
}
