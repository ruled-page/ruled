import { VT, type DisplayStyle } from '../core/types';
import { effectiveStyle } from './cellStyle';
import { displayText } from './format';
import { CELL_PADDING, excelTextWidth } from './overflow';
import type { SheetView } from './sheetView';

/** How wide text is drawn in a cell's font, in pixels. */
export type MeasureText = (text: string, style: DisplayStyle) => number;

/** The widest a column is made, as dragging its edge allows. */
const MAX_WIDTH = 1200;

/** One line of text in a font, as Excel sizes rows: 15 points for 11-point Calibri. */
const lineHeight = (style: DisplayStyle) => Math.round(style.fontSizePx * 1.36);

/**
 * The width that fits the widest value in a column, as Excel's AutoFit
 * (double-clicking the column's right edge) makes it: every cell with
 * something in it, as it is shown, in its own font, merged cells aside.
 * Undefined for a column with nothing in it.
 */
export function fittedColumnWidth(view: SheetView, styles: DisplayStyle[], col: number, date1904: boolean, locale: string, measure: MeasureText): number | undefined {
  let widest = 0;
  view.forEach({ r1: 0, c1: col, r2: Math.max(view.info.maxRow, 0), c2: col }, (row, _c, rec) => {
    if (rec.t === VT.None || view.mergeAt(row, col)) return;
    const style = effectiveStyle(view, styles, row, col, rec.style);
    const { text } = displayText(rec, style, date1904, false, locale);
    if (!text) return;
    // Numbers are fitted by Excel's own measure too, so none turns into ####.
    const width = rec.t === VT.Number ? Math.max(measure(text, style), excelTextWidth(text, style)) : measure(text, style);
    widest = Math.max(widest, width);
  });
  return widest ? Math.min(MAX_WIDTH, Math.ceil(widest + CELL_PADDING * 2 + 2)) : undefined;
}

/**
 * The height that fits a row's tallest cell, as Excel's AutoFit
 * (double-clicking the bottom edge of its number) makes it: a line of each
 * cell's font, and as many lines as wrapped text takes in its column. A row
 * with nothing in it goes back to the sheet's own height.
 */
export function fittedRowHeight(view: SheetView, styles: DisplayStyle[], row: number, date1904: boolean, locale: string, measure: MeasureText): number {
  let tallest = view.info.defaultRowHeightPx;
  view.forEach({ r1: row, c1: 0, r2: row, c2: Math.max(view.info.maxCol, 0) }, (_r, col, rec) => {
    if (rec.t === VT.None || view.mergeAt(row, col)) return;
    const style = effectiveStyle(view, styles, row, col, rec.style);
    let lines = 1;
    if (style.wrap) {
      const { text } = displayText(rec, style, date1904, false, locale);
      lines = wrappedLines(text, view.colWidth(col) - CELL_PADDING * 2, (s) => measure(s, style));
    }
    tallest = Math.max(tallest, lines * lineHeight(style));
  });
  return tallest;
}

/**
 * Text wrapped into a width as Excel wraps it: at its line breaks, then
 * between words, and a word wider than the width where it reaches the edge.
 */
export function wrapLines(text: string, width: number, measure: (s: string) => number): string[] {
  const lines: string[] = [];
  for (const paragraph of text.split('\n')) {
    let line = '';
    for (const word of paragraph.split(' ')) {
      const next = line ? `${line} ${word}` : word;
      if (measure(next) <= width) {
        line = next;
        continue;
      }
      if (line) lines.push(line);
      let rest = word;
      while (rest.length > 1 && measure(rest) > width) {
        // The most of the word that fits, and at least a character.
        let lo = 1;
        let hi = rest.length - 1;
        while (lo < hi) {
          const mid = Math.ceil((lo + hi) / 2);
          if (measure(rest.slice(0, mid)) <= width) lo = mid;
          else hi = mid - 1;
        }
        lines.push(rest.slice(0, lo));
        rest = rest.slice(lo);
      }
      line = rest;
    }
    lines.push(line);
  }
  return lines;
}

/** How many lines text takes when wrapped into a width. */
export function wrappedLines(text: string, width: number, measure: (s: string) => number): number {
  return wrapLines(text, width, measure).length;
}
