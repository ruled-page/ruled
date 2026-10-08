import type { SheetObject } from '../core/types';
import type { SheetView } from '../state/sheetView';

/**
 * Sizes of the rows or columns an object spans, for placing it over the
 * grid. Hidden ones count for nothing, as objects shrink with them in Excel;
 * ones past the grid's last row or column are counted in full.
 */
function span(shown: Int32Array, from: number, to: number, size: (i: number) => number, position: (i: number) => number): number {
  const last = shown.length ? shown[shown.length - 1] : -1;
  let total = 0;
  for (let i = from; i < to; i++) if (i > last || position(i) >= 0) total += size(i);
  return total;
}

/** Height of the sheet rows from `from` up to (not including) `to`. */
export function rowsHeight(view: SheetView, from: number, to: number): number {
  return span(view.visibleRows, from, to, (r) => view.rowHeight(r), (r) => view.gridRow(r));
}

/** Width of the sheet columns from `from` up to (not including) `to`. */
export function colsWidth(view: SheetView, from: number, to: number): number {
  return span(view.visibleCols, from, to, (c) => view.colWidth(c), (c) => view.gridCol(c));
}

/** An object's size in pixels: between its two anchors, or as the file gives it for one anchored at one cell. */
export function objectSize(view: SheetView, obj: SheetObject): { width: number; height: number } | undefined {
  const { from, to } = obj;
  if (!from) return undefined;
  if (!to) return obj.sizePx ? { width: obj.sizePx.w, height: obj.sizePx.h } : undefined;
  return {
    width: colsWidth(view, from.col, to.col) - from.colOffPx + to.colOffPx,
    height: rowsHeight(view, from.row, to.row) - from.rowOffPx + to.rowOffPx,
  };
}
