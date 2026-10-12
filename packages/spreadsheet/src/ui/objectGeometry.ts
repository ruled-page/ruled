import { MAX_COLS, MAX_ROWS } from '@sitepen/formula-engine/refs';
import type { Anchor, SheetInfo, SheetObject } from '../core/types';
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

export { objectKey } from '../core/objectKey';

/** Whether an object can be moved, resized or deleted: a chart, picture or shape on a worksheet that isn't protected. */
export function objectEditable(info: Pick<SheetInfo, 'kind' | 'protected'>, obj: SheetObject): boolean {
  return info.kind === 'worksheet' && !info.protected && !!obj.from && (obj.kind === 'chart' || obj.kind === 'image' || obj.kind === 'shape') && (!!obj.source || !!obj.chartXml);
}

/** One end of an anchor moved `by` pixels along its rows or columns, hidden ones counting for nothing; never before the first. */
function shiftLine(shown: Int32Array, line: number, off: number, by: number, size: (i: number) => number, position: (i: number) => number, max: number): [number, number] {
  const last = shown.length ? shown[shown.length - 1] : -1;
  const sizeOf = (i: number) => (i > last || position(i) >= 0 ? size(i) : 0);
  let at = off + by;
  while (at < 0 && line > 0) at += sizeOf(--line);
  at = Math.max(0, at);
  while (line < max - 1 && at >= sizeOf(line)) at -= sizeOf(line++);
  return [line, Math.round(at)];
}

/** An anchor moved `dx` pixels right and `dy` down: the cell it then sits in, and how far into it. */
export function shiftAnchor(view: SheetView, a: Anchor, dx: number, dy: number): Anchor {
  const [col, colOffPx] = shiftLine(view.visibleCols, a.col, a.colOffPx, dx, (c) => view.colWidth(c), (c) => view.gridCol(c), MAX_COLS);
  const [row, rowOffPx] = shiftLine(view.visibleRows, a.row, a.rowOffPx, dy, (r) => view.rowHeight(r), (r) => view.gridRow(r), MAX_ROWS);
  return { col, row, colOffPx, rowOffPx };
}
