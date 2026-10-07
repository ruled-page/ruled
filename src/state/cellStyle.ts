import { tableCellFormat } from '../core/tableStyles';
import type { CellBorders, DiffStyle, DisplayStyle } from '../core/types';
import type { SheetView } from './sheetView';

type Composed = WeakMap<DiffStyle, WeakMap<DisplayStyle, DisplayStyle>>;
const under: Composed = new WeakMap();
const over: Composed = new WeakMap();

function cached(cache: Composed, base: DisplayStyle, diff: DiffStyle, make: () => DisplayStyle): DisplayStyle {
  let byBase = cache.get(diff);
  if (!byBase) cache.set(diff, (byBase = new WeakMap()));
  let out = byBase.get(base);
  if (!out) byBase.set(base, (out = make()));
  return out;
}

function sides(first: CellBorders | undefined, second: CellBorders | undefined): CellBorders | undefined {
  if (!first) return second && { left: second.left, right: second.right, top: second.top, bottom: second.bottom };
  if (!second) return first;
  return {
    left: first.left ?? second.left,
    right: first.right ?? second.right,
    top: first.top ?? second.top,
    bottom: first.bottom ?? second.bottom,
  };
}

/** A table style under a cell's own format: anything the cell sets itself wins. */
export function underlay(base: DisplayStyle, diff: DiffStyle): DisplayStyle {
  return cached(under, base, diff, () => ({
    ...base,
    fill: base.fill ?? diff.fill,
    color: base.color ?? diff.color,
    bold: base.bold || !!diff.bold,
    italic: base.italic || !!diff.italic,
    underline: base.underline || !!diff.underline,
    strike: base.strike || !!diff.strike,
    border: sides(base.border, diff.border),
  }));
}

/** A conditional format over a cell's format: whatever the rule sets wins. */
export function overlay(base: DisplayStyle, diff: DiffStyle): DisplayStyle {
  return cached(over, base, diff, () => ({
    ...base,
    fill: diff.fill ?? base.fill,
    color: diff.color ?? base.color,
    bold: diff.bold ?? base.bold,
    italic: diff.italic ?? base.italic,
    underline: diff.underline ?? base.underline,
    strike: diff.strike ?? base.strike,
    border: sides(diff.border, base.border),
  }));
}

/** The style a cell is drawn with: its own format, over its table's style. */
export function effectiveStyle(view: SheetView, styles: DisplayStyle[], row: number, col: number, styleIndex = view.styleAt(row, col)): DisplayStyle {
  let style = styles[styleIndex] ?? styles[0];
  const table = view.tableAt(row, col);
  const format = table && tableCellFormat(table, row, col);
  if (format) style = underlay(style, format);
  return style;
}
