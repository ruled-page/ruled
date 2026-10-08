import { isHyperlinkFormula, literalHyperlinkTarget } from '../core/links';
import type { DisplayStyle, Hyperlink } from '../core/types';
import type { SheetView } from './sheetView';

/** A cell's link as the screen shows and follows it: one of the sheet's links, or a HYPERLINK formula's. */
export interface CellLink {
  /** An address outside the workbook; for a HYPERLINK formula, "#" and a place in the workbook goes there. */
  target?: string;
  /** A place in the workbook, or with `target` a part of that page. */
  location?: string;
  /** The link's screen tip. */
  tooltip?: string;
  /** The sheet's link, when the cell's link is one rather than a formula's. */
  link?: Hyperlink;
}

/**
 * The link on a cell, or on the merged cells it is part of: the sheet's own
 * link there, else a HYPERLINK formula's, whose address the worker works out
 * (or the formula writes, when it hasn't yet).
 */
export function cellLink(view: SheetView, row: number, col: number): CellLink | undefined {
  const merge = view.mergeAt(row, col);
  const r = merge?.r1 ?? row;
  const c = merge?.c1 ?? col;
  const link = view.linkAt(r, c);
  if (link) {
    const out: CellLink = { link };
    if (link.target !== undefined) out.target = link.target;
    if (link.location !== undefined) out.location = link.location;
    if (link.tooltip !== undefined) out.tooltip = link.tooltip;
    return out;
  }
  const rec = view.get(r, c);
  if (rec?.f === undefined || !isHyperlinkFormula(rec.f)) return undefined;
  const target = rec.link ?? literalHyperlinkTarget(rec.f);
  return target === undefined ? {} : { target };
}

const drawn = new WeakMap<DisplayStyle, Map<string, DisplayStyle>>();

/**
 * A linked cell's format, as Excel's Hyperlink cell style draws it: in the
 * link colour and underlined, unless the cell's font has a colour of its own.
 * A font in the workbook theme's hyperlink colour (the Hyperlink style's) is
 * drawn in `color` too, so links stay readable whatever the editor's theme.
 */
export function linkStyle(style: DisplayStyle, color: string): DisplayStyle {
  if (style.color && !style.linkFont) return style;
  let byColor = drawn.get(style);
  if (!byColor) drawn.set(style, (byColor = new Map()));
  let out = byColor.get(color);
  if (!out) byColor.set(color, (out = { ...style, color, underline: style.linkFont ? style.underline : true }));
  return out;
}

/** What a host's onLinkClick hears when someone follows a link to a web page or an email. */
export interface LinkClick {
  /** The address, as the browser would open it. */
  url: string;
  kind: 'web' | 'email';
  /** The sheet and cell the link is on. */
  sheet: string;
  ref: string;
}
