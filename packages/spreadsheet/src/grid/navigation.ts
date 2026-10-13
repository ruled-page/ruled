import { MAX_COLS, MAX_ROWS } from '@ruled/formula/refs';
import { VT, type Rect } from '../core/types';
import type { SheetView } from '../state/sheetView';
import type { Selection } from '../state/store';

/** What moving the selection needs to know of a sheet. */
export interface NavSheet {
  /** The sheet's rows and columns in the grid's order, hidden ones left out. */
  rows: Int32Array;
  cols: Int32Array;
  /** Whether a cell holds anything. */
  filled(row: number, col: number): boolean;
  /** The last row and column holding anything. */
  maxRow: number;
  maxCol: number;
  /** The sheet's last row and column. */
  lastRow: number;
  lastCol: number;
  /** How many of the first rows and columns are frozen. */
  frozenRows: number;
  frozenCols: number;
  /** How many rows a page down moves, and columns a page across. */
  pageRows: number;
  pageCols: number;
  /** The sheet's merged cells, and the merge a cell is part of. */
  merges: readonly Rect[];
  /** The sheet reads right to left, so the left and right arrows step to the next column and the one before. */
  rtl?: boolean;
  mergeAt(row: number, col: number): Rect | undefined;
}

/** A key, with whether Ctrl (⌘ on a Mac) is held. */
export interface NavKey {
  key: string;
  shiftKey: boolean;
  altKey: boolean;
  command: boolean;
}

/** A sheet's cells as navigation sees them, a page being `pageRows` rows and `pageCols` columns. */
export function navSheet(view: SheetView, pageRows: number, pageCols: number): NavSheet {
  // Worked out only for the keys that need it: it looks through every cell changed.
  let extent: { maxRow: number; maxCol: number } | undefined;
  const used = () => (extent ??= view.extent());
  return {
    rows: view.visibleRows,
    cols: view.visibleCols,
    filled: (row, col) => {
      const rec = view.get(row, col);
      return !!rec && rec.t !== VT.None;
    },
    get maxRow() {
      return used().maxRow;
    },
    get maxCol() {
      return used().maxCol;
    },
    lastRow: MAX_ROWS - 1,
    lastCol: MAX_COLS - 1,
    frozenRows: view.info.frozenRows,
    frozenCols: view.info.frozenCols,
    pageRows: Math.max(1, pageRows),
    pageCols: Math.max(1, pageCols),
    merges: view.info.merges,
    mergeAt: (row, col) => view.mergeAt(row, col),
    rtl: !!view.info.rightToLeft,
  };
}

/** A range grown to take in every merged cell it covers part of, as Excel's selection grows. */
export function growOverMerges(r: Rect, merges: readonly Rect[]): Rect {
  let out = r;
  for (let grew = true; grew; ) {
    grew = false;
    for (const m of merges) {
      if (m.r2 < out.r1 || m.r1 > out.r2 || m.c2 < out.c1 || m.c1 > out.c2) continue;
      if (m.r1 < out.r1 || m.r2 > out.r2 || m.c1 < out.c1 || m.c2 > out.c2) {
        out = { r1: Math.min(out.r1, m.r1), r2: Math.max(out.r2, m.r2), c1: Math.min(out.c1, m.c1), c2: Math.max(out.c2, m.c2) };
        grew = true;
      }
    }
  }
  return out;
}

/** The first index in a sorted array holding more than `v`. */
function after(arr: Int32Array, v: number): number {
  let lo = 0;
  let hi = arr.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (arr[mid] <= v) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** The line `n` shown lines on from `line` (back, for n < 0), stopping at the first or `last`; past those the grid has, the sheet's own. */
function stepLine(order: Int32Array, line: number, n: number, last: number): number {
  if (n > 0) {
    const i = after(order, line) + n - 1;
    if (i < order.length) return order[i];
    const end = order.length ? Math.max(order[order.length - 1], line) : line;
    return Math.min(last, end + (i - order.length + 1));
  }
  const i = after(order, line - 1) + n;
  return i >= 0 ? order[i] : (order[0] ?? line);
}

/**
 * Ctrl+Arrow's move, as Excel makes it: through a block of filled cells to
 * its last, or else to the next filled cell, or else to the sheet's edge.
 */
function edge(sheet: NavSheet, row: number, col: number, dr: number, dc: number): { row: number; col: number } {
  const down = dr !== 0;
  const order = down ? sheet.rows : sheet.cols;
  const last = down ? sheet.lastRow : sheet.lastCol;
  const max = down ? sheet.maxRow : sheet.maxCol;
  const d = dr || dc;
  const at = (line: number) => (down ? sheet.filled(line, col) : sheet.filled(row, line));
  const start = down ? row : col;
  let cur = start;
  let next = stepLine(order, cur, d, last);
  if (next === cur) return { row, col };
  if (at(cur) && at(next)) {
    while (next !== cur && at(next)) {
      cur = next;
      next = stepLine(order, cur, d, last);
    }
  } else {
    cur = next;
    while (!at(cur)) {
      if (d > 0 && cur > max) {
        cur = last;
        break;
      }
      next = stepLine(order, cur, d, last);
      if (next === cur) break;
      cur = next;
    }
  }
  return down ? { row: cur, col } : { row, col: cur };
}

const span = (a: { row: number; col: number }, b: { row: number; col: number }): Rect => ({
  r1: Math.min(a.row, b.row),
  r2: Math.max(a.row, b.row),
  c1: Math.min(a.col, b.col),
  c2: Math.max(a.col, b.col),
});

/** The next cell within a range, along rows (Tab) or down columns (Enter), wrapping round. */
function within(sheet: NavSheet, r: Rect, row: number, col: number, across: boolean, back: boolean): { row: number; col: number } {
  const rows = sheet.rows.filter((x) => x >= r.r1 && x <= r.r2);
  const cols = sheet.cols.filter((x) => x >= r.c1 && x <= r.c2);
  if (!rows.length || !cols.length) return { row, col };
  const cells = across ? rows.length * cols.length : cols.length * rows.length;
  const ri = Math.max(0, rows.indexOf(row));
  const ci = Math.max(0, cols.indexOf(col));
  const at = (j: number) => (across ? { row: rows[Math.floor(j / cols.length)], col: cols[j % cols.length] } : { row: rows[j % rows.length], col: cols[Math.floor(j / rows.length)] });
  let j = across ? ri * cols.length + ci : ci * rows.length + ri;
  // The cells a merge covers are passed over: the merge is its top-left cell.
  for (let n = 0; n < cells; n++) {
    j = (j + (back ? cells - 1 : 1)) % cells;
    const next = at(j);
    const m = sheet.mergeAt(next.row, next.col);
    if (!m || (m.r1 === next.row && m.c1 === next.col)) return next;
  }
  return { row, col };
}

const ARROWS: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

/**
 * How a key moves the selection, as Excel moves it, and the cell to bring
 * into view; undefined for a key that doesn't move it. Shift extends the
 * selection from the active cell, moving its far corner. Enter on a single
 * cell is left to editing.
 */
export function navigate(sel: Selection, e: NavKey, sheet: NavSheet): { selection: Selection; reveal: { row: number; col: number } } | undefined {
  const active = { row: sel.row, col: sel.col };
  const first = sel.rects[0] ?? span(active, active);
  const far = { row: active.row === first.r1 ? first.r2 : first.r1, col: active.col === first.c1 ? first.c2 : first.c1 };
  /** Where a step from the active cell starts: a merge's far side in the step's direction. */
  const merged = sheet.mergeAt(active.row, active.col);
  const stepFrom = (dr: number, dc: number) =>
    merged ? { row: dr > 0 ? merged.r2 : dr < 0 ? merged.r1 : active.row, col: dc > 0 ? merged.c2 : dc < 0 ? merged.c1 : active.col } : active;
  const from = e.shiftKey ? far : active;
  let to: { row: number; col: number } | undefined;
  const arrow = ARROWS[e.key];
  if (arrow && !e.altKey) {
    const dr = arrow[0];
    // The arrows go the way they point on screen: left is the next column on a sheet read right to left.
    const dc = sheet.rtl ? -arrow[1] : arrow[1];
    const start = e.shiftKey ? from : stepFrom(dr, dc);
    to = e.command
      ? edge(sheet, start.row, start.col, dr, dc)
      : { row: dr ? stepLine(sheet.rows, start.row, dr, sheet.lastRow) : start.row, col: dc ? stepLine(sheet.cols, start.col, dc, sheet.lastCol) : start.col };
  } else if (e.key === 'Home' && !e.altKey) {
    const firstRow = sheet.rows[after(sheet.rows, sheet.frozenRows - 1)] ?? sheet.rows[0] ?? 0;
    const firstCol = sheet.cols[after(sheet.cols, sheet.frozenCols - 1)] ?? sheet.cols[0] ?? 0;
    // Ctrl+Home goes to the first cell that scrolls; Home to the row's first cell.
    to = e.command ? { row: firstRow, col: firstCol } : { row: from.row, col: sheet.cols[0] ?? 0 };
  } else if (e.key === 'End' && e.command && !e.altKey) {
    to = { row: Math.max(0, sheet.maxRow), col: Math.max(0, sheet.maxCol) };
  } else if ((e.key === 'PageDown' || e.key === 'PageUp') && !e.command) {
    const d = e.key === 'PageDown' ? 1 : -1;
    to = e.altKey ? { row: from.row, col: stepLine(sheet.cols, from.col, d * sheet.pageCols, sheet.lastCol) } : { row: stepLine(sheet.rows, from.row, d * sheet.pageRows, sheet.lastRow), col: from.col };
  } else if (e.key === 'Backspace' && e.shiftKey && !e.command && !e.altKey) {
    return { selection: { row: active.row, col: active.col, rects: [span(active, active)] }, reveal: active };
  } else if (e.key === 'Backspace' && e.command && !e.shiftKey && !e.altKey) {
    return { selection: sel, reveal: active };
  } else if ((e.key === 'Tab' || e.key === 'Enter') && !e.command && !e.altKey) {
    const tab = e.key === 'Tab';
    // A merge selected on its own is one cell.
    const one = merged && merged.r1 === first.r1 && merged.r2 === first.r2 && merged.c1 === first.c1 && merged.c2 === first.c2;
    const range = !sel.lines && sel.rects.length === 1 && !one && (first.r1 !== first.r2 || first.c1 !== first.c2);
    if (range) {
      const next = within(sheet, first, active.row, active.col, tab, e.shiftKey);
      return { selection: { ...sel, row: next.row, col: next.col }, reveal: next };
    }
    if (!tab) return undefined;
    const start = stepFrom(0, e.shiftKey ? -1 : 1);
    to = { row: active.row, col: stepLine(sheet.cols, start.col, e.shiftKey ? -1 : 1, sheet.lastCol) };
    return cellAt(sheet, to);
  }
  if (!to) return undefined;
  if (e.shiftKey) return { selection: { row: active.row, col: active.col, rects: [growOverMerges(span(active, to), sheet.merges), ...sel.rects.slice(1)] }, reveal: to };
  return cellAt(sheet, to);
}

/** A cell selected on its own: a merge, by its top-left cell, when it is part of one. */
function cellAt(sheet: NavSheet, at: { row: number; col: number }): { selection: Selection; reveal: { row: number; col: number } } {
  const m = sheet.mergeAt(at.row, at.col);
  const to = m ? { row: m.r1, col: m.c1 } : at;
  return { selection: { row: to.row, col: to.col, rects: [m ?? span(to, to)] }, reveal: to };
}
