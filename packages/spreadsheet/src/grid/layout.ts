import type { Box } from './contract';
import type { Axis } from './axis';

/**
 * What a grid's layout is worked out from. Lines are grid lines: rows and
 * columns in view order, hidden ones left out, starting from 0.
 */
export interface LayoutSpec {
  /** The area drawn in, its scrollbars aside. */
  width: number;
  height: number;
  rows: Axis;
  cols: Axis;
  /** How many of the first rows and columns stay put while the rest scroll. */
  frozenRows: number;
  frozenCols: number;
  /** The width of the row numbers and the height of the column letters. */
  numbers: number;
  letters: number;
  /** How far the rows and columns that scroll have been scrolled. */
  scrollX: number;
  scrollY: number;
  /** The sheet reads right to left, column A at the right. */
  rtl?: boolean;
}

/** A run of lines in view: the first and last, partly shown ones included; first > last for none. */
export interface Run {
  first: number;
  last: number;
}

/** What is at a point: the corner, a column's letters, a row's number or a cell, by grid line (-1 for the headings). */
export interface Hit {
  area: 'corner' | 'letters' | 'numbers' | 'cells';
  row: number;
  col: number;
}

/**
 * Where everything is on screen. Positions are worked out from the start
 * edge (the left, or the right when the sheet reads right to left) and
 * turned into the screen's only in box() and hit(), so nothing else needs
 * to know which way the sheet reads. The column letters run along the
 * top, the row numbers down the start edge; then come the frozen rows and
 * columns, and the rest, scrolled.
 */
export class Layout {
  /** Where the lines that scroll begin: after the headings and the frozen lines. */
  readonly bodyX: number;
  readonly bodyY: number;
  /** The frozen lines, and those of the rest in view. */
  readonly frozenRowRun: Run;
  readonly frozenColRun: Run;
  readonly rowRun: Run;
  readonly colRun: Run;

  constructor(readonly spec: LayoutSpec) {
    const { rows, cols, numbers, letters } = spec;
    const fr = Math.min(spec.frozenRows, rows.count);
    const fc = Math.min(spec.frozenCols, cols.count);
    this.bodyX = numbers + cols.start(fc);
    this.bodyY = letters + rows.start(fr);
    this.frozenRowRun = { first: 0, last: fr - 1 };
    this.frozenColRun = { first: 0, last: fc - 1 };
    this.rowRun = run(rows, fr, spec.scrollY, spec.height - this.bodyY);
    this.colRun = run(cols, fc, spec.scrollX, spec.width - this.bodyX);
  }

  /** The length the scroller needs: the headings and every line. */
  get contentWidth(): number {
    return this.spec.numbers + this.spec.cols.total;
  }

  get contentHeight(): number {
    return this.spec.letters + this.spec.rows.total;
  }

  /** Where a column starts, from the start edge, where it is drawn: frozen columns don't scroll. */
  x(col: number): number {
    const { cols, numbers, scrollX } = this.spec;
    const fc = this.frozenColRun.last + 1;
    return col < fc ? numbers + cols.start(col) : this.bodyX + cols.start(col) - cols.start(fc) - scrollX;
  }

  y(row: number): number {
    const { rows, letters, scrollY } = this.spec;
    const fr = this.frozenRowRun.last + 1;
    return row < fr ? letters + rows.start(row) : this.bodyY + rows.start(row) - rows.start(fr) - scrollY;
  }

  /** A box on screen, from where it starts measured from the start edge. */
  box(x: number, y: number, width: number, height: number): Box {
    return { x: this.spec.rtl ? this.spec.width - x - width : x, y, width, height };
  }

  cell(row: number, col: number): Box {
    return this.box(this.x(col), this.y(row), this.spec.cols.size(col), this.spec.rows.size(row));
  }

  /** The box rows `y1` to `y2` and columns `x1` to `x2` take, each where it is drawn: across the freeze line, through it. */
  span(y1: number, y2: number, x1: number, x2: number): Box {
    const x = this.x(x1);
    const y = this.y(y1);
    return this.box(x, y, this.x(x2) + this.spec.cols.size(x2) - x, this.y(y2) + this.spec.rows.size(y2) - y);
  }

  /** A row's number, or a column's letters. */
  rowHeading(row: number): Box {
    return this.box(0, this.y(row), this.spec.numbers, this.spec.rows.size(row));
  }

  colHeading(col: number): Box {
    return this.box(this.x(col), 0, this.spec.cols.size(col), this.spec.letters);
  }

  corner(): Box {
    return this.box(0, 0, this.spec.numbers, this.spec.letters);
  }

  /**
   * The parts of the cells' area, each drawn clipped to its own box: the
   * frozen rows and columns, the frozen rows' other columns, the frozen
   * columns' other rows, and the rest. Parts with no lines are left out.
   */
  quadrants(): { rows: Run; cols: Run; clip: Box }[] {
    const { width, height, numbers, letters } = this.spec;
    const bands = [
      [this.frozenRowRun, letters, this.bodyY],
      [this.rowRun, this.bodyY, height],
    ] as const;
    const columns = [
      [this.frozenColRun, numbers, this.bodyX],
      [this.colRun, this.bodyX, width],
    ] as const;
    const out: { rows: Run; cols: Run; clip: Box }[] = [];
    for (const [rows, top, bottom] of bands) {
      for (const [cols, start, end] of columns) {
        if (rows.first > rows.last || cols.first > cols.last || bottom <= top || end <= start) continue;
        out.push({ rows, cols, clip: this.box(start, top, end - start, bottom - top) });
      }
    }
    return out;
  }

  /**
   * What is at a point on screen. With `near`, a point outside the cells
   * gives the nearest cell in view instead, as a drag past an edge wants.
   */
  hit(px: number, py: number, near = false): Hit {
    const { width, height, numbers, letters, rows, cols } = this.spec;
    let x = this.spec.rtl ? width - px : px;
    let y = py;
    if (near) {
      x = Math.min(Math.max(x, numbers), width - 1);
      y = Math.min(Math.max(y, letters), height - 1);
    }
    const col = x < numbers ? -1 : this.lineAt(cols, x, numbers, this.bodyX, this.frozenColRun, this.colRun, this.spec.scrollX);
    const row = y < letters ? -1 : this.lineAt(rows, y, letters, this.bodyY, this.frozenRowRun, this.rowRun, this.spec.scrollY);
    const area = row < 0 && col < 0 ? 'corner' : row < 0 ? 'letters' : col < 0 ? 'numbers' : 'cells';
    return { area, row, col };
  }

  /** The line at a distance along an axis: among the frozen lines before `body`, else among those scrolled. */
  private lineAt(axis: Axis, at: number, start: number, body: number, frozen: Run, rest: Run, scroll: number): number {
    if (at < body && frozen.last >= 0) return Math.min(axis.at(at - start), frozen.last);
    if (rest.first > rest.last) return frozen.last;
    return Math.min(Math.max(axis.at(at - body + axis.start(frozen.last + 1) + scroll), rest.first), rest.last);
  }
}

/** The lines from `first` on in view, scrolled `scroll` past it, in `room` pixels. */
function run(axis: Axis, first: number, scroll: number, room: number): Run {
  if (first >= axis.count || room <= 0) return { first, last: first - 1 };
  const from = axis.start(first) + scroll;
  return { first: Math.max(first, axis.at(from)), last: axis.at(from + room - 1) };
}

/**
 * Browsers won't make an element much taller than 17.9 million pixels
 * (Firefox) or 33.5 million (Chrome, Safari); a sheet to row 1,048,576 is
 * 21 million. A scroller is made at most this long, and its position is
 * scaled to the content's past it.
 */
export const SCROLL_LIMIT = 8_000_000;

/** How long the scroller is made for content `length` long. */
export function scrollLength(length: number): number {
  return Math.min(length, SCROLL_LIMIT);
}

/** How far the content is scrolled for a scroller scrolled to `pos`, `view` pixels of it in view. */
export function contentScroll(pos: number, length: number, view: number): number {
  const room = scrollLength(length) - view;
  return room > 0 && length > SCROLL_LIMIT ? (pos * (length - view)) / room : pos;
}

/** Where to scroll the scroller to have the content scrolled to `pos`. */
export function scrollerPosition(pos: number, length: number, view: number): number {
  const room = length - view;
  return room > 0 && length > SCROLL_LIMIT ? (pos * (scrollLength(length) - view)) / room : pos;
}
