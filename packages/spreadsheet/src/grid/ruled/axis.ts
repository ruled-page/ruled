import type { SheetView } from '../../state/sheetView';

/**
 * The rows, or the columns, a grid shows, hidden ones left out, and where
 * each starts. Most lines are the sheet's default size; only the few that
 * differ are kept, with running totals of how much they differ, so where a
 * line starts, and which line is at a distance, take a binary search even
 * at row 1,048,576. Distances are from the start of the first line.
 */
export class Axis {
  /** The lines whose size differs from the default, in order, and how much the sizes before each differ in all. */
  private readonly lines: Int32Array;
  private readonly sizes: Float64Array;
  private readonly before: Float64Array;

  /**
   * `count` lines of `base` pixels, but for those in `sizes`: [line, size]
   * pairs, by grid line, in any order.
   */
  constructor(
    readonly count: number,
    readonly base: number,
    sizes: Iterable<readonly [number, number]> = [],
  ) {
    const own = [...sizes].filter(([line, size]) => line >= 0 && line < count && size !== base).sort((a, b) => a[0] - b[0]);
    this.lines = Int32Array.from(own, ([line]) => line);
    this.sizes = Float64Array.from(own, ([, size]) => size);
    this.before = new Float64Array(own.length + 1);
    for (let i = 0; i < own.length; i++) this.before[i + 1] = this.before[i] + own[i][1] - base;
  }

  /** How many of the lines with sizes of their own come before `line`. */
  private ownBefore(line: number): number {
    let lo = 0;
    let hi = this.lines.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.lines[mid] < line) lo = mid + 1;
      else hi = mid;
    }
    return lo;
  }

  size(line: number): number {
    const i = this.ownBefore(line);
    return this.lines[i] === line ? this.sizes[i] : this.base;
  }

  /** Where a line starts. Line `count` is where the last one ends. */
  start(line: number): number {
    return line * this.base + this.before[this.ownBefore(line)];
  }

  /** The length of all the lines. */
  get total(): number {
    return this.start(this.count);
  }

  /** The line at a distance: the first or the last for one before or past them all. */
  at(px: number): number {
    let lo = 0;
    let hi = this.count - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.start(mid) <= px) lo = mid;
      else hi = mid - 1;
    }
    return Math.max(0, lo);
  }
}

/** The sheet's rows as the grid shows them. */
export function rowAxis(view: SheetView): Axis {
  const sizes: [number, number][] = [];
  for (const [row, px] of Object.entries(view.info.rowHeights)) sizes.push([view.gridRow(Number(row)), px]);
  return new Axis(view.rowCount, view.info.defaultRowHeightPx, sizes);
}

/** The sheet's columns as the grid shows them, after the row numbers. */
export function colAxis(view: SheetView): Axis {
  const sizes: [number, number][] = [];
  for (const [col, px] of view.colWidths) sizes.push([view.gridCol(col), px]);
  return new Axis(view.colCount, view.info.defaultColWidthPx, sizes);
}
