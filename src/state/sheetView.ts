import { msg, type Text } from '../i18n/message';
import { cellKey, inRect, MAX_COLS, MAX_ROWS } from '../core/refs';
import { findCell } from '../core/sheetScanner';
import type { BlockedRegion, CellNote, CellUpdate, CfCell, CfResult, DiffStyle, MergeRange, Rect, SheetGrid, SheetInfo, TableInfo } from '../core/types';
import { CellFlag, VT } from '../core/types';
import type { CellRecord } from './format';

/** Rows and columns shown beyond the used range so there is room to type. */
const EXTRA_ROWS = 200;
const EXTRA_COLS = 10;

/**
 * The UI's view of one sheet: the file's cells plus an overlay of everything
 * the worker has reported since, and the mapping between grid positions and
 * sheet rows/columns once hidden rows and columns are left out.
 */
export class SheetView {
  readonly info: SheetInfo;
  private readonly grid: SheetGrid | null;
  private readonly overlay = new Map<number, CellRecord>();
  /** Grid row → sheet row, skipping hidden rows. */
  visibleRows: Int32Array = new Int32Array();
  /** Grid column (after the row-number column) → sheet column. */
  visibleCols: Int32Array = new Int32Array();
  private rowLimit = 0;
  private colLimit = 0;
  readonly colWidths = new Map<number, number>();
  private readonly mergesByRow = new Map<number, MergeRange[]>();
  private readonly notes = new Map<number, CellNote>();
  private readonly cf = new Map<number, CfCell>();
  /** Formats referred to by conditional formatting results. */
  cfStyles: DiffStyle[] = [];
  /** Extra read-only rules, e.g. from a schema. */
  constraint?: (row: number, col: number) => Text | undefined;
  version = 0;

  constructor(info: SheetInfo, grid: SheetGrid | null) {
    this.info = info;
    this.grid = grid;
    for (const m of info.merges) {
      for (let r = m.r1; r <= m.r2; r++) {
        const list = this.mergesByRow.get(r);
        if (list) list.push(m);
        else this.mergesByRow.set(r, [m]);
      }
    }
    for (const [c, w] of Object.entries(info.colWidths)) this.colWidths.set(Number(c), w);
    for (const n of info.notes) this.notes.set(cellKey(n.row, n.col), n);
    this.extendTo(Math.max(info.maxRow, 0), Math.max(info.maxCol, 0));
  }

  /** Makes sure the given cell is inside the scrollable area. */
  extendTo(row: number, col: number): boolean {
    const rows = Math.min(MAX_ROWS, Math.max(row + 1 + EXTRA_ROWS, 500, this.rowLimit));
    const cols = Math.min(MAX_COLS, Math.max(col + 1 + EXTRA_COLS, 26, this.colLimit));
    if (rows === this.rowLimit && cols === this.colLimit) return false;
    this.rowLimit = rows;
    this.colLimit = cols;
    this.layout();
    return true;
  }

  private layout(): void {
    const hiddenRows = new Set(this.info.hiddenRows);
    const hiddenCols = new Set(this.info.hiddenCols);
    const vr: number[] = [];
    for (let r = 0; r < this.rowLimit; r++) if (!hiddenRows.has(r) && !this.filteredOut.has(r)) vr.push(r);
    const vc: number[] = [];
    for (let c = 0; c < this.colLimit; c++) if (!hiddenCols.has(c)) vc.push(c);
    this.visibleRows = Int32Array.from(vr);
    this.visibleCols = Int32Array.from(vc);
  }

  /**
   * A filter applied in this editor: the range it covers (its first row
   * holds the headings) and, per column, the values whose rows stay visible.
   * It hides rows on screen only; the file's own filter is left as it was.
   */
  filter?: { rect: Rect; columns: Map<number, Set<string>> };
  private readonly filteredOut = new Set<number>();

  /** Shows only the rows of `rect` whose text in every filtered column is allowed; no columns clears it. */
  setFilter(rect: Rect, columns: Map<number, Set<string>>, textAt: (row: number, col: number) => string): void {
    this.filter = columns.size ? { rect, columns } : undefined;
    this.filteredOut.clear();
    if (this.filter) {
      for (let row = rect.r1 + 1; row <= rect.r2; row++) {
        for (const [col, allowed] of columns) {
          if (!allowed.has(textAt(row, col))) {
            this.filteredOut.add(row);
            break;
          }
        }
      }
    }
    this.layout();
    this.version++;
  }

  /** Rows the editor's filter hides. */
  get filteredCount(): number {
    return this.filteredOut.size;
  }

  get rowCount(): number {
    return this.visibleRows.length;
  }

  get colCount(): number {
    return this.visibleCols.length;
  }

  /** Grid index of a sheet row/column, or -1 when it is hidden or out of range. */
  gridRow(row: number): number {
    return search(this.visibleRows, row);
  }

  gridCol(col: number): number {
    return search(this.visibleCols, col);
  }

  get(row: number, col: number): CellRecord | null {
    const o = this.overlay.get(cellKey(row, col));
    if (o) return o;
    const g = this.grid;
    if (!g) return null;
    const i = findCell(g, row, col);
    if (i < 0) return null;
    const rec: CellRecord = { t: g.types[i], style: g.styles[i], flags: g.flags[i] };
    if (rec.t === VT.Number || rec.t === VT.Bool) rec.n = g.nums[i];
    else if (rec.t === VT.String || rec.t === VT.Error) rec.s = g.strs[i];
    const f = g.formulas.get(i);
    if (f !== undefined && rec.flags & CellFlag.Formula) rec.f = f;
    return rec;
  }

  apply(u: CellUpdate): void {
    this.overlay.set(cellKey(u.row, u.col), { t: u.t, n: u.n, s: u.s, style: u.style, f: u.f, flags: u.flags });
    if (u.row >= this.rowLimit - 20 || u.col >= this.colLimit - 2) this.extendTo(u.row + 20, u.col + 2);
    this.version++;
  }

  /** Calls fn for each non-empty cell inside rect (file cells and overlay). */
  forEach(rect: Rect, fn: (row: number, col: number, rec: CellRecord) => void): void {
    const g = this.grid;
    const seen = new Set<number>();
    if (g) {
      for (let row = rect.r1; row <= Math.min(rect.r2, g.rowCount - 1); row++) {
        for (let i = g.rowPtr[row]; i < g.rowPtr[row + 1]; i++) {
          const col = g.cols[i];
          if (col < rect.c1 || col > rect.c2) continue;
          const key = cellKey(row, col);
          seen.add(key);
          const rec = this.get(row, col);
          if (rec) fn(row, col, rec);
        }
      }
    }
    for (const [key, rec] of this.overlay) {
      if (seen.has(key)) continue;
      const row = Math.floor(key / MAX_COLS);
      const col = key % MAX_COLS;
      if (inRect(rect, row, col)) fn(row, col, rec);
    }
  }

  /** The format a cell is drawn with: its own, else its row's, else its column's. */
  styleAt(row: number, col: number, rec: CellRecord | null = this.get(row, col)): number {
    if (rec) return rec.style;
    const own = this.info.rowStyles[row];
    if (own !== undefined) return own;
    return this.info.colStyles.find((c) => col >= c.min && col <= c.max)?.style ?? 0;
  }

  /** Takes the worker's latest conditional formatting results for this sheet. */
  setConditionalFormats(r: CfResult): void {
    this.cf.clear();
    for (const c of r.cells) this.cf.set(cellKey(c.row, c.col), c);
    this.cfStyles = r.styles;
    this.version++;
  }

  cfAt(row: number, col: number): CfCell | undefined {
    return this.cf.get(cellKey(row, col));
  }

  tableAt(row: number, col: number): TableInfo | undefined {
    for (const t of this.info.tables) if (inRect(t.rect, row, col)) return t;
    return undefined;
  }

  noteAt(row: number, col: number): CellNote | undefined {
    return this.notes.get(cellKey(row, col));
  }

  blockedAt(row: number, col: number): BlockedRegion | undefined {
    return this.info.blocked.find((b) => inRect(b, row, col));
  }

  mergeAt(row: number, col: number): MergeRange | undefined {
    return this.mergesByRow.get(row)?.find((m) => col >= m.c1 && col <= m.c2);
  }

  readOnlyReason(row: number, col: number): Text | undefined {
    if (this.info.kind !== 'worksheet') return msg('sheet.notWorksheet');
    if (this.info.protected) return msg('sheet.protected');
    const m = this.mergeAt(row, col);
    if (m && (m.r1 !== row || m.c1 !== col)) return msg('sheet.mergedPart');
    return this.blockedAt(row, col)?.reason ?? this.constraint?.(row, col);
  }

  colWidth(col: number): number {
    return this.colWidths.get(col) ?? this.info.defaultColWidthPx;
  }

  rowHeight(row: number): number {
    return this.info.rowHeights[row] ?? this.info.defaultRowHeightPx;
  }
}

function search(arr: Int32Array, v: number): number {
  let lo = 0;
  let hi = arr.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (arr[mid] === v) return mid;
    if (arr[mid] < v) lo = mid + 1;
    else hi = mid - 1;
  }
  return -1;
}
