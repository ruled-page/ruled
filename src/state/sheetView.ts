import { msg, type Text } from '../i18n/message';
import { applyLinkChange } from '../core/links';
import { cellKey, inRect, MAX_COLS, MAX_ROWS } from '../core/refs';
import { findCell } from '../core/sheetScanner';
import type { BlockedRegion, CellNote, CellUpdate, CfCell, CfResult, DiffStyle, FilterState, Hyperlink, LayoutOp, MergeRange, Rect, SheetGrid, SheetInfo, TableInfo } from '../core/types';
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
  /** Links on one row, by row; links over several rows are few, and looked through. */
  private readonly linksByRow = new Map<number, Hyperlink[]>();
  private tallLinks: Hyperlink[] = [];
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
    this.indexMerges();
    this.indexLinks();
    for (const [c, w] of Object.entries(info.colWidths)) this.colWidths.set(Number(c), w);
    // The file's own filter: the rows it hides are hidden by it, and showing them again clears it.
    if (info.filter) {
      const f = info.filter;
      this.filter = { rect: f.rect, columns: new Map(Object.entries(f.columns).map(([c, v]) => [Number(c), new Set(v)])), table: f.table };
      for (const row of info.hiddenRows) if (row > f.rect.r1 && row <= f.rect.r2) this.filteredOut.add(row);
      info.hiddenRows = info.hiddenRows.filter((row) => !this.filteredOut.has(row));
    }
    for (const n of info.notes) this.notes.set(cellKey(n.row, n.col), n);
    this.extendTo(Math.max(info.maxRow, 0), Math.max(info.maxCol, 0));
  }

  private indexMerges(): void {
    this.mergesByRow.clear();
    for (const m of this.info.merges) {
      for (let r = m.r1; r <= m.r2; r++) {
        const list = this.mergesByRow.get(r);
        if (list) list.push(m);
        else this.mergesByRow.set(r, [m]);
      }
    }
  }

  private indexLinks(): void {
    this.linksByRow.clear();
    this.tallLinks = [];
    for (const l of this.info.links) {
      const list = this.linksByRow.get(l.rect.r1);
      if (l.rect.r1 !== l.rect.r2) this.tallLinks.push(l);
      else if (list) list.push(l);
      else this.linksByRow.set(l.rect.r1, [l]);
    }
  }

  /** Bumped when sizes, hidden rows and columns, frozen panes or merges change. */
  layoutVersion = 0;

  /** Shows a change to the layout that the worker has made. */
  applyLayout(op: LayoutOp): void {
    const info = this.info;
    switch (op.kind) {
      case 'colWidths':
        for (const [c, px] of Object.entries(op.widths)) {
          if (px === null) this.colWidths.delete(Number(c));
          else this.colWidths.set(Number(c), px);
        }
        break;
      case 'rowHeights':
        for (const [r, px] of Object.entries(op.heights)) {
          if (px === null) delete info.rowHeights[Number(r)];
          else info.rowHeights[Number(r)] = px;
        }
        break;
      case 'hidden': {
        const key = op.axis === 'rows' ? 'hiddenRows' : 'hiddenCols';
        const hidden = new Set(info[key]);
        for (const [line, h] of Object.entries(op.lines)) {
          if (h) hidden.add(Number(line));
          else hidden.delete(Number(line));
        }
        info[key] = [...hidden].sort((a, b) => a - b);
        this.layout();
        break;
      }
      case 'freeze':
        info.frozenRows = op.rows;
        info.frozenCols = op.cols;
        break;
      case 'filter':
        this.applyFilter(op.filter, op.rows);
        break;
      case 'merges': {
        const same = (a: Rect, b: Rect) => a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2;
        info.merges = [...info.merges.filter((m) => !op.remove.some((r) => same(r, m))), ...op.add.filter((a) => !info.merges.some((m) => same(m, a)))];
        this.indexMerges();
        break;
      }
      case 'links':
        info.links = applyLinkChange(info.links, op).links;
        this.indexLinks();
        break;
      case 'outline': {
        const was = info.outline?.[op.axis];
        const levels = { ...was?.levels };
        for (const [line, v] of Object.entries(op.levels)) {
          if (v) levels[Number(line)] = v;
          else delete levels[Number(line)];
        }
        const collapsed = new Set(was?.collapsed);
        for (const [line, c] of Object.entries(op.collapsed)) {
          if (c) collapsed.add(Number(line));
          else collapsed.delete(Number(line));
        }
        const blank = { levels: {}, collapsed: [], summaryAfter: true };
        const next = { levels, collapsed: [...collapsed].sort((a, b) => a - b), summaryAfter: was?.summaryAfter ?? true };
        info.outline = { rows: info.outline?.rows ?? blank, cols: info.outline?.cols ?? blank, [op.axis]: next };
        break;
      }
    }
    this.layoutVersion++;
    this.version++;
  }

  /**
   * Makes sure the given cell is inside the scrollable area, with room
   * beyond it: as Excel's, the sheet goes on to its last row and column as
   * it is scrolled or moved through. True when the area grew.
   */
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
   * The filter: the range it covers (its first row holds the headings) and,
   * per column, the values whose rows stay visible; the table whose filter
   * it is, when it is one's.
   */
  filter?: { rect: Rect; columns: Map<number, Set<string>>; table?: string };
  private readonly filteredOut = new Set<number>();

  /** The rows of `rect` that `columns` would hide: those whose text in some filtered column isn't allowed. */
  static rowsFilteredBy(rect: Rect, columns: Map<number, Set<string>>, textAt: (row: number, col: number) => string): number[] {
    const out: number[] = [];
    for (let row = rect.r1 + 1; row <= rect.r2; row++) {
      for (const [col, allowed] of columns) {
        if (!allowed.has(textAt(row, col))) {
          out.push(row);
          break;
        }
      }
    }
    return out;
  }

  /** Shows a filter, or none, and the rows it hides. */
  private applyFilter(filter: FilterState | null, rows: number[]): void {
    this.filter = filter ? { rect: filter.rect, columns: new Map(Object.entries(filter.columns).map(([c, v]) => [Number(c), new Set(v)])), table: filter.table } : undefined;
    this.filteredOut.clear();
    for (const row of rows) this.filteredOut.add(row);
    this.layout();
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

  /**
   * The first and last grid rows and columns showing part of a sheet
   * rectangle, or undefined when all of it is hidden or beyond the grid.
   */
  gridSpan(rect: Rect): { y1: number; y2: number; x1: number; x2: number } | undefined {
    const [y1, y2] = span(this.visibleRows, rect.r1, rect.r2);
    const [x1, x2] = span(this.visibleCols, rect.c1, rect.c2);
    return y1 <= y2 && x1 <= x2 ? { y1, y2, x1, x2 } : undefined;
  }

  /** The first and last grid rows (or columns) showing any of the sheet's rows (or columns) a to b, or undefined when all are hidden. */
  lineSpan(axis: 'rows' | 'cols', a: number, b: number): [number, number] | undefined {
    const s = span(axis === 'rows' ? this.visibleRows : this.visibleCols, a, b);
    return s[0] <= s[1] ? s : undefined;
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
    this.overlay.set(cellKey(u.row, u.col), { t: u.t, n: u.n, s: u.s, style: u.style, f: u.f, flags: u.flags, ...(u.link !== undefined && { link: u.link }), ...(u.spill && { spill: u.spill }), ...(u.array && { array: u.array }) });
    if (u.row >= this.rowLimit - 20 || u.col >= this.colLimit - 2) this.extendTo(u.row + 20, u.col + 2);
    this.version++;
  }

  /** Cells updated since the file was read whose flags include `flag`, in no particular order. */
  flagged(flag: number): { row: number; col: number }[] {
    const out: { row: number; col: number }[] = [];
    for (const [key, rec] of this.overlay) if (rec.flags & flag) out.push({ row: Math.floor(key / MAX_COLS), col: key % MAX_COLS });
    return out;
  }

  /** Calls fn for each non-empty cell inside rect (file cells and overlay). */
  /** The last row and column in use: the file's used range, grown by cells typed or pasted beyond it. */
  extent(): { maxRow: number; maxCol: number } {
    let maxRow = this.info.maxRow;
    let maxCol = this.info.maxCol;
    for (const key of this.overlay.keys()) {
      maxRow = Math.max(maxRow, Math.floor(key / MAX_COLS));
      maxCol = Math.max(maxCol, key % MAX_COLS);
    }
    return { maxRow, maxCol };
  }

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

  /** Shows a cell's note or thread as it now is; null takes it away. */
  setNote(row: number, col: number, note: CellNote | null): void {
    if (note) this.notes.set(cellKey(row, col), note);
    else this.notes.delete(cellKey(row, col));
    this.version++;
  }

  blockedAt(row: number, col: number): BlockedRegion | undefined {
    return this.info.blocked.find((b) => inRect(b, row, col));
  }

  /** Unlocks one of the file's dynamic arrays, which the worker calculates. */
  unblockArray(rect: Rect): void {
    this.info.blocked = this.info.blocked.filter((b) => !(b.kind === 'array' && b.r1 === rect.r1 && b.c1 === rect.c1 && b.r2 === rect.r2 && b.c2 === rect.c2));
    this.version++;
  }

  /** The sheet's own link on a cell (not a HYPERLINK formula's): the first that covers it. */
  linkAt(row: number, col: number): Hyperlink | undefined {
    const inCol = (l: Hyperlink) => col >= l.rect.c1 && col <= l.rect.c2;
    return this.linksByRow.get(row)?.find(inCol) ?? this.tallLinks.find((l) => row >= l.rect.r1 && row <= l.rect.r2 && inCol(l));
  }

  mergeAt(row: number, col: number): MergeRange | undefined {
    return this.mergesByRow.get(row)?.find((m) => col >= m.c1 && col <= m.c2);
  }

  /** Whether a format leaves its cells unlocked on a protected sheet; the store knows the formats. */
  isUnlocked: (style: number) => boolean = () => false;

  readOnlyReason(row: number, col: number): Text | undefined {
    if (this.info.kind !== 'worksheet') return msg('sheet.notWorksheet');
    if (this.info.protected && !this.isUnlocked(this.styleAt(row, col))) return msg('sheet.locked');
    const m = this.mergeAt(row, col);
    if (m && (m.r1 !== row || m.c1 !== col)) return msg('sheet.mergedPart');
    return this.blockedAt(row, col)?.reason ?? this.constraint?.(row, col);
  }

  /**
   * Why a cell's note or comments can't be changed here, if they can't. A
   * protected sheet keeps them only when its protection covers objects, as
   * Excel's does, whether or not the cell is locked; a schema's read-only
   * cells keep theirs too.
   */
  commentReadOnlyReason(row: number, col: number): Text | undefined {
    if (this.info.kind !== 'worksheet') return msg('sheet.notWorksheet');
    if (this.info.protected && this.info.objectsLocked) return msg('core.reason.commentsLocked');
    const m = this.mergeAt(row, col);
    if (m && (m.r1 !== row || m.c1 !== col)) return msg('sheet.mergedPart');
    return this.constraint?.(row, col);
  }

  colWidth(col: number): number {
    return this.colWidths.get(col) ?? this.info.defaultColWidthPx;
  }

  rowHeight(row: number): number {
    return this.info.rowHeights[row] ?? this.info.defaultRowHeightPx;
  }
}

/** In a sorted array, the first index holding at least `lo` and the last holding at most `hi`. */
function span(arr: Int32Array, lo: number, hi: number): [number, number] {
  let a = 0;
  let b = arr.length;
  while (a < b) {
    const mid = (a + b) >> 1;
    if (arr[mid] < lo) a = mid + 1;
    else b = mid;
  }
  let c = a;
  let d = arr.length;
  while (c < d) {
    const mid = (c + d) >> 1;
    if (arr[mid] <= hi) c = mid + 1;
    else d = mid;
  }
  return [a, c - 1];
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
