import { CellFlag, VT, type DisplayStyle, type SheetGrid, type SheetInfo } from '../src/core/types';
import type { CellRecord } from '../src/state/format';
import { SheetView } from '../src/state/sheetView';

/** Excel's default cell format, Calibri 11pt in General. */
export const PLAIN: DisplayStyle = {
  numFmt: 'General',
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  fontSizePx: 14.7,
  fontFamily: 'Calibri',
  wrap: false,
  indent: 0,
};

export const style = (extra: Partial<DisplayStyle> = {}): DisplayStyle => ({ ...PLAIN, ...extra });

export const num = (n: number, extra: Partial<CellRecord> = {}): CellRecord => ({ t: VT.Number, n, style: 0, flags: 0, ...extra });
export const str = (s: string, extra: Partial<CellRecord> = {}): CellRecord => ({ t: VT.String, s, style: 0, flags: 0, ...extra });
export const bool = (b: boolean): CellRecord => ({ t: VT.Bool, n: b ? 1 : 0, style: 0, flags: 0 });
export const err = (s?: string): CellRecord => ({ t: VT.Error, s, style: 0, flags: 0 });

export function sheetInfo(extra: Partial<SheetInfo> = {}): SheetInfo {
  return {
    index: 0,
    name: 'Sheet1',
    kind: 'worksheet',
    state: 'visible',
    maxRow: -1,
    maxCol: -1,
    defaultColWidthPx: 64,
    defaultRowHeightPx: 20,
    colWidths: {},
    rowHeights: {},
    hiddenCols: [],
    hiddenRows: [],
    frozenCols: 0,
    frozenRows: 0,
    rowStyles: {},
    colStyles: [],
    merges: [],
    listValidations: [],
    ruleValidations: [],
    blocked: [],
    objects: [],
    notes: [],
    tables: [],
    conditionalFormats: [],
    links: [],
    protected: false,
    showGridLines: true,
    ...extra,
  };
}

/** A file's cells, as the worker sends them: [row, col, value, style?, formula?]; null is a formatted blank. */
export type FileCell = [number, number, number | string | boolean | { error: string } | null, number?, string?];

export function grid(cells: FileCell[]): SheetGrid {
  const sorted = [...cells].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const rowCount = sorted.length ? sorted[sorted.length - 1][0] + 1 : 0;
  const rowPtr = new Int32Array(rowCount + 1);
  for (const [row] of sorted) rowPtr[row + 1]++;
  for (let r = 0; r < rowCount; r++) rowPtr[r + 1] += rowPtr[r];
  const g: SheetGrid = {
    rowCount,
    rowPtr,
    cols: Int32Array.from(sorted, (c) => c[1]),
    types: new Uint8Array(sorted.length),
    nums: new Float64Array(sorted.length),
    strs: [],
    styles: Int32Array.from(sorted, (c) => c[3] ?? 0),
    flags: new Uint8Array(sorted.length),
    formulas: new Map(),
  };
  sorted.forEach(([, , v, , f], i) => {
    if (typeof v === 'number') [g.types[i], g.nums[i]] = [VT.Number, v];
    else if (typeof v === 'boolean') [g.types[i], g.nums[i]] = [VT.Bool, v ? 1 : 0];
    else if (typeof v === 'string') [g.types[i], g.strs[i]] = [VT.String, v];
    else if (v === null) g.types[i] = VT.None;
    else [g.types[i], g.strs[i]] = [VT.Error, v.error];
    if (f !== undefined) {
      g.formulas.set(i, f);
      g.flags[i] = CellFlag.Formula;
    }
  });
  return g;
}

/** A sheet whose used range is worked out from its cells. */
export function sheet(cells: FileCell[], extra: Partial<SheetInfo> = {}): SheetView {
  const maxRow = Math.max(-1, ...cells.map((c) => c[0]));
  const maxCol = Math.max(-1, ...cells.map((c) => c[1]));
  return new SheetView(sheetInfo({ maxRow, maxCol, ...extra }), grid(cells));
}
