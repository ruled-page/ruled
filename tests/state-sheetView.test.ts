import { describe, expect, it } from 'vitest';
import { MAX_COLS, MAX_ROWS } from '../src/core/refs';
import { CellFlag, VT, type DiffStyle, type TableInfo } from '../src/core/types';
import { english, msg } from '../src/i18n';
import { effectiveStyle, overlay, underlay } from '../src/state/cellStyle';
import { workbookReaders } from '../src/state/readers';
import { SheetView } from '../src/state/sheetView';
import { grid, sheet, sheetInfo, style } from './state-fixtures';

const update = (row: number, col: number, value: number | string, extra: { style?: number; f?: string; flags?: number } = {}) => ({
  sheet: 0,
  row,
  col,
  t: typeof value === 'number' ? VT.Number : VT.String,
  n: typeof value === 'number' ? value : undefined,
  s: typeof value === 'string' ? value : undefined,
  style: extra.style ?? 0,
  f: extra.f,
  flags: extra.flags ?? 0,
});

describe("a sheet's cells", () => {
  it("reads each kind of value from the file's grid", () => {
    const view = sheet([
      [0, 0, 1.5, 2],
      [0, 3, 'text'],
      [2, 1, true],
      [2, 2, { error: '#N/A' }],
      [3, 0, 6, 0, 'B2*C2'],
    ]);
    expect(view.get(0, 0)).toEqual({ t: VT.Number, n: 1.5, style: 2, flags: 0 });
    expect(view.get(0, 3)).toEqual({ t: VT.String, s: 'text', style: 0, flags: 0 });
    expect(view.get(2, 1)).toMatchObject({ t: VT.Bool, n: 1 });
    expect(view.get(2, 2)).toMatchObject({ t: VT.Error, s: '#N/A' });
    expect(view.get(3, 0)).toMatchObject({ n: 6, f: 'B2*C2', flags: CellFlag.Formula });
    // Gaps, rows past the grid and negative rows are empty.
    expect(view.get(0, 1)).toBeNull();
    expect(view.get(1, 0)).toBeNull();
    expect(view.get(50, 0)).toBeNull();
    expect(view.get(-1, 0)).toBeNull();
  });

  it('reads a formatted blank as an empty cell with its format', () => {
    const view = sheet([[0, 0, null, 4]]);
    expect(view.get(0, 0)).toEqual({ t: VT.None, style: 4, flags: 0 });
    expect(view.styleAt(0, 0)).toBe(4);
  });

  it('only reports a formula for cells flagged as holding one', () => {
    const g = grid([[0, 0, 6, 0, 'B2*C2']]);
    g.flags[0] = 0;
    expect(new SheetView(sheetInfo({ maxRow: 0, maxCol: 0 }), g).get(0, 0)?.f).toBeUndefined();
  });

  it('shows what the worker reported over what the file had', () => {
    const view = sheet([[0, 0, 1]]);
    const before = view.version;
    view.apply(update(0, 0, 'changed', { style: 3 }));
    view.apply(update(5, 5, 7, { f: 'A1+6', flags: CellFlag.Formula }));
    expect(view.get(0, 0)).toMatchObject({ t: VT.String, s: 'changed', style: 3 });
    expect(view.get(5, 5)).toMatchObject({ t: VT.Number, n: 7, f: 'A1+6' });
    expect(view.version).toBe(before + 2);
  });

  it('works without a grid, for sheets the worker has not sent', () => {
    const view = new SheetView(sheetInfo(), null);
    expect(view.get(0, 0)).toBeNull();
    const seen: number[][] = [];
    view.forEach({ r1: 0, c1: 0, r2: 10, c2: 10 }, (r, c) => seen.push([r, c]));
    expect(seen).toEqual([]);
    view.apply(update(1, 1, 'x'));
    view.forEach({ r1: 0, c1: 0, r2: 10, c2: 10 }, (r, c) => seen.push([r, c]));
    expect(seen).toEqual([[1, 1]]);
  });

  it('visits each non-empty cell in a range once, edits included', () => {
    const view = sheet([
      [0, 0, 'a'],
      [0, 5, 'outside'],
      [1, 1, 'b'],
      [4, 1, 'below'],
    ]);
    view.apply(update(1, 1, 'b2'));
    view.apply(update(2, 2, 'new'));
    view.apply(update(9, 9, 'far'));
    const seen: [number, number, string | undefined][] = [];
    view.forEach({ r1: 0, c1: 0, r2: 2, c2: 3 }, (r, c, rec) => seen.push([r, c, rec.s]));
    expect(seen).toEqual([
      [0, 0, 'a'],
      [1, 1, 'b2'],
      [2, 2, 'new'],
    ]);
  });
});

describe('the scrollable area', () => {
  it('leaves room beyond the used range, and at least 500 rows by 26 columns', () => {
    expect(new SheetView(sheetInfo(), null)).toMatchObject({ rowCount: 500, colCount: 26 });
    const big = new SheetView(sheetInfo({ maxRow: 999, maxCol: 49 }), null);
    expect(big.rowCount).toBe(1000 + 200);
    expect(big.colCount).toBe(50 + 10);
  });

  it('grows when asked to show a cell past it, and only then', () => {
    const view = new SheetView(sheetInfo(), null);
    expect(view.extendTo(10, 5)).toBe(false);
    expect(view.extendTo(600, 40)).toBe(true);
    expect(view.rowCount).toBe(801);
    expect(view.colCount).toBe(51);
    // It never shrinks back.
    expect(view.extendTo(0, 0)).toBe(false);
    expect(view.rowCount).toBe(801);
  });

  it("stops at Excel's last row and column", () => {
    const view = new SheetView(sheetInfo(), null);
    view.extendTo(MAX_ROWS + 10, MAX_COLS + 10);
    expect(view.rowCount).toBe(MAX_ROWS);
    expect(view.colCount).toBe(MAX_COLS);
  });

  it('grows as edits arrive near its edge', () => {
    const view = new SheetView(sheetInfo(), null);
    view.apply(update(490, 0, 'x'));
    expect(view.rowCount).toBe(490 + 20 + 1 + 200);
    view.apply(update(0, 25, 'y'));
    expect(view.colCount).toBe(25 + 2 + 1 + 10);
  });

  it('maps grid positions to sheet rows and columns around hidden ones', () => {
    const view = new SheetView(sheetInfo({ hiddenRows: [1, 2], hiddenCols: [0] }), null);
    expect(Array.from(view.visibleRows.slice(0, 3))).toEqual([0, 3, 4]);
    expect(view.gridRow(3)).toBe(1);
    expect(view.gridRow(2)).toBe(-1);
    expect(view.gridRow(10_000)).toBe(-1);
    expect(view.gridCol(0)).toBe(-1);
    expect(view.gridCol(1)).toBe(0);
    expect(view.gridCol(25)).toBe(24);
  });

  it('uses custom sizes, else the defaults', () => {
    const view = new SheetView(sheetInfo({ colWidths: { 2: 120 }, rowHeights: { 4: 33 } }), null);
    expect(view.colWidth(2)).toBe(120);
    expect(view.colWidth(3)).toBe(64);
    expect(view.rowHeight(4)).toBe(33);
    expect(view.rowHeight(5)).toBe(20);
  });
});

describe("the editor's filter", () => {
  const view = () =>
    sheet([
      [0, 0, 'Region'],
      [1, 0, 'North'],
      [2, 0, 'South'],
      [3, 0, 'North'],
      [4, 0, 'East'],
      [6, 0, 'Not in the table'],
    ]);
  const textAt = (v: SheetView) => (row: number, col: number) => v.get(row, col)?.s ?? '';

  it('hides table rows whose value is not allowed, never the heading row or rows outside', () => {
    const v = view();
    v.setFilter({ r1: 0, c1: 0, r2: 4, c2: 0 }, new Map([[0, new Set(['North'])]]), textAt(v));
    expect(v.filteredCount).toBe(2);
    expect(Array.from(v.visibleRows.slice(0, 4))).toEqual([0, 1, 3, 5]);
    expect(v.filter?.columns.get(0)).toEqual(new Set(['North']));
  });

  it('needs a row to pass every filtered column', () => {
    const v = sheet([
      [0, 0, 'Region'],
      [0, 1, 'Size'],
      [1, 0, 'North'],
      [1, 1, 'L'],
      [2, 0, 'North'],
      [2, 1, 'S'],
    ]);
    v.setFilter(
      { r1: 0, c1: 0, r2: 2, c2: 1 },
      new Map([
        [0, new Set(['North'])],
        [1, new Set(['L'])],
      ]),
      textAt(v),
    );
    expect(v.gridRow(1)).toBe(1);
    expect(v.gridRow(2)).toBe(-1);
  });

  it('clears when no columns are filtered', () => {
    const v = view();
    v.setFilter({ r1: 0, c1: 0, r2: 4, c2: 0 }, new Map([[0, new Set(['East'])]]), textAt(v));
    const version = v.version;
    v.setFilter({ r1: 0, c1: 0, r2: 4, c2: 0 }, new Map(), textAt(v));
    expect(v.filter).toBeUndefined();
    expect(v.filteredCount).toBe(0);
    expect(v.gridRow(2)).toBe(2);
    expect(v.version).toBe(version + 1);
  });
});

describe('what a cell is drawn and edited with', () => {
  const merged = { r1: 1, c1: 1, r2: 2, c2: 3 };
  const info = sheetInfo({
    maxRow: 5,
    maxCol: 5,
    rowStyles: { 4: 7 },
    colStyles: [{ min: 2, max: 3, style: 8 }],
    merges: [merged, { r1: 4, c1: 0, r2: 4, c2: 1 }, { r1: 2, c1: 4, r2: 2, c2: 5 }],
    notes: [{ row: 0, col: 0, kind: 'note', entries: [{ text: 'hello' }] }],
    blocked: [{ r1: 5, c1: 0, r2: 5, c2: 5, kind: 'pivot', reason: msg('sheet.protected') }],
  });

  it("takes the cell's own format, else its row's, else its column's", () => {
    const view = new SheetView(info, grid([[4, 2, 1, 5]]));
    expect(view.styleAt(4, 2)).toBe(5);
    expect(view.styleAt(4, 3)).toBe(7);
    expect(view.styleAt(0, 3)).toBe(8);
    expect(view.styleAt(0, 4)).toBe(0);
    // A record passed in is used as it is.
    expect(view.styleAt(0, 0, { t: VT.String, s: 'x', style: 9, flags: 0 })).toBe(9);
  });

  it('finds the merge, note, table and blocked region at a cell', () => {
    const table = { name: 'T', rect: { r1: 0, c1: 4, r2: 3, c2: 5 } } as TableInfo;
    const view = new SheetView({ ...info, tables: [table] }, null);
    expect(view.mergeAt(2, 3)).toBe(merged);
    // Two merges can share a row.
    expect(view.mergeAt(2, 5)).toEqual({ r1: 2, c1: 4, r2: 2, c2: 5 });
    expect(view.mergeAt(2, 0)).toBeUndefined();
    expect(view.mergeAt(4, 1)).toEqual({ r1: 4, c1: 0, r2: 4, c2: 1 });
    expect(view.noteAt(0, 0)?.entries[0].text).toBe('hello');
    expect(view.noteAt(0, 1)).toBeUndefined();
    expect(view.tableAt(3, 5)).toBe(table);
    expect(view.tableAt(4, 5)).toBeUndefined();
    expect(view.blockedAt(5, 2)?.kind).toBe('pivot');
    expect(view.blockedAt(4, 2)).toBeUndefined();
  });

  it("keeps the latest conditional formatting results", () => {
    const view = new SheetView(info, null);
    view.setConditionalFormats({ sheet: 0, styles: [{ fill: '#FF0000' }], cells: [{ row: 1, col: 1, style: 0 }], skipped: 0 });
    expect(view.cfAt(1, 1)).toEqual({ row: 1, col: 1, style: 0 });
    view.setConditionalFormats({ sheet: 0, styles: [], cells: [{ row: 2, col: 2, icon: { set: '3Flags', index: 1 } }], skipped: 0 });
    expect(view.cfAt(1, 1)).toBeUndefined();
    expect(view.cfAt(2, 2)?.icon).toEqual({ set: '3Flags', index: 1 });
    expect(view.cfStyles).toEqual([]);
  });

  it('says why a cell cannot be edited', () => {
    const reason = (v: SheetView, row: number, col: number) => english.text(v.readOnlyReason(row, col));
    const view = new SheetView(info, null);
    expect(view.readOnlyReason(0, 0)).toBeUndefined();
    // The top-left cell of a merge is the one to edit.
    expect(view.readOnlyReason(1, 1)).toBeUndefined();
    expect(reason(view, 2, 3)).toMatch(/^Part of a merged cell/);
    expect(view.readOnlyReason(5, 0)).toEqual(msg('sheet.protected'));
    view.constraint = (row) => (row === 3 ? 'Schema says no' : undefined);
    expect(view.readOnlyReason(3, 0)).toBe('Schema says no');
    expect(view.readOnlyReason(0, 0)).toBeUndefined();
    expect(reason(new SheetView({ ...info, protected: true }, null), 0, 0)).toBe(english.text(msg('sheet.protected')));
    expect(new SheetView({ ...info, kind: 'chartsheet' }, null).readOnlyReason(0, 0)).toEqual(msg('sheet.notWorksheet'));
  });
});

describe('layering formats', () => {
  const thin = { width: 1, pattern: 'solid', color: '#000000' } as const;
  const thick = { width: 3, pattern: 'solid', color: '#FF0000' } as const;

  it("puts a table style under the cell's own format", () => {
    const own = style({ fill: '#FFFF00', bold: false, border: { left: thin } });
    const table: DiffStyle = { fill: '#DDEBF7', color: '#1F4E78', bold: true, italic: true, underline: true, strike: true, border: { left: thick, top: thick } };
    expect(underlay(own, table)).toMatchObject({
      fill: '#FFFF00',
      color: '#1F4E78',
      bold: true,
      italic: true,
      underline: true,
      strike: true,
      border: { left: thin, top: thick, right: undefined, bottom: undefined },
    });
    // Nothing set on either side stays unset.
    expect(underlay(style(), {})).toMatchObject({ fill: undefined, bold: false, border: undefined });
    expect(underlay(style(), { border: { bottom: thin } }).border).toEqual({ left: undefined, right: undefined, top: undefined, bottom: thin });
    expect(underlay(own, { fill: '#DDEBF7' }).border).toBe(own.border);
  });

  it("puts a conditional format over the cell's own format", () => {
    const own = style({ fill: '#FFFF00', color: '#000000', bold: true, italic: true, border: { left: thin, top: thin } });
    expect(overlay(own, { fill: '#FFC7CE', color: '#9C0006', bold: false, border: { left: thick } })).toMatchObject({
      fill: '#FFC7CE',
      color: '#9C0006',
      bold: false,
      italic: true,
      border: { left: thick, top: thin },
    });
    expect(overlay(own, {})).toMatchObject({ fill: '#FFFF00', bold: true, border: own.border });
  });

  it('makes each combination once, so the grid can compare styles by identity', () => {
    const own = style();
    const diff: DiffStyle = { fill: '#EEEEEE' };
    expect(underlay(own, diff)).toBe(underlay(own, diff));
    expect(overlay(own, diff)).toBe(overlay(own, diff));
    expect(overlay(own, diff)).not.toBe(underlay(own, diff));
    expect(overlay(style(), diff)).not.toBe(overlay(own, diff));
  });

  it("draws a cell with its table's style under it and conditional formatting over it", () => {
    const styles = [style(), style({ italic: true }), style({ fill: '#00FF00' })];
    const table: TableInfo = {
      name: 'Sales',
      rect: { r1: 0, c1: 0, r2: 3, c2: 1 },
      headerRows: 1,
      totalsRows: 0,
      columns: ['Region', 'Amount'],
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: false,
      showColumnStripes: false,
      look: { elements: { wholeTable: { fill: '#DDEBF7' }, headerRow: { bold: true, fill: '#4472C4', color: '#FFFFFF' } }, rowStripe: 1, colStripe: 1 },
    };
    const view = sheet(
      [
        [0, 0, 'Region', 1],
        [1, 0, 'North', 2],
        [1, 1, 5],
      ],
      { tables: [table] },
    );
    view.setConditionalFormats({ sheet: 0, styles: [{ color: '#9C0006' }], cells: [{ row: 1, col: 1, style: 0 }, { row: 2, col: 1, bar: { from: 0, to: 1, color: '#638EC6', gradient: true } }], skipped: 0 });
    expect(effectiveStyle(view, styles, 0, 0)).toMatchObject({ bold: true, italic: true, fill: '#4472C4', color: '#FFFFFF' });
    // The cell's own fill beats the table's.
    expect(effectiveStyle(view, styles, 1, 0)).toMatchObject({ fill: '#00FF00', bold: false });
    expect(effectiveStyle(view, styles, 1, 1)).toMatchObject({ fill: '#DDEBF7', color: '#9C0006' });
    // A data bar alone changes no format.
    expect(effectiveStyle(view, styles, 2, 1)).toMatchObject({ fill: '#DDEBF7', color: undefined });
    // Outside the table and without results: the cell's format, or the default for an unknown one.
    expect(effectiveStyle(view, styles, 5, 5)).toBe(styles[0]);
    expect(effectiveStyle(view, styles, 5, 5, 2)).toBe(styles[2]);
    expect(effectiveStyle(view, styles, 5, 5, 99)).toBe(styles[0]);
  });
});

describe('reading cells for a host', () => {
  const dates = [style(), style({ numFmt: 'yyyy-mm-dd' })];
  const data = sheet(
    [
      [0, 0, 'Item'],
      [0, 1, 'Due'],
      [1, 0, 'Bolts'],
      [1, 1, 46023, 1],
      [2, 0, 'Nuts'],
      [2, 2, true],
    ],
    { name: 'Data' },
  );
  const empty = new SheetView(sheetInfo({ name: 'Empty', index: 1 }), null);
  const read = workbookReaders([data, empty], dates, false);

  it('reads one cell by sheet name, ignoring case, with dates as ISO text', () => {
    expect(read.cell('data', 'B2')).toBe('2026-01-01');
    expect(read.cell('Data', '$A$3')).toBe('Nuts');
    expect(read.cell('DATA', 'Z99')).toBeNull();
  });

  it('reads a range row by row, cut to the used range', () => {
    expect(read.range('Data', 'A1:B2')).toEqual([
      ['Item', 'Due'],
      ['Bolts', '2026-01-01'],
    ]);
    expect(read.range('Data', 'B2:Z1000')).toEqual([
      ['2026-01-01', null],
      [null, true],
    ]);
    // A range wholly past the used range still gives its first row and column.
    expect(read.range('Data', 'E10:F20')).toEqual([[null]]);
  });

  it("reads a sheet's used range when no range is given, and nothing from an empty sheet", () => {
    expect(read.range('Data')).toEqual([
      ['Item', 'Due', null],
      ['Bolts', '2026-01-01', null],
      ['Nuts', null, true],
    ]);
    expect(read.range('Empty')).toEqual([]);
  });

  it("reads a cell whose format is unknown with the workbook's default one", () => {
    const odd = sheet([[0, 0, 46023, 42]], { name: 'Odd' });
    expect(workbookReaders([odd], dates, false).cell('Odd', 'A1')).toBe(46023);
  });

  it('reads dates in the 1904 date system', () => {
    expect(workbookReaders([data], dates, true).cell('Data', 'B2')).toBe('2030-01-02');
  });

  it('explains what is wrong with a sheet name or reference', () => {
    expect(() => read.cell('Nowhere', 'A1')).toThrow('There is no sheet called "Nowhere".');
    expect(() => read.range('Nowhere')).toThrow('There is no sheet called "Nowhere".');
    for (const bad of ['B', 'A1:B2', 'Data!A1']) expect(() => read.cell('Data', bad), bad).toThrow(`"${bad}" isn't a cell reference.`);
    expect(() => read.range('Data', 'A1:B2:C3')).toThrow(`"A1:B2:C3" isn't a range.`);
  });

  it('turns down references off the edge of the sheet rather than reading empty cells there', () => {
    for (const bad of ['A0', 'XFE1', 'A1048577']) expect(() => read.cell('Data', bad), bad).toThrow(`"${bad}" isn't a cell reference.`);
    // Before, A0:B2 came back with an extra row of nulls in front.
    for (const bad of ['A0:B2', 'A1:XFE2']) expect(() => read.range('Data', bad), bad).toThrow(`"${bad}" isn't a range.`);
    expect(read.cell('Data', 'XFD1048576')).toBeNull();
    expect(read.range('Data', 'A1:XFD1048576')).toHaveLength(3);
  });
});
