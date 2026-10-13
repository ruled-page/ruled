import ExcelJS from 'exceljs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { Axis, colAxis, rowAxis } from '../src/grid/ruled/axis';
import { Painter, type PaintInput } from '../src/grid/ruled/draw';
import { contentScroll, Layout, SCROLL_LIMIT, scrollerPosition, scrollLength, type LayoutSpec } from '../src/grid/ruled/layout';
import type { WorkbookStore } from '../src/state/store';
import { DEFAULT_PALETTE } from '../src/ui/gridPalette';
import { Recorder } from './canvasRecorder';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

describe("a grid's rows or columns", () => {
  it('starts each line after those before it, the few of their own size included', () => {
    const a = new Axis(10, 20, [
      [3, 50],
      [1, 0],
      [7, 20],
      [12, 99],
    ]);
    expect([0, 1, 2, 3, 4].map((i) => a.size(i))).toEqual([20, 0, 20, 50, 20]);
    expect([0, 1, 2, 3, 4, 10].map((i) => a.start(i))).toEqual([0, 20, 20, 40, 90, 210]);
    expect(a.total).toBe(210);
  });

  it('finds the line at a distance, the first or last outside them all', () => {
    const a = new Axis(10, 20, [[3, 50]]);
    expect([-5, 0, 19, 20, 59, 60, 109, 110, 5000].map((px) => a.at(px))).toEqual([0, 0, 0, 1, 2, 3, 3, 4, 9]);
    expect(new Axis(0, 20).at(10)).toBe(0);
  });

  it('reaches row 1,048,576 by the default size, with the changes before it', () => {
    const a = new Axis(1_048_576, 20, [[500_000, 40]]);
    expect(a.total).toBe(1_048_576 * 20 + 20);
    expect(a.at(a.start(1_000_000) + 5)).toBe(1_000_000);
  });
});

describe('where the grid puts things', () => {
  const spec = (over: Partial<LayoutSpec> = {}): LayoutSpec => ({
    width: 400,
    height: 300,
    rows: new Axis(1000, 20),
    cols: new Axis(100, 64),
    frozenRows: 0,
    frozenCols: 0,
    numbers: 40,
    letters: 24,
    scrollX: 0,
    scrollY: 0,
    ...over,
  });

  it('puts the cells after the headings, scrolled, and knows which are in view', () => {
    const l = new Layout(spec({ scrollX: 70, scrollY: 25 }));
    expect(l.cell(1, 1)).toEqual({ x: 40 + 64 - 70, y: 24 + 20 - 25, width: 64, height: 20 });
    expect(l.rowRun).toEqual({ first: 1, last: 15 });
    expect(l.colRun).toEqual({ first: 1, last: 6 });
    expect([l.contentWidth, l.contentHeight]).toEqual([40 + 6400, 24 + 20000]);
    expect(l.corner()).toEqual({ x: 0, y: 0, width: 40, height: 24 });
    expect(l.rowHeading(1)).toEqual({ x: 0, y: 19, width: 40, height: 20 });
    expect(l.colHeading(1)).toEqual({ x: 34, y: 0, width: 64, height: 24 });
    expect(l.quadrants()).toEqual([{ rows: l.rowRun, cols: l.colRun, clip: { x: 40, y: 24, width: 360, height: 276 } }]);
  });

  it('keeps frozen rows and columns in place, in their own parts, while the rest scrolls under them', () => {
    const l = new Layout(spec({ frozenRows: 2, frozenCols: 1, scrollX: 64, scrollY: 40 }));
    expect([l.bodyX, l.bodyY]).toEqual([104, 64]);
    expect(l.cell(0, 0)).toEqual({ x: 40, y: 24, width: 64, height: 20 });
    // The first rows and columns past the frozen ones have scrolled away.
    expect(l.cell(4, 2)).toEqual({ x: 104, y: 64, width: 64, height: 20 });
    expect([l.frozenRowRun, l.rowRun, l.frozenColRun, l.colRun]).toEqual([
      { first: 0, last: 1 },
      { first: 4, last: 15 },
      { first: 0, last: 0 },
      { first: 2, last: 6 },
    ]);
    expect(l.quadrants().map((q) => q.clip)).toEqual([
      { x: 40, y: 24, width: 64, height: 40 },
      { x: 104, y: 24, width: 296, height: 40 },
      { x: 40, y: 64, width: 64, height: 236 },
      { x: 104, y: 64, width: 296, height: 236 },
    ]);
  });

  it('says what is at a point, and the nearest cell in view for a point outside them', () => {
    const l = new Layout(spec({ frozenRows: 2, frozenCols: 1, scrollX: 64, scrollY: 40 }));
    expect(l.hit(10, 10)).toEqual({ area: 'corner', row: -1, col: -1 });
    expect(l.hit(50, 10)).toEqual({ area: 'letters', row: -1, col: 0 });
    expect(l.hit(110, 10)).toEqual({ area: 'letters', row: -1, col: 2 });
    expect(l.hit(10, 50)).toEqual({ area: 'numbers', row: 1, col: -1 });
    expect(l.hit(110, 70)).toEqual({ area: 'cells', row: 4, col: 2 });
    expect(l.hit(-50, 999, true)).toEqual({ area: 'cells', row: 15, col: 0 });
    expect(l.hit(999, -50, true)).toEqual({ area: 'cells', row: 0, col: 6 });
  });

  it('mirrors everything for a sheet that reads right to left, column A at the right', () => {
    const l = new Layout(spec({ rtl: true }));
    expect(l.corner()).toEqual({ x: 360, y: 0, width: 40, height: 24 });
    expect(l.cell(0, 0)).toEqual({ x: 296, y: 24, width: 64, height: 20 });
    expect(l.hit(300, 30)).toEqual({ area: 'cells', row: 0, col: 0 });
    expect(l.hit(390, 30)).toEqual({ area: 'numbers', row: 0, col: -1 });
  });

  it('has no lines to scroll when the frozen ones are all there are, or there is no room', () => {
    const l = new Layout(spec({ rows: new Axis(2, 20), frozenRows: 2 }));
    expect(l.rowRun.first > l.rowRun.last).toBe(true);
    expect(l.hit(100, 200, true).row).toBe(1);
    expect(new Layout(spec({ height: 20 })).rowRun.first > new Layout(spec({ height: 20 })).rowRun.last).toBe(true);
    expect(new Layout(spec({ height: 20 })).quadrants()).toEqual([]);
  });

  it('scales the scroll position for content longer than browsers let an element be', () => {
    expect(scrollLength(1000)).toBe(1000);
    expect(scrollLength(30_000_000)).toBe(SCROLL_LIMIT);
    expect(contentScroll(100, 1000, 300)).toBe(100);
    expect(scrollerPosition(100, 1000, 300)).toBe(100);
    const long = 21_000_000;
    const end = SCROLL_LIMIT - 500;
    expect(contentScroll(end, long, 500)).toBeCloseTo(long - 500);
    expect(scrollerPosition(long - 500, long, 500)).toBeCloseTo(end);
    expect(scrollerPosition(contentScroll(1234, long, 500), long, 500)).toBeCloseTo(1234);
  });
});

describe("drawing a sheet in Ruled's grid", () => {
  let store: WorkbookStore;
  afterEach(() => store?.dispose());

  async function open(): Promise<WorkbookStore> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet');
    ws.getColumn(1).width = 12;
    ws.getColumn(2).width = 5;
    ws.getColumn(4).hidden = true;
    ws.getCell('A1').value = 'Title';
    ws.getCell('A1').font = { bold: true };
    ws.getCell('A1').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC000' } };
    ws.getCell('A1').border = { top: { style: 'thin', color: { argb: 'FF0000FF' } } };
    ws.mergeCells('A1:C1');
    ws.getCell('A2').value = 'A long piece of text that runs on into the empty cells beside it';
    ws.getRow(2).height = 30;
    ws.getCell('B3').value = 12345.678;
    ws.getCell('A4').value = 'one two three four five six seven eight nine ten';
    ws.getCell('A4').alignment = { wrapText: true, vertical: 'top' };
    ws.getCell('A5').value = 'struck';
    ws.getCell('A5').font = { strike: true, underline: true, italic: true };
    ws.getCell('A6').value = 'right';
    ws.getCell('A6').alignment = { horizontal: 'right', indent: 1 };
    ws.getCell('B6').border = { left: { style: 'medium', color: { argb: 'FFFF0000' } } };
    ws.getCell('A7').border = { right: { style: 'thin', color: { argb: 'FF00FF00' } } };
    ws.getCell('A8').border = { bottom: { style: 'thin', color: { argb: 'FF00FFFF' } } };
    ws.getCell('C9').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00B050' } };
    ws.getCell('A11').value = 'centred overflow';
    ws.getCell('A11').alignment = { horizontal: 'center' };
    ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];
    store = await testStore();
    await opened(store, new Uint8Array(await wb.xlsx.writeBuffer()), 'draw.xlsx');
    return store;
  }

  function input(s: WorkbookStore, over: Partial<LayoutSpec> = {}, extra: Partial<PaintInput> = {}): PaintInput {
    const view = s.view!;
    const layout = new Layout({
      width: 600,
      height: 300,
      rows: rowAxis(view),
      cols: colAxis(view),
      frozenRows: 1,
      frozenCols: 1,
      numbers: 40,
      letters: 24,
      scrollX: 0,
      scrollY: 0,
      ...over,
    });
    return { view, layout, styles: s.styles, palette: DEFAULT_PALETTE, date1904: false, locale: 'en-US', fullNumbers: false, highlightEdits: true, ...extra };
  }

  it('sizes rows and columns as the sheet does, hidden ones left out', async () => {
    const s = await open();
    const view = s.view!;
    const rows = rowAxis(view);
    const cols = colAxis(view);
    expect(rows.size(1)).toBe(40);
    expect(rows.size(2)).toBe(view.info.defaultRowHeightPx);
    expect(cols.size(1)).toBe(view.colWidth(1));
    // D is hidden: the fourth grid column is E.
    expect(view.visibleCols[3]).toBe(4);
    expect(cols.count).toBe(view.colCount);
  });

  it('draws the cells, their text, fills, borders and grid lines, the headings and the freeze lines', async () => {
    const s = await open();
    const r = new Recorder();
    const p = new Painter(r.ctx);
    // Measured before the palette is known, text falls back on a plain font.
    expect(p.measure('abc', s.styles[0])).toBe(3 * s.styles[0].fontSizePx * 0.5);
    p.draw(input(s));
    const texts = r.texts().map((t) => t[0]);
    for (const t of ['Title', 'A long piece of text that runs on into the empty cells beside it', 'struck', 'right', 'centred overflow', 'A', 'B', 'C', 'E', '1', '2']) expect(texts).toContain(t);
    // D is hidden.
    expect(texts).not.toContain('D');
    // The number is too wide for B, so it is shortened as Excel shortens it.
    const number = texts.find((t) => /^[\d.E+#]+$/.test(t) && t !== '1' && t !== '2' && !/^\d+$/.test(t));
    expect(number).toBeDefined();
    expect(number).not.toBe('12345.678');
    // Wrapped text is drawn line by line, and only the lines that reach into the row.
    const lines = r.texts().filter((t) => /^(one|two|three|four|five|six|seven|eight|nine|ten)/.test(t[0]));
    expect(lines.length).toBeGreaterThan(0);
    expect(lines.length).toBeLessThan(10);
    // The merge's fill, the fill of C9, and borders in their colours.
    const fills = r.rects().map((x) => String(x[4]).toLowerCase());
    for (const colour of ['#ffc000', '#00b050', '#ff0000', '#00ff00', '#00ffff', '#0000ff', DEFAULT_PALETTE.freezeLine, DEFAULT_PALETTE.line, DEFAULT_PALETTE.headerBg]) expect(fills).toContain(colour.toLowerCase());
    // Strikethrough and underline under struck text.
    expect(r.calls.filter((c) => c.op === 'fillRect' && c.args[3] === 1).length).toBeGreaterThan(2);
  });

  it('runs text into empty neighbours once in each part of the grid, and draws no grid lines under it', async () => {
    const s = await open();
    const r = new Recorder();
    new Painter(r.ctx).draw(input(s, { frozenRows: 0, frozenCols: 0 }));
    expect(r.texts().filter((t) => t[0].startsWith('A long piece'))).toHaveLength(1);
    // Across the freeze line, each side draws its part.
    const frozen = new Recorder();
    new Painter(frozen.ctx).draw(input(s));
    expect(frozen.texts().filter((t) => t[0].startsWith('A long piece'))).toHaveLength(2);
    // Grid lines down the left of B2 and C2 are left out under the text; B3's is drawn.
    const view = s.view!;
    const layout = input(s, { frozenRows: 0, frozenCols: 0 }).layout;
    const lineAt = (gr: number, gc: number) => r.rects().some(([x, y, w, h, c]) => x === layout.x(gc) && y === layout.y(gr) && w === 1 && h === layout.spec.rows.size(gr) && c === DEFAULT_PALETTE.line);
    expect(view.gridRow(1)).toBe(1);
    expect(lineAt(1, 1)).toBe(false);
    expect(lineAt(2, 1)).toBe(true);
  });

  it('fills changed cells and schema labels, and draws again from scratch after reset or a new palette', async () => {
    const s = await open();
    await s.input(9, 4, 'changed');
    const r = new Recorder();
    const p = new Painter(r.ctx);
    p.draw(input(s, {}, { isLabel: (row, col) => row === 9 && col === 5 }));
    const fills = r.rects().map((x) => x[4]);
    expect(fills).toContain(DEFAULT_PALETTE.editedFill);
    expect(fills).toContain(DEFAULT_PALETTE.labelFill);
    p.reset();
    const palette = { ...DEFAULT_PALETTE, bg: '#101010' };
    p.draw(input(s, {}, { palette, highlightEdits: false }));
    expect(r.rects().some((x) => x[4] === '#101010')).toBe(true);
  });

  it('draws without grid lines on a sheet that hides them, and mirrored for one read right to left', async () => {
    const s = await open();
    s.view!.info.showGridLines = false;
    const r = new Recorder();
    new Painter(r.ctx).draw(input(s, { rtl: true, scrollY: 20, scrollX: 10 }));
    // Text in the frozen column A starts from the right.
    const right = r.texts().find((t) => t[0] === 'struck');
    expect(right?.[3]).toContain('right');
    expect(r.rects().some((x) => x[4] === DEFAULT_PALETTE.line && x[3] === 20 && x[2] === 1 && (x[1] as number) > 24)).toBe(false);
  });

  it('tints the selection but its active cell, outlines it, marks the headings it covers and draws the marks', async () => {
    const s = await open();
    const r = new Recorder();
    const p = new Painter(r.ctx);
    const selection = { row: 2, col: 1, rects: [{ r1: 2, r2: 3, c1: 1, c2: 2 }] };
    const marks = [
      { rect: { r1: 6, r2: 6, c1: 1, c2: 1 }, color: 'rgba(1, 2, 3, 0.5)', style: 'dashed' as const },
      { rect: { r1: 0, r2: 0, c1: 3, c2: 3 }, color: '#000', style: 'solid' as const },
    ];
    p.draw(input(s, { frozenRows: 0, frozenCols: 0 }, { selection, marks, fillHandle: true }));
    // The tint leaves out the active cell, filling between the two outlines.
    expect(r.calls.some((c) => c.op === 'fill' && c.args[0] === 'evenodd' && c.style === DEFAULT_PALETTE.selectionFill)).toBe(true);
    // Four sides of the outline, then the fill handle at its corner, edged in the selection's text colour.
    expect(r.rects().filter((x) => x[4] === DEFAULT_PALETTE.selection)).toHaveLength(5);
    expect(r.rects().filter((x) => x[4] === DEFAULT_PALETTE.onSelection && x[2] === 8)).toHaveLength(1);
    // B and C, and rows 3 and 4, in the headings' selected colour.
    expect(r.rects().filter((x) => x[4] === DEFAULT_PALETTE.headerSelectedBg)).toHaveLength(4);
    // The dashed mark is outlined solid in its colour; D is hidden, so the second mark isn't drawn.
    expect(r.calls.filter((c) => c.op === 'strokeRect').map((c) => c.style)).toEqual(['rgba(1, 2, 3, 1)']);
    expect(r.args('setLineDash')).toContainEqual([[4, 3]]);

    // Whole columns, then the whole sheet: their headings in the selection's colour, their text in the selected text colour.
    const lines = new Recorder();
    new Painter(lines.ctx).draw(input(s, {}, { selection: { ...selection, lines: 'cols' } }));
    expect(lines.rects().filter((x) => x[4] === DEFAULT_PALETTE.selection).length).toBeGreaterThan(4);
    expect(lines.texts().find((t) => t[0] === 'B')?.[3]).toContain(DEFAULT_PALETTE.headerSelectedText);
    for (const whole of ['rows', 'all'] as const) {
      const w = new Recorder();
      new Painter(w.ctx).draw(input(s, {}, { selection: { ...selection, lines: whole } }));
      expect(w.texts().find((t) => t[0] === '3')?.[3]).toContain(DEFAULT_PALETTE.headerSelectedText);
    }
  });

  it("gives a cell's text as drawn, a number shown as #### in full and a merge's covered cells empty", async () => {
    const s = await open();
    const view = s.view!;
    const p = new Painter(new Recorder().ctx);
    const i = input(s);
    expect(p.cellText(i, 0, 0)).toBe('Title');
    expect(p.cellText(i, 0, 1)).toBe('');
    // A date too wide for its column is drawn as ####, but said in full.
    await s.input(2, 1, '2026-12-31');
    p.reset();
    view.colWidths.set(1, 20);
    expect(p.cellText(input(s), 2, 1)).toMatch(/2026/);
  });
});
