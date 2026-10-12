import ExcelJS from 'exceljs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { wrappedLines } from '../src/state/autofit';
import type { WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** Seven pixels a character, as the grid would measure a plain font. */
const seven = (s: string) => s.length * 7;

describe('wrapping text into lines', () => {
  it('breaks at spaces where the next word would pass the width, and at line breaks', () => {
    expect(wrappedLines('one two three', 7 * 7, seven)).toBe(2);
    expect(wrappedLines('one two three', 7 * 13, seven)).toBe(1);
    expect(wrappedLines('one\ntwo three four', 7 * 9, seven)).toBe(3);
    // A word longer than the width keeps its own line.
    expect(wrappedLines('extraordinarily', 7 * 3, seven)).toBe(1);
    expect(wrappedLines('', 50, seven)).toBe(1);
  });
});

describe('AutoFit', () => {
  let store: WorkbookStore;
  afterEach(() => store?.dispose());

  /** Sheet!A1:B3 short text, C1 a long one, D1:E1 merged over long text. */
  async function open(): Promise<WorkbookStore> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet');
    ws.addRow(['ab', 'abcd', 'a much longer piece of text', 'merged text running on and on']);
    ws.addRow(['abc', 'x']);
    ws.addRow(['a', 'xyz']);
    ws.mergeCells('D1:E1');
    store = await testStore();
    await opened(store, new Uint8Array(await wb.xlsx.writeBuffer()), 'fit.xlsx');
    store.measureText = seven;
    return store;
  }

  it("fits a column to its widest value, and the other selected columns to their own, as one step", async () => {
    await open();
    await store.autofit('cols', 2);
    // 27 characters, with the cell's padding either side.
    expect(store.view!.colWidth(2)).toBe(Math.ceil(27 * 7 + 4.5 * 2 + 2));
    store.selectLines('cols', [[0, 1]], { row: 0, col: 0 });
    await store.autofit('cols', 1);
    expect([store.view!.colWidth(0), store.view!.colWidth(1)]).toEqual([3 * 7 + 11, 4 * 7 + 11]);
    expect(store.history.map((h) => h.action)).toHaveLength(2);
    await store.undo();
    expect(store.view!.colWidth(0)).toBe(64);
  });

  it('leaves a column with nothing in it, or only merged cells, as it is', async () => {
    await open();
    await store.autofit('cols', 3);
    await store.autofit('cols', 7);
    expect([store.view!.colWidth(3), store.view!.colWidth(7)]).toEqual([64, 64]);
    expect(store.history).toHaveLength(0);
  });

  it('fits a row to the lines its wrapped text takes, and an empty one back to the sheet’s height', async () => {
    await open();
    store.setSelection({ row: 0, col: 2, rects: [{ r1: 0, c1: 2, r2: 0, c2: 2 }] });
    await store.formatCells({ wrap: true });
    // 27 characters at 7 pixels in a 64-pixel column: two words a line at most.
    const lines = wrappedLines('a much longer piece of text', 64 - 9, seven);
    await store.autofit('rows', 0);
    expect(store.view!.rowHeight(0)).toBe(lines * Math.round(store.styles[store.view!.get(0, 2)!.style].fontSizePx * 1.36));
    await store.autofit('rows', 9);
    expect(store.view!.rowHeight(9)).toBe(store.view!.info.defaultRowHeightPx);
  });
});
