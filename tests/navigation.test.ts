import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { english } from '../src/i18n';
import type { WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const lines = wb.addWorksheet('Lines');
  lines.addRow(['Item', 'Qty', 'Price', 'Total']);
  lines.addRow(['Bolts', 2, 3, { formula: 'B2*C2', result: 6 }]);
  lines.addRow(['Nuts', 4, 0.5, { formula: 'B3*C3', result: 2 }]);
  lines.addRow(['Washers', 1, 1, 1]);
  wb.addWorksheet("Bob's sheet").addRow(['x']);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

async function open(): Promise<WorkbookStore> {
  const store = await testStore();
  await opened(store, bytes, 'lines.xlsx');
  return store;
}

const text = (store: WorkbookStore, row: number, col: number) => store.view?.get(row, col);

describe('going to a reference from the name box', () => {
  it('selects a cell or range on this sheet', async () => {
    const store = await open();
    expect(store.goTo('c3')).toBe(true);
    expect(store.selection).toMatchObject({ row: 2, col: 2, rects: [{ r1: 2, c1: 2, r2: 2, c2: 2 }] });
    expect(store.goTo(' $B$2:D3 ')).toBe(true);
    expect(store.selection.rects[0]).toEqual({ r1: 1, c1: 1, r2: 2, c2: 3 });
    expect(store.revealFocus).toBe(true);
    store.dispose();
  });

  it('switches to the sheet named, quoted or not', async () => {
    const store = await open();
    expect(store.goTo("'Bob''s sheet'!A1")).toBe(true);
    expect(store.active).toBe(1);
    expect(store.goTo('lines!D4')).toBe(true);
    expect(store.active).toBe(0);
    expect(store.selection.row).toBe(3);
    store.dispose();
  });

  it('turns down anything that is not a reference to a sheet here', async () => {
    const store = await open();
    for (const bad of ['', 'Nowhere!A1', 'hello', 'A0', 'XFE1', 'A1048577', 'A1:B2:C3']) expect(store.goTo(bad), bad).toBe(false);
    expect(store.selection.row).toBe(0);
    store.dispose();
  });
});

describe('filling from the keyboard', () => {
  it('copies the top row down through the selection, adjusting formulas', async () => {
    const store = await open();
    store.goTo('C2:D4');
    await store.fillDown();
    expect(text(store, 3, 2)?.n).toBe(3);
    expect(text(store, 3, 3)?.f).toBe('B4*C4');
    expect(text(store, 3, 3)?.n).toBe(3);
    expect(english.text(store.history.at(-1)?.action)).toBe('Fill down');
    store.dispose();
  });

  it('copies the cell above, or to the left, into a single selected row or column', async () => {
    const store = await open();
    store.goTo('B4');
    await store.fillDown();
    expect(text(store, 3, 1)?.n).toBe(4);
    store.goTo('C4');
    await store.fillDown(true);
    expect(text(store, 3, 2)?.n).toBe(4);
    expect(english.text(store.history.at(-1)?.action)).toBe('Fill right');
    store.dispose();
  });

  it('does nothing in the first row', async () => {
    const store = await open();
    store.goTo('B1');
    await store.fillDown();
    expect(store.history).toHaveLength(0);
    store.dispose();
  });
});

describe('column widths from the keyboard', () => {
  it('reads and sets widths in Excel character units', async () => {
    const store = await open();
    const before = store.colVersion;
    await store.setColumnWidth(1, 20);
    expect(store.view?.colWidth(1)).toBe(140);
    expect(store.columnWidth(1)).toBe(20);
    expect(store.colVersion).toBe(before + 1);
    await store.setColumnWidth(1, Number.NaN);
    expect(store.columnWidth(1)).toBe(20);
    store.dispose();
  });
});
