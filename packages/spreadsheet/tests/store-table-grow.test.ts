import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import type { WorkbookStore } from '../src/state/store';
import { parts } from './helpers';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/**
 * Orders!A1:C3 is table "Lines" (Item, Qty, Double) without a totals row,
 * Double being the calculated column [@Qty]*2; E1 = SUM(Lines[Qty]).
 * Totals!A1:B4 is table "Summed", which has a totals row.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const orders = wb.addWorksheet('Orders');
  orders.addTable({
    name: 'Lines',
    ref: 'A1',
    headerRow: true,
    columns: [{ name: 'Item' }, { name: 'Qty' }, { name: 'Double' }],
    rows: [
      ['Bolts', 2, { formula: 'Lines[[#This Row],[Qty]]*2', result: 4 }],
      ['Nuts', 5, { formula: 'Lines[[#This Row],[Qty]]*2', result: 10 }],
    ],
  });
  orders.getCell('E1').value = { formula: 'SUM(Lines[Qty])', result: 7 };
  wb.addWorksheet('Totals').addTable({
    name: 'Summed',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    columns: [{ name: 'Item', totalsRowLabel: 'Total' }, { name: 'Qty', totalsRowFunction: 'sum' }],
    rows: [
      ['a', 1],
      ['b', 2],
    ],
  });
  const files = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
  const path = Object.keys(files).find((k) => k.startsWith('xl/tables/') && strFromU8(files[k]).includes('name="Lines"'))!;
  // ExcelJS doesn't write calculated columns.
  files[path] = strToU8(
    strFromU8(files[path]).replace(/(<tableColumn [^>]*name="Double"[^>]*?)\s*\/>/, '$1><calculatedColumnFormula>Lines[[#This Row],[Qty]]*2</calculatedColumnFormula></tableColumn>'),
  );
  return zipSync(files);
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => store?.dispose());

async function open(): Promise<WorkbookStore> {
  store = await testStore();
  await opened(store, bytes, 'orders.xlsx');
  return store;
}

const cell = (row: number, col: number) => store.sheets[store.active].get(row, col);
const table = () => store.view!.info.tables[0];
const actions = () => store.history.map((h) => english.text(h.action));

describe('an Excel table without a totals row', () => {
  it('takes in a row typed just below it, with its calculated column, and its references count it', async () => {
    await open();
    await store.input(3, 0, 'Washers');
    expect(table().rect).toEqual({ r1: 0, c1: 0, r2: 3, c2: 2 });
    expect(cell(3, 0)?.s).toBe('Washers');
    expect(cell(3, 2)?.f).toBe('Lines[[#This Row],[Qty]]*2');
    await store.input(3, 1, '4');
    // The table grew once, not again for a cell in a row it already has.
    expect(table().rect.r2).toBe(3);
    expect(cell(3, 2)?.n).toBe(8);
    expect(cell(0, 4)?.n).toBe(11);
    expect(actions()).toEqual(['Added 1 row to Lines', 'Edit', 'Edit']);

    const saved = await store.snapshot();
    const xml = Object.entries(parts(saved!.bytes)).find(([k, v]) => k.startsWith('xl/tables/') && v.includes('name="Lines"'))![1];
    expect(xml).toMatch(/ref="A1:C4"/);
  });

  it('takes in as many rows as are pasted below it', async () => {
    await open();
    await store.paste({ row: 3, col: 0 }, [['Pins', '1'], ['Clips', '3']], { r1: 3, c1: 0, r2: 3, c2: 0 });
    expect(table().rect.r2).toBe(4);
    expect(cell(4, 2)?.f).toBe('Lines[[#This Row],[Qty]]*2');
  });

  it('is left as it is by typing beside or further below it, or by clearing a cell', async () => {
    await open();
    await store.input(3, 3, 'Note');
    await store.input(4, 0, 'Gap');
    await store.input(3, 0, '');
    expect(table().rect.r2).toBe(2);
  });

  it('shrinks back when the typing is undone, one step at a time', async () => {
    await open();
    await store.input(3, 0, 'Washers');
    await store.undo();
    expect(cell(3, 0)?.s).toBeUndefined();
    expect(table().rect.r2).toBe(3);
    await store.undo();
    expect(table().rect.r2).toBe(2);
    expect(cell(3, 2)).toBeFalsy();
  });
});

describe('an Excel table with a totals row', () => {
  it('stays the same size when typing below it', async () => {
    await open();
    store.setActive(1);
    await store.input(4, 0, 'c');
    expect(table().rect).toEqual({ r1: 0, c1: 0, r2: 3, c2: 1 });
    expect(cell(4, 0)?.s).toBe('c');
  });
});
