import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { english } from '../src/i18n';
import type { WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** Data!A1:C6 filled with "r1c1"…; Locked is protected. */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  for (let r = 1; r <= 6; r++) data.addRow([1, 2, 3].map((c) => `r${r}c${c}`));
  const locked = wb.addWorksheet('Locked');
  locked.getCell('A1').value = 1;
  await locked.protect('', {});
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => store?.dispose());

async function open(): Promise<WorkbookStore> {
  store = await testStore();
  await opened(store, bytes, 'data.xlsx');
  return store;
}

const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const select = (r: ReturnType<typeof rect>) => store.setSelection({ row: r.r1, col: r.c1, rects: [r] });
const actions = () => store.history.map((h) => english.text(h.action));
const toasts = () => store.toasts.map((t) => english.text(t.text));
async function saved() {
  const s = await store.snapshot();
  return WorkbookSession.open(s!.bytes, 'data.xlsx').model.sheets[0].info;
}

describe('sizes', () => {
  it('set the width of every selected column, and undo as one step', async () => {
    await open();
    select(rect(0, 1, 0, 2));
    await store.setColumnWidth(1, 20);
    expect([store.view!.colWidth(1), store.view!.colWidth(2)]).toEqual([140, 140]);
    // A column outside the selection is set on its own.
    await store.setColumnWidth(4, 5);
    expect(store.view!.colWidth(4)).toBe(35);
    expect(actions()).toEqual(['Width of 2 columns', 'Column width']);
    expect((await saved()).colWidths).toMatchObject({ 1: 140, 2: 140, 4: 35 });
    await store.undo();
    await store.undo();
    expect(store.view!.colWidth(1)).toBe(store.view!.info.defaultColWidthPx);
    expect(store.editedCount).toBe(0);
    await store.redo();
    expect(store.view!.colWidth(2)).toBe(140);
    expect(store.editedCount).toBe(1);
  });

  it('keep a width dragged to in the grid, and set row heights in points', async () => {
    await open();
    select(rect(2, 0));
    await store.resizedColumn(0, 99.6);
    expect(store.view!.colWidth(0)).toBe(100);
    await store.setRowHeight(2, 30);
    expect(store.view!.rowHeight(2)).toBe(40);
    expect(store.rowHeightPt(2)).toBe(30);
    const info = await saved();
    expect([info.colWidths[0], info.rowHeights[2]]).toEqual([100, 40]);
  });
});

describe('hiding', () => {
  it('hides the selected rows and shows them again from the rows around them', async () => {
    await open();
    select(rect(1, 0, 2, 0));
    await store.hideLines('rows');
    expect([store.view!.gridRow(1), store.view!.gridRow(2), store.view!.gridRow(3)]).toEqual([-1, -1, 1]);
    // Selecting the row below finds the hidden ones just above it.
    select(rect(3, 0));
    expect(store.hasHiddenNear('rows')).toBe(true);
    await store.unhideLines('rows');
    expect(store.view!.gridRow(2)).toBe(2);
    expect(actions()).toEqual(['Hide 2 rows', 'Unhide 2 rows']);
    select(rect(0, 1, 5, 2));
    await store.hideLines('cols');
    expect((await saved()).hiddenCols).toEqual([1, 2]);
    expect(store.hasHiddenNear('rows')).toBe(false);
  });

  it('shows hidden columns at the start of the sheet from the first one left', async () => {
    await open();
    select(rect(0, 0, 0, 1));
    await store.hideLines('cols');
    select(rect(0, 2));
    await store.unhideLines('cols');
    expect(store.view!.info.hiddenCols).toEqual([]);
  });
});

describe('frozen panes', () => {
  it('freeze at the active cell, the top row or the first column, and unfreeze', async () => {
    await open();
    select(rect(2, 1));
    await store.freezePanes('cell');
    expect([store.view!.info.frozenRows, store.view!.info.frozenCols]).toEqual([2, 1]);
    await store.freezePanes('topRow');
    expect([store.view!.info.frozenRows, store.view!.info.frozenCols]).toEqual([1, 0]);
    const info = await saved();
    expect([info.frozenRows, info.frozenCols]).toEqual([1, 0]);
    await store.freezePanes('none');
    expect(store.view!.info.frozenRows).toBe(0);
    expect(actions()).toEqual(['Freeze panes', 'Freeze panes', 'Unfreeze panes']);
  });
});

describe('merged cells', () => {
  it('asks before clearing the other cells, then merges, and undo brings them back', async () => {
    await open();
    select(rect(0, 0, 1, 1));
    const merging = store.mergeCells();
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    expect(english.text(store.confirmRequest!.body)).toBe('Merged cells keep only the top-left value. 3 other cells will be cleared.');
    store.confirmRequest!.resolve('confirm');
    await merging;
    expect(store.view!.info.merges).toEqual([rect(0, 0, 1, 1)]);
    expect([store.view!.get(0, 0)?.s, store.view!.get(1, 1)?.s]).toEqual(['r1c1', undefined]);
    expect(store.view!.readOnlyReason(1, 1)).toBeTruthy();
    await store.undo();
    expect(store.view!.info.merges).toEqual([]);
    expect(store.view!.get(1, 1)?.s).toBe('r2c2');
  });

  it('take in merges the selection covers part of, and unmerge them', async () => {
    await open();
    select(rect(4, 0, 4, 1));
    const first = store.mergeCells();
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    store.confirmRequest!.resolve('confirm');
    await first;
    select(rect(4, 1, 5, 2));
    const second = store.mergeCells();
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    store.confirmRequest!.resolve('confirm');
    await second;
    expect(store.view!.info.merges).toEqual([rect(4, 0, 5, 2)]);
    select(rect(5, 2));
    expect(store.mergesInSelection()).toHaveLength(1);
    await store.unmergeCells();
    expect(store.view!.info.merges).toEqual([]);
    expect((await saved()).merges).toEqual([]);
  });

  it('are refused, with the reason, on a protected sheet', async () => {
    await open();
    store.setActive(1);
    select(rect(0, 0, 1, 1));
    await store.mergeCells();
    expect(toasts()).toEqual(['This sheet is protected in Excel, so it is read-only here.']);
    expect(store.undoStack).toEqual([]);
  });
});

describe('formats', () => {
  it('toggle as the active cell has them, step decimals, and clear', async () => {
    await open();
    select(rect(0, 0, 1, 1));
    expect(store.activeStyle.bold).toBe(false);
    await store.toggleFormat('bold');
    expect(store.activeStyle.bold).toBe(true);
    expect(store.styles[store.view!.styleAt(1, 1)].bold).toBe(true);
    await store.toggleFormat('bold');
    expect(store.activeStyle.bold).toBe(false);
    await store.stepDecimals(1);
    expect(store.activeStyle.numFmt).toBe('0.0');
    await store.formatCells({ fill: '#FF0000' }, { borders: { where: 'outside' } });
    expect(store.activeStyle).toMatchObject({ fill: '#FF0000', border: { left: {}, top: {} } });
    await store.formatCells({}, { clear: true });
    expect(store.activeStyle).toMatchObject({ bold: false, numFmt: 'General' });
    expect(store.activeStyle.fill).toBeUndefined();
    expect(actions()).toEqual(['Format cells', 'Format cells', 'Format cells', 'Format cells', 'Clear formats']);
    await store.undo();
    expect(store.activeStyle.fill).toBe('#FF0000');
  });
});
