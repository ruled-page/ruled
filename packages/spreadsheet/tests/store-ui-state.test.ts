import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import { WorkbookStore, type StoreOptions } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const lines = wb.addWorksheet('Lines');
  lines.addRow(['Item', 'Qty']);
  lines.addRow(['Bolts', 2]);
  lines.getColumn(1).width = 20;
  wb.addWorksheet('Other').getCell('C3').value = 'x';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => {
  store?.dispose();
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

async function open(options: StoreOptions = {}): Promise<WorkbookStore> {
  store = await testStore(options);
  await opened(store, bytes, 'lines.xlsx');
  return store;
}

describe('toasts', () => {
  it('keeps the latest four, newest last', () => {
    store = new WorkbookStore({ drafts: false });
    for (let i = 1; i <= 6; i++) store.toast('info', `Message ${i}`);
    expect(store.toasts.map((t) => t.text)).toEqual(['Message 3', 'Message 4', 'Message 5', 'Message 6']);
  });

  it('can be dismissed; others go by themselves after five seconds, but errors stay until dismissed', () => {
    vi.useFakeTimers();
    store = new WorkbookStore({ drafts: false });
    store.toast('info', 'Saved');
    store.toast('error', 'Failed');
    store.toast('warn', 'Careful');
    store.dismissToast(store.toasts[2].id);
    expect(store.toasts.map((t) => t.text)).toEqual(['Saved', 'Failed']);
    vi.advanceTimersByTime(5000);
    expect(store.toasts.map((t) => t.text)).toEqual(['Failed']);
    vi.advanceTimersByTime(60_000);
    expect(store.toasts.map((t) => t.text)).toEqual(['Failed']);
    store.dismissToast(store.toasts[0].id);
    expect(store.toasts).toEqual([]);
  });
});

describe('editor state', () => {
  it('sets each piece of UI state and tells subscribers', () => {
    store = new WorkbookStore({ drafts: false });
    const seen: number[] = [];
    const unsubscribe = store.subscribe(() => seen.push(store.getVersion()));
    store.set('panel', 'report');
    expect(store.reportOpen).toBe(true);
    store.set('findOpen', true);
    store.set('menu', { x: 1, y: 2, kind: 'row' });
    expect([store.findOpen, store.menu]).toEqual([true, { x: 1, y: 2, kind: 'row' }]);
    expect(seen).toEqual([1, 2, 3]);
    unsubscribe();
    store.set('panel', null);
    expect(seen).toHaveLength(3);
    expect(store.reportOpen).toBe(false);
  });

  it('turns on the features a host does not hide', () => {
    store = new WorkbookStore({ drafts: false, features: { open: false } });
    expect(store.features).toEqual({ open: false, save: true, download: true, schemas: true, report: true, charts: true, format: true, csv: true, autoComplete: true });
    expect(new WorkbookStore({ features: { charts: false } }).features.charts).toBe(false);
    expect(store.warnOnUnload).toBe(true);
    expect(new WorkbookStore({ warnOnUnload: false }).warnOnUnload).toBe(false);
  });

  it('follows the browser’s locale and language unless told otherwise', () => {
    vi.stubGlobal('navigator', { language: 'de-DE' });
    store = new WorkbookStore({ drafts: false });
    expect(store.locale).toBe('de-DE');
    expect(store.t.language).toBe('de-DE');
    expect(store.saveLabel).toBe('Als Excel speichern');
    const fixed = new WorkbookStore({ locale: 'en-GB', language: 'fr', translations: { fr: { 'toolbar.saveAsExcel': 'Enregistrer' } } });
    expect(fixed.locale).toBe('en-GB');
    expect(fixed.saveLabel).toBe('Enregistrer');
  });

  it('changes language while open', () => {
    store = new WorkbookStore({ drafts: false, language: 'en' });
    const version = store.version;
    store.setLanguage('de');
    expect(store.saveLabel).toBe('Als Excel speichern');
    expect(store.version).toBe(version + 1);
  });

  it('changes the theme while open, so the grid reads its colours again', () => {
    const icons = { open: 'M0 0h24v24H0z' };
    store = new WorkbookStore({ drafts: false, theme: { accent: '#f00' }, icons });
    expect(store.theme).toEqual({ accent: '#f00' });
    expect(store.icons).toBe(icons);
    store.setTheme({ accent: '#0f0' });
    expect(store.theme).toEqual({ accent: '#0f0' });
    store.setTheme(undefined);
    expect(store.theme).toBeUndefined();
    expect(store.themeVersion).toBe(2);
  });

  it('shows numbers as the host, then the schema, then the user’s toggle say', () => {
    store = new WorkbookStore({ drafts: false, numbers: { paste: 'strict' } });
    expect(store.numbers).toEqual({ display: 'formatted', copy: 'formatted', paste: 'strict', reviewBeforeSave: false });
    store.set('exactNumbers', true);
    expect(store.numbers).toMatchObject({ display: 'full', copy: 'raw', paste: 'strict' });
    store.set('exactNumbers', false);
    expect(store.numbers).toMatchObject({ display: 'formatted', copy: 'formatted' });
  });
});

describe('moving around the workbook', () => {
  it('starts at the top-left cell of each sheet', async () => {
    await open();
    expect(store.selection).toEqual({ row: 0, col: 0, rects: [{ r1: 0, c1: 0, r2: 0, c2: 0 }] });
    store.setSelection({ row: 1, col: 1, rects: [{ r1: 1, c1: 1, r2: 1, c2: 1 }] });
    store.setActive(1);
    expect(store.selection.row).toBe(0);
    store.setActive(0);
    expect(store.selection.row).toBe(1);
  });

  it('reveals a cell, switching sheet if needed, and can hand the grid the keyboard', async () => {
    await open();
    const seq = store.revealSeq;
    store.reveal(1, 2, 2);
    expect(store.active).toBe(1);
    expect(store.selection.rects).toEqual([{ r1: 2, c1: 2, r2: 2, c2: 2 }]);
    expect(store.revealSeq).toBe(seq + 1);
    expect(store.revealFocus).toBe(false);
    store.reveal(1, 0, 0);
    expect(store.active).toBe(1);
    store.focusGrid();
    expect(store.revealFocus).toBe(true);
    expect(store.revealSeq).toBe(seq + 3);
  });

  it('reads column widths in characters, with the file’s own widths', async () => {
    await open();
    expect(store.columnWidth(0)).toBeCloseTo(20, 0);
    await store.setColumnWidth(1, 10000);
    expect(store.view?.colWidth(1)).toBe(2000);
    await store.setColumnWidth(1, -1);
    expect(store.view?.colWidth(1)).toBe(2000);
    expect(new WorkbookStore({ drafts: false }).columnWidth(0)).toBe(9.14);
  });

  it('outlines the references in a formula being typed, by sheet, redrawing only when they change', async () => {
    await open();
    store.setActive(1);
    const listener = vi.fn();
    store.subscribe(listener);
    store.showFormulaRefs("=SUM(A1:B2)+lines!C3+'Other'!D4+Nowhere!E5+[1]Lines!A1+A1:B2");
    expect(store.formulaRefs).toEqual([
      { sheet: 1, rect: { r1: 0, c1: 0, r2: 1, c2: 1 }, color: 0 },
      { sheet: 0, rect: { r1: 2, c1: 2, r2: 2, c2: 2 }, color: 1 },
      { sheet: 1, rect: { r1: 3, c1: 3, r2: 3, c2: 3 }, color: 2 },
      { sheet: 1, rect: { r1: 0, c1: 0, r2: 1, c2: 1 }, color: 0 },
    ]);
    expect(listener).toHaveBeenCalledTimes(1);
    store.showFormulaRefs('=SUM(A1:B2)+lines!C3+\'Other\'!D4+Nowhere!E5+[1]Lines!A1+A1:B2+1');
    expect(listener).toHaveBeenCalledTimes(1);
    store.showFormulaRefs('A1');
    expect(store.formulaRefs).toEqual([]);
    store.showFormulaRefs(undefined);
    expect(listener).toHaveBeenCalledTimes(2);
  });
});

describe('stopping the editor', () => {
  it('drops work still waiting on the worker, quietly', async () => {
    await open();
    const typing = store.input(1, 1, '5');
    // Until the edit has gone to the worker.
    for (let i = 0; i < 100 && !store.busy; i++) await Promise.resolve();
    expect(store.busy).toBe(true);
    const listener = vi.fn();
    store.subscribe(listener);
    store.dispose();
    await typing;
    expect(store.toasts).toEqual([]);
    expect(listener).not.toHaveBeenCalled();
  });

  it('says no workbook is open once stopped', async () => {
    await open();
    store.dispose();
    await store.input(1, 1, '5');
    expect(store.toasts.map((t) => english.text(t.text))).toEqual(['No workbook is open.']);
  });
});

describe('drafts without browser storage', () => {
  it('opens and edits as usual, offering nothing back', async () => {
    expect(typeof indexedDB).toBe('undefined');
    await open({ drafts: true });
    expect(store.status).toBe('ready');
    await store.input(1, 1, '5');
    // Past the moment the draft would be written.
    await new Promise((r) => setTimeout(r, 450));
    expect(store.confirmRequest).toBeUndefined();
    expect(store.toasts).toEqual([]);
    store.dispose();

    await open({ drafts: true });
    await new Promise((r) => setTimeout(r, 20));
    expect(store.confirmRequest).toBeUndefined();
    expect(store.editedCount).toBe(0);
  });
});
