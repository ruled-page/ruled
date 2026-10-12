import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import type { WorkbookSchema } from '../src/schema/types';
import type { WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** Lines!A1:B4: Item and Qty, with three records. */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const lines = wb.addWorksheet('Lines');
  lines.addRow(['Item', 'Qty']);
  lines.addRow(['Bolts', 2]);
  lines.addRow(['Nuts', 4]);
  lines.addRow(['Bolts', 1]);
  wb.addWorksheet('Notes');
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

const schema: WorkbookSchema = {
  id: 'lines',
  name: 'Order lines',
  sheets: [
    {
      name: 'Lines',
      tables: [{ key: 'lines', at: 'A1', fields: [{ key: 'item', label: 'Item', type: 'text' }, { key: 'qty', label: 'Qty', type: 'integer', min: 0 }] }],
    },
  ],
};

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => store?.dispose());

async function open(schemas: WorkbookSchema[] = []): Promise<WorkbookStore> {
  store = await testStore({ schemas });
  await opened(store, bytes, 'lines.xlsx');
  return store;
}

/** What has been said to screen readers and not shown, newest last. */
const said = () => store.spoken.map((s) => english.text(s.text));
const last = () => said().at(-1);

describe('what screen readers are told', () => {
  it('says which file opened and how many sheets it has', async () => {
    await open();
    expect(last()).toBe('Opened lines.xlsx, 2 sheets');
  });

  it('says how many rows a filter hides, and when it is cleared', async () => {
    await open();
    const rect = { r1: 0, c1: 0, r2: 3, c2: 1 };
    await store.setFilter(rect, 0, new Set(['Nuts']));
    expect(last()).toBe('Filtered: 2 rows hidden');
    await store.clearFilter();
    expect(last()).toBe('Filter cleared: every row shown');
  });

  it('says what undo and redo undid and redid, structural changes too', async () => {
    await open();
    store.setSelection({ row: 1, col: 0, rects: [{ r1: 1, c1: 0, r2: 1, c2: 1 }] });
    await store.formatCells({ bold: true });
    expect(last()).toBe('Format cells');
    await store.undo();
    expect(last()).toBe('Undone: Format cells');
    await store.redo();
    expect(last()).toBe('Redone: Format cells');

    // Typing is heard as it is typed, so it isn't said again; undoing it is.
    await store.input(1, 0, 'Washers');
    expect(last()).toBe('Redone: Format cells');
    await store.undo();
    expect(last()).toBe('Undone: Edit');

    await store.restructure({ type: 'insertRows', sheet: 0, at: 2, count: 2 });
    const inserted = last();
    expect(inserted).toMatch(/2 rows/);
    await store.undo();
    expect(last()).toBe(`Undone: ${inserted}`);
    await store.redo();
    expect(last()).toBe(`Redone: ${inserted}`);
  });

  it('says what is wrong with a value just typed, once its checks have run', async () => {
    await open([schema]);
    await store.input(1, 1, '-3');
    expect(last()).toMatch(/^B2: error: /);
    // A good value says nothing more.
    const before = said().length;
    await store.input(2, 1, '5');
    expect(said()).toHaveLength(before);
    // Messages shown are said by where they are shown, not again here.
    store.toast('error', 'The disk is full');
    expect(said()).toHaveLength(before);
  });
});

describe('messages', () => {
  it('stay while the pointer or focus is on them, and warnings stay longer than news', async () => {
    await open();
    vi.useFakeTimers();
    try {
      store.toast('info', 'Saved');
      store.toast('warn', 'Some cells were skipped');
      vi.advanceTimersByTime(5000);
      expect(store.toasts.map((t) => english.text(t.text))).toEqual(['Some cells were skipped']);
      store.holdToasts(true);
      vi.advanceTimersByTime(20_000);
      expect(store.toasts).toHaveLength(1);
      store.holdToasts(false);
      vi.advanceTimersByTime(1000);
      expect(store.toasts).toHaveLength(0);
    } finally {
      vi.useRealTimers();
    }
  });
});
