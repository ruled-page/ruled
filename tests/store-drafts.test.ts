import ExcelJS from 'exceljs';
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import { deleteDraft, fileKey, type Draft } from '../src/state/drafts';
import type { Answer, ConfirmRequest, StoreOptions, WorkbookStore } from '../src/state/store';
import { sample } from './helpers';
import { testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

// Node has no IndexedDB, so drafts are kept in a map here instead.
const kept = vi.hoisted(() => new Map<string, Draft>());
vi.mock('../src/state/drafts', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/state/drafts')>()),
  loadDraft: vi.fn(async (key: string) => kept.get(key)),
  saveDraft: vi.fn(async (draft: Draft) => void kept.set(draft.key, structuredClone(draft))),
  deleteDraft: vi.fn(async (key: string) => void kept.delete(key)),
}));

async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const lines = wb.addWorksheet('Lines');
  lines.addRow(['Item', 'Qty', 'Price', 'Total']);
  lines.addRow(['Bolts', 2, 3, { formula: 'B2*C2', result: 6 }]);
  lines.addRow(['Nuts', 4, 0.5, { formula: 'B3*C3', result: 2 }]);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
let key: string;
beforeAll(async () => {
  bytes = await workbook();
  key = await fileKey(bytes.slice().buffer as ArrayBuffer);
});

let store: WorkbookStore;
beforeEach(() => {
  kept.clear();
  vi.mocked(deleteDraft).mockClear();
});
afterEach(() => store?.dispose());

/** Opens the workbook; when a draft is offered, `reply` answers it. */
async function open(options: StoreOptions = {}, reply?: Answer): Promise<ConfirmRequest | undefined> {
  store = await testStore({ drafts: true, ...options });
  const opening = store.open('lines.xlsx', bytes.slice().buffer as ArrayBuffer);
  let asked: ConfirmRequest | undefined;
  if (reply) {
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    asked = store.confirmRequest;
    asked?.resolve(reply);
  }
  await opening;
  for (let i = 0; i < 200 && !store.calc; i++) await new Promise((r) => setTimeout(r, 5));
  return asked;
}

const cell = (row: number, col: number) => store.view?.get(row, col);
const draft = (ops: Draft['ops'], editedCount = ops.length): Draft => ({ key, fileName: 'lines.xlsx', updatedAt: Date.UTC(2026, 9, 1, 9, 30), editedCount, ops });

describe('drafts of unsaved changes', () => {
  it('keeps each edit shortly after it is made, and offers them back when the file is reopened', async () => {
    await open();
    await store.input(1, 1, '5');
    store.setSelection({ row: 2, col: 0, rects: [{ r1: 2, c1: 0, r2: 2, c2: 0 }] });
    await store.insertRows('above');
    await store.input(2, 0, 'Washers');
    await vi.waitFor(() => expect(kept.get(key)?.ops).toHaveLength(3));
    expect(kept.get(key)).toMatchObject({ fileName: 'lines.xlsx', editedCount: 3 });
    expect(kept.get(key)?.ops[1]).toEqual({ type: 'structure', op: { type: 'insertRows', sheet: 0, at: 2, count: 1 } });
    store.dispose();

    const asked = await open({}, 'confirm');
    expect(english.text(asked?.title)).toBe('Restore unsaved changes?');
    expect(english.text(asked?.body)).toMatch(/^This file was edited on .+ and not saved \(3 changed cells\)\. Restore those changes\?$/);
    expect([cell(1, 1)?.n, cell(1, 3)?.n, cell(2, 0)?.s, cell(3, 0)?.s]).toEqual([5, 15, 'Washers', 'Nuts']);
    expect(store.history.map((h) => english.text(h.action))).toEqual(['Restored unsaved changes', 'Inserted 1 row', 'Restored unsaved changes']);
    // Replaying doesn't move the selection.
    expect(store.selection.row).toBe(0);
    expect(store.undoStack).toHaveLength(3);
  });

  it('forgets the draft once every change is undone', async () => {
    kept.set(key, draft([{ type: 'input', sheet: 0, row: 1, col: 1, text: '5' }]));
    await open({}, 'confirm');
    expect(cell(1, 1)?.n).toBe(5);
    await store.undo();
    await vi.waitFor(() => expect(kept.has(key)).toBe(false));
  });

  it('deletes the draft when the user turns it down', async () => {
    kept.set(key, draft([{ type: 'input', sheet: 0, row: 1, col: 1, text: '5' }]));
    await open({}, 'cancel');
    expect(cell(1, 1)?.n).toBe(2);
    expect(store.editedCount).toBe(0);
    expect(kept.has(key)).toBe(false);
  });

  it('offers nothing for an empty draft', async () => {
    kept.set(key, draft([]));
    await open();
    expect(store.confirmRequest).toBeUndefined();
  });

  it('deletes the draft once the changes are saved', async () => {
    kept.set(key, draft([{ type: 'input', sheet: 0, row: 1, col: 1, text: '5' }]));
    const onSave = vi.fn();
    await open({ onSave }, 'confirm');
    await store.download();
    expect(onSave).toHaveBeenCalledOnce();
    expect(kept.has(key)).toBe(false);
  });

  it('deletes the draft when another file is opened over unsaved changes', async () => {
    await open();
    await store.input(1, 1, '5');
    const opening = store.open('budget.xlsx', sample('budget.xlsx').buffer as ArrayBuffer);
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    store.confirmRequest?.resolve('confirm');
    await opening;
    expect(deleteDraft).toHaveBeenCalledWith(key);
    expect(store.fileName).toBe('budget.xlsx');
  });
});
