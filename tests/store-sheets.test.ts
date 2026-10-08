import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import { fileKey, type Draft } from '../src/state/drafts';
import type { Answer, StoreOptions, WorkbookStore } from '../src/state/store';
import type { WorkbookSchema } from '../src/schema/types';
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

/** Prices!A1:B3 (Item, Price), Totals!A1 = SUM(Prices!B2:B3), Archive hidden, and an empty Notes. */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const prices = wb.addWorksheet('Prices');
  prices.addRow(['Item', 'Price']);
  prices.addRow(['Bolts', 3]);
  prices.addRow(['Nuts', 0.5]);
  wb.addWorksheet('Totals').getCell('A1').value = { formula: 'SUM(Prices!B2:B3)', result: 3.5 };
  wb.addWorksheet('Archive', { state: 'hidden' }).getCell('A1').value = 'old';
  wb.addWorksheet('Notes');
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

const schema: WorkbookSchema = {
  id: 'prices',
  name: 'Price list',
  sheets: [
    { name: 'Prices', tables: [{ key: 'prices', at: 'A1', fields: [{ key: 'item', label: 'Item', type: 'text' }, { key: 'price', label: 'Price', type: 'number' }] }] },
    { name: { pattern: '^Tot' }, optional: true },
  ],
};

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
beforeEach(() => kept.clear());
afterEach(() => store?.dispose());

async function open(options: StoreOptions = {}, from = bytes, reply?: Answer): Promise<WorkbookStore> {
  store = await testStore(options);
  const opening = store.open('prices.xlsx', from.slice().buffer as ArrayBuffer);
  if (reply) {
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    store.confirmRequest?.resolve(reply);
  }
  await opening;
  for (let i = 0; i < 200 && !store.calc; i++) await new Promise((r) => setTimeout(r, 5));
  return store;
}

/** Answers the confirmation the call asks for. */
async function answering<T>(reply: Answer, call: Promise<T>): Promise<T> {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  store.confirmRequest?.resolve(reply);
  return call;
}

const names = () => store.sheets.map((s) => s.info.name);
const active = () => store.sheets[store.active].info.name;
const toasts = () => store.toasts.map((t) => english.text(t.text));
const actions = () => store.history.map((h) => english.text(h.action));
const totals = () => store.sheets[names().indexOf('Totals')].get(0, 0);

describe('adding a sheet', () => {
  it('adds one after the active sheet, named as Excel names it, and shows it; undo and redo take it away and back', async () => {
    await open();
    expect(store.newSheetName).toBe('Sheet1');
    expect(await store.addSheet()).toBe(true);
    expect(names()).toEqual(['Prices', 'Sheet1', 'Totals', 'Archive', 'Notes']);
    expect(active()).toBe('Sheet1');
    expect(actions()).toEqual(['Added sheet “Sheet1”']);
    expect(store.editedCount).toBe(1);
    expect(store.newSheetName).toBe('Sheet2');
    await store.undo();
    expect(names()).toEqual(['Prices', 'Totals', 'Archive', 'Notes']);
    expect(active()).toBe('Prices');
    await store.redo();
    expect(active()).toBe('Sheet1');
  });

  it("names it in the editor's language", async () => {
    await open({ language: 'de' });
    await store.addSheet(4);
    expect(names()[4]).toBe('Tabelle1');
  });
});

describe('renaming a sheet', () => {
  it('renames it, with formulas that use it following', async () => {
    await open();
    store.setActive(1);
    store.setSelection({ row: 0, col: 0, rects: [{ r1: 0, c1: 0, r2: 0, c2: 0 }] });
    expect(await store.renameSheet(0, 'Price list')).toBe(true);
    expect(names()[0]).toBe('Price list');
    expect(totals()?.f).toBe("SUM('Price list'!B2:B3)");
    expect(active()).toBe('Totals');
    expect(actions()).toEqual(['Renamed sheet “Prices” to “Price list”']);
    await store.undo();
    expect(totals()?.f).toBe('SUM(Prices!B2:B3)');
  });

  it("refuses a name Excel wouldn't take, and changes nothing for the same name", async () => {
    await open();
    expect(await store.renameSheet(0, 'Totals')).toBe(false);
    expect(await store.renameSheet(0, 'A/B')).toBe(false);
    expect(toasts()).toEqual(["There's already a sheet called “Totals” (capital letters don't make a name different).", "A sheet name can't contain : \\ / ? * [ or ], or a line break."]);
    expect(await store.renameSheet(0, 'Prices')).toBe(true);
    expect(store.undoStack).toEqual([]);
  });
});

describe('deleting a sheet', () => {
  it('asks first, then deletes it and shows the next sheet; formulas that used it show #REF!', async () => {
    await open();
    expect(await answering('cancel', store.deleteSheet(0))).toBe(false);
    expect(names()).toHaveLength(4);
    expect(await answering('confirm', store.deleteSheet(0))).toBe(true);
    expect(names()).toEqual(['Totals', 'Archive', 'Notes']);
    expect(active()).toBe('Totals');
    expect(totals()?.f).toBe('SUM(#REF!)');
    expect(toasts()).toEqual(['1 formula now shows #REF! because it used cells that were deleted or pasted over.']);
    await store.undo();
    expect(names()[0]).toBe('Prices');
    expect(active()).toBe('Prices');
    expect(totals()?.f).toBe('SUM(Prices!B2:B3)');
  });

  it("deletes an empty sheet without asking, and won't delete the last visible one", async () => {
    await open();
    expect(await store.deleteSheet(3)).toBe(true);
    expect(store.confirmRequest).toBeUndefined();
    await answering('confirm', store.deleteSheet(1));
    expect(await store.deleteSheet(0)).toBe(false);
    expect(toasts().at(-1)).toBe("A workbook must keep at least one visible sheet, so this one can't be deleted.");
  });
});

describe('moving a sheet', () => {
  it('moves past the next visible sheet, keeping the active sheet and selections with their sheets', async () => {
    await open();
    store.setSelection({ row: 2, col: 1, rects: [{ r1: 2, c1: 1, r2: 2, c2: 1 }] });
    expect(await store.moveSheetBy(0, 1)).toBe(true);
    expect(names()).toEqual(['Totals', 'Prices', 'Archive', 'Notes']);
    expect(active()).toBe('Prices');
    expect(store.selection.row).toBe(2);
    // The hidden Archive is passed over.
    await store.moveSheetBy(1, 1);
    expect(names()).toEqual(['Totals', 'Archive', 'Notes', 'Prices']);
    expect(await store.moveSheetBy(3, 1)).toBe(false);
    await store.undo();
    await store.undo();
    expect(names()).toEqual(['Prices', 'Totals', 'Archive', 'Notes']);
    expect(active()).toBe('Prices');
    expect(store.selection.row).toBe(2);
    await store.dropSheet(3, 0);
    expect(names()).toEqual(['Notes', 'Prices', 'Totals', 'Archive']);
    expect(await store.dropSheet(0, 1)).toBe(false);
    expect(actions().at(-1)).toBe('Moved sheet “Notes”');
  });
});

describe('hiding and unhiding sheets', () => {
  it('hides the active sheet, showing the next visible one, and unhides one, showing it', async () => {
    await open();
    expect(store.hiddenSheets).toEqual([2]);
    await store.hideSheet(0);
    expect(store.sheets[0].info.state).toBe('hidden');
    expect(active()).toBe('Totals');
    expect(store.hiddenSheets).toEqual([0, 2]);
    await store.unhideSheet(2);
    expect(active()).toBe('Archive');
    expect(actions()).toEqual(['Hid sheet “Prices”', 'Unhid sheet “Archive”']);
    await store.hideSheet(1);
    await store.hideSheet(2);
    expect(await store.hideSheet(3)).toBe(false);
    expect(toasts()).toEqual(["A workbook must keep at least one visible sheet, so this one can't be hidden."]);
  });
});

describe('a schema that finds its sheets by name', () => {
  it("keeps a sheet it checks matching it, and the sheets it needs", async () => {
    await open({ schemas: [schema] });
    expect(store.schema).toBeDefined();
    expect(await store.renameSheet(0, 'Items')).toBe(false);
    expect(toasts()).toEqual(['The “Price list” schema finds this sheet by its name, which has to stay "Prices".']);
    expect(await answering('confirm', store.deleteSheet(0))).toBe(false);
    expect(toasts().at(-1)).toBe("The “Price list” schema checks this sheet, so it can't be deleted.");
    // A name the schema's pattern still matches, and a sheet it can do without.
    expect(await store.renameSheet(1, 'Totals 2026')).toBe(true);
    expect(store.schema?.binding.sheets.map((b) => b.name)).toEqual(['Prices', 'Totals 2026']);
    expect(await answering('confirm', store.deleteSheet(1))).toBe(true);
    expect(store.schema).toBeDefined();
    // Moving and hiding a sheet it checks leave it applying.
    await store.moveSheetBy(0, 1);
    expect(store.schema?.binding.sheets[0].sheet).toBe(2);
    await store.hideSheet(1);
    expect(store.schema).toBeDefined();
  });
});

describe('a workbook whose structure is protected', () => {
  it('says why its sheets stay as they are', async () => {
    const files = unzipSync(bytes);
    files['xl/workbook.xml'] = strToU8(strFromU8(files['xl/workbook.xml']).replace('<sheets>', '<workbookProtection lockStructure="1"/><sheets>'));
    await open({}, zipSync(files));
    expect(store.structureLocked).toBe(true);
    expect(await store.addSheet()).toBe(false);
    expect(await store.moveSheetBy(0, 1)).toBe(false);
    expect(toasts()).toEqual([
      "The workbook's structure is protected, so sheets can't be added, deleted, renamed, moved, hidden or unhidden.",
      "The workbook's structure is protected, so sheets can't be added, deleted, renamed, moved, hidden or unhidden.",
    ]);
  });
});

describe('drafts', () => {
  it('brings back changes to sheets along with cell edits, in order', async () => {
    await open({ drafts: true });
    await store.addSheet();
    await store.input(0, 0, 'Fresh');
    await store.renameSheet(1, 'New');
    await new Promise((r) => setTimeout(r, 500));
    const key = await fileKey(bytes.slice().buffer as ArrayBuffer);
    expect(kept.get(key)?.ops.map((o) => o.type)).toEqual(['structure', 'input', 'structure']);
    store.dispose();
    await open({ drafts: true }, bytes, 'confirm');
    for (let i = 0; i < 200 && names()[1] !== 'New'; i++) await new Promise((r) => setTimeout(r, 5));
    expect(names()).toEqual(['Prices', 'New', 'Totals', 'Archive', 'Notes']);
    expect(store.sheets[1].get(0, 0)?.s).toBe('Fresh');
  });
});
