import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import type { Answer, WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** Data!A2:A5 take whole numbers from 1 to 10; B2:B5 pick from Yes or No; C has none. Locked is protected. */
let bytes: Uint8Array;
beforeAll(async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  for (let r = 2; r <= 5; r++) {
    ws.getCell(`A${r}`).dataValidation = { type: 'whole', operator: 'between', formulae: [1, 10], showErrorMessage: true };
    ws.getCell(`B${r}`).dataValidation = { type: 'list', formulae: ['"Yes,No"'], showErrorMessage: true };
  }
  await wb.addWorksheet('Locked').protect('', {});
  bytes = new Uint8Array(await wb.xlsx.writeBuffer());
});

let store: WorkbookStore;
afterEach(() => store?.dispose());
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const select = (r: ReturnType<typeof rect>) => store.setSelection({ row: r.r1, col: r.c1, rects: [r] });
const toasts = () => store.toasts.map((t) => english.text(t.text));

async function open(): Promise<WorkbookStore> {
  store = await testStore();
  await opened(store, bytes, 'rules.xlsx');
  return store;
}

/** Waits for the store to ask something, and gives the answer. */
async function answer(reply: Answer): Promise<string> {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  if (!store.confirmRequest) throw new Error('Nothing was asked');
  const asked = english.text(store.confirmRequest.body);
  store.confirmRequest.resolve(reply);
  return asked;
}

describe("Excel's Data Validation dialog", () => {
  it("opens on the selection's rule, or a new one where it has none", async () => {
    await open();
    select(rect(2, 0, 3, 0));
    await store.openValidationDialog();
    expect(store.validationDialog).toMatchObject({ rects: [rect(2, 0, 3, 0)], rule: { kind: 'whole', formula1: '1' } });
    select(rect(2, 2));
    await store.openValidationDialog();
    expect(store.validationDialog).toEqual({ rects: [rect(2, 2)], rule: undefined });
  });

  it('asks before erasing several rules, or taking one to cells without it', async () => {
    await open();
    select(rect(1, 0, 1, 1));
    const several = store.openValidationDialog();
    expect(await answer('confirm')).toMatch(/more than one type of validation/);
    await several;
    expect(store.validationDialog).toEqual({ rects: [rect(1, 0, 1, 1)] });
    store.set('validationDialog', undefined);
    select(rect(1, 1, 1, 2));
    const extend = store.openValidationDialog();
    expect(await answer('cancel')).toMatch(/extend Data Validation/);
    await extend;
    expect(store.validationDialog).toBeUndefined();
  });

  it('sets a rule on the selection, cutting it out of the rules there, and clears it, each a step that undoes', async () => {
    await open();
    const view = store.view!;
    const yesNo = view.info.listValidations[0];
    await store.setValidation([rect(2, 0, 2, 1)], { ...yesNo, rects: [] });
    expect(view.info.ruleValidations[0].rects).toEqual([rect(1, 0), rect(3, 0, 4, 0)]);
    expect(view.info.listValidations.map((l) => l.rects)).toEqual([[rect(1, 1), rect(3, 1, 4, 1)], [rect(2, 0, 2, 1)]]);
    await store.setValidation([rect(0, 0, 9, 9)], undefined);
    expect([view.info.listValidations, view.info.ruleValidations]).toEqual([[], []]);
    expect(store.history.map((h) => english.text(h.action))).toEqual(['Data validation', 'Clear data validation']);
    await store.undo();
    await store.undo();
    expect(view.info.listValidations).toEqual([yesNo]);
  });

  it('resolves list sources by the sheets and names there are', async () => {
    await open();
    store.setActive(1);
    expect(store.listContext()).toEqual({ sheetIndex: new Map([['data', 0], ['locked', 1]]), ownSheet: 1, names: [] });
    expect(store.validationPlace([rect(3, 3)])).toEqual({ active: { row: 0, col: 0 }, rects: [rect(3, 3)], locale: 'en-US', date1904: false });
  });

  it('is refused on a protected sheet', async () => {
    await open();
    store.setActive(1);
    await store.openValidationDialog();
    expect(store.validationDialog).toBeUndefined();
    expect(toasts()).toEqual(['This sheet is protected in Excel, so it is read-only here.']);
  });
});
