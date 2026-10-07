import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import type { Answer, ConfirmRequest, WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

type Validation = Record<string, unknown>;

/**
 * Orders row 2 has a rule in each column:
 *   A2 list "North,South", stopping with its own title and message;
 *   B2 whole number 1 to 10, warning with a message;
 *   C2 list from Lists!A1:A3, information only;
 *   D2 text of at most 4 characters, stopping with Excel's default message;
 *   E2 list from a formula over the first two of Lists!A1:A3;
 *   F2 list "Yes,No" that doesn't alert at all.
 * G1:G3 = 1, 5, 2 are formatted when over 3.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const orders = wb.addWorksheet('Orders');
  const lists = wb.addWorksheet('Lists');
  orders.addRow(['Region', 'Qty', 'Status', 'Code', 'Product', 'Gift']);
  const rules: Record<string, Validation> = {
    A2: { type: 'list', formulae: ['"North,South"'], showErrorMessage: true, errorTitle: 'Region', error: 'Pick a region.' },
    B2: { type: 'whole', operator: 'between', formulae: [1, 10], showErrorMessage: true, errorStyle: 'warning', error: 'Usually 1 to 10.' },
    C2: { type: 'list', formulae: ['Lists!$A$1:$A$3'], showErrorMessage: true, errorStyle: 'information' },
    D2: { type: 'textLength', operator: 'lessThanOrEqual', formulae: [4], showErrorMessage: true },
    E2: { type: 'list', formulae: ['OFFSET(Lists!$A$1,0,0,2,1)'], showErrorMessage: true },
    F2: { type: 'list', formulae: ['"Yes,No"'] },
  };
  for (const [ref, rule] of Object.entries(rules)) orders.getCell(ref).dataValidation = rule as unknown as ExcelJS.DataValidation;
  ['Open', 'Closed', 'Held'].forEach((v, i) => (lists.getCell(`A${i + 1}`).value = v));
  [1, 5, 2].forEach((v, i) => (orders.getCell(`G${i + 1}`).value = v));
  orders.addConditionalFormatting({
    ref: 'G1:G3',
    rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: ['3'], priority: 1, style: { font: { bold: true } } }],
  });
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
  await opened(store, bytes, 'orders.xlsx');
  return store;
}

const cell = (row: number, col: number, sheet = store.active) => store.sheets[sheet].get(row, col);
const toasts = () => store.toasts.map((t) => [t.kind, english.text(t.text)]);

async function answer(reply: Answer): Promise<ConfirmRequest> {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  const asked = store.confirmRequest;
  if (!asked) throw new Error('Nothing was asked');
  asked.resolve(reply);
  return asked;
}

describe('typing into cells with Excel validation', () => {
  it('refuses a value a stopping list does not allow, in the workbook’s own words', async () => {
    await open();
    await store.input(1, 0, 'East');
    expect(toasts()).toEqual([['error', 'Region: Pick a region.']]);
    expect(cell(1, 0)).toBeFalsy();
    await store.input(1, 0, " north ");
    expect(cell(1, 0)?.s).toBe(' north ');
    await store.input(1, 0, "'South");
    expect(cell(1, 0)?.s).toBe('South');
  });

  it('asks before keeping a value a warning rule does not allow', async () => {
    await open();
    const typing = store.input(1, 1, '20');
    const asked = await answer('cancel');
    await typing;
    expect(english.text(asked.title)).toBe('Value not allowed');
    expect(english.text(asked.body)).toBe('Usually 1 to 10. Keep this value anyway?');
    expect(english.text(asked.confirmLabel)).toBe('Keep value');
    expect(cell(1, 1)).toBeFalsy();

    const again = store.input(1, 1, '20');
    await answer('confirm');
    await again;
    expect(cell(1, 1)?.n).toBe(20);
    await store.input(1, 1, '7');
    expect(cell(1, 1)?.n).toBe(7);
  });

  it('only informs for an information rule, and keeps the value', async () => {
    await open();
    await store.input(1, 2, 'Lost');
    expect(toasts()).toEqual([['info', "This value doesn't match the data validation restrictions defined for this cell."]]);
    expect(cell(1, 2)?.s).toBe('Lost');
    await store.input(1, 2, 'held');
    expect(toasts()).toHaveLength(1);
  });

  it('stops a value that breaks a rule, with Excel’s message when the workbook has none', async () => {
    await open();
    await store.input(1, 3, 'ABCDE');
    expect(toasts()).toEqual([['error', "This value doesn't match the data validation restrictions defined for this cell."]]);
    expect(cell(1, 3)).toBeFalsy();
    await store.input(1, 3, 'ABCD');
    expect(cell(1, 3)?.s).toBe('ABCD');
  });

  it('checks a list worked out by a formula', async () => {
    await open();
    await store.input(1, 4, 'Held');
    expect(toasts()).toHaveLength(1);
    expect(cell(1, 4)).toBeFalsy();
    await store.input(1, 4, 'Closed');
    expect(cell(1, 4)?.s).toBe('Closed');
  });

  it('leaves alone lists that do not alert, formulas and blanks', async () => {
    await open();
    await store.input(1, 5, 'Maybe');
    await store.input(1, 0, '="East"');
    await store.input(1, 3, '');
    expect(store.toasts).toEqual([]);
    expect(cell(1, 5)?.s).toBe('Maybe');
    expect(cell(1, 0)?.f).toBe('"East"');
  });
});

describe('list options', () => {
  it('gives a list’s values, read again when its source sheet changes', async () => {
    await open();
    const inline = store.listAt(0, 1, 0)!;
    expect(store.listOptions(inline)).toEqual(['North', 'South']);
    const ranged = store.listAt(0, 1, 2)!;
    const first = store.listOptions(ranged);
    expect(first).toEqual(['Open', 'Closed', 'Held']);
    expect(store.listOptions(ranged)).toBe(first);

    store.setActive(1);
    await store.input(1, 0, 'Shut');
    expect(store.listOptions(ranged)).toEqual(['Open', 'Shut', 'Held']);
  });

  it('has none where no list applies', async () => {
    await open();
    expect(store.listAt(0, 5, 0)).toBeUndefined();
    expect(store.listAt(1, 0, 0)).toBeUndefined();
    expect(store.listAt(7, 0, 0)).toBeUndefined();
  });

  it('works out a formula list in the worker, showing the last values while new ones load', async () => {
    await open();
    const dv = store.listAt(0, 1, 4)!;
    expect(dv.source.kind).toBe('formula');
    expect(store.listOptions(dv)).toBeUndefined();
    await vi.waitFor(() => expect(store.listOptions(dv)).toEqual(['Open', 'Closed']));

    store.setActive(1);
    await store.input(0, 0, 'New');
    // The values change with the data; until the worker answers, the old ones show.
    expect(store.listOptions(dv)).toEqual(['Open', 'Closed']);
    await vi.waitFor(() => expect(store.listOptions(dv)).toEqual(['New', 'Closed']));
  });

  it('finds the rule covering a cell, with its place on the sheet', async () => {
    await open();
    expect(store.ruleAt(0, 1, 3)).toMatchObject({ index: 1, rule: { kind: 'textLength' } });
    expect(store.ruleAt(0, 2, 3)).toBeUndefined();
    expect(store.ruleAt(9, 0, 0)).toBeUndefined();
  });
});

describe('conditional formatting', () => {
  it('is worked out in the worker when a sheet opens, and again after its values change', async () => {
    await open();
    await vi.waitFor(() => expect(store.view?.cfAt(1, 6)).toBeDefined());
    expect(store.view?.cfAt(0, 6)).toBeUndefined();
    await store.input(0, 6, '9');
    await vi.waitFor(() => expect(store.view?.cfAt(0, 6)).toBeDefined());
    await store.input(1, 6, '0');
    await vi.waitFor(() => expect(store.view?.cfAt(1, 6)).toBeUndefined());
  });

  it('says once when a rule covers too many cells to show', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Big');
    ws.getCell('A1').value = 1;
    ws.getCell('Z10000').value = 2;
    ws.addConditionalFormatting({
      ref: 'A1:Z10000',
      rules: [{ type: 'expression', formulae: ['A1>1'], priority: 1, style: { font: { bold: true } } }],
    });
    store = await testStore();
    await opened(store, new Uint8Array(await wb.xlsx.writeBuffer()), 'big.xlsx');
    await vi.waitFor(() => expect(toasts()).toEqual([['info', '1 conditional format rule on Big covers too many cells to show here.']]));
    await store.input(0, 0, '5');
    await new Promise((r) => setTimeout(r, 50));
    expect(toasts()).toHaveLength(1);
  });
});
