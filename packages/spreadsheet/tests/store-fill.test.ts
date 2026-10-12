import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { SeriesOptions } from '../src/core/types';
import { english } from '../src/i18n';
import type { WorkbookSchema } from '../src/schema/types';
import type { Answer, ConfirmRequest, StoreOptions, WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** A day as Excel numbers it, in the 1900 or 1904 date system. */
const serial = (y: number, m: number, d: number, date1904 = false) => (Date.UTC(y, m - 1, d) - (date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30))) / 86_400_000;

/**
 * Data: A1 = 1 and A3 = 99; B1 = 31 Jan 2026 as a date; C1 = "Item 1";
 * D1 = "Montag"; E1 = "South"; F1 = 2 and F2 = F1*2; L1:N1 = 1, 3, 4.
 * Locked is protected, with A1 = 1 and only A3 unlocked.
 */
async function workbook(date1904 = false): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  if (date1904) wb.properties.date1904 = true;
  const data = wb.addWorksheet('Data');
  data.getCell('A1').value = 1;
  data.getCell('A3').value = 99;
  data.getCell('B1').value = serial(2026, 1, 31, date1904);
  data.getCell('B1').numFmt = 'yyyy-mm-dd';
  data.getCell('C1').value = 'Item 1';
  data.getCell('D1').value = 'Montag';
  data.getCell('E1').value = 'South';
  data.getCell('F1').value = 2;
  data.getCell('F2').value = { formula: 'F1*2', result: 4 };
  ['L1', 'M1', 'N1'].forEach((ref, i) => (data.getCell(ref).value = [1, 3, 4][i]));
  const locked = wb.addWorksheet('Locked');
  locked.getCell('A1').value = 1;
  locked.getCell('A3').protection = { locked: false };
  await locked.protect('', {});
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => {
  store?.dispose();
  vi.unstubAllGlobals();
});

async function open(options: StoreOptions = {}, file = bytes): Promise<WorkbookStore> {
  store = await testStore(options);
  await opened(store, file, 'data.xlsx');
  return store;
}

const cell = (ref: string) => {
  const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
  const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  return store.view?.get(Number(m[2]) - 1, col);
};
const values = (...refs: string[]) => refs.map((r) => cell(r)?.n ?? cell(r)?.s);
const toasts = () => store.toasts.map((t) => english.text(t.text));
const lastAction = () => english.text(store.history.at(-1)?.action);
const series = (options: Partial<SeriesOptions> = {}): SeriesOptions => ({ by: 'columns', type: 'linear', dateUnit: 'day', step: 1, trend: false, ...options });

async function answer(reply: Answer): Promise<ConfirmRequest> {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  const asked = store.confirmRequest;
  if (!asked) throw new Error('Nothing was asked');
  asked.resolve(reply);
  return asked;
}

describe('the Fill Series dialog', () => {
  it('starts in rows for a wide selection, and with dates when the first cell is one', async () => {
    await open();
    store.goTo('A1:D1');
    expect(store.seriesStart()).toEqual({ by: 'rows', type: 'linear', dateUnit: 'day', step: 1, trend: false });
    store.goTo('B1:B4');
    expect(store.seriesStart()).toMatchObject({ by: 'columns', type: 'date' });
    store.goTo('A1');
    expect(store.seriesStart().by).toBe('columns');
  });

  it('fills a series down the selection, undone in one step', async () => {
    await open();
    store.goTo('A1:A5');
    await store.fillSeries(series({ step: 2 }));
    expect(values('A1', 'A2', 'A3', 'A4', 'A5')).toEqual([1, 3, 5, 7, 9]);
    expect(lastAction()).toBe('Fill series');
    await store.undo();
    expect(values('A2', 'A3', 'A4', 'A5')).toEqual([undefined, 99, undefined, undefined]);
  });

  it('leaves the cells after the stop value as they were', async () => {
    await open();
    store.goTo('A1:A5');
    await store.fillSeries(series({ stop: 2 }));
    expect(values('A2', 'A3', 'A4')).toEqual([2, 99, undefined]);
  });

  it('runs on from one selected cell until the stop value, as Excel does', async () => {
    await open();
    store.goTo('A1');
    await store.fillSeries(series({ stop: 4 }));
    expect(values('A2', 'A3', 'A4', 'A5')).toEqual([2, 3, 4, undefined]);
    store.goTo('A1');
    await store.fillSeries(series({ by: 'rows', step: -1, stop: -1 }));
    expect(values('B1', 'C1')).toEqual([0, -1]);
    expect(cell('D1')?.s).toBe('Montag');
  });

  it('steps dates by month, giving the new cells the first cell’s date format', async () => {
    await open();
    store.goTo('B1:B3');
    await store.fillSeries(series({ type: 'date', dateUnit: 'month' }));
    expect(values('B2', 'B3')).toEqual([serial(2026, 2, 28), serial(2026, 3, 31)]);
    expect(store.styles[cell('B3')!.style].numFmt).toBe('yyyy-mm-dd');
  });

  it('steps dates in a workbook that counts from 1904', async () => {
    await open({}, await workbook(true));
    expect(store.date1904).toBe(true);
    store.goTo('B1:B3');
    await store.fillSeries(series({ type: 'date', dateUnit: 'month' }));
    expect(values('B2', 'B3')).toEqual([serial(2026, 2, 28, true), serial(2026, 3, 31, true)]);
  });

  it('replaces the starting numbers with a trend fitted to them', async () => {
    await open();
    store.goTo('L1:P1');
    await store.fillSeries(series({ by: 'rows', trend: true }));
    expect(values('L1', 'M1', 'N1', 'O1', 'P1')).toEqual([1.16666666666667, 2.66666666666667, 4.16666666666667, 5.66666666666667, 7.16666666666667]);
  });

  it('says so, and changes nothing, when no row or column starts with a number', async () => {
    await open();
    store.goTo('K1:K4');
    await store.fillSeries(series());
    store.goTo('C1:C3');
    await store.fillSeries(series());
    expect(toasts()).toEqual([
      'Nothing was filled: a series starts from a number or date in the first cell of each row or column, and none of them starts with one.',
      'Nothing was filled: a series starts from a number or date in the first cell of each row or column, and none of them starts with one.',
    ]);
    expect(store.history).toHaveLength(0);
  });

  it('continues text with AutoFill as dragging would', async () => {
    await open();
    store.goTo('C1:C3');
    await store.fillSeries(series({ type: 'autofill' }));
    expect(values('C2', 'C3')).toEqual(['Item 2', 'Item 3']);
  });

  it('asks before replacing a formula with a number', async () => {
    await open();
    store.goTo('F1:F3');
    const filling = store.fillSeries(series());
    expect(english.text((await answer('cancel')).title)).toBe('Replace a formula?');
    await filling;
    expect(cell('F2')?.f).toBe('F1*2');
  });

  it('continues formulas with AutoFill without asking about them', async () => {
    await open();
    store.goTo('F1:F4');
    await store.fillSeries(series({ type: 'autofill' }));
    expect(store.confirmRequest).toBeUndefined();
    expect([cell('F3')?.n, cell('F4')?.f]).toEqual([2, 'F3*2']);
  });

  it("leaves a protected sheet's locked cells, as filling by dragging does", async () => {
    await open();
    store.setActive(1);
    store.goTo('A1:A4');
    await store.fillSeries(series());
    expect(values('A2', 'A3', 'A4')).toEqual([undefined, 3, undefined]);
    expect(toasts()).toEqual(['2 cells were left unchanged. This cell is locked: its sheet is protected in Excel.']);
  });
});

describe("a schema's read-only cells", () => {
  /** Lines!A1:C4 = Item, Qty, Price with three records; Item and Price are read-only. */
  async function lines(): Promise<Uint8Array> {
    const wb = new ExcelJS.Workbook();
    const sheet = wb.addWorksheet('Lines');
    sheet.addRow(['Item', 'Qty', 'Price']);
    sheet.addRow(['Bolts', 2, 3]);
    sheet.addRow(['Nuts', 4, 5]);
    sheet.addRow(['Saw', 1, 7]);
    return new Uint8Array(await wb.xlsx.writeBuffer());
  }
  const schema: WorkbookSchema = {
    id: 'lines',
    name: 'Lines',
    sheets: [
      {
        name: 'Lines',
        tables: [
          {
            key: 'lines',
            at: 'A1',
            fields: [
              { key: 'item', label: 'Item', type: 'text', readOnly: true },
              { key: 'qty', label: 'Qty', type: 'integer' },
              { key: 'price', label: 'Price', type: 'number', readOnly: true },
            ],
          },
        ],
      },
    ],
  };
  const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
  const left = '4 cells were left unchanged. “Item” is read-only.';

  it('are left as they were by dragging the fill handle', async () => {
    await open({ schemas: [schema] }, await lines());
    await store.fill(rect(1, 0, 1, 2), rect(1, 0, 3, 2));
    expect(values('A3', 'A4', 'B3', 'B4', 'C3', 'C4')).toEqual(['Nuts', 'Saw', 2, 2, 5, 7]);
    expect(toasts()).toEqual([left]);
  });

  it('are left as they were by filling down', async () => {
    await open({ schemas: [schema] }, await lines());
    store.goTo('A2:C4');
    await store.fillDown();
    expect(values('A3', 'B3', 'C4')).toEqual(['Nuts', 2, 7]);
    expect(toasts()).toEqual([left]);
  });

  it('are left as they were by pasting, from here or elsewhere, and by clearing', async () => {
    await open({ schemas: [schema] }, await lines());
    // Copied here: A2:C2 over A3:C3.
    store.noteCopy(rect(1, 0, 1, 2));
    await store.paste({ row: 2, col: 0 }, [['Bolts', '2', '3']], rect(2, 0));
    expect(values('A3', 'B3', 'C3')).toEqual(['Nuts', 2, 5]);
    // Pasted from another app.
    await store.paste({ row: 3, col: 0 }, [['Drill', '9', '99']], rect(3, 0));
    expect(values('A4', 'B4', 'C4')).toEqual(['Saw', 9, 7]);
    await store.clear([rect(1, 0, 1, 2)]);
    expect(values('A2', 'B2', 'C2')).toEqual(['Bolts', undefined, 3]);
  });

  it('are left as they were by Fill Series', async () => {
    await open({ schemas: [schema] }, await lines());
    store.goTo('B2:C4');
    await store.fillSeries(series());
    expect(values('B3', 'B4', 'C3', 'C4')).toEqual([3, 4, 5, 7]);
    expect(toasts()).toEqual(['2 cells were left unchanged. “Price” is read-only.']);
  });
});

describe('lists that filling continues', () => {
  const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

  it("continues the host's lists, and takes new ones while open", async () => {
    await open({ customLists: [['North', 'East', 'South', 'West']] });
    await store.fill(rect(0, 4), rect(0, 4, 2, 4));
    expect(values('E2', 'E3')).toEqual(['West', 'North']);
    store.setCustomLists([['South', 'Middle']]);
    await store.fill(rect(0, 4), rect(0, 4, 1, 4));
    expect(values('E2')).toEqual(['Middle']);
  });

  it('continues day names in the editor’s language', async () => {
    await open({ language: 'de' });
    await store.fill(rect(0, 3), rect(0, 3, 2, 3));
    expect(values('D2', 'D3')).toEqual(['Dienstag', 'Mittwoch']);
  });

  it("keeps the user's own lists in the browser, for each user, and continues them", async () => {
    const items = new Map<string, string>();
    vi.stubGlobal('localStorage', { getItem: (k: string) => items.get(k) ?? null, setItem: (k: string, v: string) => void items.set(k, v), removeItem: (k: string) => void items.delete(k) });
    await open({ user: 'ana' });
    store.setUserLists([['South', 'Sea', 'Shore']]);
    expect(store.listsNotKept).toBe(false);
    expect(JSON.parse(items.get('wbe-custom-lists:ana')!)).toEqual([['South', 'Sea', 'Shore']]);
    store.goTo('E1:E3');
    await store.fillSeries(series({ type: 'autofill' }));
    expect(values('E2', 'E3')).toEqual(['Sea', 'Shore']);
    store.setUser('ben');
    expect(store.userLists).toEqual([]);
    store.setUser('ana');
    const deleting = store.deleteUserList(0);
    expect(english.text((await answer('confirm')).body)).toBe('Filling will no longer continue “South, Sea, Shore”.');
    expect(await deleting).toBe(true);
    expect(items.has('wbe-custom-lists:ana')).toBe(false);
  });

  it("keeps the user's lists until the page closes where the browser won't store them", async () => {
    await open();
    store.setUserLists([['South', 'Sea']]);
    expect(store.listsNotKept).toBe(true);
    expect(store.userLists).toEqual([['South', 'Sea']]);
    await store.fill(rect(0, 4), rect(0, 4, 1, 4));
    expect(values('E2')).toEqual(['Sea']);
  });

  it('imports the text the selected cells show', async () => {
    await open();
    store.goTo('A1:E1');
    expect(store.selectedTexts()).toEqual(['1', '2026-01-31', 'Item 1', 'Montag', 'South']);
  });
});
