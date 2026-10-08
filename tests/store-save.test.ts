import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { CellFlag, VT } from '../src/core/types';
import { english } from '../src/i18n';
import type { WorkbookSchema } from '../src/schema/types';
import { SaveConflictError, type Answer, type ConfirmRequest, type SavedWorkbook, type StoreOptions, type WorkbookStore } from '../src/state/store';
import { has, sample } from './helpers';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** Lines!A1:D4 as in the other store tests, plus E2 = 3 shown as "3.00". */
async function workbook(hideFirst = false): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  if (hideFirst) wb.addWorksheet('Hidden').state = 'hidden';
  const lines = wb.addWorksheet('Lines');
  lines.addRow(['Item', 'Qty', 'Price', 'Total']);
  lines.addRow(['Bolts', 2, 3, { formula: 'B2*C2', result: 6 }]);
  lines.addRow(['Nuts', 4, 0.5, { formula: 'B3*C3', result: 2 }]);
  lines.addRow(['Washers', 1, 1, { formula: 'B4*C4', result: 1 }]);
  lines.getCell('E2').value = 3;
  lines.getCell('E2').numFmt = '0.00';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

/** A workbook made macro-enabled, with a stand-in VBA project. */
async function xlsm(): Promise<Uint8Array> {
  const files = unzipSync(await workbook());
  const text = (p: string) => strFromU8(files[p]);
  files['xl/vbaProject.bin'] = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4]);
  files['[Content_Types].xml'] = strToU8(
    text('[Content_Types].xml')
      .replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
      .replace('<Default Extension="xml"', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"'),
  );
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    text('xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
    ),
  );
  return zipSync(files);
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => {
  store?.dispose();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

async function open(options: StoreOptions = {}, file = bytes, name = 'lines.xlsx'): Promise<WorkbookStore> {
  store = await testStore(options);
  await opened(store, file, name);
  return store;
}

const toasts = () => store.toasts.map((t) => english.text(t.text));
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

/** Stands in for the browser's download link, recording each file it would download. */
function downloads(): { name: string; type: string }[] {
  const files: { name: string; type: string }[] = [];
  const blobs = new Map<string, Blob>();
  vi.spyOn(URL, 'createObjectURL').mockImplementation((blob) => {
    const url = `blob:${blobs.size}`;
    blobs.set(url, blob as Blob);
    return url;
  });
  vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => undefined);
  const link = { href: '', download: '', click: () => files.push({ name: link.download, type: blobs.get(link.href)!.type }) };
  vi.stubGlobal('document', { createElement: () => link });
  return files;
}

async function answer(reply: Answer): Promise<ConfirmRequest> {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  const asked = store.confirmRequest;
  if (!asked) throw new Error('Nothing was asked');
  asked.resolve(reply);
  return asked;
}

const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

describe('saving without a host', () => {
  it('downloads the workbook and says how many cells changed', async () => {
    const files = downloads();
    await open();
    expect(store.saveLabel).toBe('Save as Excel');
    expect(store.hasSaveHandler).toBe(false);
    await store.input(1, 1, '5');
    await store.download();
    expect(files).toEqual([{ name: 'lines.xlsx', type: XLSX }]);
    expect(toasts()).toEqual(['Saved lines.xlsx with 1 changed cell.']);
    expect(store.unsaved).toBe(false);
    expect(store.busy).toBe(false);
    await store.input(1, 1, '6');
    expect(store.unsaved).toBe(true);
  });

  it('says when the workbook was saved unchanged', async () => {
    const files = downloads();
    await open();
    await store.download();
    expect(files).toHaveLength(1);
    expect(toasts()).toEqual(['Saved lines.xlsx (unchanged).']);
  });

  it('downloads a macro workbook as one', async () => {
    const files = downloads();
    await open({}, await xlsm(), 'lines.xlsm');
    await store.download();
    // Blob types are lower-cased.
    expect(files[0]).toEqual({ name: 'lines.xlsm', type: 'application/vnd.ms-excel.sheet.macroenabled.12' });
  });
});

describe('results that may be out of date', () => {
  /** B2 uses a function the editor can't calculate, and C2 reads it. */
  async function chain(): Promise<Uint8Array> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Rates');
    ws.addRow(['Input', 'Lookup', 'Doubled']);
    ws.addRow([1, { formula: '_xludf.RATE_FOR(A2)', result: 0.25 }, { formula: 'B2*2', result: 0.5 }]);
    return new Uint8Array(await wb.xlsx.writeBuffer());
  }

  it('asks before saving them, and tells the host which cells they are', async () => {
    const saves: SavedWorkbook[] = [];
    await open({ onSave: (f) => void saves.push(f) }, await chain(), 'rates.xlsx');
    await store.download();
    // Nothing changed, so Excel's results stand and nothing is asked.
    expect(saves.map((f) => f.outOfDate)).toEqual([[]]);

    await store.input(1, 0, '2');
    const saving = store.download();
    expect(english.text((await answer('cancel')).title)).toBe('Save results that may be out of date?');
    await saving;
    expect(saves).toHaveLength(1);

    const again = store.download();
    expect(english.text((await answer('confirm')).body)).toMatch(/^2 formulas may be out of date\./);
    await again;
    expect(saves[1].outOfDate).toEqual([
      { sheet: 'Rates', ref: 'B2' },
      { sheet: 'Rates', ref: 'C2' },
    ]);
  });
});

describe('saving to the host', () => {
  it('hands over the file, its changes and the history in the editor’s language', async () => {
    let got: SavedWorkbook | undefined;
    let busy = false;
    await open({
      user: 'Ana',
      onSave: (f) => {
        busy = store.busy;
        got = f;
      },
    });
    expect(store.saveLabel).toBe('Save');
    await store.input(1, 1, '5');
    await store.download();
    expect(busy).toBe(true);
    expect(got).toMatchObject({ fileName: 'lines.xlsx', editedCount: 1, history: [{ user: 'Ana', action: 'Edit', cells: 1, refs: ['Lines!B2'] }] });
    expect(got!.changes).toEqual([{ sheet: 'Lines', ref: 'B2', before: '2', after: '5' }]);
    expect(got!.cell('Lines', 'D2')).toBe(15);
    expect(got!.schema).toBeUndefined();
    expect(toasts()).toEqual(['Saved lines.xlsx.']);
    expect(store.unsaved).toBe(false);
  });

  it('reads ranges that take in cells typed past what the file used', async () => {
    let got: SavedWorkbook | undefined;
    await open({ onSave: (f) => void (got = f) });
    await store.input(4, 0, 'Screws');
    await store.input(0, 5, 'Note');
    await store.download();
    expect(got!.cell('Lines', 'A5')).toBe('Screws');
    expect(got!.range('Lines', 'A4:A5')).toEqual([['Washers'], ['Screws']]);
    expect(got!.range('Lines')[0]).toEqual(['Item', 'Qty', 'Price', 'Total', null, 'Note']);
  });

  it('shows the host’s error and keeps the changes', async () => {
    const files = downloads();
    await open({
      onSave: () => {
        throw new Error('The server is down.');
      },
    });
    await store.input(1, 1, '5');
    await store.download();
    expect(toasts()).toEqual(['Not saved: The server is down.']);
    expect(store.unsaved).toBe(true);
    expect(files).toHaveLength(0);
    expect(store.busy).toBe(false);
  });

  it('offers the changes as a copy when someone else saved first', async () => {
    const files = downloads();
    await open({
      onSave: async () => {
        throw new SaveConflictError();
      },
    });
    await store.input(1, 1, '5');
    const saving = store.download();
    const asked = await answer('confirm');
    await saving;
    expect(english.text(asked.title)).toBe('This workbook was changed elsewhere');
    expect(english.text(asked.body)).toMatch(/^Someone else saved a newer version of this workbook\. Your changes are still here/);
    expect(files).toEqual([{ name: 'lines (my changes).xlsx', type: XLSX }]);
    expect(store.unsaved).toBe(true);
    expect(store.toasts).toHaveLength(0);
  });

  it('explains a conflict in the host’s words, and keeps editing when asked to', async () => {
    const files = downloads();
    // A conflict raised by another copy of the class, as across bundles, is known by its name.
    const conflict = Object.assign(new Error('Ana saved it at 10:02.'), { name: 'SaveConflictError' });
    await open({ onSave: () => Promise.reject(conflict) });
    const saving = store.download();
    const asked = await answer('cancel');
    await saving;
    expect(english.text(asked.body)).toMatch(/^Ana saved it at 10:02\. Your changes/);
    expect(files).toHaveLength(0);
  });

  it('labels the button as the host asks, in the editor’s language', async () => {
    await open({ onSave: () => undefined, saveLabel: { en: 'Send to ITAS', de: 'An ITAS senden' }, language: 'de' });
    expect(store.saveLabel).toBe('An ITAS senden');
  });
});

describe('exporting a copy', () => {
  it('downloads without counting as saved when a host handles saving', async () => {
    const files = downloads();
    const onSave = vi.fn();
    await open({ onSave });
    await store.input(1, 1, '5');
    await store.exportCopy();
    expect(files).toEqual([{ name: 'lines.xlsx', type: XLSX }]);
    expect(toasts()).toEqual(['Downloaded lines.xlsx.']);
    expect(onSave).not.toHaveBeenCalled();
    expect(store.unsaved).toBe(true);
  });

  it('counts as saved when the download is the only way to save', async () => {
    downloads();
    await open();
    await store.input(1, 1, '5');
    await store.exportCopy();
    expect(store.unsaved).toBe(false);
  });

  it('does nothing when there is no workbook to save', async () => {
    const files = downloads();
    const onSave = vi.fn();
    store = await testStore({ onSave });
    expect(await store.save()).toBeUndefined();
    await store.exportCopy();
    await store.download();
    expect(files).toHaveLength(0);
    expect(onSave).not.toHaveBeenCalled();
    expect(toasts()).toEqual(Array(3).fill('Could not save: No workbook is open.'));
  });
});

describe('reviewing changes before saving', () => {
  it('lists each change, old and new, and saves only when confirmed', async () => {
    const onSave = vi.fn();
    await open({ onSave, numbers: { reviewBeforeSave: true } });
    await store.input(1, 1, '5');
    // Same value, different format.
    store.noteCopy(rect(1, 4));
    await store.paste({ row: 1, col: 2 }, [['3.00']], rect(1, 2));

    const saving = store.download();
    const asked = await answer('cancel');
    await saving;
    expect(english.text(asked.title)).toBe('Save 2 changed cells?');
    expect(asked.details?.header.map((h) => english.text(h))).toEqual(['Cell', 'Was', 'Now']);
    expect(asked.details?.rows).toEqual([
      ['Lines!B2', '2', '5'],
      ['Lines!C2', '3', '3 (format changed)'],
    ]);
    expect(onSave).not.toHaveBeenCalled();

    const again = store.download();
    await answer('confirm');
    await again;
    expect(onSave).toHaveBeenCalledOnce();
  });

  it('shows the first 300 and counts the rest', async () => {
    downloads();
    await open({ numbers: { reviewBeforeSave: true } });
    await store.paste({ row: 10, col: 0 }, Array.from({ length: 301 }, (_, i) => [String(i + 1)]), rect(10, 0));
    const saving = store.download();
    const asked = await answer('confirm');
    await saving;
    expect(asked.details?.rows).toHaveLength(300);
    expect(asked.details?.more).toBe(1);
  });

  it('does not ask when nothing changed', async () => {
    const onSave = vi.fn();
    await open({ onSave, numbers: { reviewBeforeSave: true } });
    await store.download();
    expect(store.confirmRequest).toBeUndefined();
    expect(onSave).toHaveBeenCalledOnce();
  });
});

describe('schema checks before saving', () => {
  const schema = (more: Partial<WorkbookSchema> = {}): WorkbookSchema => ({
    id: 'lines',
    name: 'Order lines',
    sheets: [
      {
        name: 'Lines',
        outside: 'editable',
        tables: [
          {
            key: 'lines',
            at: 'A1',
            fields: [
              { key: 'item', label: 'Item', type: 'text', required: true },
              { key: 'qty', label: 'Qty', type: 'integer', min: 0 },
              { key: 'price', label: 'Price', type: 'number' },
              { key: 'total', label: 'Total', type: 'number' },
            ],
          },
        ],
      },
    ],
    ...more,
  });

  it('refuses to save while the workbook breaks the schema, showing the issues', async () => {
    const onSave = vi.fn();
    await open({ onSave, schemas: [schema()] });
    expect(store.schema?.schema.id).toBe('lines');
    await store.input(1, 1, '-1');
    store.set('panel', null);
    await store.download();
    expect(onSave).not.toHaveBeenCalled();
    expect(store.panel).toBe('issues');
    expect(toasts().at(-1)).toBe('Not saved. 1 cell breaks the “Order lines” schema.');
  });

  it('asks first when the schema allows saving with errors', async () => {
    const onSave = vi.fn();
    await open({ onSave, schemas: [schema({ saveWithErrors: 'confirm' })] });
    await store.input(1, 1, '-1');
    const saving = store.download();
    const asked = await answer('cancel');
    await saving;
    expect(english.text(asked.title)).toBe('Save with errors?');
    expect(english.text(asked.body)).toBe('1 cell breaks the “Order lines” schema. Save the file anyway?');
    expect(onSave).not.toHaveBeenCalled();
  });

  it('stops the save when the schema’s last check says why', async () => {
    const onSave = vi.fn();
    await open({ onSave, schemas: [schema({ beforeSave: () => 'Approval is missing.' })] });
    await store.download();
    expect(onSave).not.toHaveBeenCalled();
    expect(toasts().at(-1)).toBe('Not saved. Approval is missing.');
  });

  it('hands the schema’s summary and values to the host', async () => {
    let got: SavedWorkbook | undefined;
    await open({ onSave: (f) => void (got = f), schemas: [schema()] });
    await store.download();
    expect(got?.schema).toEqual({ id: 'lines', name: 'Order lines', version: undefined, errors: 0, warnings: 0 });
    expect(got?.values?.Lines.tables.lines[0]).toEqual({ item: 'Bolts', qty: 2, price: 3, total: 6 });
  });
});

describe('opening a workbook', () => {
  it('refuses files that are not Excel workbooks by name', async () => {
    store = await testStore();
    await store.open('notes.csv', new ArrayBuffer(4));
    expect(store.status).toBe('empty');
    expect(toasts()).toEqual(['Only Excel workbooks (.xlsx or .xlsm) can be opened.']);
  });

  it('shows why a file could not be opened', async () => {
    const cases: [Uint8Array, StoreOptions, RegExp][] = [
      [strToU8('hello'), {}, /not an Excel workbook/],
      [await xlsm(), { macros: 'block' }, /contains macros/],
      [bytes, { limits: { maxFileBytes: 1000 } }, /the most this editor accepts is/],
    ];
    for (const [file, options, error] of cases) {
      store = await testStore(options);
      await store.open('book.xlsm', file.slice().buffer as ArrayBuffer);
      expect(store.status).toBe('error');
      expect(english.text(store.error)).toMatch(error);
      expect(store.sheets).toEqual([]);
      store.dispose();
    }
  });

  it('says it is reading the file while it loads', async () => {
    store = await testStore();
    const opening = store.open('lines.xlsx', bytes.slice().buffer as ArrayBuffer);
    expect(store.status).toBe('loading');
    expect(english.text(store.stage)).toBe('Reading lines.xlsx…');
    await opening;
    expect(store.status).toBe('ready');
    expect(store.stage).toBe('');
  });

  it('starts on the first visible sheet', async () => {
    await open({}, await workbook(true));
    expect(store.active).toBe(1);
    expect(store.view?.info.name).toBe('Lines');
  });

  it('asks before discarding unsaved changes for another file', async () => {
    await open();
    await store.input(1, 1, '5');
    const opening = store.open('other.xlsx', sample('budget.xlsx').buffer as ArrayBuffer);
    const asked = await answer('cancel');
    await opening;
    expect(english.text(asked.body)).toBe('You have unsaved changes in lines.xlsx. Opening another file discards them.');
    expect(store.fileName).toBe('lines.xlsx');
    expect(store.view?.get(1, 1)?.n).toBe(5);

    const again = store.open('other.xlsx', sample('budget.xlsx').buffer as ArrayBuffer);
    await answer('confirm');
    await again;
    expect(store.fileName).toBe('other.xlsx');
    expect([store.editedCount, store.undoStack.length, store.history.length]).toEqual([0, 0, 0]);
  });

  it('keeps only the last of two files opened in quick succession', async () => {
    store = await testStore();
    const first = store.open('lines.xlsx', bytes.slice().buffer as ArrayBuffer);
    const second = store.open('budget.xlsx', sample('budget.xlsx').buffer as ArrayBuffer);
    await Promise.all([first, second]);
    expect(store.status).toBe('ready');
    expect(store.fileName).toBe('budget.xlsx');
    expect(store.error).toBeUndefined();
  });
});

describe('live recalculation turned off', () => {
  it('keeps Excel’s values until the user turns it on', async () => {
    await open({ recalc: 'off' });
    expect(store.calc?.off).toBe(true);
    expect(english.text(store.findings.find((f) => f.kind === 'formulaEngine')?.title)).toBe('Live recalculation is off');
    await store.input(1, 1, '5');
    expect(store.view?.get(1, 3)?.n).toBe(6);
    await store.input(5, 0, '=2+3');
    expect(has(store.view?.get(5, 0)?.flags, CellFlag.Unsupported)).toBe(true);
    const overwriting = store.input(2, 3, '7');
    await answer('confirm');
    await overwriting;
    await store.input(5, 1, '=D3*2');

    const starting: boolean[] = [];
    store.subscribe(() => starting.push(store.recalcStarting));
    await store.enableRecalc();
    expect(starting[0]).toBe(true);
    // What was typed meanwhile is calculated, not taken as the engine disagreeing with Excel.
    expect(store.view?.get(1, 3)).toMatchObject({ n: 15, flags: CellFlag.Formula });
    expect(store.calc?.mismatched).toBe(0);
    expect(store.recalcStarting).toBe(false);
    expect(store.calc?.off).toBeUndefined();
    expect(store.view?.get(5, 0)).toMatchObject({ t: VT.Number, n: 5 });
    expect(store.view?.get(5, 1)?.n).toBe(14);
    expect(store.findings.some((f) => english.text(f.title) === 'Live recalculation is off')).toBe(false);
    await store.input(1, 1, '6');
    expect(store.view?.get(1, 3)?.n).toBe(18);
  });

  it('says so when there is no workbook to recalculate', async () => {
    store = await testStore();
    await store.enableRecalc();
    expect(toasts()).toEqual(['No workbook is open.']);
    expect(store.busy).toBe(false);
  });
});

describe('the list of changed cells', () => {
  it('is read again on request, and after each edit while the Changes panel is open', async () => {
    await open();
    await store.input(1, 1, '5');
    expect(store.changeList).toEqual([]);
    await store.refreshChanges();
    expect(store.changeList).toEqual([{ sheet: 'Lines', ref: 'B2', before: '2', after: '5' }]);
    store.set('panel', 'changes');
    await store.input(2, 1, '1');
    await vi.waitFor(() => expect(store.changeList.map((c) => c.ref)).toEqual(['B2', 'B3']));
  });

  it('is left as it was without a workbook', async () => {
    store = await testStore();
    await store.refreshChanges();
    expect(store.changeList).toEqual([]);
  });
});
