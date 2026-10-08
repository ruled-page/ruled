import ExcelJS from 'exceljs';
import { strToU8 } from 'fflate';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { VT } from '../src/core/types';
import { english } from '../src/i18n';
import type { StoreOptions, WorkbookStore } from '../src/state/store';
import { loadedValue } from './helpers';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/** Text as one byte a character, as Windows-1252 writes the characters it shares with Latin-1. */
const latin1 = (s: string) => Uint8Array.from(s, (c) => c.charCodeAt(0));

/** A price list as a German Excel saves it: Windows-1252, semicolons, decimal commas, day-first dates. */
const GERMAN = latin1('Artikel;Preis;Gültig ab;Notiz\r\nSchraube;1,50;31.12.2025;"=HYPERLINK(""http://example.com"")"\r\nMutter;0,25;01.02.2026;\r\n');

let store: WorkbookStore;
afterEach(() => {
  store?.dispose();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

async function open(bytes: Uint8Array, name: string, options: StoreOptions = {}): Promise<WorkbookStore> {
  store = await testStore(options);
  await opened(store, bytes, name);
  return store;
}

const cell = (row: number, col: number) => store.sheets[store.active].get(row, col);
const shown = (row: number, col: number) => store.shownText(store.sheets[store.active], row, col);
const toasts = () => store.toasts.map((t) => english.text(t.text));

describe('opening a CSV file in the editor', () => {
  it('shows it as one sheet, read in the locale, with formula-like text kept as text', async () => {
    await open(GERMAN, 'preise.csv', { locale: 'de-DE' });
    expect(store.status).toBe('ready');
    expect(store.fileName).toBe('preise.csv');
    expect(store.sheets.map((s) => s.info.name)).toEqual(['preise']);
    expect(shown(0, 2)).toBe('Gültig ab');
    expect(cell(1, 1)).toMatchObject({ t: VT.Number, n: 1.5 });
    expect(shown(1, 2)).toBe('31/12/2025');
    expect(cell(1, 3)).toMatchObject({ t: VT.String, s: '=HYPERLINK("http://example.com")' });
    expect(cell(1, 3)?.f).toBeUndefined();
    expect(store.calc?.formulaCells).toBe(0);
  });

  it('edits it like any workbook and saves it as .xlsx, named after the CSV file', async () => {
    await open(strToU8('Item,Qty\nBolts,2\n'), 'order.csv');
    await store.input(1, 1, '5');
    await store.input(2, 0, 'Nuts');
    await store.input(2, 1, '=B2*2');
    const saved = await store.snapshot();
    expect(saved?.fileName).toBe('order.xlsx');
    expect(saved?.editedCount).toBe(3);
    const reopened = WorkbookSession.open(saved!.bytes, 'order.xlsx');
    expect(['A1', 'B2', 'A3', 'B3'].map((ref) => loadedValue(reopened, 'order', ref))).toEqual(['Item', 5, 'Nuts', 10]);
  });

  it('says why a CSV file was refused', async () => {
    await open(strToU8('a,b,c\n1,2,3\n'), 'big.csv', { limits: { maxCells: 4 } });
    expect(store.status).toBe('error');
    expect(english.text(store.error)).toBe('This CSV file has more than 4 values, the most this editor can hold.');
    expect(toasts()).toEqual([]);
  });
});

/**
 * Lines: values of each kind with their formats, text a spreadsheet app
 * would run as formulas, and a gap before the last column; Other is a second
 * sheet.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const lines = wb.addWorksheet('Lines');
  lines.addRow(['Item', 'Price', 'Share', 'Due', 'Done', 'Total', 'Note']);
  lines.addRow(['Bolts, M6', 1234.5, 0.125, new Date(Date.UTC(2026, 2, 4)), true, { formula: 'B2*2', result: 2469 }, '=HYPERLINK("http://example.com")']);
  lines.addRow(['Say "hi"', -5, null, null, false, { formula: 'B3*2', result: -10 }, '-2+3+cmd|calc']);
  lines.getCell('A4').value = 'two\nlines';
  lines.getCell('I4').value = '@SUM(1)';
  for (const ref of ['B2', 'B3', 'F2', 'F3']) lines.getCell(ref).numFmt = '#,##0.00';
  lines.getCell('C2').numFmt = '0.0%';
  lines.getCell('D2').numFmt = 'yyyy-mm-dd';
  wb.addWorksheet('Other').getCell('A1').value = 'x';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let lines: Uint8Array;
beforeAll(async () => {
  lines = await workbook();
});

/** Stands in for the browser's download link, recording each file it would download, with its bytes. */
function downloads(): { name: string; type: string; bytes: Promise<Uint8Array> }[] {
  const files: { name: string; type: string; bytes: Promise<Uint8Array> }[] = [];
  const blobs = new Map<string, Blob>();
  vi.spyOn(URL, 'createObjectURL').mockImplementation((blob) => {
    const url = `blob:${blobs.size}`;
    blobs.set(url, blob as Blob);
    return url;
  });
  vi.spyOn(URL, 'revokeObjectURL').mockImplementation(() => undefined);
  const link = {
    href: '',
    download: '',
    click: () => {
      const blob = blobs.get(link.href)!;
      files.push({ name: link.download, type: blob.type, bytes: blob.arrayBuffer().then((b) => new Uint8Array(b)) });
    },
  };
  vi.stubGlobal('document', { createElement: () => link });
  return files;
}

const csvLines = async (options = {}) => (await store.csv(options))?.text.split('\r\n');

describe('the sheet shown, as CSV', () => {
  it('has each value as shown, formula results, escaped formulas, quotes where needed and CRLF line ends', async () => {
    await open(lines, 'lines.xlsx');
    expect(await csvLines()).toEqual([
      'Item,Price,Share,Due,Done,Total,Note,,',
      `"Bolts, M6","1,234.50",12.5%,2026-03-04,TRUE,"2,469.00","'=HYPERLINK(""http://example.com"")",,`,
      `"Say ""hi""",-5.00,,,FALSE,-10.00,'-2+3+cmd|calc,,`,
      `"two\nlines",,,,,,,,'@SUM(1)`,
      '',
    ]);
  });

  it('downloads it as UTF-8 with a byte order mark, named after the workbook and the sheet, without counting as a save', async () => {
    const files = downloads();
    await open(lines, 'lines.xlsx');
    await store.input(0, 0, 'Thing');
    await store.exportCsv();
    expect(files.map((f) => [f.name, f.type])).toEqual([['lines - Lines.csv', 'text/csv;charset=utf-8']]);
    const bytes = await files[0].bytes;
    expect(Array.from(bytes.slice(0, 3))).toEqual([0xef, 0xbb, 0xbf]);
    expect(new TextDecoder().decode(bytes.slice(3)).split('\r\n')[0]).toBe('Thing,Price,Share,Due,Done,Total,Note,,');
    expect(toasts()).toEqual(['Downloaded lines - Lines.csv: the values on Lines, as shown. A CSV file holds one sheet, without formulas or formatting.']);
    expect(store.unsaved).toBe(true);
  });

  it('names the file after the workbook alone when it has one sheet', async () => {
    const files = downloads();
    await open(strToU8('a,b\n'), 'order.csv');
    await store.exportCsv();
    expect(files.map((f) => f.name)).toEqual(['order.csv']);
  });

  it("separates values with the locale's list separator and writes its decimal comma, as Excel does there", async () => {
    await open(lines, 'lines.xlsx', { locale: 'de-DE' });
    expect((await csvLines())?.slice(0, 2)).toEqual([
      'Item;Price;Share;Due;Done;Total;Note;;',
      `Bolts, M6;1.234,50;12,5%;2026-03-04;TRUE;2.469,00;"'=HYPERLINK(""http://example.com"")";;`,
    ]);
  });

  it('takes another delimiter, and leaves formulas as they are when told to, through the options or the call', async () => {
    await open(lines, 'lines.xlsx', { csvExport: { escapeFormulas: false } });
    expect((await csvLines())?.[2]).toBe('"Say ""hi""",-5.00,,,FALSE,-10.00,-2+3+cmd|calc,,');
    expect((await csvLines({ escapeFormulas: true, delimiter: '\t' }))?.[2]).toBe(`"Say ""hi"""\t-5.00\t\t\tFALSE\t-10.00\t'-2+3+cmd|calc\t\t`);
  });

  it('has edits and the results they recalculate', async () => {
    await open(lines, 'lines.xlsx');
    await store.input(2, 1, '7');
    expect((await csvLines())?.[2]).toBe(`"Say ""hi""",7.00,,,FALSE,14.00,'-2+3+cmd|calc,,`);
  });

  it('is nothing without a workbook, and empty for an empty sheet', async () => {
    const files = downloads();
    store = await testStore();
    expect(await store.csv()).toBeUndefined();
    await store.exportCsv();
    expect(files).toEqual([]);
    await opened(store, strToU8(''), 'empty.csv');
    expect((await store.csv())?.text).toBe('');
  });
});
