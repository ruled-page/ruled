import { strToU8 } from 'fflate';
import { afterEach, describe, expect, it, vi } from 'vitest';
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
