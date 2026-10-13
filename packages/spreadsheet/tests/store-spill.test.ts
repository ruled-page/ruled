import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { parseA1, rangeToA1 } from '@ruled/formula/refs';
import { CellFlag, VT } from '../src/core/types';
import { english } from '../src/i18n';
import { editText } from '../src/state/format';
import { workbookReaders } from '../src/state/readers';
import type { WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

const templates = () => new Uint8Array(readFileSync(join(__dirname, 'corpus', 'templates.xlsx')));

let store: WorkbookStore;
afterEach(() => store?.dispose());

async function open(): Promise<WorkbookStore> {
  store = await testStore();
  await opened(store, templates(), 'templates.xlsx');
  store.setActive(store.sheets.findIndex((v) => v.info.name === 'Arrays'));
  return store;
}

const cell = (ref: string) => store.view!.get(parseA1(ref)!.row, parseA1(ref)!.col);
const select = (ref: string) => {
  const { row, col } = parseA1(ref)!;
  store.setSelection({ row, col, rects: [{ r1: row, c1: col, r2: row, c2: col }] });
};

describe("the editor's view of dynamic arrays", () => {
  it("shows a file's spills as calculated here, unlocked, and keeps arrays it can't calculate locked", async () => {
    await open();
    const view = store.view!;
    expect(cell('A3')).toMatchObject({ t: VT.String, s: 'Hardware', flags: CellFlag.Spilled });
    expect(rangeToA1(cell('A3')!.spill!)).toBe('A2:A4');
    expect(rangeToA1(cell('A2')!.spill!)).toBe('A2:A4');
    expect(view.readOnlyReason(2, 0)).toBeUndefined();
    // B2:B4 reads the spill as A2#, and spills as it does.
    expect(view.readOnlyReason(3, 1)).toBeUndefined();
    expect(english.text(view.readOnlyReason(6, 6)!)).toMatch(/array formula/);
    // A spilled cell holds nothing to edit.
    expect(editText(cell('A3'), store.styles[0], false)).toBe('');
  });

  it('spills a typed formula, and gives its values to the host as values', async () => {
    await open();
    select('K1');
    await store.input(0, 10, '=SEQUENCE(2,2)*10');
    expect(cell('L2')).toMatchObject({ t: VT.Number, n: 40, flags: CellFlag.Spilled });
    expect(workbookReaders(store.sheets, store.styles, store.date1904).range('Arrays', 'K1:L2')).toEqual([
      [10, 20],
      [30, 40],
    ]);
    // Typing into one of its cells blocks it, and undoing that spills it again.
    await store.input(1, 10, 'x');
    expect(cell('K1')).toMatchObject({ t: VT.Error, s: '#SPILL!' });
    expect(cell('L2')).toMatchObject({ t: VT.None, flags: 0 });
    await store.undo();
    expect(cell('L2')).toMatchObject({ t: VT.Number, n: 40, flags: CellFlag.Spilled });
  });

  it('copies spilled cells as their values', async () => {
    await open();
    store.setSelection({ row: 2, col: 0, rects: [{ r1: 2, c1: 0, r2: 3, c2: 0 }] });
    expect(store.selectionText({ r1: 2, c1: 0, r2: 3, c2: 0 })).toBe('Hardware\nSafety');
    store.noteCopy({ r1: 2, c1: 0, r2: 3, c2: 0 });
    select('M1');
    await store.pasteFromClipboard();
    expect(cell('M1')).toMatchObject({ t: VT.String, s: 'Hardware', flags: CellFlag.Edited });
  });
});

describe('schema checks on spilled cells', () => {
  it('read their values, which come from a formula', async () => {
    store = await testStore({
      schemas: [
        {
          id: 'arrays',
          name: 'Arrays',
          sheets: [
            {
              name: 'Arrays',
              cells: [
                { key: 'second', ref: 'A3', type: 'text', required: true, formula: 'require' },
                { key: 'cell', ref: 'H3', type: 'integer', max: 4 },
              ],
            },
          ],
        },
      ],
    });
    await opened(store, templates(), 'templates.xlsx');
    const arrays = store.sheets.findIndex((v) => v.info.name === 'Arrays');
    expect(store.schema!.cellIssues(arrays, 2, 0)).toEqual([]);
    expect(store.schema!.cellIssues(arrays, 2, 7).map((i) => english.text(i.message))).toEqual(['Must be at most 4.']);
  });
});
