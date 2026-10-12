import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import type { Answer, FindQuery, WorkbookStore } from '../src/state/store';
import { sample } from './helpers';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/**
 * Fruit!A1:D5 = Name, Qty, Price, Total / Pear 3 2 =B2*C2 / apple 1 5 =B3*C3 / Fig 2 1 =B4*C4 / Banana 10 0.5 =B5*C5,
 * F1 = SUM(D2:D5), and A8 = "Lone" on its own.
 * Table!A1:B4 is table "Items" (Item, Qty) with a totals row; Filtered!A1:B4 has the file's own filter.
 * Locked is protected.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const fruit = wb.addWorksheet('Fruit');
  fruit.addRow(['Name', 'Qty', 'Price', 'Total']);
  const rows: [string, number, number][] = [
    ['Pear', 3, 2],
    ['apple', 1, 5],
    ['Fig', 2, 1],
    ['Banana', 10, 0.5],
  ];
  rows.forEach(([name, qty, price], i) => fruit.addRow([name, qty, price, { formula: `B${i + 2}*C${i + 2}`, result: qty * price }]));
  fruit.getCell('F1').value = { formula: 'SUM(D2:D5)', result: 18 };
  fruit.getCell('A8').value = 'Lone';
  const table = wb.addWorksheet('Table');
  table.addTable({
    name: 'Items',
    ref: 'A1',
    headerRow: true,
    totalsRow: true,
    columns: [{ name: 'Item', totalsRowLabel: 'Total' }, { name: 'Qty', totalsRowFunction: 'sum' }],
    rows: [
      ['b', 2],
      ['a', 1],
    ],
  });
  const filtered = wb.addWorksheet('Filtered');
  filtered.addRow(['Region', 'Sales']);
  filtered.addRow(['North', 1]);
  filtered.addRow(['South', 2]);
  filtered.addRow(['North', 3]);
  filtered.autoFilter = 'A1:B4';
  const locked = wb.addWorksheet('Locked');
  locked.getCell('A1').value = 'Pear';
  await locked.protect('', {});
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
  await opened(store, bytes, 'fruit.xlsx');
  return store;
}

const cell = (row: number, col: number, sheet = store.active) => store.sheets[sheet].get(row, col);
const column = (col: number, r1: number, r2: number) => Array.from({ length: r2 - r1 + 1 }, (_, i) => cell(r1 + i, col)?.s ?? cell(r1 + i, col)?.n);
const toasts = () => store.toasts.map((t) => english.text(t.text));
const actions = () => store.history.map((h) => english.text(h.action));
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const select = (r: ReturnType<typeof rect>) => store.setSelection({ row: r.r1, col: r.c1, rects: [r] });

describe('inserting and deleting rows and columns', () => {
  it('inserts as many rows as are selected, above or below, moving formulas with them', async () => {
    await open();
    select(rect(2, 1, 3, 1));
    expect(await store.insertRows('above')).toBe(true);
    expect(cell(4, 0)?.s).toBe('apple');
    expect(cell(4, 3)?.f).toBe('B5*C5');
    expect(cell(0, 5)?.f).toBe('SUM(D2:D7)');
    // The new rows are selected whole, as Excel selects them, keeping the active column.
    expect(store.selection).toEqual({ row: 2, col: 1, rects: [rect(2, 0, 3, 5)], lines: 'rows' });

    select(rect(1, 0));
    await store.insertRows('below');
    expect(cell(2, 0)).toBeFalsy();
    expect(cell(3, 0)).toBeFalsy();
    expect(actions()).toEqual(['Inserted 2 rows', 'Inserted 1 row']);
    expect(store.editedCount).toBe(2);
  });

  it('inserts columns left or right of the selection', async () => {
    await open();
    select(rect(0, 1, 0, 2));
    await store.insertCols('left');
    expect(cell(0, 3)?.s).toBe('Qty');
    expect(cell(1, 5)?.f).toBe('D2*E2');
    expect(store.selection.rects).toEqual([rect(0, 1, 7, 2)]);
    select(rect(0, 0));
    await store.insertCols('right');
    expect(cell(0, 2)).toBeFalsy();
    expect(actions()).toEqual(['Inserted 2 columns', 'Inserted 1 column']);
  });

  it('deletes rows and columns, selecting where they were', async () => {
    await open();
    store.setSelection({ row: 2, col: 3, rects: [rect(2, 3, 3, 3)] });
    await store.deleteRows();
    expect(column(0, 1, 3)).toEqual(['Pear', 'Banana', undefined]);
    expect(cell(0, 5)?.f).toBe('SUM(D2:D3)');
    expect(store.selection).toEqual({ row: 2, col: 3, rects: [rect(2, 3)] });

    store.setSelection({ row: 1, col: 2, rects: [rect(1, 2)] });
    await store.deleteCols();
    expect(cell(1, 2)?.f).toBe('B2*#REF!');
    expect(store.selection).toEqual({ row: 1, col: 2, rects: [rect(1, 2)] });
    expect(actions()).toEqual(['Deleted 2 rows', 'Deleted 1 column']);
    expect(toasts()).toEqual(['2 formulas now show #REF! because they used cells that were deleted or pasted over.']);
  });

  it('turns down a change the sheet does not allow, leaving nothing to undo', async () => {
    await open();
    store.setActive(3);
    expect(await store.insertRows('above')).toBe(false);
    expect(toasts()).toEqual(['This sheet is protected in Excel, so it is read-only here.']);
    expect(store.undoStack).toHaveLength(0);
    expect(store.history).toHaveLength(0);
  });

  it('does nothing before a workbook is open', async () => {
    store = await testStore();
    expect(await store.insertRows('above')).toBe(false);
  });

  it('refreshes the list of changes when it is showing', async () => {
    await open();
    await store.input(1, 1, '4');
    store.set('panel', 'changes');
    select(rect(0, 0));
    await store.insertRows('above');
    await vi.waitFor(() => expect(store.changeList.map((c) => c.ref)).toContain('B3'));
  });
});

describe('undoing and redoing structural changes', () => {
  it('goes back to the workbook as it was, and forward again', async () => {
    await open();
    await store.input(1, 1, '4');
    select(rect(1, 0, 2, 0));
    await store.insertRows('above');
    // The edited cell and the insertion.
    expect(store.editedCount).toBe(2);

    await store.undo();
    expect(cell(1, 0)?.s).toBe('Pear');
    expect(cell(1, 1)?.n).toBe(4);
    expect(cell(0, 5)?.f).toBe('SUM(D2:D5)');
    expect(store.editedCount).toBe(1);

    await store.redo();
    expect(cell(3, 0)?.s).toBe('Pear');
    expect(cell(0, 5)?.f).toBe('SUM(D4:D7)');

    await store.undo();
    await store.undo();
    expect(cell(1, 1)?.n).toBe(3);
    expect(store.editedCount).toBe(0);
    expect(actions()).toEqual(['Edit', 'Inserted 2 rows', 'Undo: Inserted 2 rows', 'Redo: Inserted 2 rows', 'Undo: Inserted 2 rows', 'Undo']);
  });

  it('refreshes the list of changes when it is showing', async () => {
    await open();
    select(rect(1, 0));
    await store.insertRows('above');
    store.set('panel', 'changes');
    await store.undo();
    await vi.waitFor(() => expect(store.changeList).toEqual([]));
    await store.redo();
    await vi.waitFor(() => expect(store.changeList).toHaveLength(1));
  });

  it('brings the engine back after undoing, so formulas recalculate', async () => {
    await open();
    select(rect(1, 0));
    await store.deleteRows();
    await store.undo();
    await store.input(1, 1, '10');
    expect(cell(1, 3)?.n).toBe(20);
    expect(cell(0, 5)?.n).toBe(32);
  });

  it('recalculates a total that used a deleted row', async () => {
    await open();
    select(rect(1, 0));
    await store.deleteRows();
    await vi.waitFor(() => expect(store.calc).toBeDefined());
    expect(cell(0, 5)).toMatchObject({ f: 'SUM(D2:D4)', n: 12 });
  });

  it('gives up on older steps once the workbook can no longer go back to them', async () => {
    await open();
    select(rect(20, 0));
    for (let i = 0; i < 21; i++) await store.insertRows('above');
    expect(store.undoStack).toHaveLength(21);
    for (let i = 0; i < 20; i++) await store.undo();
    expect(store.undoStack).toHaveLength(1);
    await store.undo();
    expect(toasts().at(-1)).toBe("That change can't be undone any more.");
    expect(store.undoStack).toHaveLength(0);
    expect(store.redoStack).toHaveLength(20);
  });
});

describe('sorting', () => {
  it('sorts the block around the active cell by its column, keeping the headings', async () => {
    await open();
    store.setSelection({ row: 2, col: 1, rects: [rect(2, 1)] });
    await store.sort(false);
    expect(column(0, 0, 4)).toEqual(['Name', 'apple', 'Fig', 'Pear', 'Banana']);
    // Formulas move with their rows and still refer to their own row.
    expect(cell(4, 3)).toMatchObject({ f: 'B5*C5', n: 5 });
    await store.sort(true, 0);
    expect(column(0, 1, 4)).toEqual(['Pear', 'Fig', 'Banana', 'apple']);
    expect(actions()).toEqual(['Sort A to Z', 'Sort Z to A']);
    await store.undo();
    expect(column(0, 1, 4)).toEqual(['apple', 'Fig', 'Pear', 'Banana']);
  });

  it('sorts a selected range from its first row, unless that row looks like headings', async () => {
    await open();
    // Data only: every row is sorted.
    select(rect(1, 0, 4, 3));
    await store.sort(false);
    expect(column(0, 1, 4)).toEqual(['apple', 'Banana', 'Fig', 'Pear']);
    expect(cell(1, 3)).toMatchObject({ f: 'B2*C2', n: 5 });
    await store.undo();
    // Text over a column of numbers is taken as headings, as Excel takes it.
    select(rect(0, 0, 4, 3));
    await store.sort(false);
    expect(column(0, 0, 4)).toEqual(['Name', 'apple', 'Banana', 'Fig', 'Pear']);
    await store.undo();
    // Text over text, all plain, is data.
    select(rect(0, 0, 4, 0));
    await store.sort(false);
    expect(column(0, 0, 4)).toEqual(['apple', 'Banana', 'Fig', 'Name', 'Pear']);
  });

  it('says so when there is nothing around the active cell to sort', async () => {
    await open();
    store.setSelection({ row: 10, col: 10, rects: [rect(10, 10)] });
    expect(await store.sort(false)).toBeUndefined();
    store.setSelection({ row: 7, col: 0, rects: [rect(7, 0)] });
    await store.sort(false);
    expect(toasts()).toEqual(['Select a cell in a block of data to sort it.', 'Select a cell in a block of data to sort it.']);
    expect(store.history).toHaveLength(0);
  });

  it('sorts a table without its totals row', async () => {
    await open();
    store.setActive(1);
    store.setSelection({ row: 1, col: 0, rects: [rect(1, 0)] });
    expect(store.dataRegion()).toEqual(rect(0, 0, 2, 1));
    await store.sort(false);
    expect(column(0, 0, 3)).toEqual(['Item', 'a', 'b', 'Total']);
  });
});

describe('the block a sort or filter applies to', () => {
  it('is the selection when more than one cell is selected', async () => {
    await open();
    select(rect(1, 0, 2, 1));
    expect(store.dataRegion()).toEqual(rect(1, 0, 2, 1));
  });

  it('grows from the active cell while any neighbouring cell, diagonals included, is filled', async () => {
    await open();
    await store.input(5, 4, 'touching corner');
    store.setSelection({ row: 2, col: 1, rects: [rect(2, 1)] });
    expect(store.dataRegion()).toEqual(rect(0, 0, 5, 5));
    expect(store.dataRegion(0, 7, 0)).toEqual(rect(7, 0));
    expect(store.dataRegion(9)).toBeUndefined();
  });

  it("is the file's own filter range around the active cell", async () => {
    await open();
    store.setActive(2);
    expect(store.dataRegion(2, 2, 1)).toEqual(rect(0, 0, 3, 1));
  });
});

describe('filtering', () => {
  it('lists the values in a column with how often each appears, blanks last', async () => {
    await open();
    await store.input(2, 1, '');
    await store.input(3, 1, '10');
    select(rect(1, 1));
    expect(store.filterChoices(1)).toEqual({
      rect: rect(0, 0, 4, 3),
      values: [
        { text: '3', count: 1 },
        { text: '10', count: 2 },
        { text: '', count: 1 },
      ],
      allowed: undefined,
    });
    expect(store.filterChoices(9)).toBeUndefined();
  });

  it('hides rows whose values are not allowed, column by column, and shows them all again', async () => {
    await open();
    store.setActive(2);
    select(rect(1, 0));
    const choices = store.filterChoices(0)!;
    expect(choices.values).toEqual([
      { text: 'North', count: 2 },
      { text: 'South', count: 1 },
    ]);
    await store.setFilter(choices.rect, 0, new Set(['North']));
    const view = store.view!;
    expect(view.filteredCount).toBe(1);
    expect(view.gridRow(2)).toBe(-1);
    expect(store.filterChoices(0)?.allowed).toEqual(new Set(['North']));

    await store.setFilter(choices.rect, 1, new Set(['3']));
    expect(view.filteredCount).toBe(2);
    await store.setFilter(choices.rect, 0, undefined);
    expect(view.filteredCount).toBe(2);
    expect(view.filter?.columns.has(0)).toBe(false);

    // Copying leaves out the rows the filter hides.
    expect(store.selectionText(choices.rect)).toBe('Region\tSales\nNorth\t3');

    await store.clearFilter();
    expect(view.filter).toBeUndefined();
    expect(view.filteredCount).toBe(0);
    const version = store.version;
    await store.clearFilter();
    expect(store.version).toBe(version);
  });
});

describe('find and replace', () => {
  const query = (text: string, more: Partial<FindQuery> = {}): FindQuery => ({ text, matchCase: false, wholeCell: false, lookIn: 'values', scope: 'sheet', ...more });

  it('finds matching cells in reading order on this sheet or in the whole workbook', async () => {
    await open();
    expect(store.find(query('a'))).toEqual([
      { sheet: 0, row: 0, col: 0 },
      { sheet: 0, row: 0, col: 3 },
      { sheet: 0, row: 1, col: 0 },
      { sheet: 0, row: 2, col: 0 },
      { sheet: 0, row: 4, col: 0 },
    ]);
    expect(store.find(query('pear', { scope: 'workbook' })).map((h) => h.sheet)).toEqual([0, 3]);
    expect(store.find(query(''))).toEqual([]);
  });

  it('matches case or the whole cell when asked to', async () => {
    await open();
    expect(store.find(query('A', { matchCase: true }))).toEqual([]);
    expect(store.find(query('apple', { matchCase: true })).length).toBe(1);
    expect(store.find(query('pea', { wholeCell: true }))).toEqual([]);
    expect(store.find(query('PEAR', { wholeCell: true }))).toEqual([{ sheet: 0, row: 1, col: 0 }]);
  });

  it('searches what cells show, or the formulas they hold', async () => {
    await open();
    expect(store.find(query('B2'))).toEqual([]);
    expect(store.find(query('B2', { lookIn: 'formulas' }))).toEqual([{ sheet: 0, row: 1, col: 3 }]);
    expect(store.find(query('18'))).toEqual([{ sheet: 0, row: 0, col: 5 }]);
  });

  it('skips sheets that are not worksheets', async () => {
    store = await testStore();
    await opened(store, sample('kitchen-sink.xlsx'), 'kitchen-sink.xlsx');
    const chart = store.sheets.findIndex((s) => s.info.kind !== 'worksheet');
    expect(chart).toBeGreaterThanOrEqual(0);
    const hits = store.find(query('e', { scope: 'workbook' }));
    expect(hits.length).toBeGreaterThan(0);
    expect(hits.some((h) => h.sheet === chart)).toBe(false);
  });

  it('replaces text in every matching cell as one step that undoes at once', async () => {
    await open();
    expect(await store.replace(query('a'), 'o')).toBe(5);
    expect(column(0, 0, 4)).toEqual(['Nome', 'Peor', 'opple', 'Fig', 'Bonono']);
    expect(cell(0, 3)?.s).toBe('Totol');
    expect(actions()).toEqual(['Replace']);
    await store.undo();
    expect(column(0, 0, 4)).toEqual(['Name', 'Pear', 'apple', 'Fig', 'Banana']);
  });

  it('replaces inside formulas, only matching case when asked, with ~ before a wildcard meaning it literally', async () => {
    await open();
    expect(await store.replace(query('b2~*', { matchCase: true }), 'B3*')).toBe(0);
    expect(await store.replace(query('b2~*'), 'B3*')).toBe(1);
    expect(cell(1, 3)).toMatchObject({ f: 'B3*C2', n: 2 });
  });

  it("finds and replaces with Excel's wildcards: * for any run and ? for any one character", async () => {
    await open();
    expect(store.find(query('?ear', { wholeCell: true })).map((h) => h.row)).toEqual([1]);
    expect(store.find(query('b*n*a')).map((h) => h.row)).toEqual([4]);
    expect(store.find(query('p*r', { wholeCell: true, matchCase: true }))).toEqual([]);
    // Each match goes, left to right; a trailing star takes the rest of the cell.
    expect(await store.replace(query('a?'), '-', { sheet: 0, row: 4, col: 0 })).toBe(1);
    expect(cell(4, 0)?.s).toBe('B--a');
    expect(await store.replace(query('i*'), 'ig tree', { sheet: 0, row: 3, col: 0 })).toBe(1);
    expect(cell(3, 0)?.s).toBe('Fig tree');
  });

  it('replaces the whole cell, or just one cell', async () => {
    await open();
    expect(await store.replace(query('fig', { wholeCell: true }), 'Kiwi')).toBe(1);
    expect(cell(3, 0)?.s).toBe('Kiwi');
    expect(await store.replace(query('a'), 'A', { sheet: 0, row: 4, col: 0 })).toBe(1);
    expect(cell(4, 0)?.s).toBe('BAnAnA');
    expect(cell(1, 0)?.s).toBe('Pear');
  });

  it('leaves read-only cells and says how many', async () => {
    await open();
    expect(await store.replace(query('pear', { scope: 'workbook' }), 'Quince')).toBe(1);
    expect(toasts()).toEqual(['1 matching cell is read-only here and was left as it was.']);
    expect(cell(0, 0, 3)?.s).toBe('Pear');
  });

  it('changes cells on several sheets in one step', async () => {
    await open();
    expect(await store.replace(query('e', { scope: 'workbook', wholeCell: true }), 'x')).toBe(0);
    expect(await store.replace(query('north', { scope: 'workbook' }), 'East')).toBe(2);
    expect([cell(1, 0, 2)?.s, cell(3, 0, 2)?.s]).toEqual(['East', 'East']);
    // "Name" already reads that way, so it is left alone.
    expect(await store.replace(query('n', { scope: 'workbook' }), 'N')).toBe(3);
    expect([cell(0, 0, 0)?.s, cell(4, 0, 0)?.s, cell(7, 0, 0)?.s, cell(0, 0, 2)?.s]).toEqual(['Name', 'BaNaNa', 'LoNe', 'RegioN']);
    expect(store.undoStack).toHaveLength(2);
    await store.undo();
    expect([cell(4, 0, 0)?.s, cell(0, 0, 2)?.s, cell(1, 0, 2)?.s]).toEqual(['Banana', 'Region', 'East']);
  });
});

describe("dragging the selection's border", () => {
  /** Waits for the store to ask something, and gives the answer. */
  async function answer(reply: Answer): Promise<void> {
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    if (!store.confirmRequest) throw new Error('Nothing was asked');
    store.confirmRequest.resolve(reply);
  }

  it('moves the cells, with references to them following, and selects where they went', async () => {
    await open();
    select(rect(1, 1));
    expect(await store.dragSelection(1, 6, false)).toBe(true);
    expect([cell(1, 1), cell(1, 6)?.n]).toEqual([null, 3]);
    expect([cell(1, 3)?.f, cell(1, 3)?.n]).toEqual(['G2*C2', 6]);
    expect(store.selection.rects).toEqual([rect(1, 6)]);
  });

  it('copies them with Ctrl, their formulas shifting as when pasted', async () => {
    await open();
    select(rect(1, 3));
    expect(await store.dragSelection(1, 4, true)).toBe(true);
    expect([cell(1, 3)?.f, cell(1, 4)?.f, cell(1, 4)?.n]).toEqual(['B2*C2', 'C2*D2', 12]);
    expect(store.selection.rects).toEqual([rect(1, 4)]);
  });

  it('asks before replacing cells that hold something, as Excel does', async () => {
    await open();
    select(rect(1, 1));
    const kept = store.dragSelection(1, 2, false);
    await answer('cancel');
    expect(await kept).toBe(false);
    expect([cell(1, 1)?.n, cell(1, 2)?.n]).toEqual([3, 2]);
    const replaced = store.dragSelection(1, 2, false);
    await answer('confirm');
    expect(await replaced).toBe(true);
    expect([cell(1, 1), cell(1, 2)?.n]).toEqual([null, 3]);
  });

  it("doesn't ask when the only cells in the way are the ones moving", async () => {
    await open();
    select(rect(6, 0, 7, 0));
    expect(await store.dragSelection(7, 0, false)).toBe(true);
    expect(store.confirmRequest).toBeUndefined();
    expect([cell(7, 0), cell(8, 0)?.s]).toEqual([null, 'Lone']);
  });

  it('turns down a drop off the sheet, or a selection of several areas', async () => {
    await open();
    select(rect(1, 1));
    expect(await store.dragSelection(-1, 1, false)).toBe(false);
    store.setSelection({ row: 1, col: 1, rects: [rect(1, 1), rect(3, 1)] });
    expect(await store.dragSelection(1, 6, false)).toBe(false);
    expect(cell(1, 1)?.n).toBe(3);
  });
});

describe('inserting and deleting cells', () => {
  it('inserts cells, moving those below down with references following, and selects them', async () => {
    await open();
    select(rect(1, 1));
    expect(await store.insertCells('down')).toBe(true);
    expect(column(1, 1, 5)).toEqual([undefined, 3, 1, 2, 10]);
    expect([cell(1, 3)?.f, cell(1, 3)?.n]).toEqual(['B3*C2', 6]);
    expect(store.selection.rects).toEqual([rect(1, 1)]);
    expect(actions()).toEqual(['Inserted cells B2']);
  });

  it('deletes cells, closing the gap, as one step that undoes at once', async () => {
    await open();
    select(rect(1, 0, 2, 0));
    expect(await store.deleteCells('up')).toBe(true);
    expect(column(0, 0, 4)).toEqual(['Name', 'Fig', 'Banana', undefined, undefined]);
    expect(actions()).toEqual(['Deleted cells A2:A3']);
    await store.undo();
    expect(column(0, 0, 4)).toEqual(['Name', 'Pear', 'apple', 'Fig', 'Banana']);
  });

  it('inserts or deletes whole rows or columns when asked to', async () => {
    await open();
    const each = async (step: () => Promise<boolean>) => {
      select(rect(1, 0));
      expect(await step()).toBe(true);
    };
    await each(() => store.insertCells('rows'));
    await each(() => store.deleteCells('cols'));
    await each(() => store.insertCells('cols'));
    await each(() => store.deleteCells('rows'));
    expect(actions()).toEqual(['Inserted 1 row', 'Deleted 1 column', 'Inserted 1 column', 'Deleted 1 row']);
    expect(column(1, 0, 1)).toEqual(['Qty', 3]);
  });

  it('inserts or deletes selected rows straight away from the keyboard, and asks about cells', async () => {
    await open();
    store.selectLines('rows', [[1, 1]], { row: 1, col: 0 });
    store.changeCells('insert');
    await vi.waitFor(() => expect(actions()).toEqual(['Inserted 1 row']));
    store.selectLines('cols', [[0, 0]], { row: 0, col: 0 });
    store.changeCells('delete');
    await vi.waitFor(() => expect(actions()).toEqual(['Inserted 1 row', 'Deleted 1 column']));
    // The column that took its place is selected, so the keys go on working on whole columns.
    expect(store.selection.lines).toBe('cols');
    store.selectLines('rows', [[1, 1]], { row: 1, col: 0 });
    store.changeCells('delete');
    await vi.waitFor(() => expect(store.selection).toMatchObject({ lines: 'rows', row: 1 }));
    select(rect(2, 1, 3, 1));
    store.changeCells('delete');
    expect(store.cellsDialog).toBe('delete');
  });

  it('turns down several ranges at once, and protected sheets', async () => {
    await open();
    store.setSelection({ row: 1, col: 1, rects: [rect(1, 1), rect(3, 1)] });
    store.changeCells('insert');
    expect(await store.deleteCells('up')).toBe(false);
    expect(store.cellsDialog).toBeUndefined();
    store.setActive(3);
    store.changeCells('insert');
    expect(toasts()).toEqual([
      'Cells are inserted and deleted one range at a time: select a single range.',
      'Cells are inserted and deleted one range at a time: select a single range.',
      'This sheet is protected in Excel, so it is read-only here.',
    ]);
  });

  it('moves cells in between others with Shift+drag, selecting where they land', async () => {
    await open();
    select(rect(1, 0, 1, 3));
    expect(await store.dragInsert(4, 0, 'down')).toBe(true);
    expect(column(0, 0, 4)).toEqual(['Name', 'apple', 'Fig', 'Pear', 'Banana']);
    expect([cell(3, 3)?.f, cell(0, 5)?.f, cell(0, 5)?.n]).toEqual(['B4*C4', 'SUM(D2:D5)', 18]);
    expect(store.selection.rects).toEqual([rect(3, 0, 3, 3)]);
    expect(actions()).toEqual(['Moved A2:D2']);
  });

  it('leaves cells dropped on their own edge, or a selection of several areas, where they are', async () => {
    await open();
    select(rect(1, 0, 1, 3));
    expect(await store.dragInsert(2, 0, 'down')).toBe(false);
    expect(await store.dragInsert(1, 4, 'right')).toBe(false);
    store.setSelection({ row: 1, col: 1, rects: [rect(1, 1), rect(3, 1)] });
    expect(await store.dragInsert(6, 1, 'down')).toBe(false);
    expect(actions()).toEqual([]);
  });
});
