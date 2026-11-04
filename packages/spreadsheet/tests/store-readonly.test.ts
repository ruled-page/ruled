import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import { objectKey } from '../src/core/objectKey';
import type { ChartEdit } from '../src/core/chartEdit';
import type { PivotSetup } from '../src/core/pivotTable';
import type { Draft } from '../src/state/drafts';
import type { StoreOptions, WorkbookStore } from '../src/state/store';
import type { WorkerRequest } from '../src/worker/protocol';
import { MENU_COMMANDS, SHORTCUT_COMMANDS, shortcutOn, viewOnly, type MenuCommandId } from '../src/ui/menuCommands';
import { applyLiveOptions } from '../src/editor/liveOptions';
import { handleFor } from '../src/editor/handle';
import { WorkbookSession } from '../src/core/session';
import { outlineGroups, type Axis } from '../src/state/outline';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 20_000 });

// Node has no IndexedDB, so drafts are kept in a map here instead.
const kept = vi.hoisted(() => new Map<string, Draft>());
vi.mock('../src/state/drafts', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/state/drafts')>()),
  loadDraft: vi.fn(async (key: string) => kept.get(key)),
  saveDraft: vi.fn(async (draft: Draft) => void kept.set(draft.key, structuredClone(draft))),
  deleteDraft: vi.fn(async (key: string) => void kept.delete(key)),
}));

/** Prices!A1:B4 with a formula, Totals, a hidden Archive and an empty Notes. */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const prices = wb.addWorksheet('Prices');
  prices.addRow(['Item', 'Price']);
  prices.addRow(['Bolts', 3]);
  prices.addRow(['Nuts', 0.5]);
  prices.addRow(['Total', { formula: 'SUM(B2:B3)', result: 3.5 }]);
  prices.getCell('A2').note = 'Steel';
  wb.addWorksheet('Totals').getCell('A1').value = { formula: 'SUM(Prices!B2:B3)', result: 3.5 };
  wb.addWorksheet('Archive', { state: 'hidden' }).getCell('A1').value = 'old';
  wb.addWorksheet('Notes');
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => {
  store?.dispose();
  kept.clear();
});

async function open(options: StoreOptions = {}): Promise<WorkbookStore> {
  store = await testStore(options);
  await store.open('prices.xlsx', bytes.slice().buffer as ArrayBuffer);
  for (let i = 0; i < 200 && !store.calc; i++) await new Promise((r) => setTimeout(r, 5));
  return store;
}

const select = (r1: number, c1: number, r2 = r1, c2 = c1) => store.setSelection({ row: r1, col: c1, rects: [{ r1, c1, r2, c2 }] });

/** What the user sees of the workbook: each sheet's name, state, cells and layout. */
function state() {
  return store.sheets.map((v) => {
    const cells: string[] = [];
    v.forEach({ r1: 0, c1: 0, r2: 20, c2: 10 }, (row, col, rec) => cells.push(`${row},${col}:${JSON.stringify(rec)}`));
    return { name: v.info.name, state: v.info.state, cells, info: JSON.stringify({ ...v.info, index: undefined }), notes: JSON.stringify(v.noteAt(1, 0) ?? null) };
  });
}

/** The worker requests that change the workbook. */
const CHANGES = new Set<WorkerRequest['type']>(['edit', 'structure', 'undoStructure', 'redoStructure']);

/** Makes the store read-only and counts what it asks the worker to change from then on. */
function readOnlyAndWatch(): WorkerRequest['type'][] {
  const sent: WorkerRequest['type'][] = [];
  const request = (store as unknown as { request: (req: WorkerRequest) => Promise<unknown> }).request.bind(store);
  (store as unknown as { request: (req: WorkerRequest) => Promise<unknown> }).request = (req) => {
    if (CHANGES.has(req.type)) sent.push(req.type);
    return request(req);
  };
  store.setReadOnly(true);
  return sent;
}

describe('a read-only workbook', () => {
  it('refuses every edit, quietly, without changing anything or marking it unsaved', async () => {
    await open();
    // Steps to undo and redo, and a chart, made while it could be edited.
    await store.input(1, 1, '4');
    await store.input(2, 1, '5');
    await store.undo();
    select(0, 0, 3, 1);
    await store.insertChart('column');
    const chart = store.view!.info.objects[0];
    expect(chart).toBeDefined();
    const editedCount = store.editedCount;
    const undoSteps = store.undoStack.length;
    const redoSteps = store.redoStack.length;
    const before = state();
    const saved = await store.snapshot();
    const sent = readOnlyAndWatch();
    store.toasts = [];

    const rect = { r1: 1, c1: 1, r2: 2, c2: 1 };
    select(1, 1, 2, 1);
    const calls: [string, () => Promise<unknown> | unknown][] = [
      ['type', () => store.input(1, 1, '9')],
      ['clear (Delete)', () => store.clear([rect])],
      ['paste', () => store.paste({ row: 1, col: 1 }, [['7'], ['8']], rect)],
      ['paste from the menu', () => store.pasteFromClipboard()],
      ['paste special', () => store.pasteSpecial('values')],
      [
        'cut and paste (moving cut cells)',
        () => {
          store.noteCopy(rect, true);
          return store.paste({ row: 5, col: 5 }, [], { r1: 5, c1: 5, r2: 5, c2: 5 }, true);
        },
      ],
      ['fill down (Ctrl+D)', () => store.fillDown()],
      ['fill right (Ctrl+R)', () => store.fillDown(true)],
      ['fill handle', () => store.fill(rect, { r1: 1, c1: 1, r2: 6, c2: 1 })],
      ['fill series', () => store.fillSeries(store.seriesStart())],
      ['drag cells', () => store.dragSelection(8, 3, false)],
      ['drag a copy of cells', () => store.dragSelection(8, 3, true)],
      ['drag cells in between', () => store.dragInsert(0, 0, 'down')],
      ['insert rows', () => store.insertRows('above')],
      ['delete rows', () => store.deleteRows()],
      ['insert columns', () => store.insertCols('left')],
      ['delete columns', () => store.deleteCols()],
      ['insert cells', () => store.insertCells('down')],
      ['delete cells', () => store.deleteCells('up')],
      ['Ctrl++', () => store.changeCells('insert')],
      ['hide rows', () => store.hideLines('rows')],
      ['unhide rows', () => store.unhideLines('rows')],
      ['row height', () => store.setRowHeight(1, 40)],
      ['column width', () => store.setColumnWidth(1, 30)],
      ['dragging a column edge', () => store.resizedColumn(1, 200)],
      ['dragging a row edge', () => store.resizedRow(1, 60)],
      ['autofit', () => store.autofit('cols', 1)],
      ['group rows', () => store.groupLines('rows')],
      ['freeze panes', () => store.freezePanes('cell')],
      ['merge', () => store.mergeCells()],
      ['unmerge', () => store.unmergeCells()],
      ['bold', () => store.toggleFormat('bold')],
      ['format', () => store.formatCells({ fill: 'FFFF0000' })],
      ['decimals', () => store.stepDecimals(1)],
      ['borders', () => store.drawBorders('all')],
      ['new note', () => store.setNote(0, 0, 'hello')],
      ['edit note', () => store.setNote(1, 0, 'changed')],
      ['delete note', () => store.deleteComment(1, 0)],
      ['new comment', () => store.startThread(0, 1, 'hi')],
      ['link', () => store.setLink({ address: 'https://example.com' } as Parameters<WorkbookStore['setLink']>[0])],
      ['remove links', () => store.removeLinks()],
      ['sort', () => store.sort(false)],
      ['custom sort', () => store.sortBy({ rect: { r1: 0, c1: 0, r2: 3, c2: 1 }, keys: [], header: true, caseSensitive: false, across: false })],
      ['replace', () => store.replace({ text: 'Bolts', matchCase: false, wholeCell: false, lookIn: 'formulas', scope: 'workbook' }, 'Screws')],
      ['checkboxes', () => store.insertCheckboxes()],
      ['tick', () => store.toggleCheckboxes({ row: 1, col: 1 })],
      ['data validation', () => store.setValidation([rect], undefined)],
      ['conditional formats', () => store.setConditionalFormats([])],
      ['clear conditional formats', () => store.clearConditionalFormats('sheet')],
      ['define a name', () => store.changeName({ to: { name: 'Prices', ref: 'Prices!$B$2:$B$3' } as never }, english.text('x') as never)],
      ['insert a chart', () => store.insertChart('bar')],
      ['move a chart', () => store.placeObject(objectKey(chart), { left: 5, top: 5, right: 5, bottom: 5 })],
      ['edit a chart', () => store.editChart(objectKey(chart), { title: 'Changed' } as ChartEdit)],
      ['delete a chart', () => store.deleteObject(objectKey(chart))],
      ['insert a pivot table', () => store.insertPivot('Prices!A1:B3')],
      ['refresh pivot tables', () => store.refreshPivots(true)],
      ['pivot fields', () => store.setPivotFields('xl/pivotTables/pivotTable1.xml', {} as PivotSetup)],
      ['add a sheet', () => store.addSheet()],
      ['rename a sheet', () => store.renameSheet(0, 'Renamed')],
      ['delete a sheet', () => store.deleteSheet(3)],
      ['move a sheet', () => store.moveSheet(0, 1)],
      ['move a sheet (Alt+Shift+arrow)', () => store.moveSheetBy(0, 1)],
      ['drop a sheet tab', () => store.dropSheet(0, 2)],
      ['hide a sheet', () => store.hideSheet(1)],
      ['unhide a sheet', () => store.unhideSheet(2)],
      ['undo (Ctrl+Z)', () => store.undo()],
      ['redo (Ctrl+Y)', () => store.redo()],
    ];
    for (const [what, call] of calls) {
      await call();
      expect({ what, sent, confirm: store.confirmRequest?.title, toasts: store.toasts.map((t) => english.text(t.text)) }).toEqual({ what, sent: [], confirm: undefined, toasts: [] });
    }
    await new Promise((r) => setTimeout(r, 20));
    expect(state()).toEqual(before);
    expect(store.editedCount).toBe(editedCount);
    expect(store.undoStack.length).toBe(undoSteps);
    expect(store.redoStack.length).toBe(redoSteps);
    expect(store.cutRange).toBeUndefined();
    const after = await store.snapshot();
    expect(after?.editedCount).toBe(saved?.editedCount);
    expect(after?.changes).toEqual(saved?.changes);
  });

  it('says once in a while, to screen readers only, that edits are refused', async () => {
    await open({ readOnly: true });
    store.spoken = [];
    await store.input(1, 1, '1');
    await store.input(1, 1, '2');
    await store.clear([{ r1: 1, c1: 1, r2: 1, c2: 1 }]);
    expect(store.spoken.map((s) => english.text(s.text))).toEqual(['This workbook is read-only.']);
    expect(store.toasts).toEqual([]);
  });

  it("doesn't open the dialogs, menus and modes that edit, and closes them when it becomes read-only", async () => {
    await open();
    store.set('sortDialog', { rect: { r1: 0, c1: 0, r2: 3, c2: 1 }, header: true });
    store.set('nameManagerOpen', true);
    store.set('seriesDialog', true);
    store.setReadOnly(true);
    expect([store.sortDialog, store.nameManagerOpen, store.seriesDialog]).toEqual([undefined, false, false]);
    store.openSortDialog();
    store.openCfManager();
    store.openCfGallery();
    store.openCfDialog('greaterThan');
    store.openInsertPivot();
    store.showFieldList();
    await store.openValidationDialog();
    store.set('renameSheetFor', 0);
    store.set('cellsDialog', 'insert' as never);
    expect([store.sortDialog, store.cfManager, store.cfGallery, store.cfDialog, store.pivotDialog, store.validationDialog, store.renameSheetFor, store.cellsDialog]).toEqual(Array(8).fill(undefined));
    expect(store.panel).not.toBe('pivot');
    // Finding, the list of shortcuts and filtering still open.
    store.set('findOpen', true);
    store.set('shortcutsOpen', true);
    store.set('filterPicker', { col: 0, rect: { r1: 0, c1: 0, r2: 3, c2: 1 }, x: 0, y: 0 });
    store.set('conditionDialog', { col: 1, rect: { r1: 0, c1: 0, r2: 3, c2: 1 }, mode: 'custom' });
    expect([store.findOpen, store.shortcutsOpen, !!store.filterPicker, !!store.conditionDialog]).toEqual([true, true, true, true]);
    // Groups collapse and expand, but aren't changed.
    expect([store.outlineLocked('rows'), store.outlineViewLocked('rows')]).toEqual([true, false]);
    expect(store.editableObjects()).toEqual([]);
  });

  it('still selects, copies and finds', async () => {
    await open({ readOnly: true });
    select(1, 0, 2, 1);
    expect(store.selection.rects[0]).toEqual({ r1: 1, c1: 0, r2: 2, c2: 1 });
    expect(store.selectionText(store.selection.rects[0])).toBe('Bolts\t3\nNuts\t0.5');
    store.noteCopy(store.selection.rects[0], false);
    expect(store.find({ text: 'Nuts', matchCase: false, wholeCell: false, lookIn: 'values', scope: 'sheet' })).toEqual([{ sheet: 0, row: 2, col: 0 }]);
    expect(store.goTo('Totals!A1', { focus: false })).toBe(true);
    expect(store.sheets[store.active].info.name).toBe('Totals');
  });

  it('marks every cell read-only, without a note at each, and edits again once it is turned off', async () => {
    await open({ readOnly: true });
    expect(english.text(store.view!.readOnlyReason(1, 1)!)).toBe('This workbook is read-only.');
    expect(english.text(store.view!.commentReadOnlyReason(1, 0)!)).toBe('This workbook is read-only.');
    store.setReadOnly(false);
    expect(store.view!.readOnlyReason(1, 1)).toBeUndefined();
    await store.input(1, 1, '9');
    expect(store.view!.get(1, 1)?.n).toBe(9);
    expect(store.editedCount).toBe(1);
  });

  it("hides the toolbar's editing parts without changing the features the host gave", async () => {
    await open({ readOnly: true, features: { csv: false, format: true } });
    expect(store.features).toMatchObject({ format: false, schemas: false, sheetTabButtons: false, sheetMenus: false, undoButtons: false, recordButtons: false, csv: false });
    expect(store.features).toMatchObject({ open: true, save: true, download: true, report: true, statusBar: true, shortcuts: true, exactNumbers: true });
    store.setReadOnly(false);
    expect(store.features).toMatchObject({ format: true, sheetMenus: true, undoButtons: true, csv: false });
  });

  it("doesn't offer back a draft, and keeps it for when the workbook can be edited", async () => {
    await open({ drafts: true });
    await store.input(1, 1, '9');
    await new Promise((r) => setTimeout(r, 600));
    expect(kept.size).toBe(1);
    store.dispose();
    await open({ drafts: true, readOnly: true });
    await new Promise((r) => setTimeout(r, 50));
    expect(store.confirmRequest).toBeUndefined();
    expect(store.view!.get(1, 1)?.n).toBe(3);
    expect(store.editedCount).toBe(0);
    expect(kept.size).toBe(1);
  });

  it('follows the readOnly option through the handle and live options', async () => {
    await open();
    const handle = handleFor(store, () => undefined);
    expect(handle.readOnly).toBe(false);
    applyLiveOptions(handle, {}, { readOnly: true });
    expect(store.readOnly).toBe(true);
    applyLiveOptions(handle, { readOnly: true }, {});
    expect(handle.readOnly).toBe(false);
    handle.setReadOnly(true);
    expect(store.readOnly).toBe(true);
  });
});

describe('menu entries and shortcuts in a read-only workbook', () => {
  it('keeps only the entries that do not edit or only change what shows, and plugin commands', () => {
    const kept = (Object.keys(MENU_COMMANDS) as MenuCommandId[]).filter(viewOnly);
    expect(kept).toEqual([
      'menu.copy',
      'menu.editNote',
      'menu.showComments',
      'outline.menu.hideDetail',
      'outline.menu.showDetail',
      'outline.menu.showLevel',
      'menu.filter',
      'menu.filterOn',
      'menu.clearFilter',
      'links.menu.open',
      'menu.find',
    ]);
    expect(viewOnly('menu.freezePanes')).toBe(false);
    expect(viewOnly('plugin/command')).toBe(true);
  });

  it('turns off the shortcuts of entries that edit, whatever the menus option keeps', () => {
    const on = SHORTCUT_COMMANDS.filter((id) => shortcutOn(undefined, id, true));
    expect(on).toEqual(['menu.copy', 'menu.find', 'menu.filterOn', 'menu.editNote', 'menu.showComments', 'links.menu.open']);
    expect(shortcutOn({ shortcuts: ['menu.paste'] }, 'menu.paste', true)).toBe(false);
    expect(shortcutOn({ cell: ['menu.copy'], header: [] }, 'menu.copy', true)).toBe(true);
    expect(shortcutOn({ cell: [], header: [] }, 'menu.copy', true)).toBe(false);
    expect(shortcutOn(undefined, 'menu.paste')).toBe(true);
  });
});

describe('filtering and outlines in a read-only workbook', () => {
  const data = { r1: 0, c1: 0, r2: 3, c2: 1 };
  const shown = (axis: Axis, a: number, b: number) =>
    Array.from({ length: b - a + 1 }, (_, i) => a + i).filter((l) => (axis === 'rows' ? store.view!.gridRow(l) : store.view!.gridCol(l)) >= 0);
  /** Checks nothing was changed: no worker edit, undo step, unsaved mark or draft, and the saved file is the one opened. */
  async function unchanged(sent: WorkerRequest['type'][], original: Uint8Array): Promise<void> {
    await new Promise((r) => setTimeout(r, 50));
    expect(sent).toEqual([]);
    expect([store.undoStack.length, store.redoStack.length, store.editedCount]).toEqual([0, 0, 0]);
    expect(kept.size).toBe(0);
    const saved = (await store.snapshot())!;
    expect(saved.changes).toEqual([]);
    const sheets = (b: Uint8Array) => WorkbookSession.open(b, 'x.xlsx').loadResult().sheets.map((s) => ({ hiddenRows: s.hiddenRows, hiddenCols: s.hiddenCols, outline: s.outline, filter: s.filter, autoFilter: s.autoFilter }));
    expect(sheets(saved.bytes)).toEqual(sheets(original));
  }

  it('filters by values and conditions, clears a column and the filter, for viewing only', async () => {
    await open({ drafts: true });
    const sent = readOnlyAndWatch();
    select(1, 0);
    await store.toggleFilter();
    expect(store.view!.filter?.rect).toEqual(data);
    await store.setFilter(data, 0, new Set(['Bolts', 'Total']));
    expect(shown('rows', 0, 3)).toEqual([0, 1, 3]);
    await store.setFilterCondition(data, 1, { kind: 'custom', and: true, rules: [{ operator: 'greaterThan', value: '1' }] });
    expect(shown('rows', 0, 3)).toEqual([0, 1, 3]);
    expect(store.view!.filteredCount).toBe(1);
    // Clearing column A leaves column B's condition.
    await store.setFilter(data, 0, undefined);
    expect(shown('rows', 0, 3)).toEqual([0, 1, 3]);
    // Show all, as the status bar's link and the menu's Clear do.
    await store.clearFilter();
    expect([store.view!.filtering, shown('rows', 0, 3)]).toEqual([false, [0, 1, 2, 3]]);
    await store.setFilter(data, 0, new Set(['Nuts']));
    expect(shown('rows', 0, 3)).toEqual([0, 2]);
    expect(store.toasts).toEqual([]);
    expect(store.spoken.map((s) => english.text(s.text))).not.toContain('This workbook is read-only.');
    await unchanged(sent, bytes);
    // Sorting is still refused.
    await store.sortBy({ rect: data, keys: [{ col: 0, descending: true }], header: true, caseSensitive: false, across: false });
    expect(store.view!.get(1, 0)?.s).toBe('Bolts');
    await unchanged(sent, bytes);
    // Turning the filter buttons off shows every row.
    await store.toggleFilter();
    expect([store.view!.filter, shown('rows', 0, 3)]).toEqual([undefined, [0, 1, 2, 3]]);
  });

  it("starts from the file's filter, and shows it again once the workbook can be edited", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Stock');
    for (const row of [['Item', 'Qty'], ['Bolts', 3], ['Nuts', 5], ['Pins', 7]]) ws.addRow(row);
    ws.autoFilter = 'A1:B4';
    const original = new Uint8Array(await wb.xlsx.writeBuffer());
    store = await testStore({ readOnly: true });
    await store.open('stock.xlsx', original.slice().buffer as ArrayBuffer);
    const sent = readOnlyAndWatch();
    expect(store.view!.filter?.rect).toEqual(data);
    await store.setFilter(data, 0, new Set(['Pins']));
    expect(shown('rows', 0, 3)).toEqual([0, 3]);
    // Filter off, then on again, from the menu.
    await store.toggleFilter();
    expect(store.view!.filter).toBeUndefined();
    await unchanged(sent, original);
    store.setReadOnly(false);
    expect([store.view!.filter?.rect, store.view!.filtering, shown('rows', 0, 3)]).toEqual([data, false, [0, 1, 2, 3]]);
    // Editable again, filtering is a change to the file.
    await store.setFilter(data, 0, new Set(['Nuts']));
    expect([store.undoStack.length, store.editedCount]).toEqual([1, 1]);
  });

  it('collapses and expands groups and shows levels for viewing only, and shows the file\'s outline again once editable', async () => {
    const original = new Uint8Array(readFileSync(join(__dirname, 'corpus', 'outline.xlsx')));
    store = await testStore({ drafts: true });
    await opened(store, original, 'outline.xlsx');
    const sent = readOnlyAndWatch();
    const group = (axis: Axis, level: number, start: number) => outlineGroups(store.view!.info.outline?.[axis]).find((g) => g.level === level && g.start === start)!;
    expect(shown('rows', 0, 7)).toEqual([0, 1, 2, 3, 6, 7]);
    await store.toggleGroup('rows', group('rows', 1, 1));
    expect(shown('rows', 0, 7)).toEqual([0, 7]);
    await store.toggleGroup('rows', group('rows', 1, 1));
    await store.toggleGroup('rows', group('rows', 2, 4));
    expect(shown('rows', 0, 7)).toEqual([0, 1, 2, 3, 4, 5, 6, 7]);
    await store.showOutlineLevel('rows', 1);
    expect(shown('rows', 0, 7)).toEqual([0, 7]);
    await store.toggleGroup('cols', group('cols', 1, 5));
    expect(shown('cols', 0, 9)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    // Hide and Show Detail at the active cell.
    select(0, 2);
    await store.showDetail('cols', false);
    expect(shown('cols', 0, 9)).toEqual([0, 4, 5, 6, 7, 8, 9]);
    // Grouping and clearing the outline are still refused.
    select(8, 0, 9, 0);
    await store.groupLines('rows');
    await store.clearOutline('rows');
    expect(store.view!.info.outline!.rows.levels[8]).toBeUndefined();
    expect(store.toasts).toEqual([]);
    await unchanged(sent, original);
    store.setReadOnly(false);
    expect([shown('rows', 0, 7), shown('cols', 0, 9)]).toEqual([
      [0, 1, 2, 3, 6, 7],
      [0, 1, 2, 3, 4, 8, 9],
    ]);
    expect(store.view!.info.outline!.rows.collapsed).toEqual([6]);
  });
});
