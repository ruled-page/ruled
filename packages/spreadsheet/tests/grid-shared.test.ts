import ExcelJS from 'exceljs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { withAlpha } from '../src/grid/color';
import { editorKey, type GridKey } from '../src/grid/keys';
import { cellOptions } from '../src/grid/lists';
import { sheetMarks } from '../src/grid/marks';
import { growOverMerges, navigate, navSheet, type NavKey, type NavSheet } from '../src/grid/navigation';
import { spokenText } from '../src/grid/spoken';
import type { Selection, WorkbookStore } from '../src/state/store';
import { DEFAULT_PALETTE, DEFAULT_REF_COLORS } from '../src/ui/gridPalette';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

describe('colours with an opacity', () => {
  it('reads hex and rgb() colours, and leaves others as they are', () => {
    expect(withAlpha('#abc', 0.5)).toBe('rgba(170, 187, 204, 0.5)');
    expect(withAlpha('#1e7145', 0)).toBe('rgba(30, 113, 69, 0)');
    expect(withAlpha('#1e714580', 1)).toBe('rgba(30, 113, 69, 1)');
    expect(withAlpha('rgba(30, 113, 69, 0.1)', 1)).toBe('rgba(30, 113, 69, 1)');
    expect(withAlpha('rgb(1 2 3)', 0.2)).toBe('rgba(1, 2, 3, 0.2)');
    expect(withAlpha('red', 0.5)).toBe('red');
    expect(withAlpha('#12345', 0.5)).toBe('#12345');
  });
});

describe('moving the selection with the keyboard, as Excel does', () => {
  /** Rows 0 to 99 but 5, columns 0 to 19 but 3; B2:B4 and B8 filled, and A1:C1. */
  const filled = new Set(['1:1', '2:1', '3:1', '7:1', '0:0', '0:1', '0:2']);
  const sheet: NavSheet = {
    rows: Int32Array.from({ length: 100 }, (_, i) => i).filter((r) => r !== 5),
    cols: Int32Array.from({ length: 20 }, (_, i) => i).filter((c) => c !== 3),
    filled: (r, c) => filled.has(`${r}:${c}`),
    maxRow: 7,
    maxCol: 2,
    lastRow: 1_048_575,
    lastCol: 16_383,
    frozenRows: 1,
    frozenCols: 1,
    pageRows: 10,
    pageCols: 4,
    merges: [],
    mergeAt: () => undefined,
  };
  const at = (row: number, col: number): Selection => ({ row, col, rects: [{ r1: row, r2: row, c1: col, c2: col }] });
  const key = (k: string, mods: Partial<NavKey> = {}): NavKey => ({ key: k, shiftKey: false, altKey: false, command: false, ...mods });
  const to = (sel: Selection, k: NavKey) => {
    const moved = navigate(sel, k, sheet);
    return moved && [moved.selection.row, moved.selection.col, moved.selection.rects[0], moved.reveal];
  };

  it('steps over hidden rows and columns, stopping at the first, and goes on past the rows shown so far', () => {
    expect(to(at(4, 2), key('ArrowDown'))?.slice(0, 2)).toEqual([6, 2]);
    expect(to(at(6, 2), key('ArrowUp'))?.slice(0, 2)).toEqual([4, 2]);
    expect(to(at(2, 2), key('ArrowRight'))?.slice(0, 2)).toEqual([2, 4]);
    expect(to(at(0, 0), key('ArrowUp'))?.slice(0, 2)).toEqual([0, 0]);
    expect(to(at(0, 0), key('ArrowLeft'))?.slice(0, 2)).toEqual([0, 0]);
    expect(to(at(99, 0), key('ArrowDown'))?.slice(0, 2)).toEqual([100, 0]);
    expect(to(at(200, 0), key('ArrowDown'))?.slice(0, 2)).toEqual([201, 0]);
  });

  it('extends the selection with Shift, the active cell staying and the far corner moving', () => {
    const down = navigate(at(2, 2), key('ArrowDown', { shiftKey: true }), sheet)!;
    expect(down.selection).toEqual({ row: 2, col: 2, rects: [{ r1: 2, r2: 3, c1: 2, c2: 2 }] });
    expect(down.reveal).toEqual({ row: 3, col: 2 });
    const back = navigate(down.selection, key('ArrowUp', { shiftKey: true }), sheet)!;
    const up = navigate(back.selection, key('ArrowUp', { shiftKey: true }), sheet)!;
    expect(up.selection.rects[0]).toEqual({ r1: 1, r2: 2, c1: 2, c2: 2 });
  });

  it('jumps to the edge of the data with Ctrl+Arrow, or the sheet edge past it', () => {
    // Through a block of filled cells to its last.
    expect(to(at(1, 1), key('ArrowDown', { command: true }))?.slice(0, 2)).toEqual([3, 1]);
    // From the end of a block, to the next filled cell.
    expect(to(at(3, 1), key('ArrowDown', { command: true }))?.slice(0, 2)).toEqual([7, 1]);
    // Past the last, to the sheet's last row; back up, to the first.
    expect(to(at(7, 1), key('ArrowDown', { command: true }))?.slice(0, 2)).toEqual([1_048_575, 1]);
    expect(to(at(50, 4), key('ArrowUp', { command: true }))?.slice(0, 2)).toEqual([0, 4]);
    expect(to(at(0, 0), key('ArrowRight', { command: true }))?.slice(0, 2)).toEqual([0, 2]);
    expect(to(at(0, 2), key('ArrowLeft', { command: true }))?.slice(0, 2)).toEqual([0, 0]);
    expect(to(at(0, 0), key('ArrowUp', { command: true }))?.slice(0, 2)).toEqual([0, 0]);
    // Shift with Ctrl+Arrow selects to the edge.
    expect(to(at(1, 1), key('ArrowDown', { command: true, shiftKey: true }))?.[2]).toEqual({ r1: 1, r2: 3, c1: 1, c2: 1 });
  });

  it('goes home, to the first cell that scrolls, to the last cell used, and by pages', () => {
    expect(to(at(40, 9), key('Home'))?.slice(0, 2)).toEqual([40, 0]);
    expect(to(at(40, 9), key('Home', { command: true }))?.slice(0, 2)).toEqual([1, 1]);
    expect(to(at(40, 9), key('End', { command: true }))?.slice(0, 2)).toEqual([7, 2]);
    expect(to(at(0, 0), key('PageDown'))?.slice(0, 2)).toEqual([11, 0]);
    expect(to(at(11, 0), key('PageUp'))?.slice(0, 2)).toEqual([0, 0]);
    expect(to(at(0, 0), key('PageDown', { altKey: true }))?.slice(0, 2)).toEqual([0, 5]);
  });

  it('moves along with Tab, and round a range with Tab and Enter, leaving Enter on one cell to editing', () => {
    expect(to(at(2, 2), key('Tab'))?.slice(0, 2)).toEqual([2, 4]);
    expect(to(at(2, 2), key('Tab', { shiftKey: true }))?.slice(0, 2)).toEqual([2, 1]);
    expect(navigate(at(2, 2), key('Enter'), sheet)).toBeUndefined();
    const range: Selection = { row: 1, col: 1, rects: [{ r1: 1, r2: 2, c1: 1, c2: 2 }] };
    expect(to(range, key('Tab'))?.slice(0, 3)).toEqual([1, 2, range.rects[0]]);
    expect(to({ ...range, row: 1, col: 2 }, key('Tab'))?.slice(0, 2)).toEqual([2, 1]);
    expect(to({ ...range, row: 2, col: 2 }, key('Tab'))?.slice(0, 2)).toEqual([1, 1]);
    expect(to(range, key('Enter'))?.slice(0, 2)).toEqual([2, 1]);
    expect(to(range, key('Enter', { shiftKey: true }))?.slice(0, 2)).toEqual([2, 2]);
    // A range of hidden lines only has nowhere to go.
    expect(to({ row: 5, col: 3, rects: [{ r1: 5, r2: 5, c1: 3, c2: 4 }] }, key('Tab'))?.slice(0, 2)).toEqual([5, 3]);
  });

  it('treats a merge as one cell, as Excel does, and grows a selection to cover the merges it touches', () => {
    // E10:G11 and H12:H14 merged.
    const merges = [
      { r1: 9, r2: 10, c1: 4, c2: 6 },
      { r1: 11, r2: 13, c1: 7, c2: 7 },
    ];
    const merged: NavSheet = { ...sheet, merges, mergeAt: (r, c) => merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) };
    const go = (sel: Selection, k: NavKey) => navigate(sel, k, merged)!;
    // Into a merge, onto its top-left cell, the merge selected.
    expect(go(at(8, 5), key('ArrowDown')).selection).toEqual({ row: 9, col: 4, rects: [merges[0]] });
    // Out of one, from its far side.
    const inMerge: Selection = { row: 9, col: 4, rects: [merges[0]] };
    expect(go(inMerge, key('ArrowRight')).selection.col).toBe(7);
    expect(go(inMerge, key('ArrowDown')).selection.row).toBe(11);
    expect(go(inMerge, key('ArrowLeft')).selection.col).toBe(2);
    expect(go(inMerge, key('Tab')).selection.col).toBe(7);
    // Shift grows over the merges it reaches.
    expect(go(at(8, 4), key('ArrowDown', { shiftKey: true })).selection.rects[0]).toEqual({ r1: 8, r2: 10, c1: 4, c2: 6 });
    expect(growOverMerges({ r1: 10, r2: 12, c1: 6, c2: 7 }, merges)).toEqual({ r1: 9, r2: 13, c1: 4, c2: 7 });
    // Round a range, the cells a merge covers passed over.
    const range: Selection = { row: 9, col: 4, rects: [{ r1: 9, r2: 10, c1: 4, c2: 7 }] };
    expect(go(range, key('Tab')).selection.col).toBe(7);
    expect(go({ ...range, col: 7 }, key('Tab')).selection).toMatchObject({ row: 10, col: 7 });
    expect(go({ ...range, row: 10, col: 7 }, key('Tab')).selection).toMatchObject({ row: 9, col: 4 });
    // A merge on its own is one cell: Enter edits it.
    expect(navigate(inMerge, key('Enter'), merged)).toBeUndefined();
  });

  it('steps the way the arrows point on a sheet read right to left, Tab still going on to the next column', () => {
    const rtl: NavSheet = { ...sheet, rtl: true };
    expect(navigate(at(2, 2), key('ArrowLeft'), rtl)?.selection.col).toBe(4);
    expect(navigate(at(2, 2), key('ArrowRight'), rtl)?.selection.col).toBe(1);
    expect(navigate(at(2, 2), key('Tab'), rtl)?.selection.col).toBe(4);
  });

  it('shrinks the selection to the active cell with Shift+Backspace, and brings it into view with Ctrl+Backspace', () => {
    const range: Selection = { row: 1, col: 1, rects: [{ r1: 1, r2: 4, c1: 1, c2: 2 }] };
    expect(navigate(range, key('Backspace', { shiftKey: true }), sheet)?.selection).toEqual(at(1, 1));
    const shown = navigate(range, key('Backspace', { command: true }), sheet)!;
    expect(shown.selection).toBe(range);
    expect(shown.reveal).toEqual({ row: 1, col: 1 });
    expect(navigate(range, key('Backspace'), sheet)).toBeUndefined();
    expect(navigate(range, key('x'), sheet)).toBeUndefined();
  });
});

describe('what the grids share', () => {
  let store: WorkbookStore;
  afterEach(() => store?.dispose());

  async function open(): Promise<WorkbookStore> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet');
    ws.getCell('A1').value = 'Merged';
    ws.mergeCells('A1:B1');
    ws.getCell('C1').value = { text: 'site', hyperlink: 'https://example.com' };
    ws.getCell('D1').value = { formula: '1+1', result: 2 };
    ws.getCell('E1').note = 'A note';
    ws.getCell('E1').value = 'noted';
    ws.getCell('F1').dataValidation = { type: 'list', allowBlank: true, formulae: ['"red,green,blue"'] };
    ws.getCell('A3').value = 1;
    ws.getCell('A4').value = 2;
    ws.autoFilter = 'A3:B4';
    store = await testStore();
    await opened(store, new Uint8Array(await wb.xlsx.writeBuffer()), 'shared.xlsx');
    return store;
  }

  it('says what a cell is: its reference, what it shows, and what else is true of it', async () => {
    const s = await open();
    const view = s.view!;
    expect(spokenText(s, view, 0, 1, view.mergeAt(0, 1), '', false)).toBe('A1, Merged, merged A1:B1');
    expect(spokenText(s, view, 0, 2, undefined, 'site', false)).toContain('https://example.com');
    expect(spokenText(s, view, 0, 3, undefined, '2', false)).toContain('formula');
    expect(spokenText(s, view, 0, 4, undefined, 'noted', false)).toContain('A note');
    expect(spokenText(s, view, 0, 5, undefined, '', true)).toBe('F1, Empty, has a list');
    await s.input(1, 0, 'changed');
    expect(spokenText(s, view, 1, 0, undefined, 'changed', false)).toContain('changed');
  });

  it('offers a list cell’s values, and none for other cells', async () => {
    const s = await open();
    expect(cellOptions(s, s.view!, 0, 5)).toEqual(['red', 'green', 'blue']);
    expect(cellOptions(s, s.view!, 0, 6)).toBeUndefined();
    // A cell under a merge can't be edited.
    expect(cellOptions(s, s.view!, 0, 1)).toBeUndefined();
  });

  it('marks a cut, a landing place, a spill and the references of a formula being typed', async () => {
    const s = await open();
    await s.input(6, 0, '=SEQUENCE(3)');
    s.setSelection({ row: 6, col: 0, rects: [{ r1: 6, r2: 6, c1: 0, c2: 0 }] });
    s.noteCopy({ r1: 2, r2: 3, c1: 0, c2: 0 }, true);
    s.showFormulaRefs('=A3+B4');
    const marks = sheetMarks(s, s.view!, s.selection, DEFAULT_PALETTE, DEFAULT_REF_COLORS, { r1: 9, r2: 9, c1: 0, c2: 0 });
    expect(marks.map((m) => m.style)).toEqual(['dashed', 'dashed', 'solid', 'solid', 'solid']);
    expect(marks[2].rect).toEqual({ r1: 6, r2: 8, c1: 0, c2: 0 });
    expect(marks[3].color).toBe(withAlpha(DEFAULT_REF_COLORS[0], 0.1));
    // Something in the way of the spill dashes its outline.
    await s.input(8, 0, 'blocked');
    expect(sheetMarks(s, s.view!, s.selection, DEFAULT_PALETTE, DEFAULT_REF_COLORS).find((m) => m.rect.r1 === 6)?.style).toBe('dashed');
  });

  it("takes the editor's own keys, and leaves others to the grid", async () => {
    const s = await open();
    const press = (k: Partial<GridKey>) => editorKey(s, { key: '', keyCode: 0, shiftKey: false, ctrlKey: false, metaKey: false, altKey: false, ...k }, (row, col) => (row === 2 && col === 0 ? { col, x: 1, y: 2 } : undefined));
    const calls = (name: string) => vi.spyOn(s as unknown as Record<string, () => unknown>, name).mockImplementation(() => Promise.resolve());
    const toggle = calls('toggleCheckboxes');
    const follow = calls('followLink');
    const group = calls('groupSelection');
    const refresh = calls('refreshPivots');
    const chart = calls('insertChart');
    const filter = calls('toggleFilter');
    const change = calls('changeCells');
    const select = vi.spyOn(s, 'selectObject');

    expect(press({ key: 'x' })).toBe(false);
    expect(press({ key: ' ' })).toBe(false);
    vi.spyOn(s, 'checkboxAt').mockReturnValueOnce(false);
    expect(press({ key: ' ' })).toBe(true);
    expect(toggle).toHaveBeenCalled();
    expect(press({ key: 'Enter', ctrlKey: true, metaKey: true })).toBe(true);
    expect(follow).toHaveBeenCalled();
    expect(press({ key: 'ArrowRight', altKey: true, shiftKey: true })).toBe(true);
    expect(group).toHaveBeenCalledWith(1);
    s.setSelection({ row: 2, col: 0, rects: [{ r1: 2, r2: 3, c1: 0, c2: 1 }] });
    expect(press({ key: ' ', shiftKey: true })).toBe(true);
    expect(s.selection.lines).toBe('rows');
    expect(press({ key: 'a', ctrlKey: true, metaKey: true })).toBe(true);
    expect(s.selection.lines).toBe('all');
    s.setSelection({ row: 2, col: 0, rects: [{ r1: 2, r2: 2, c1: 0, c2: 0 }] });
    expect(press({ key: 'ArrowDown', altKey: true })).toBe(true);
    expect(s.filterPicker).toEqual({ col: 0, x: 1, y: 2 });
    s.setSelection({ row: 5, col: 5, rects: [{ r1: 5, r2: 5, c1: 5, c2: 5 }] });
    expect(press({ key: 'ArrowDown', altKey: true })).toBe(false);
    expect(press({ key: '5', keyCode: 53, ctrlKey: true, altKey: true })).toBe(false);
    expect(select).not.toHaveBeenCalled();
    expect(press({ key: 'F5', altKey: true })).toBe(true);
    expect(refresh).toHaveBeenCalledWith(false);
    expect(press({ key: 'F1', altKey: true })).toBe(true);
    expect(chart).toHaveBeenCalledWith('column');
    expect(press({ key: 'L', ctrlKey: true, metaKey: true, shiftKey: true })).toBe(true);
    expect(filter).toHaveBeenCalled();
    expect(press({ key: '+', ctrlKey: true, metaKey: true })).toBe(true);
    expect(press({ key: '-', ctrlKey: true, metaKey: true })).toBe(true);
    expect(change.mock.calls).toEqual([['insert'], ['delete']]);
  });

  it('finds the filled cells and the used range of a sheet, the latter only when asked', async () => {
    const s = await open();
    const view = s.view!;
    const extent = vi.spyOn(view, 'extent');
    const sheet = navSheet(view, 0, 7);
    expect([sheet.pageRows, sheet.pageCols, sheet.frozenRows]).toEqual([1, 7, 0]);
    expect(sheet.filled(0, 0)).toBe(true);
    expect(sheet.filled(5, 5)).toBe(false);
    expect(extent).not.toHaveBeenCalled();
    // F1 has a list but no value.
    expect([sheet.maxRow, sheet.maxCol]).toEqual([3, 4]);
    expect(extent).toHaveBeenCalledTimes(1);
  });
});
