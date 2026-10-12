import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { english } from '../src/i18n';
import { outlineGroups, type Axis } from '../src/state/outline';
import type { WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/**
 * tests/corpus/outline.xlsx. On "Report", rows 2-7 are grouped (level 1)
 * with their total in row 8, and inside them North's lines 2-3 and South's
 * 5-6 (level 2), with totals in rows 4 and 7; South's is collapsed. Columns
 * B:D and F:H are grouped, with totals in E and I; F:H is collapsed.
 * "Locked" is protected, letting rows be formatted but not columns.
 */
const bytes = new Uint8Array(readFileSync(join(__dirname, 'corpus', 'outline.xlsx')));

let store: WorkbookStore;
afterEach(() => store?.dispose());

async function open(): Promise<WorkbookStore> {
  store = await testStore();
  await opened(store, bytes, 'outline.xlsx');
  return store;
}

const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const select = (r: ReturnType<typeof rect>) => store.setSelection({ row: r.r1, col: r.c1, rects: [r] });
const actions = () => store.history.map((h) => english.text(h.action));
const toasts = () => store.toasts.map((t) => english.text(t.text));
const view = () => store.view!;
/** The rows (or columns) of a range that show, by sheet number. */
const showing = (axis: Axis, a: number, b: number) => Array.from({ length: b - a + 1 }, (_, i) => a + i).filter((l) => (axis === 'rows' ? view().gridRow(l) : view().gridCol(l)) >= 0);
const groups = (axis: Axis) => outlineGroups(view().info.outline?.[axis]);
const group = (axis: Axis, level: number, start: number) => groups(axis).find((g) => g.level === level && g.start === start)!;
const levels = (axis: Axis) => view().info.outline?.[axis].levels;
async function saved() {
  const s = await store.snapshot();
  return WorkbookSession.open(s!.bytes, 'outline.xlsx').loadResult().sheets;
}
async function asked() {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  return store.confirmRequest!;
}

describe('collapsing and expanding', () => {
  it("collapses and expands a group from its button, leaving the groups inside it that were collapsed as they were", async () => {
    await open();
    expect(groups('rows')).toEqual([
      { level: 1, start: 1, end: 6 },
      { level: 2, start: 1, end: 2 },
      { level: 2, start: 4, end: 5 },
    ]);
    expect(showing('rows', 0, 7)).toEqual([0, 1, 2, 3, 6, 7]);
    await store.toggleGroup('rows', group('rows', 1, 1));
    expect(showing('rows', 0, 7)).toEqual([0, 7]);
    expect(view().info.outline!.rows.collapsed).toEqual([6, 7]);
    await store.toggleGroup('rows', group('rows', 1, 1));
    // South's lines stay hidden: their group was collapsed before.
    expect(showing('rows', 0, 7)).toEqual([0, 1, 2, 3, 6, 7]);
    await store.toggleGroup('rows', group('rows', 2, 4));
    expect(showing('rows', 0, 7)).toEqual([0, 1, 2, 3, 4, 5, 6, 7]);
    expect(actions()).toEqual(['Hide detail', 'Show detail', 'Show detail']);
    // Each is one step to undo.
    await store.undo();
    await store.undo();
    expect(showing('rows', 0, 7)).toEqual([0, 7]);
    expect(store.editedCount).toBe(2);
    await store.undo();
    expect(store.editedCount).toBe(0);
  });

  it('collapses and expands groups of columns', async () => {
    await open();
    expect(showing('cols', 0, 9)).toEqual([0, 1, 2, 3, 4, 8, 9]);
    await store.toggleGroup('cols', group('cols', 1, 5));
    await store.toggleGroup('cols', group('cols', 1, 1));
    expect(showing('cols', 0, 9)).toEqual([0, 4, 5, 6, 7, 8, 9]);
    const sheet = (await saved())[0];
    expect([sheet.hiddenCols, sheet.outline!.cols.collapsed]).toEqual([[1, 2, 3], [4]]);
  });

  it('shows the outline to a level, as the level buttons do', async () => {
    await open();
    await store.showOutlineLevel('rows', 1);
    expect(showing('rows', 0, 7)).toEqual([0, 7]);
    await store.showOutlineLevel('rows', 2);
    expect(showing('rows', 0, 7)).toEqual([0, 3, 6, 7]);
    expect(view().info.outline!.rows.collapsed).toEqual([3, 6]);
    // Expanding a group shown to level 2 shows its lines.
    await store.toggleGroup('rows', group('rows', 2, 1));
    expect(showing('rows', 0, 7)).toEqual([0, 1, 2, 3, 6, 7]);
    await store.showOutlineLevel('rows', 3);
    expect(showing('rows', 0, 7)).toEqual([0, 1, 2, 3, 4, 5, 6, 7]);
    expect(view().info.outline!.rows.collapsed).toEqual([]);
    expect(actions()).toEqual(['Show outline level 1', 'Show outline level 2', 'Show detail', 'Show outline level 3']);
  });

  it('hides and shows the detail of the group at the active cell, moving the active cell out of the rows it hides', async () => {
    await open();
    select(rect(2, 3));
    expect(store.detailGroup('rows', false)).toEqual({ level: 2, start: 1, end: 2 });
    expect(store.detailGroup('rows', true)).toBeUndefined();
    await store.showDetail('rows', false);
    expect(showing('rows', 0, 3)).toEqual([0, 3]);
    // The total below the group, where its button is.
    expect([store.selection.row, store.selection.col]).toEqual([3, 3]);
    expect(store.detailGroup('rows', true)).toEqual({ level: 2, start: 1, end: 2 });
    await store.showDetail('rows', true);
    expect(showing('rows', 0, 3)).toEqual([0, 1, 2, 3]);
    // On South's total, Show detail opens South.
    select(rect(6, 0));
    await store.showDetail('rows', true);
    expect(showing('rows', 4, 5)).toEqual([4, 5]);
    // Outside any group, and on no group's total, there is no detail.
    select(rect(9, 0));
    expect([store.detailGroup('rows', false), store.detailGroup('rows', true)]).toEqual([undefined, undefined]);
  });
});

describe('grouping', () => {
  it('groups and ungroups the selected rows, one level at a time up to Excel\'s seven', async () => {
    await open();
    select(rect(9, 0, 10, 9));
    for (let i = 0; i < 7; i++) await store.groupLines('rows');
    expect(levels('rows')).toMatchObject({ 9: 7, 10: 7 });
    await store.groupLines('rows');
    expect(toasts()).toEqual(['Groups can go at most 7 levels deep, as in Excel.']);
    expect(levels('rows')).toMatchObject({ 9: 7, 10: 7 });
    await store.groupLines('rows', -1);
    expect(levels('rows')).toMatchObject({ 9: 6, 10: 6 });
    expect(actions().slice(-2)).toEqual(['Group 2 rows', 'Ungroup 2 rows']);
    select(rect(12, 0));
    await store.groupLines('rows', -1);
    expect(toasts()).toContain("The selected rows aren't in a group.");
    expect((await saved())[0].outline!.rows.levels).toMatchObject({ 9: 6, 10: 6 });
  });

  it('takes the collapsed mark off a total whose group is ungrouped', async () => {
    await open();
    select(rect(4, 0, 5, 9));
    await store.groupLines('rows', -1);
    expect(levels('rows')).toEqual({ 1: 2, 2: 2, 3: 1, 4: 1, 5: 1, 6: 1 });
    expect(view().info.outline!.rows.collapsed).toEqual([]);
    // The rows the group hid stay hidden, as in Excel.
    expect(view().info.hiddenRows).toEqual([4, 5]);
  });

  it('groups whole rows or columns with the shortcut, and asks which for a range of cells', async () => {
    await open();
    select(rect(0, 10, view().info.maxRow, 11));
    await store.groupSelection(1);
    expect(levels('cols')).toMatchObject({ 10: 1, 11: 1 });
    select(rect(9, 1, 10, 2));
    const grouping = store.groupSelection(1);
    const ask = await asked();
    expect([english.text(ask.title), english.text(ask.confirmLabel), english.text(ask.alternateLabel!)]).toEqual(['Group', 'Rows', 'Columns']);
    ask.resolve('confirm');
    await grouping;
    expect(levels('rows')).toMatchObject({ 9: 1, 10: 1 });
    const ungrouping = store.groupSelection(-1);
    (await asked()).resolve('cancel');
    await ungrouping;
    expect(levels('rows')).toMatchObject({ 9: 1, 10: 1 });
  });

  it("clears the outline, leaving the rows its groups hid hidden, as Excel does", async () => {
    await open();
    await store.clearOutline('rows');
    expect([levels('rows'), view().info.outline!.rows.collapsed, view().info.hiddenRows]).toEqual([{}, [], [4, 5]]);
    expect(actions()).toEqual(['Clear outline']);
    const sheet = (await saved())[0];
    expect([sheet.outline!.rows.levels, sheet.outline!.cols.levels]).toEqual([{}, { 1: 1, 2: 1, 3: 1, 5: 1, 6: 1, 7: 1 }]);
  });

  it('groups rows on a protected sheet that lets them be formatted, but not its columns', async () => {
    await open();
    store.setActive(2);
    expect([store.outlineLocked('rows'), store.outlineLocked('cols')]).toEqual([false, true]);
    select(rect(5, 0, 6, 0));
    await store.groupLines('rows');
    expect(levels('rows')).toMatchObject({ 5: 1, 6: 1 });
    await store.toggleGroup('rows', group('rows', 1, 1));
    expect(showing('rows', 0, 4)).toEqual([0, 4]);
    await store.groupLines('cols');
    expect(toasts()).toEqual(['This sheet is protected in Excel, so it is read-only here.']);
    expect(levels('cols')).toEqual({});
  });
});

describe('outlines and inserted rows', () => {
  it('keep their groups as rows are inserted, and show them again as they were when the insert is undone', async () => {
    await open();
    select(rect(9, 0, 10, 9));
    await store.groupLines('rows');
    await store.toggleGroup('rows', group('rows', 2, 1));
    select(rect(0, 0));
    await store.insertRows('above');
    expect(levels('rows')).toEqual({ 2: 2, 3: 2, 4: 1, 5: 2, 6: 2, 7: 1, 10: 1, 11: 1 });
    expect(showing('rows', 1, 8)).toEqual([1, 4, 7, 8]);
    await store.undo();
    // Back in the workbook from before the insert, which has the changes the file didn't.
    expect(levels('rows')).toEqual({ 1: 2, 2: 2, 3: 1, 4: 2, 5: 2, 6: 1, 9: 1, 10: 1 });
    expect(showing('rows', 0, 7)).toEqual([0, 3, 6, 7]);
  });
});
