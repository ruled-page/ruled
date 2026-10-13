import { afterEach, describe, expect, it, vi } from 'vitest';
import type { PivotSetup } from '../src/core/pivotTable';
import { english } from '../src/i18n';
import { addField, fieldsInUse, moveChip, moveTo, removeChip, removeField, setShown, setSummary, shownItems, VALUES_FIELD, type DataName } from '../src/state/pivotFields';
import type { WorkbookStore } from '../src/state/store';
import { sample } from './helpers';
import { opened, testStore } from './storeHarness';

const NAMES = ['Region', 'Product', 'Units', 'Price', 'Revenue'];
const name: DataName = (fn, field) => english.text({ key: `pivot.dataName.${fn}`, params: { field: NAMES[field] } });
const base: PivotSetup = { rows: [0], cols: [], pages: [], data: [{ field: 2, fn: 'sum', name: 'Sum of Units' }], hidden: {} };

describe('ticking and unticking fields', () => {
  it('sums numbers among the values and puts anything else down the rows, the Values field across once there are two', () => {
    let s = addField(base, 1, false, name);
    expect(s.rows).toEqual([0, 1]);
    s = addField(s, 4, true, name);
    expect(s.data.map((d) => d.name)).toEqual(['Sum of Units', 'Sum of Revenue']);
    expect(s.cols).toEqual([VALUES_FIELD]);
    // The same field twice is named apart, as Excel names it.
    expect(addField(s, 4, true, name).data[2].name).toBe('Sum of Revenue2');
    expect([...fieldsInUse(s)]).toEqual([0, 1, 2, 4]);
    s = removeField({ ...s, hidden: { 1: ['sgizmo'] } }, 4);
    expect([s.cols, s.hidden[1]]).toEqual([[], ['sgizmo']]);
    expect(removeField(s, 1).hidden).toEqual({});
  });
});

describe('moving fields between the areas', () => {
  it('moves a field from one area to another, or within one, and a field dragged in from the list', () => {
    expect(moveChip(base, { area: 'rows', index: 0 }, { area: 'cols', at: 0 }, name)).toMatchObject({ rows: [], cols: [0] });
    const two = { ...base, rows: [0, 1] };
    expect(moveChip(two, { area: 'rows', index: 0 }, { area: 'rows', at: 2 }, name).rows).toEqual([1, 0]);
    expect(moveTo(two, 1, 'rows', 0, name).rows).toEqual([1, 0]);
    expect(moveTo(two, 1, 'pages', 0, name)).toMatchObject({ rows: [0], pages: [{ field: 1 }] });
    // A text field dragged among the values is counted; one dragged out of them leaves them.
    const counted = moveChip(two, { area: 'rows', index: 1 }, { area: 'data', at: 0 }, name, 'count');
    expect(counted).toMatchObject({ rows: [0], cols: [VALUES_FIELD] });
    expect(counted.data.map((d) => d.name)).toEqual(['Count of Product', 'Sum of Units']);
    expect(moveChip(counted, { area: 'data', index: 1 }, { area: 'data', at: 0 }, name).data.map((d) => d.field)).toEqual([2, 1]);
    expect(moveChip(counted, { area: 'data', index: 0 }, { area: 'cols', at: 0 }, name)).toMatchObject({ rows: [0], cols: [1], data: [{ field: 2 }] });
    // The Values field goes down the rows or across the columns, nowhere else.
    expect(moveChip(counted, { area: 'cols', index: 0 }, { area: 'rows', at: 1 }, name)).toMatchObject({ rows: [0, VALUES_FIELD], cols: [] });
    expect(moveTo(counted, VALUES_FIELD, 'pages', 0, name)).toBe(counted);
    expect(moveTo(counted, VALUES_FIELD, 'data', 0, name)).toBe(counted);
    expect(moveChip(counted, { area: 'rows', index: 5 }, { area: 'cols', at: 0 }, name)).toBe(counted);
  });

  it('removes a field from the one area its menu is in', () => {
    const both = { ...base, rows: [0, 2], hidden: { 2: ['n5'] } };
    expect(removeChip(both, 'rows', 1)).toMatchObject({ rows: [0], data: [{ field: 2 }], hidden: {} });
    expect(removeChip(both, 'data', 0)).toMatchObject({ rows: [0, 2], data: [], hidden: { 2: ['n5'] } });
    expect(removeChip({ ...both, cols: [VALUES_FIELD] }, 'cols', 0).cols).toEqual([VALUES_FIELD]);
  });
});

describe('summaries and filters', () => {
  it('renames a value summarised another way only when it had its summary’s own name', () => {
    expect(setSummary(base, 0, 'average', name).data[0]).toEqual({ field: 2, fn: 'average', name: 'Average of Units' });
    expect(setSummary({ ...base, data: [{ field: 2, fn: 'sum', name: 'Units sold' }] }, 0, 'max', name).data[0].name).toBe('Units sold');
    expect(setSummary(base, 0, 'sum', name)).toBe(base);
  });

  it('hides the items not shown, and a report filter showing one picks it', () => {
    const all = ['seast', 'snorth', 'ssouth', 'swest'];
    const s = setShown(base, 0, ['seast', 'snorth'], all);
    expect(s.hidden).toEqual({ 0: ['ssouth', 'swest'] });
    expect(shownItems(s, 0, all)).toEqual(['seast', 'snorth']);
    expect(setShown(s, 0, all, all).hidden).toEqual({});
    const paged = { ...base, rows: [], pages: [{ field: 0 }] };
    const one = setShown(paged, 0, ['snorth'], all);
    expect(one).toMatchObject({ pages: [{ field: 0, item: 'snorth' }], hidden: {} });
    expect(shownItems(one, 0, all)).toEqual(['snorth']);
    expect(setShown(one, 0, ['snorth', 'swest'], all)).toMatchObject({ pages: [{ field: 0 }], hidden: { 0: ['seast', 'ssouth'] } });
    expect(setShown(one, 0, all, all)).toMatchObject({ pages: [{ field: 0 }], hidden: {} });
  });
});

describe('the PivotTable Fields pane’s store', () => {
  let store: WorkbookStore;
  afterEach(() => store?.dispose());
  vi.setConfig({ testTimeout: 10_000 });

  it('knows the pivot table the active cell is in, and changes its fields as one undoable step', async () => {
    store = await testStore();
    await opened(store, sample('pivot.xlsx'), 'pivot.xlsx');
    store.setActive(1);
    store.setSelection({ row: 20, col: 5, rects: [{ r1: 20, c1: 5, r2: 20, c2: 5 }] });
    expect(store.activePivot).toBeUndefined();
    store.setSelection({ row: 4, col: 1, rects: [{ r1: 4, c1: 1, r2: 4, c2: 1 }] });
    const pivot = store.activePivot!;
    expect(pivot.name).toBe('UnitsByRegion');
    store.showFieldList();
    expect(store.panel).toBe('pivot');

    await store.setPivotFields(pivot.path, { ...pivot.setup, cols: [1] });
    expect(store.activePivot?.setup.cols).toEqual([1]);
    expect(store.sheets[1].info.blocked.find((b) => b.kind === 'pivot')).toMatchObject({ c2: 4 });
    expect(store.history.map((h) => english.text(h.action))).toEqual(['Change the fields of UnitsByRegion']);
    await store.undo();
    expect(store.activePivot?.setup.cols).toEqual([]);
    // One gone, or one whose fields can't be changed, is left alone.
    const steps = store.history.length;
    await store.setPivotFields('xl/pivotTables/none.xml', pivot.setup);
    expect(store.history).toHaveLength(steps);
  });
});
