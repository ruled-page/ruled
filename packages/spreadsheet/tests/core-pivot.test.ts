import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { english } from '../src/i18n';
import { open, parts } from './helpers';

/** The demo's pivot.xlsx: Data!A1:E101 (Region, Product, Units, Price, Revenue), and on Pivot!A3:B8 the units summed by region. */
const shown = (s: WorkbookSession, ref: string) => {
  const [col, row] = [ref.charCodeAt(0) - 65, Number(ref.slice(1)) - 1];
  const c = s.current(1, row, col).content;
  return c.kind === 'empty' ? null : 'value' in c ? c.value : c;
};
const column = (s: WorkbookSession, col: string, from: number, to: number) => Array.from({ length: to - from + 1 }, (_, i) => shown(s, `${col}${from + i}`));

describe('a pivot table worked out from its source', () => {
  it('matches what Excel saved, so refreshing an untouched one changes nothing', () => {
    const s = open('pivot.xlsx');
    const before = [column(s, 'A', 3, 8), column(s, 'B', 3, 8)];
    expect(before).toEqual([
      ['Row Labels', 'East', 'North', 'South', 'West', 'Grand Total'],
      ['Sum of Units', 658, 1315, 1132, 1449, 4554],
    ]);
    const r = s.apply({ type: 'refreshPivots' });
    expect(r.rejected).toBeUndefined();
    expect(r.undo).toEqual([]);
    expect([column(s, 'A', 3, 8), column(s, 'B', 3, 8)]).toEqual(before);
  });

  it('follows its source as it is edited, in the same step, and is not counted as an edit', () => {
    const s = open('pivot.xlsx');
    // Data!C2 is a North row's units.
    const was = shown(s, 'B5') as number;
    const units = s.current(0, 1, 2).content;
    const r = s.apply({ type: 'input', sheet: 0, row: 1, col: 2, text: String(('value' in units ? (units.value as number) : 0) + 100) });
    expect([shown(s, 'B5'), shown(s, 'B8')]).toEqual([was + 100, 4654]);
    expect(r.updates.some((u) => u.sheet === 1 && u.row === 4 && u.col === 1)).toBe(true);
    // One edit: the cell typed in, and the pivot table noted as worked out again.
    expect(s.editedCount).toBe(2);
    expect(s.changes().map((c) => (c.note ? english.text(c.note) : `${c.sheet}!${c.ref}`))).toEqual(['Data!C2', 'Pivot table “UnitsByRegion” worked out again from its source']);
    s.apply({ type: 'restore', cells: r.undo });
    expect(shown(s, 'B5')).toBe(was);
  });

  it('grows with a new item, moving its read-only region and its saved location, and keeps the cache refreshing on open', () => {
    const s = open('pivot.xlsx');
    s.apply({ type: 'input', sheet: 0, row: 1, col: 0, text: 'Central' });
    expect(column(s, 'A', 3, 9)).toEqual(['Row Labels', 'East', 'North', 'South', 'West', 'Central', 'Grand Total']);
    const r = s.apply({ type: 'input', sheet: 0, row: 2, col: 0, text: 'Central' });
    expect(r.blocked).toBeUndefined();
    expect(s.model.sheets[1].info.blocked.find((b) => b.kind === 'pivot')).toMatchObject({ r1: 2, c1: 0, r2: 8, c2: 1 });
    const p = parts(s.save());
    expect(p['xl/pivotTables/pivotTable1.xml']).toContain('<location ref="A3:B9" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>');
    expect(p['xl/pivotCache/pivotCacheDefinition1.xml']).toContain('refreshOnLoad="1"');
  });

  it("won't grow over cells that have something in them, and says so", () => {
    const s = open('pivot.xlsx');
    s.apply({ type: 'input', sheet: 1, row: 8, col: 0, text: 'note' });
    const r = s.apply({ type: 'input', sheet: 0, row: 1, col: 0, text: 'Central' });
    expect(r.warnings.map((w) => english.text(w))).toContain('Pivot table “UnitsByRegion” wasn\'t worked out again: it would cover A3:B9, where there are other cells.');
    expect(shown(s, 'A9')).toBe('note');
  });
});

import { strToU8, zipSync } from 'fflate';
import { sample } from './helpers';

describe('a pivot table whose source is a name', () => {
  it('is worked out from the range the name gives', () => {
    const p = parts(sample('pivot.xlsx'));
    p['xl/workbook.xml'] = p['xl/workbook.xml'].replace(/<definedNames>|(<\/sheets>)/, (m) => (m === '<definedNames>' ? '<definedNames><definedName name="Sales">Data!$A$1:$E$101</definedName>' : `${m}<definedNames><definedName name="Sales">Data!$A$1:$E$101</definedName></definedNames>`));
    p['xl/pivotCache/pivotCacheDefinition1.xml'] = p['xl/pivotCache/pivotCacheDefinition1.xml'].replace('<worksheetSource ref="A1:E101" sheet="Data"/>', '<worksheetSource name="Sales"/>');
    const s = WorkbookSession.open(zipSync(Object.fromEntries(Object.entries(p).map(([k, v]) => [k, strToU8(v)]))), 'pivot.xlsx', 'en-US');
    s.apply({ type: 'input', sheet: 0, row: 1, col: 0, text: 'Central' });
    expect(column(s, 'A', 3, 9)).toEqual(['Row Labels', 'East', 'North', 'South', 'West', 'Central', 'Grand Total']);
  });

  it("shows values in their data field's number format", () => {
    const p = parts(sample('pivot.xlsx'));
    p['xl/pivotTables/pivotTable1.xml'] = p['xl/pivotTables/pivotTable1.xml'].replace('<dataField name="Sum of Units" fld="2"', '<dataField name="Sum of Units" fld="2" numFmtId="4"');
    const s = WorkbookSession.open(zipSync(Object.fromEntries(Object.entries(p).map(([k, v]) => [k, strToU8(v)]))), 'pivot.xlsx', 'en-US');
    const r = s.apply({ type: 'refreshPivots' });
    expect(r.newStyles.map((n) => n.style.numFmt)).toContain('#,##0.00');
    expect(s.model.styles.get(s.current(1, 4, 1).style).numFmt).toBe('#,##0.00');
  });

  it('is refreshed on demand, and there being none to refresh is said', () => {
    const none = open('orders.xlsx');
    expect(english.text(none.apply({ type: 'refreshPivots' }).rejected!)).toBe('There are no pivot tables to refresh.');
    // The Data sheet has none of its own; the Pivot sheet has one.
    const s = open('pivot.xlsx');
    expect(english.text(s.apply({ type: 'refreshPivots', sheet: 0 }).rejected!)).toBe('There are no pivot tables to refresh.');
    expect(s.apply({ type: 'refreshPivots', sheet: 1 }).rejected).toBeUndefined();
  });
});

import { afterEach, vi } from 'vitest';
import type { WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

describe('refreshing from the editor', () => {
  let store: WorkbookStore;
  afterEach(() => store?.dispose());
  vi.setConfig({ testTimeout: 10_000 });

  it('knows when the active cell is in a pivot table, and refreshes them as one step', async () => {
    store = await testStore();
    await opened(store, sample('pivot.xlsx'), 'pivot.xlsx');
    store.setActive(1);
    store.setSelection({ row: 4, col: 1, rects: [{ r1: 4, c1: 1, r2: 4, c2: 1 }] });
    expect(store.inPivot).toBe(true);
    store.setSelection({ row: 20, col: 5, rects: [{ r1: 20, c1: 5, r2: 20, c2: 5 }] });
    expect(store.inPivot).toBe(false);
    await store.refreshPivots();
    await store.refreshPivots(true);
    expect(store.history.map((h) => english.text(h.action))).toEqual([]);
  });

  it('moves a grown pivot table read-only region on the page', async () => {
    store = await testStore();
    await opened(store, sample('pivot.xlsx'), 'pivot.xlsx');
    await store.input(1, 0, 'Central');
    expect(store.sheets[1].info.blocked.find((b) => b.kind === 'pivot')).toMatchObject({ r1: 2, r2: 8 });
  });
});
