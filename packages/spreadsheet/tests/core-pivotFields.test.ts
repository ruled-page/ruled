import { strToU8, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import type { PivotSetup } from '../src/core/pivotTable';
import { WorkbookSession } from '../src/core/session';
import type { PivotInfo } from '../src/core/types';
import { english } from '../src/i18n';
import { open, parts, sample } from './helpers';

/** The demo's pivot.xlsx: Data!A1:E101 (Region, Product, Units, Price, Revenue), and on Pivot!A3:B8 the units summed by region. */
const PATH = 'xl/pivotTables/pivotTable1.xml';
const pivotOf = (s: WorkbookSession): PivotInfo => s.loadResult().sheets[1].pivots![0];
const set = (s: WorkbookSession, change: Partial<PivotSetup>) => s.apply({ type: 'layout', sheet: 1, op: { kind: 'pivot', path: PATH, setup: { ...pivotOf(s).setup, ...change } } });
/** The Pivot sheet's cells, row by row, '.' where empty. */
const grid = (s: WorkbookSession, rows: number, cols: number) =>
  Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => {
      const k = s.current(1, r, c).content;
      return k.kind === 'empty' ? '.' : 'value' in k ? k.value : '?';
    }),
  );
const reopen = (bytes: Uint8Array) => WorkbookSession.open(bytes, 'pivot.xlsx', 'en-US');

describe('a pivot table as the PivotTable Fields pane shows it', () => {
  it('lists its fields with their items, which are numbers, and where each is', () => {
    const p = pivotOf(open('pivot.xlsx'));
    expect(p).toMatchObject({ path: PATH, name: 'UnitsByRegion', regions: [{ r1: 2, c1: 0, r2: 7, c2: 1 }] });
    expect(p.kept).toBeUndefined();
    expect(p.fields.map((f) => [f.name, f.numeric])).toEqual([
      ['Region', false],
      ['Product', false],
      ['Units', true],
      ['Price', true],
      ['Revenue', true],
    ]);
    expect(p.fields[0].items).toEqual([
      { key: 'seast', label: 'East' },
      { key: 'snorth', label: 'North' },
      { key: 'ssouth', label: 'South' },
      { key: 'swest', label: 'West' },
    ]);
    expect(p.setup).toEqual({ rows: [0], cols: [], pages: [], data: [{ field: 2, fn: 'sum', name: 'Sum of Units' }], hidden: {} });
  });

  it("says why one Excel's features the editor doesn't work out can't be changed", () => {
    const p = parts(sample('pivot.xlsx'));
    p[PATH] = p[PATH].replace('<dataField name="Sum of Units" fld="2"', '<dataField name="Sum of Units" fld="2" subtotal="median"');
    const s = reopen(zipSync(Object.fromEntries(Object.entries(p).map(([k, v]) => [k, strToU8(v)]))));
    expect(english.text(pivotOf(s).kept!)).toContain("summarizes values in a way the editor doesn't work out");
    expect(english.text(set(s, { cols: [1] }).rejected!)).toContain("summarizes values in a way the editor doesn't work out");
  });
});

describe('changing a pivot table’s fields', () => {
  it('lays it out again with a field across the columns, as one step that undoes and redoes', () => {
    const s = open('pivot.xlsx');
    const r = set(s, { cols: [1] });
    expect(r.rejected).toBeUndefined();
    expect(grid(s, 9, 5).slice(2)).toEqual([
      ['Sum of Units', 'Column Labels', '.', '.', '.'],
      ['Row Labels', 'Gadget', 'Gizmo', 'Widget', 'Grand Total'],
      ['East', 86, 307, 265, 658],
      ['North', 461, 607, 247, 1315],
      ['South', 450, 284, 398, 1132],
      ['West', 516, 370, 563, 1449],
      ['Grand Total', 1513, 1568, 1473, 4554],
    ]);
    // Its read-only region and the pane's lists follow it.
    expect(r.blocked?.[0].regions.find((b) => b.kind === 'pivot')).toMatchObject({ r1: 2, c1: 0, r2: 8, c2: 4 });
    expect(r.pivots?.[0].list[0].setup.cols).toEqual([1]);
    // It is a change of its own, not of its cells.
    expect(s.changes().map((c) => (c.note ? english.text(c.note) : c.ref))).toEqual(['Pivot table “UnitsByRegion”: fields changed']);

    const undone = s.apply({ type: 'restore', cells: r.undo, layout: r.layout!.undo });
    expect([undone.skipped, undone.rejected]).toEqual([0, undefined]);
    expect(grid(s, 9, 5)[2]).toEqual(['Row Labels', 'Sum of Units', '.', '.', '.']);
    expect([s.editedCount, pivotOf(s).setup.cols]).toEqual([0, []]);
    s.apply({ type: 'restore', cells: r.redo, layout: r.layout!.applied });
    expect(grid(s, 9, 5)[3]).toEqual(['Row Labels', 'Gadget', 'Gizmo', 'Widget', 'Grand Total']);
  });

  it('hides the items filtered out, and picks a report filter’s item above it, moving down to make room', () => {
    const s = open('pivot.xlsx');
    set(s, { hidden: { 0: ['swest', 'seast'] } });
    expect(grid(s, 6, 2).slice(2)).toEqual([
      ['Row Labels', 'Sum of Units'],
      ['North', 1315],
      ['South', 1132],
      ['Grand Total', 2447],
    ]);
    // Two filters need A1:B2 and a row between, so the table moves from A3 to A4.
    set(s, { rows: [], pages: [{ field: 0, item: 'snorth' }, { field: 1 }], hidden: {} });
    expect(grid(s, 5, 2)).toEqual([
      ['Region', 'North'],
      ['Product', '(All)'],
      ['.', '.'],
      ['Sum of Units', '.'],
      [1315, '.'],
    ]);
    // Its read-only cells are its table and its filters, not the cells beside them.
    expect(s.model.sheets[1].info.blocked.filter((b) => b.kind === 'pivot').map(({ r1, c1, r2, c2 }) => ({ r1, c1, r2, c2 }))).toEqual([
      { r1: 3, c1: 0, r2: 4, c2: 0 },
      { r1: 0, c1: 0, r2: 1, c2: 1 },
    ]);
  });

  it('summarises several values each its own way, the Values field across', () => {
    const s = open('pivot.xlsx');
    set(s, { data: [{ field: 2, fn: 'sum', name: 'Sum of Units' }, { field: 4, fn: 'max', name: 'Max of Revenue' }] });
    expect(grid(s, 4, 3)[2]).toEqual(['Row Labels', 'Sum of Units', 'Max of Revenue']);
    expect(pivotOf(s).setup.cols).toEqual([-2]);
  });

  it("won't grow over other cells, and says so", () => {
    const s = open('pivot.xlsx');
    s.apply({ type: 'input', sheet: 1, row: 2, col: 3, text: 'note' });
    expect(english.text(set(s, { cols: [1] }).rejected!)).toContain('it would cover A3:E9');
    expect(grid(s, 4, 3)[2]).toEqual(['Row Labels', 'Sum of Units', '.']);
  });

  it('can have no values and no fields, as a new pivot table has none', () => {
    const s = open('pivot.xlsx');
    set(s, { data: [] });
    expect(grid(s, 9, 2).slice(2).map((r) => r[1])).toEqual(['.', '.', '.', '.', '.', '.', '.']);
    expect(s.model.sheets[1].info.blocked.find((b) => b.kind === 'pivot')).toMatchObject({ r1: 2, c2: 0 });
    set(s, { rows: [], data: [] });
    expect(grid(s, 4, 2)[2]).toEqual(['.', '.']);
  });
});

describe('a pivot table whose fields were changed, saved', () => {
  it('has its definition written again, its cache refreshed on opening, and opens laid out the same', () => {
    const s = open('pivot.xlsx');
    set(s, { rows: [1], cols: [], pages: [{ field: 0, item: 'snorth' }], data: [{ field: 2, fn: 'sum', name: 'Sum of Units' }, { field: 4, fn: 'average', name: 'Average of Revenue' }], hidden: { 1: ['sgizmo'] } });
    const bytes = s.save();
    const p = parts(bytes);
    const def = p[PATH];
    expect(def).toContain('<location ref="A3:C6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1" rowPageCount="1" colPageCount="1"/>');
    expect(def).toContain('<pivotField showAll="0" axis="axisPage"><items count="5"><item x="3"/><item x="0"/><item x="2"/><item x="1"/><item t="default"/></items></pivotField>');
    // The products, new to the rows, sorted as Excel first lists them, Gizmo hidden.
    expect(def).toContain('<pivotField showAll="0" axis="axisRow"><items count="4"><item x="2"/><item x="0" h="1"/><item x="1"/><item t="default"/></items></pivotField>');
    expect(def).toContain('<rowFields count="1"><field x="1"/></rowFields>');
    expect(def).toContain('<colFields count="1"><field x="-2"/></colFields><pageFields count="1"><pageField fld="0" item="1" hier="-1"/></pageFields>');
    expect(def).toContain('<dataField name="Average of Revenue" fld="4" subtotal="average" baseField="0" baseItem="0"/>');
    expect(def).not.toContain('<rowItems');
    expect(p['xl/pivotCache/pivotCacheDefinition1.xml']).toContain('refreshOnLoad="1"');

    const again = reopen(bytes);
    expect(pivotOf(again).setup).toEqual(pivotOf(s).setup);
    expect(again.apply({ type: 'refreshPivots' }).undo).toEqual([]);
  });

  it('writes the cache again from the source, keeping the shared items’ order, with records where it had them', () => {
    const p = parts(sample('pivot.xlsx'));
    const cachePath = 'xl/pivotCache/pivotCacheDefinition1.xml';
    p[cachePath] = p[cachePath].replace('saveData="0"', 'r:id="rId1"');
    p['xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'] =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>';
    p['xl/pivotCache/pivotCacheRecords1.xml'] = '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0"/>';
    p['[Content_Types].xml'] = p['[Content_Types].xml'].replace(
      '</Types>',
      '<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/></Types>',
    );
    const s = reopen(zipSync(Object.fromEntries(Object.entries(p).map(([k, v]) => [k, strToU8(v)]))));
    s.apply({ type: 'input', sheet: 0, row: 1, col: 0, text: 'Central' });
    set(s, { rows: [0, 3] });
    const saved = parts(s.save());
    const cache = saved[cachePath];
    expect(cache).toContain('recordCount="100"');
    expect(cache).toContain('<sharedItems count="5"><s v="North"/><s v="West"/><s v="South"/><s v="East"/><s v="Central"/></sharedItems>');
    // Prices, now on the rows, list their values.
    expect(cache).toMatch(/<cacheField name="Price" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="7.25" maxValue="30" count="3"><n v="[\d.]+"\/>/);
    const records = saved['xl/pivotCache/pivotCacheRecords1.xml'];
    expect(records).toMatch(/^<\?xml[^>]*\?>\n<pivotCacheRecords [^>]*count="100"><r><x v="4"\/><x v="\d"\/><n v="\d+"\/><x v="\d"\/><n v="[\d.]+"\/><\/r>/);
    expect(reopen(s.save()).apply({ type: 'refreshPivots' }).undo).toEqual([]);
  });
});

describe('Insert PivotTable', () => {
  const insert = (s: WorkbookSession, sheet: number, source: string, row: number, col: number) => s.apply({ type: 'layout', sheet, op: { kind: 'insertPivot', source, row, col } });
  const pivots = (s: WorkbookSession, sheet: number) => s.loadResult().sheets[sheet].pivots ?? [];

  it('makes an empty pivot table with a field for each heading, covering Excel’s outline of one until fields are placed', () => {
    const s = open('pivot.xlsx');
    const r = insert(s, 1, 'Data!$A$1:$E$101', 12, 0);
    expect(r.rejected).toBeUndefined();
    const made = pivots(s, 1)[1];
    expect(made).toMatchObject({ path: 'xl/pivotTables/pivotTable2.xml', name: 'PivotTable1', regions: [{ r1: 12, c1: 0, r2: 29, c2: 2 }] });
    expect(made.fields.map((f) => f.name)).toEqual(['Region', 'Product', 'Units', 'Price', 'Revenue']);
    expect(made.setup).toEqual({ rows: [], cols: [], pages: [], data: [], hidden: {} });
    expect(r.layout?.undo).toEqual([{ sheet: 1, op: { kind: 'removePivot', path: made.path } }]);
    expect(s.changes().map((c) => english.text(c.note!))).toEqual(['Pivot table “PivotTable1” inserted']);

    // Its fields placed, a value takes its source's number format, as Excel 365 gives it.
    s.apply({ type: 'layout', sheet: 1, op: { kind: 'pivot', path: made.path, setup: { ...made.setup, rows: [1], data: [{ field: 4, fn: 'sum', name: 'Sum of Revenue' }] } } });
    // Gadget's 1,513 units at $30 make $45,390.
    expect(grid(s, 17, 2).slice(12)).toEqual([
      ['Row Labels', 'Sum of Revenue'],
      ['Gadget', 45390],
      ['Gizmo', 11368],
      ['Widget', 18412.5],
      ['Grand Total', 75170.5],
    ]);
    expect(s.model.styles.get(s.current(1, 13, 1).style).numFmt).toBe(s.model.styles.get(s.current(0, 1, 4).style).numFmt);

    // Undone, it goes, and its cells with it; redone, it comes back where it was, as it was named.
    const undone = s.apply({ type: 'restore', cells: r.undo, layout: r.layout!.undo });
    expect(undone.rejected).toBeUndefined();
    expect(pivots(s, 1)).toHaveLength(1);
    expect(s.model.sheets[1].info.blocked.filter((b) => b.kind === 'pivot')).toHaveLength(1);
    s.apply({ type: 'restore', cells: r.redo, layout: r.layout!.applied });
    expect(pivots(s, 1)[1]).toMatchObject({ path: made.path, name: 'PivotTable1' });
  });

  it('is saved as new parts the workbook finds, and opens as it was', () => {
    const s = open('pivot.xlsx');
    insert(s, 1, 'Data!$A$1:$E$101', 12, 0);
    const path = pivots(s, 1)[1].path;
    s.apply({ type: 'layout', sheet: 1, op: { kind: 'pivot', path, setup: { rows: [0], cols: [1], pages: [], data: [{ field: 2, fn: 'sum', name: 'Sum of Units' }], hidden: {} } } });
    const bytes = s.save();
    const p = parts(bytes);
    expect(p['xl/workbook.xml']).toMatch(/<pivotCaches><pivotCache cacheId="1" r:id="[^"]+"\/><pivotCache cacheId="2" r:id="rId\d+"\/><\/pivotCaches>/);
    expect(p['xl/worksheets/_rels/sheet2.xml.rels']).toContain('Target="../pivotTables/pivotTable2.xml"');
    expect(p['xl/pivotTables/_rels/pivotTable2.xml.rels']).toContain('Target="../pivotCache/pivotCacheDefinition2.xml"');
    expect(p['xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels']).toContain('Target="pivotCacheRecords2.xml"');
    expect(p['[Content_Types].xml']).toContain('<Override PartName="/xl/pivotTables/pivotTable2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>');
    expect(p['xl/pivotTables/pivotTable2.xml']).toContain('name="PivotTable1" cacheId="2"');
    expect(p['xl/pivotCache/pivotCacheDefinition2.xml']).toMatch(/<worksheetSource ref="A1:E101" sheet="Data"\/>.*recordCount|recordCount="100".*<worksheetSource ref="A1:E101" sheet="Data"\/>/);
    const again = reopen(bytes);
    expect(pivots(again, 1).map((q) => q.name)).toEqual(['UnitsByRegion', 'PivotTable1']);
    expect(again.apply({ type: 'refreshPivots' }).undo).toEqual([]);
    // A rebuilt workbook, as after a sheet is added, keeps it.
    const rebuilt = s.restructure({ type: 'addSheet', at: 0, name: 'First' });
    expect(pivots(rebuilt, 2).map((q) => q.name)).toEqual(['UnitsByRegion', 'PivotTable1']);
  });

  it('takes the data a name gives by the name, and an empty one is saved too', () => {
    const p = parts(sample('pivot.xlsx'));
    p['xl/workbook.xml'] = p['xl/workbook.xml'].replace('</sheets>', '</sheets><definedNames><definedName name="Sales">Data!$A$1:$E$101</definedName></definedNames>');
    const s = reopen(zipSync(Object.fromEntries(Object.entries(p).map(([k, v]) => [k, strToU8(v)]))));
    expect(insert(s, 1, 'Sales', 12, 4).rejected).toBeUndefined();
    const saved = parts(s.save());
    expect(saved['xl/pivotCache/pivotCacheDefinition2.xml']).toContain('<worksheetSource name="Sales"/>');
    expect(saved['xl/pivotTables/pivotTable2.xml']).toContain('<location ref="E13:G30"');
  });

  it("says why it can't be made: no such data, data without headings or rows, or no room", () => {
    const s = open('pivot.xlsx');
    const why = (r: ReturnType<typeof insert>) => english.text(r.rejected!);
    expect(why(insert(s, 1, 'Nowhere', 12, 0))).toBe("That isn't a range or a table: Nowhere");
    expect(why(insert(s, 1, 'Data!$A$1:$E$1', 12, 0))).toBe("A pivot table's data needs its headings and at least one row under them.");
    s.apply({ type: 'input', sheet: 0, row: 0, col: 1, text: '' });
    expect(why(insert(s, 1, 'Data!$A$1:$E$101', 12, 0))).toBe("Every column of a pivot table's data needs a heading. Fill in the empty ones, or choose other cells.");
    expect(why(insert(s, 1, 'Data!$A$1:$A$101', 0, 3))).toBe("There's no room for the pivot table at D1:F18: there are other cells there. Choose somewhere empty.");
    expect(english.text(s.apply({ type: 'layout', sheet: 1, op: { kind: 'removePivot', path: 'xl/pivotTables/pivotTable1.xml' } }).rejected!)).toBe('That pivot table is no longer there.');
  });

  it('names headings that repeat apart, and numbers its name after the others', () => {
    const s = open('pivot.xlsx');
    s.apply({ type: 'input', sheet: 0, row: 0, col: 1, text: 'Region' });
    insert(s, 1, 'Data!$A$1:$B$101', 12, 0);
    insert(s, 1, 'Data!$A$1:$B$101', 12, 4);
    expect(pivots(s, 1).slice(1).map((q) => [q.name, q.fields.map((f) => f.name)])).toEqual([
      ['PivotTable1', ['Region', 'Region2']],
      ['PivotTable2', ['Region', 'Region2']],
    ]);
  });
});

describe('Show Values As', () => {
  /**
   * Regions down, products across, units summed:
   *          Gadget Gizmo Widget  Total
   * East        86    307    265    658
   * North      461    607    247   1315
   * South      450    284    398   1132
   * West       516    370    563   1449
   * Total     1513   1568   1473   4554
   */
  const shownAs = (value: Partial<PivotSetup['data'][number]>, rows = [0], cols = [1]) => {
    const s = open('pivot.xlsx');
    s.apply({ type: 'input', sheet: 1, row: 1, col: 3, text: '' });
    const r = set(s, { rows, cols, data: [{ field: 2, fn: 'sum', name: 'Sum of Units', ...value }] });
    expect(r.rejected).toBeUndefined();
    return s;
  };
  /** The values part of the table: B5:E9 with products across. */
  const values = (s: WorkbookSession, rows = 5, cols = 4, top = 4) => grid(s, top + rows, 1 + cols).slice(top).map((r) => r.slice(1));
  const close = (rows: (number | string)[][]) => rows.map((r) => r.map((v) => (typeof v === 'number' ? expect.closeTo(v, 9) : v)));

  it('measures each value against the grand total, its row’s or its column’s, or both as an index', () => {
    expect(values(shownAs({ showAs: 'percentOfTotal' }))[0]).toEqual(close([[86 / 4554, 307 / 4554, 265 / 4554, 658 / 4554]])[0]);
    expect(values(shownAs({ showAs: 'percentOfRow' }))[1]).toEqual(close([[461 / 1315, 607 / 1315, 247 / 1315, 1]])[0]);
    expect(values(shownAs({ showAs: 'percentOfCol' })).map((r) => r[0])).toEqual(close([[86 / 1513, 461 / 1513, 450 / 1513, 516 / 1513, 1]])[0]);
    expect(values(shownAs({ showAs: 'index' }))[0][0]).toBeCloseTo((86 * 4554) / (658 * 1513), 9);
    // Percentages show as Excel shows them.
    const s = shownAs({ showAs: 'percentOfTotal' });
    expect(s.model.styles.get(s.current(1, 4, 1).style).numFmt).toBe('0.00%');
  });

  it('measures against another item of a base field, the one before or after, leaving the base item and totals empty', () => {
    const north = values(shownAs({ showAs: 'difference', baseField: 0, baseItem: 'snorth' }));
    expect(north).toEqual([
      [86 - 461, 307 - 607, 265 - 247, 658 - 1315],
      ['.', '.', '.', '.'],
      [450 - 461, 284 - 607, 398 - 247, 1132 - 1315],
      [516 - 461, 370 - 607, 563 - 247, 1449 - 1315],
      ['.', '.', '.', '.'],
    ]);
    expect(values(shownAs({ showAs: 'percent', baseField: 0, baseItem: 'previous' })).map((r) => r[0])).toEqual(close([[1, 461 / 86, 450 / 461, 516 / 450, '.']])[0]);
    expect(values(shownAs({ showAs: 'percentDiff', baseField: 0, baseItem: 'next' })).map((r) => r[0])).toEqual(close([[(86 - 461) / 461, (461 - 450) / 450, (450 - 516) / 516, '.', '.']])[0]);
    // A base field the cells aren't under leaves them all empty.
    expect(values(shownAs({ showAs: 'difference', baseField: 3, baseItem: 'previous' }))[0]).toEqual(['.', '.', '.', '.']);
  });

  it('runs totals down a base field, and ranks its items', () => {
    expect(values(shownAs({ showAs: 'runTotal', baseField: 0, baseItem: 'previous' })).map((r) => r[0])).toEqual([86, 547, 997, 1513, '.']);
    expect(values(shownAs({ showAs: 'percentOfRunningTotal', baseField: 0, baseItem: 'previous' })).map((r) => r[3])).toEqual(close([[658 / 4554, 1973 / 4554, 3105 / 4554, 1, '.']])[0]);
    expect(values(shownAs({ showAs: 'rankDescending', baseField: 0, baseItem: 'previous' })).map((r) => r[0])).toEqual([4, 2, 3, 1, '.']);
    expect(values(shownAs({ showAs: 'rankAscending', baseField: 1, baseItem: 'previous' }))[0]).toEqual([1, 3, 2, '.']);
  });

  it('measures against the parent row’s total, or a parent item’s', () => {
    // Regions with their products under them, in one column: East, Gadget, Gizmo, Widget, North…
    // With nothing across, the values start a row higher.
    const parentRow = values(shownAs({ showAs: 'percentOfParentRow' }, [0, 1], []), 6, 1, 3).map((r) => r[0]);
    expect(parentRow).toEqual(close([[658 / 4554, 86 / 658, 307 / 658, 265 / 658, 1315 / 4554, 461 / 1315]])[0]);
    const parent = values(shownAs({ showAs: 'percentOfParent', baseField: 0, baseItem: 'previous' }, [0, 1], []), 6, 1, 3).map((r) => r[0]);
    expect(parent).toEqual(close([[1, 86 / 658, 307 / 658, 265 / 658, 1, 461 / 1315]])[0]);
    expect(values(shownAs({ showAs: 'percentOfParentCol' }), 1, 4)[0]).toEqual(close([[86 / 658, 307 / 658, 265 / 658, 1]])[0]);
  });

  it('is saved in the file format’s attribute, or Excel 2010’s extension, and read back', () => {
    const s = shownAs({ showAs: 'difference', baseField: 0, baseItem: 'snorth' });
    const def = parts(s.save())[PATH];
    // North is the second of the regions as the definition lists them.
    expect(def).toContain('<dataField name="Sum of Units" fld="2" showDataAs="difference" baseField="0" baseItem="1"/>');
    expect(pivotOf(reopen(s.save())).setup.data).toEqual([{ field: 2, fn: 'sum', name: 'Sum of Units', showAs: 'difference', baseField: 0, baseItem: 'snorth' }]);
    const ranked = shownAs({ showAs: 'rankDescending', baseField: 0, baseItem: 'previous' });
    expect(parts(ranked.save())[PATH]).toContain(
      '<dataField name="Sum of Units" fld="2" baseField="0" baseItem="1048828"><extLst><ext uri="{E15A36E0-9728-4e99-A89B-3F7291B0FE68}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataField pivotShowAs="rankDescending"/></ext></extLst></dataField>',
    );
    const again = reopen(ranked.save());
    expect(pivotOf(again).setup.data[0]).toMatchObject({ showAs: 'rankDescending', baseItem: 'previous' });
    expect(again.apply({ type: 'refreshPivots' }).undo).toEqual([]);
  });
});
