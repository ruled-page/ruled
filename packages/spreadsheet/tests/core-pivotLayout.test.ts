import { describe, expect, it } from 'vitest';
import { layoutPivot, summarize, type PivotField, type PivotModel, type PivotValue } from '../src/core/pivotTable';

const field = (name: string, extra: Partial<PivotField> = {}): PivotField => ({
  name,
  items: [],
  compact: true,
  outline: true,
  subtotalTop: true,
  subtotals: true,
  sort: 'manual',
  showAll: false,
  ...extra,
});

const model = (extra: Partial<PivotModel> = {}): PivotModel => ({
  name: 'P',
  location: { ref: { r1: 2, c1: 0, r2: 2, c2: 0 }, firstHeaderRow: 1, firstDataRow: 1, firstDataCol: 1 },
  fields: [field('Region', { axis: 'row' }), field('Product', { axis: 'row' }), field('Units')],
  rows: [0, 1],
  cols: [],
  pages: [],
  data: [{ name: 'Sum of Units', field: 2, fn: 'sum' }],
  rowGrand: true,
  colGrand: true,
  captions: { data: 'Values', grand: 'Grand Total', rowHeader: 'Row Labels', colHeader: 'Column Labels' },
  ...extra,
});

const RECORDS: PivotValue[][] = [
  ['East', 'Gadget', 1],
  ['East', 'Gizmo', 2],
  ['North', 'Gadget', 4],
];

/** The laid-out table as rows of what each cell shows, empty cells as ''. */
function grid(m: PivotModel, records = RECORDS): string[][] {
  const out = layoutPivot(m, records, (v) => String(v));
  const top = Math.min(0, ...out.cells.map((c) => c.row));
  const rows = Array.from({ length: out.height - top }, () => Array.from({ length: out.width }, () => ''));
  for (const c of out.cells) rows[c.row - top][c.col] = typeof c.value === 'object' && c.value ? c.value.error : String(c.value);
  return rows;
}

describe('laying out a pivot table', () => {
  it('in compact form: one column of labels, each item over its details with its subtotal', () => {
    expect(grid(model())).toEqual([
      ['Row Labels', 'Sum of Units'],
      ['East', '3'],
      ['Gadget', '1'],
      ['Gizmo', '2'],
      ['North', '4'],
      ['Gadget', '4'],
      ['Grand Total', '7'],
    ]);
  });

  it('in outline form: a column a field, headed by its name', () => {
    const outline = model({ fields: [field('Region', { axis: 'row', compact: false }), field('Product', { axis: 'row', compact: false }), field('Units')] });
    expect(grid(outline)).toEqual([
      ['Region', 'Product', 'Sum of Units'],
      ['East', '', '3'],
      ['', 'Gadget', '1'],
      ['', 'Gizmo', '2'],
      ['North', '', '4'],
      ['', 'Gadget', '4'],
      ['Grand Total', '', '7'],
    ]);
  });

  it('in tabular form: labels beside their first detail, and subtotals below', () => {
    const tabular = model({ fields: [field('Region', { axis: 'row', compact: false, outline: false }), field('Product', { axis: 'row', compact: false, outline: false }), field('Units')] });
    expect(grid(tabular)).toEqual([
      ['Region', 'Product', 'Sum of Units'],
      ['East', 'Gadget', '1'],
      ['', 'Gizmo', '2'],
      ['East Total', '', '3'],
      ['North', 'Gadget', '4'],
      ['North Total', '', '4'],
      ['Grand Total', '', '7'],
    ]);
  });

  it('with subtotals below and none at all, and no grand totals', () => {
    const below = model({ fields: [field('Region', { axis: 'row', subtotalTop: false }), field('Product', { axis: 'row' }), field('Units')], rowGrand: false });
    expect(grid(below)).toEqual([
      ['Row Labels', 'Sum of Units'],
      ['East', ''],
      ['Gadget', '1'],
      ['Gizmo', '2'],
      ['East Total', '3'],
      ['North', ''],
      ['Gadget', '4'],
      ['North Total', '4'],
    ]);
    const none = model({ fields: [field('Region', { axis: 'row', subtotals: false }), field('Product', { axis: 'row' }), field('Units')] });
    expect(grid(none).slice(1, 3)).toEqual([
      ['East', ''],
      ['Gadget', '1'],
    ]);
  });

  it('across the columns too, with a total column, leaving empty where nothing falls', () => {
    expect(grid(model({ rows: [0], cols: [1], fields: [field('Region', { axis: 'row' }), field('Product', { axis: 'col' }), field('Units')] }))).toEqual([
      ['Sum of Units', 'Column Labels', '', ''],
      ['Row Labels', 'Gadget', 'Gizmo', 'Grand Total'],
      ['East', '1', '2', '3'],
      ['North', '4', '', '4'],
      ['Grand Total', '5', '2', '7'],
    ]);
  });

  it('with several values across, or down the rows', () => {
    const data = [
      { name: 'Sum of Units', field: 2, fn: 'sum' as const },
      { name: 'Count of Units', field: 2, fn: 'count' as const },
    ];
    expect(grid(model({ rows: [1], data }))).toEqual([
      ['Row Labels', 'Sum of Units', 'Count of Units'],
      ['Gadget', '5', '2'],
      ['Gizmo', '2', '1'],
      ['Grand Total', '7', '3'],
    ]);
    expect(grid(model({ rows: [0, -2], data }))).toEqual([
      ['Row Labels', ''],
      ['East', ''],
      ['Sum of Units', '3'],
      ['Count of Units', '2'],
      ['North', ''],
      ['Sum of Units', '4'],
      ['Count of Units', '1'],
      ['Total Sum of Units', '7'],
      ['Total Count of Units', '3'],
    ]);
  });

  it('filtered by a report filter, hidden items, and sorted, with new items after the known ones', () => {
    const items = [{ value: 'North', hidden: false, collapsed: false }, { value: 'East', hidden: false, collapsed: false }];
    const pages = model({ rows: [1], pages: [{ field: 0, item: 1 }], fields: [field('Region', { axis: 'page', items }), field('Product', { axis: 'row' }), field('Units')] });
    expect(grid(pages)).toEqual([
      ['Region', 'East'],
      ['', ''],
      ['Row Labels', 'Sum of Units'],
      ['Gadget', '1'],
      ['Gizmo', '2'],
      ['Grand Total', '3'],
    ]);
    const several = model({ rows: [1], pages: [{ field: 0 }], fields: [field('Region', { axis: 'page', items: [{ ...items[0], hidden: true }, items[1]] }), field('Product', { axis: 'row' }), field('Units')] });
    expect(grid(several).slice(0, 1)).toEqual([['Region', '(Multiple Items)']]);
    expect(grid(several).at(-1)).toEqual(['Grand Total', '3']);
    const known = model({ rows: [0], fields: [field('Region', { axis: 'row', items }), field('Product'), field('Units')] });
    expect(grid(known, [...RECORDS, ['Central', 'Gizmo', 8]]).map((r) => r[0])).toEqual(['Row Labels', 'North', 'East', 'Central', 'Grand Total']);
    const sorted = model({ rows: [0], fields: [field('Region', { axis: 'row', items, sort: 'descending' }), field('Product'), field('Units')] });
    expect(grid(sorted).map((r) => r[0])).toEqual(['Row Labels', 'North', 'East', 'Grand Total']);
    // A collapsed item shows only its subtotal; a blank one is "(blank)".
    const collapsed = model({ fields: [field('Region', { axis: 'row', items: [{ value: 'East', hidden: false, collapsed: true }] }), field('Product', { axis: 'row' }), field('Units')] });
    expect(grid(collapsed, [...RECORDS, [null, 'Gadget', 1]]).map((r) => r[0])).toEqual(['Row Labels', 'East', 'North', 'Gadget', '(blank)', 'Gadget', 'Grand Total']);
  });
});

describe('laying out less usual pivot tables', () => {
  it('in tabular form with values down the rows, and items of every kind sorted', () => {
    const data = [
      { name: 'Sum of Units', field: 2, fn: 'sum' as const },
      { name: 'Max of Units', field: 2, fn: 'max' as const },
    ];
    const tabular = model({
      rows: [0, -2],
      data,
      fields: [field('Key', { axis: 'row', compact: false, outline: false, sort: 'ascending' }), field('Product'), field('Units')],
    });
    const records: PivotValue[][] = [
      ['b', 'x', 1],
      [2, 'x', 3],
      [true, 'x', 4],
      ['a', 'x', 5],
    ];
    expect(grid(tabular, records).map((r) => r.join('|'))).toEqual([
      'Key|Values|',
      '2|Sum of Units|3',
      '|Max of Units|3',
      'a|Sum of Units|5',
      '|Max of Units|5',
      'b|Sum of Units|1',
      '|Max of Units|1',
      'true|Sum of Units|4',
      '|Max of Units|4',
      'Total Sum of Units||13',
      'Total Max of Units||5',
    ]);
  });
});

describe("a pivot table's summaries", () => {
  const values: PivotValue[] = [2, 4, 'x', null, 6];
  it('sum, count and average as Excel does, counting text only where Count does', () => {
    expect([summarize('sum', values), summarize('count', values), summarize('countNums', values), summarize('average', values)]).toEqual([12, 4, 3, 4]);
    expect([summarize('max', values), summarize('min', values), summarize('product', values)]).toEqual([6, 2, 48]);
    expect(summarize('stdDev', values)).toBeCloseTo(2, 10);
    expect(summarize('var', values)).toBeCloseTo(4, 10);
    expect(summarize('stdDevp', values)).toBeCloseTo(Math.sqrt(8 / 3), 10);
    expect(summarize('varp', values)).toBeCloseTo(8 / 3, 10);
  });

  it('give errors where Excel would, and carry an error in the data', () => {
    expect([summarize('average', ['x']), summarize('stdDev', [1]), summarize('var', [1]), summarize('stdDevp', []), summarize('varp', [])]).toEqual([
      { error: '#DIV/0!' },
      { error: '#DIV/0!' },
      { error: '#DIV/0!' },
      { error: '#DIV/0!' },
      { error: '#DIV/0!' },
    ]);
    expect([summarize('max', []), summarize('min', []), summarize('product', [])]).toEqual([0, 0, 0]);
    expect(summarize('sum', [1, { error: '#N/A' }])).toEqual({ error: '#N/A' });
    expect(summarize('count', [1, { error: '#N/A' }])).toBe(2);
  });
});

import { readPivot } from '../src/core/pivotTable';
import { parseXml } from '../src/core/xml';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const TABLE = (fields: string, more = '') =>
  `<pivotTableDefinition ${NS} name="T" outline="1"><location ref="A3:B6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="2">${fields}</pivotFields>` +
  `<rowFields count="1"><field x="0"/></rowFields>${more}</pivotTableDefinition>`;
const CACHE = (shared: string, extra = '', source = '<worksheetSource ref="A1:B9" sheet="Data"/>') =>
  `<pivotCacheDefinition ${NS}><cacheSource type="worksheet">${source}</cacheSource><cacheFields count="2"><cacheField name="Key" numFmtId="14"><sharedItems>${shared}</sharedItems>${extra}</cacheField><cacheField name="Amount" numFmtId="0"><sharedItems/></cacheField></cacheFields></pivotCacheDefinition>`;
const read = (table: string, cache?: string) => readPivot(parseXml(table), cache === undefined ? undefined : parseXml(cache));

describe("reading a pivot table's definition", () => {
  it('reads every kind of cached item, its fields, filters and data fields', () => {
    const fields =
      '<pivotField axis="axisRow" compact="0" outline="0" subtotalTop="0" defaultSubtotal="0" sortType="ascending" showAll="0"><items count="6">' +
      '<item x="0"/><item x="1" h="1"/><item x="2" sd="0"/><item x="3" n="Missing"/><item x="4"/><item x="5"/><item t="default"/></items></pivotField><pivotField dataField="1" name="Takings"/>';
    const m = read(
      TABLE(fields, '<colFields count="1"><field x="-2"/></colFields><pageFields count="1"><pageField fld="1" item="0"/></pageFields><dataFields count="1"><dataField name="Max of Takings" fld="1" subtotal="max" numFmtId="4"/></dataFields>'),
      CACHE('<s v="a"/><n v="2"/><b v="1"/><m/><d v="2026-01-05T00:00:00"/><e v="#N/A"/>'),
    );
    expect(m.unsupported).toBeUndefined();
    expect(m.fields[0].items.map((i) => [i.value, i.hidden, i.collapsed, i.name])).toEqual([
      ['a', false, false, undefined],
      [2, true, false, undefined],
      [true, false, true, undefined],
      [null, false, false, 'Missing'],
      [46027, false, false, undefined],
      [{ error: '#N/A' }, false, false, undefined],
    ]);
    expect(m.fields[0]).toMatchObject({ axis: 'row', compact: false, outline: false, subtotalTop: false, subtotals: false, sort: 'ascending', showAll: false, numFmtId: 14 });
    expect(m.fields[1].name).toBe('Takings');
    expect([m.rows, m.cols, m.pages, m.data]).toEqual([[0], [-2], [{ field: 1, item: 0 }], [{ name: 'Max of Takings', field: 1, fn: 'max', numFmtId: 4 }]]);
    expect(m.source).toEqual({ sheet: 'Data', ref: 'A1:B9', name: undefined });
  });

  it("knows the pivot tables it can't work out, which keep what Excel saved", () => {
    const plain = '<pivotField axis="axisRow"/><pivotField dataField="1"/>';
    expect(read(TABLE(plain)).unsupported).toBe('cache');
    expect(read(TABLE(plain), CACHE('', '', '')).unsupported).toBe('source');
    expect(read(TABLE(plain), CACHE('', '<fieldGroup base="0"/>')).unsupported).toBe('grouped');
    expect(read(TABLE('<pivotField axis="axisRow"><items count="1"><item x="0" f="1"/></items></pivotField><pivotField/>'), CACHE('<s v="a"/>')).unsupported).toBe('calculatedItem');
    expect(read(TABLE(plain, '<dataFields count="1"><dataField fld="1" showDataAs="percentOfSomething"/></dataFields>'), CACHE('')).unsupported).toBe('showDataAs');
    expect(read(TABLE(plain, '<dataFields count="1"><dataField fld="1" subtotal="median"/></dataFields>'), CACHE('')).unsupported).toBe('function');
  });
});
