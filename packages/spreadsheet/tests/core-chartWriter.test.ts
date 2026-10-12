import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { parseChart, type ChartContext } from '../src/core/charts';
import { chartFromRange, chartPartXml, type ChartCell, type NewChartKind } from '../src/core/chartWriter';
import { WorkbookSession } from '../src/core/session';
import type { Anchor, LayoutOp, Rect } from '../src/core/types';
import { parseXml } from '../src/core/xml';
import { english } from '../src/i18n';
import type { WorkbookStore } from '../src/state/store';
import { open as openSample, parts } from './helpers';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

const rect = (r1: number, c1: number, r2: number, c2: number): Rect => ({ r1, c1, r2, c2 });

/** A grid of cells as Insert Chart sees them; numbers in `dates` columns are dates. */
function grid(rows: (string | number | null)[][], dates: number[] = []) {
  return (row: number, col: number): ChartCell => {
    const value = rows[row]?.[col] ?? null;
    const date = typeof value === 'number' && dates.includes(col);
    return { value, date, format: date ? 'mmm yyyy' : 'General', text: value === null ? '' : String(value) };
  };
}

const SALES = [
  [null, 'Q1', 'Q2'],
  ['North', 10, 12],
  ['South', 7, 9],
  ['East', 4, 5],
];

describe("Insert Chart's guess at what to plot", () => {
  it('names series from a first row of text, and takes categories from a first column of text, both with an empty corner', () => {
    const chart = chartFromRange(rect(0, 0, 3, 2), grid(SALES), 'Sales Data', 'column', 'Chart Title');
    expect(chart.series.map((s) => [s.name?.ref, s.name?.text, s.values.ref, s.values.points])).toEqual([
      ["'Sales Data'!$B$1", 'Q1', "'Sales Data'!$B$2:$B$4", [10, 7, 4]],
      ["'Sales Data'!$C$1", 'Q2', "'Sales Data'!$C$2:$C$4", [12, 9, 5]],
    ]);
    expect(chart.series[0].categories).toEqual({ ref: "'Sales Data'!$A$2:$A$4", points: ['North', 'South', 'East'], numeric: false, format: 'General' });
    // Two series: Excel's placeholder title and a legend.
    expect([chart.title, chart.legend]).toEqual(['Chart Title', true]);
  });

  it('runs series along the rows when the data is wider than it is tall', () => {
    const wide = [
      [null, 'Jan', 'Feb', 'Mar', 'Apr'],
      ['Sales', 1, 2, 3, 4],
    ];
    const chart = chartFromRange(rect(0, 0, 1, 4), grid(wide), 'S', 'line', 'Chart Title');
    expect(chart.series).toHaveLength(1);
    expect(chart.series[0]).toMatchObject({ name: { ref: 'S!$A$2', text: 'Sales' }, values: { ref: 'S!$B$2:$E$2' }, categories: { ref: 'S!$B$1:$E$1' } });
    // One series is titled with its name, and needs no legend.
    expect([chart.title, chart.legend]).toEqual([undefined, false]);
  });

  it('takes dates down the first column as categories, and plots plain numbers with no labels at all', () => {
    const months = [
      ['Month', 'Revenue'],
      [46023, 120],
      [46054, 130],
    ];
    const dated = chartFromRange(rect(0, 0, 2, 1), grid(months, [0]), 'S', 'column', 'Chart Title');
    expect(dated.series[0]).toMatchObject({ name: { text: 'Revenue' }, categories: { points: [46023, 46054], numeric: true, format: 'mmm yyyy' } });
    // A single column's heading names its series.
    const headed = chartFromRange(rect(0, 3, 2, 3), grid([[null, null, null, 'Total'], [null, null, null, 6], [null, null, null, 2]]), 'S', 'column', 'Chart Title');
    expect(headed.series).toEqual([{ name: { ref: 'S!$D$1', text: 'Total' }, values: { ref: 'S!$D$2:$D$3', points: [6, 2], format: 'General' } }]);
    const plain = chartFromRange(rect(0, 0, 2, 0), grid([[1], [2], [3]]), 'S', 'pie', 'Chart Title');
    expect(plain.series).toEqual([{ values: { ref: 'S!$A$1:$A$3', points: [1, 2, 3], format: 'General' } }]);
    // A pie always has a legend, for its slices.
    expect(plain.legend).toBe(true);
  });

  it("gives a scatter chart its first column of numbers as the x values", () => {
    const xy = [
      ['x', 'y'],
      [1, 5],
      [2, 7],
    ];
    const chart = chartFromRange(rect(0, 0, 2, 1), grid(xy), 'S', 'scatter', 'Chart Title');
    expect(chart.series).toHaveLength(1);
    expect(chart.series[0]).toMatchObject({ name: { text: 'y' }, categories: { ref: 'S!$A$2:$A$3', numeric: true }, values: { ref: 'S!$B$2:$B$3' } });
  });
});

const CTX: ChartContext = { theme: [], sheetIndex: new Map([['sales data', 0]]), names: [] };

describe('the chart part', () => {
  it.each<[NewChartKind, string, boolean]>([
    ['column', 'bar', false],
    ['bar', 'bar', true],
    ['line', 'line', false],
    ['area', 'area', false],
    ['pie', 'pie', false],
    ['doughnut', 'doughnut', false],
    ['scatter', 'scatter', false],
  ])('for a %s chart reads back as that chart, with its series where they came from', (kind, drawn, horizontal) => {
    const xml = chartPartXml(chartFromRange(rect(0, 0, 3, 2), grid(SALES), 'Sales Data', kind, 'Chart Title'));
    const spec = parseChart(parseXml(xml), CTX);
    expect(spec.unsupported).toBeUndefined();
    expect(spec.groups[0]).toMatchObject({ kind: drawn, horizontal });
    const s = spec.groups[0].series[0];
    expect(s.values).toMatchObject({ formula: "'Sales Data'!$B$2:$B$4", ref: { sheet: 0, rect: rect(1, 1, 3, 1) }, cache: [10, 7, 4] });
    expect(s.categories?.cache).toEqual(['North', 'South', 'East']);
    expect(s.name?.cache).toEqual(['Q1']);
    expect(spec.legend?.position).toBe('bottom');
  });

  it('escapes what it writes, and keeps a single series titled by its name', () => {
    const xml = chartPartXml(chartFromRange(rect(0, 0, 1, 1), grid([['A & B', 'Sales <net>'], ['x', 3]]), 'S', 'column', 'Chart Title'));
    expect(xml).toContain('<c:v>Sales &lt;net&gt;</c:v>');
    const spec = parseChart(parseXml(xml), { ...CTX, sheetIndex: new Map([['s', 0]]) });
    expect(spec.title?.cache).toEqual(['Sales <net>']);
  });
});

/** Data!A1:C4: regions down the side, two quarters across. */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.addRow([null, 'Q1', 'Q2']);
  ws.addRow(['North', 10, 12]);
  ws.addRow(['South', 7, 9]);
  ws.addRow(['East', 4, 5]);
  wb.addWorksheet('Empty');
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

const at = (row: number, col: number): Anchor => ({ row, col, rowOffPx: 0, colOffPx: 0 });
const insert = (s: WorkbookSession, sheet: number, kind: NewChartKind, r: Rect) =>
  s.apply({ type: 'layout', sheet, op: { kind: 'insertChart', chart: kind, rect: r, from: at(r.r1, r.c2 + 2), to: at(r.r1 + 14, r.c2 + 9), title: 'Chart Title' } });
const charts = (s: WorkbookSession, sheet = 0) => s.model.sheets[sheet].info.objects.filter((o) => o.kind === 'chart');

describe('a chart inserted into a workbook', () => {
  it('is drawn at once, saved as Excel saves one, read back, and undone', () => {
    const s = WorkbookSession.open(bytes, 'charts.xlsx', 'en-US');
    const done = insert(s, 0, 'column', rect(0, 0, 3, 2));
    const [chart] = charts(s);
    expect(chart).toMatchObject({ name: 'Chart 1', title: 'Chart Title', from: at(0, 4), to: at(14, 11) });
    expect(english.text(chart.chartType!)).toBe('Column');
    expect(chart.chart?.groups[0].series).toHaveLength(2);
    // What the page applies, and what redo applies again, is the sheet's objects with the chart.
    expect(done.layout?.applied[0].op.kind).toBe('objects');
    expect(s.changes().map((c) => c.note && english.text(c.note))).toContain('Charts, pictures or shapes changed');

    const saved = s.save();
    const p = parts(saved);
    expect(p['xl/worksheets/sheet1.xml']).toMatch(/<drawing xmlns:r="[^"]+" r:id="rId\d+"\/>/);
    expect(p['xl/worksheets/_rels/sheet1.xml.rels']).toContain('Target="../drawings/drawing1.xml"');
    expect(p['xl/drawings/drawing1.xml']).toContain('<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row>');
    expect(p['xl/drawings/_rels/drawing1.xml.rels']).toContain('Target="../charts/chart1.xml"');
    expect(p['[Content_Types].xml']).toContain('PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"');
    expect(p['[Content_Types].xml']).toContain('PartName="/xl/drawings/drawing1.xml"');
    const again = WorkbookSession.open(saved, 'charts.xlsx', 'en-US');
    expect(charts(again)).toEqual([expect.objectContaining({ name: 'Chart 1', from: at(0, 4), source: { drawing: 'xl/drawings/drawing1.xml', anchor: 0 } })]);
    expect(charts(again)[0].chart?.groups[0].series[1].values.formula).toBe('Data!$C$2:$C$4');

    s.apply({ type: 'restore', cells: [], layout: done.layout!.undo });
    expect(charts(s)).toEqual([]);
    expect(Object.keys(parts(s.save())).some((k) => k.includes('chart'))).toBe(false);
  });

  it('joins a sheet that has charts already, and the file charts can be moved and deleted', () => {
    const s = openSample('kitchen-sink.xlsx');
    expect(charts(s)).toHaveLength(2);
    insert(s, 0, 'pie', rect(4, 0, 7, 1));
    const drawing = () => parts(s.save())['xl/drawings/drawing1.xml'];
    expect(charts(s).map((c) => c.name)).toEqual([charts(s)[0].name, charts(s)[1].name, 'Chart 3']);
    // Each frame in a drawing has an id of its own.
    const ids = [...drawing().matchAll(/cNvPr id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);

    const [first, second, mine] = s.model.sheets[0].info.objects.filter((o) => o.kind === 'chart');
    const objects = s.model.sheets[0].info.objects.filter((o) => o !== first).map((o) => (o === second ? { ...o, from: at(30, 8), to: at(44, 15) } : o));
    s.apply({ type: 'layout', sheet: 0, op: { kind: 'objects', objects } satisfies LayoutOp });
    const saved = s.save();
    const t = WorkbookSession.open(saved, 'k.xlsx', 'en-US');
    expect(charts(t).map((c) => [c.title, c.from])).toEqual([
      [second.title, at(30, 8)],
      [mine.title, mine.from],
    ]);
    // The deleted chart's part and its relationship are gone too.
    const p = parts(saved);
    expect(Object.keys(p).filter((k) => /^xl\/charts\/chart\d+\.xml$/.test(k))).toHaveLength(3);
  });

  it('is refused for a range with no numbers', () => {
    const s = WorkbookSession.open(bytes, 'charts.xlsx', 'en-US');
    const r = insert(s, 1, 'column', rect(0, 0, 3, 3));
    expect(english.text(r.rejected!)).toMatch(/nothing to chart/);
    expect(charts(s, 1)).toEqual([]);
  });
});

describe('Insert Chart in the editor', () => {
  let store: WorkbookStore;
  afterEach(() => store?.dispose());

  it("charts the selection, or the data around the active cell, at Excel's default size beside it, as one step", async () => {
    store = await testStore();
    await opened(store, bytes, 'charts.xlsx');
    store.setSelection({ row: 0, col: 0, rects: [rect(0, 0, 3, 2)] });
    await store.insertChart('line');
    const chart = () => store.view!.info.objects.find((o) => o.kind === 'chart');
    // 480 by 288 pixels from E1: seven 64-pixel columns and 32 more, fourteen 20-pixel rows and 8 more.
    expect(chart()).toMatchObject({ name: 'Chart 1', from: at(0, 4), to: { col: 11, colOffPx: 32, row: 14, rowOffPx: 8 } });
    expect(english.text(store.history.at(-1)!.action)).toBe('Insert chart');
    await store.undo();
    expect(chart()).toBeUndefined();

    store.setSelection({ row: 2, col: 1, rects: [rect(2, 1, 2, 1)] });
    await store.insertChart('pie');
    expect(chart()?.chart?.groups[0].series[0].values.formula).toBe('Data!$B$2:$B$4');
  });
});
