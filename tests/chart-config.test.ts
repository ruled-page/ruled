import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import type { ChartConfiguration, LegendItem } from 'chart.js';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { SheetObject } from '../src/core/types';
import { english } from '../src/i18n';
import { plotChart, type ChartPlot } from '../src/state/charts';
import { SheetView } from '../src/state/sheetView';
import { axisBounds, chartConfig, paletteColor, parsePalette, type ChartLook } from '../src/ui/chartConfig';
import { colsWidth, objectSize, rowsHeight } from '../src/ui/objectGeometry';
import { sheetIndex } from './helpers';
import { sheetInfo } from './state-fixtures';

const session = WorkbookSession.open(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'charts.xlsx'))), 'charts.xlsx', 'en-US');
const r = session.loadResult();
const src = { sheets: r.sheets.map((info, i) => new SheetView(info, r.grids[i])), styles: r.styles, date1904: r.date1904 };
const plot = (title: string): ChartPlot => {
  const spec = r.sheets[sheetIndex(session, 'Charts')].objects.find((o) => o.title === title)?.chart;
  if (!spec) throw new Error(`no chart titled ${title}`);
  return plotChart(spec, src);
};

const LOOK: ChartLook = { palette: [], background: '#fff', text: '#111', muted: '#555', gridline: '#ddd', axis: '#ccc', font: 'Calibri' };
const config = (p: ChartPlot, look: Partial<ChartLook> = {}) => chartConfig(p, { ...LOOK, ...look }, english.text);
/** Options and datasets loosely, to look into them. */
type Loose = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any
const parts = (c: ChartConfiguration) => ({ options: c.options as Loose, datasets: c.data.datasets as Loose[], scales: (c.options as Loose).scales as Loose });
const tick = (scale: Loose, v: number) => scale.ticks.callback(v);
const ACCENTS_2007 = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646'];

describe('the editor’s look for charts', () => {
  it('reads the colours token', () => {
    expect(parsePalette('auto')).toEqual([]);
    expect(parsePalette('  ')).toEqual([]);
    expect(parsePalette('#123456 #abcdef')).toEqual(['#123456', '#abcdef']);
    expect(parsePalette('rgb(1, 2, 3), #fff,hsl(10 20% 30%)')).toEqual(['rgb(1, 2, 3)', '#fff', 'hsl(10 20% 30%)']);
  });

  it('repeats a palette darker, then lighter, past its end', () => {
    const palette = ['#808080', 'red'];
    expect([0, 1, 2, 3, 4, 5].map((i) => paletteColor(palette, i))).toEqual(['#808080', 'red', '#4d4d4d', 'red', '#a6a6a6', 'red']);
  });

  it('starts value axes where Excel does', () => {
    expect(axisBounds([10, 50, null])).toEqual({ beginAtZero: true });
    expect(axisBounds([36_000, 42_000])).toEqual({ beginAtZero: false, suggestedMin: 33_000 });
    expect(axisBounds([-42_000, -36_000])).toEqual({ beginAtZero: false, suggestedMax: -33_000 });
    expect(axisBounds([-50, -10])).toEqual({ beginAtZero: true });
    expect(axisBounds([-5, 5])).toEqual({ beginAtZero: true });
    expect(axisBounds([0, 0])).toEqual({ beginAtZero: true });
    expect(axisBounds([])).toEqual({ beginAtZero: true });
  });
});

describe('Chart.js configurations', () => {
  it('draws clustered columns with the file’s colours, then the workbook theme’s, and its axes', () => {
    const c = config(plot('Quarterly sales'));
    const { options, datasets, scales } = parts(c);
    expect(c.type).toBe('bar');
    expect(c.data.labels).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
    expect(options.indexAxis).toBe('x');
    expect(datasets.map((d) => [d.type, d.label, d.backgroundColor[0], d.stack])).toEqual([
      ['bar', 'North', '#1F77B4', undefined],
      ['bar', 'South', ACCENTS_2007[1], undefined],
      ['bar', 'East', ACCENTS_2007[2], undefined],
    ]);
    // Three bars and a gap of one and a half bars to each category.
    expect(datasets[0].categoryPercentage).toBeCloseTo(3 / 4.5);
    expect(scales.y).toMatchObject({ type: 'linear', position: 'left', min: 0, max: 200, beginAtZero: true, grace: 0, ticks: { stepSize: 50 }, title: { display: true, text: 'Units' } });
    expect(tick(scales.y, 150)).toBe('$150');
    expect(scales.x).toMatchObject({ type: 'category', position: 'bottom', reverse: false, offset: true, title: { text: 'Quarter' }, grid: { display: false } });
    expect(options.plugins.title).toMatchObject({ display: true, text: 'Quarterly sales' });
    expect(options.plugins.legend).toMatchObject({ display: true, position: 'bottom' });
    expect(options).toMatchObject({ animation: false, events: [], responsive: true, maintainAspectRatio: false });
  });

  it('stacks, and stacks to 100%', () => {
    const stacked = parts(config(plot('Stacked columns')));
    expect(stacked.datasets.map((d) => d.stack)).toEqual(['g0', 'g0', 'g0']);
    expect(stacked.datasets[0].categoryPercentage).toBeCloseTo(1 / 2.5);
    const share = parts(config(plot('Share by quarter')));
    const q1 = share.datasets.map((d) => d.data[0]);
    expect(q1.reduce((a: number, b: number) => a + b)).toBeCloseTo(1);
    expect(q1[0]).toBeCloseTo(120 / 270);
    expect(share.scales.y).toMatchObject({ min: 0, max: 1, grace: 0 });
    expect(tick(share.scales.y, 0.5)).toBe('50%');
  });

  it('runs bar charts sideways, first category at the bottom', () => {
    const { options, scales } = parts(config(plot('Horizontal bars')));
    expect(options.indexAxis).toBe('y');
    expect(scales.y).toMatchObject({ type: 'category', position: 'left', reverse: true });
    expect(scales.x).toMatchObject({ type: 'linear', position: 'bottom' });
    expect(options.plugins.legend.position).toBe('top');
  });

  it('draws lines with their width, dashes, markers and smoothing, over any bars', () => {
    const { datasets } = parts(config(plot('Sales trend')));
    expect(datasets[0]).toMatchObject({ type: 'line', borderColor: '#C00000', borderWidth: 3.3, pointRadius: 3.5, pointStyle: 'circle', tension: 0, borderDash: [], fill: false, order: -1 });
    expect(datasets[1]).toMatchObject({ borderDash: [6, 4], pointRadius: 0, pointStyle: 'line', tension: 0.4, borderWidth: 2.5 });
    expect(parts(config(plot('Visits by month'))).datasets[0].spanGaps).toBe(true);
  });

  it('fills areas edge to edge, later ones in front, stacked ones on the one below', () => {
    const area = parts(config(plot('Area')));
    expect(area.datasets.map((d) => [d.type, d.fill, d.order, d.pointRadius])).toEqual([
      ['line', 'origin', -0, 0],
      ['line', 'origin', -1, 0],
      ['line', 'origin', -2, 0],
    ]);
    expect(area.scales.x.offset).toBe(false);
    expect(parts(config(plot('Stacked area'))).datasets.map((d) => [d.fill, d.stack, d.order])).toEqual([
      ['origin', 'g0', 0],
      ['-1', 'g0', 0],
      ['-1', 'g0', 0],
    ]);
  });

  it('draws pies and doughnuts with their slice colours, rotation and hole', () => {
    const pie = config(plot('North by quarter'));
    const { options, datasets } = parts(pie);
    expect(pie.type).toBe('pie');
    expect(datasets[0].backgroundColor).toEqual(['#FF0000', ACCENTS_2007[1], '#00B050', ACCENTS_2007[3]]);
    expect(datasets[0]).toMatchObject({ data: [120, 150, 130, 170], borderColor: '#fff' });
    expect(options).toMatchObject({ rotation: 90, cutout: 0 });
    expect(options.plugins.legend.labels.usePointStyle).toBe(false);
    const doughnut = config(plot('Units by product'));
    expect(doughnut.type).toBe('doughnut');
    expect((doughnut.options as Loose).cutout).toBe('40%');
  });

  it('places scatter points by x and y, with or without lines', () => {
    const scatter = config(plot('Ice cream sales'));
    const { datasets, scales } = parts(scatter);
    expect(scatter.type).toBe('scatter');
    expect(datasets[0].data.slice(0, 2)).toEqual([
      { x: 12, y: 210 },
      { x: 15, y: 260 },
    ]);
    expect(datasets[0]).toMatchObject({ showLine: false, pointRadius: 3.5 });
    expect(scales.x).toMatchObject({ type: 'linear', position: 'bottom', min: 10, title: { text: 'Temperature' } });
    expect(parts(config(plot('Sales by temperature'))).datasets[0].showLine).toBe(true);
  });

  it('puts a combination’s line on the secondary axis, and lists the legend in Excel’s order', () => {
    const { datasets, scales, options } = parts(config(plot('Sales and margin')));
    expect(datasets.map((d) => [d.type, d.yAxisID])).toEqual([
      ['bar', 'y'],
      ['line', 'y2'],
    ]);
    expect(scales.y2).toMatchObject({ position: 'right', grid: { drawOnChartArea: false }, title: { text: 'Margin' } });
    expect(tick(scales.y2, 0.12)).toBe('12%');
    const sort = options.plugins.legend.labels.sort as (a: LegendItem, b: LegendItem) => number;
    expect(sort({ datasetIndex: 1 } as LegendItem, { datasetIndex: 0 } as LegendItem)).toBeGreaterThan(0);
  });

  it('takes the host’s palette over the theme’s, and leaves deleted legend entries out', () => {
    const p = plot('Stacked columns');
    expect(parts(config(p, { palette: ['#000000'] })).datasets.map((d) => d.backgroundColor[0])).toEqual(['#000000', '#000000', '#4d4d4d']);
    expect(parts(config({ ...p, accents: [] })).datasets[0].backgroundColor[0]).toBe('#4472C4');
    const { options } = parts(config({ ...p, legend: { position: 'right', deleted: [1] } }));
    const filter = options.plugins.legend.labels.filter as (item: LegendItem) => boolean;
    expect([0, 1, 2].map((datasetIndex) => filter({ datasetIndex } as LegendItem))).toEqual([true, false, true]);
    const pie = parts(config({ ...plot('North by quarter'), legend: { position: 'right', deleted: [0] } })).options.plugins.legend.labels;
    expect(pie.filter({ index: 0 } as LegendItem)).toBe(false);
    expect(pie.sort({ index: 2 } as LegendItem, { index: 1 } as LegendItem)).toBeGreaterThan(0);
  });

  it('plots empty cells as zero when the chart asks, varies a lone series’ colours, and handles log axes', () => {
    const p = plot('East');
    const zero = parts(config({ ...p, blanks: 'zero', groups: [{ ...p.groups[0], kind: 'line', varyColors: true, series: [{ ...p.groups[0].series[0], values: [1, null] }] }] }));
    expect(zero.datasets[0].data).toEqual([1, 0]);
    const varied = parts(config({ ...p, groups: [{ ...p.groups[0], varyColors: true }] }));
    expect(varied.datasets[0].backgroundColor).toEqual(ACCENTS_2007.slice(0, 4));
    const log = parts(config({ ...p, axes: { ...p.axes, value: { reversed: true, hidden: true, gridlines: true, logBase: 10, format: '0' } } }));
    expect(log.scales.y).toMatchObject({ type: 'logarithmic', reverse: true, display: false });
    expect(log.scales.y.beginAtZero).toBeUndefined();
    const bare = parts(config({ ...p, title: undefined, legend: undefined, axes: {} }));
    expect(bare.options.plugins.title.display).toBe(false);
    expect(bare.options.plugins.legend).toMatchObject({ display: false, position: 'right' });
    expect(bare.scales.y.grid.display).toBe(true);
  });

  it('draws an empty chart as empty columns', () => {
    const c = config({ title: undefined, categories: [], groups: [], axes: {}, blanks: 'gap', accents: [] });
    expect(c.type).toBe('line');
    expect(c.data.datasets).toEqual([]);
  });
});

describe('where objects sit over the grid', () => {
  const view = new SheetView(sheetInfo({ maxRow: 20, maxCol: 10, colWidths: { 1: 100 }, rowHeights: { 2: 40 }, hiddenRows: [3], hiddenCols: [2] }), null);
  const obj = (from: [number, number, number, number], to?: [number, number, number, number], size?: { w: number; h: number }): SheetObject => ({
    kind: 'chart',
    name: 'Chart 1',
    from: { row: from[0], col: from[1], rowOffPx: from[2], colOffPx: from[3] },
    to: to && { row: to[0], col: to[1], rowOffPx: to[2], colOffPx: to[3] },
    sizePx: size,
  });

  it('measures rows and columns, leaving out hidden ones and counting ones past the grid', () => {
    expect(rowsHeight(view, 0, 5)).toBe(20 + 20 + 40 + 20);
    expect(colsWidth(view, 0, 4)).toBe(64 + 100 + 64);
    expect(rowsHeight(view, 1_000, 1_002)).toBe(40);
  });

  it('sizes an object between its anchors, or as the file gives it', () => {
    expect(objectSize(view, obj([1, 0, 5, 10], [5, 3, 10, 20]))).toEqual({ width: 64 + 100 - 10 + 20, height: 20 + 40 + 20 - 5 + 10 });
    expect(objectSize(view, obj([1, 0, 0, 0], undefined, { w: 300, h: 200 }))).toEqual({ width: 300, height: 200 });
    expect(objectSize(view, obj([1, 0, 0, 0]))).toBeUndefined();
    expect(objectSize(view, { kind: 'chart', name: 'x' })).toBeUndefined();
    expect(rowsHeight(new SheetView(sheetInfo(), null), 0, 1)).toBe(20);
  });
});
