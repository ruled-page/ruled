import type { ChartConfiguration, LegendItem } from 'chart.js';
import { describe, expect, it } from 'vitest';
import { chartKindOf } from '../src/core/chartEdit';
import { objectKey } from '../src/core/objectKey';
import type { SheetObject } from '../src/core/types';
import { english } from '../src/i18n';
import { describeChart, plotChart, type ChartPlot } from '../src/state/charts';
import { SheetView } from '../src/state/sheetView';
import { chartConfig, type ChartLook } from '../src/ui/chartConfig';
import { open } from './helpers';

// The demo's chart gallery: one chart of each of the less common kinds, made by openpyxl.
const session = open('chart-gallery.xlsx');
const r = session.loadResult();
const src = { sheets: r.sheets.map((info, i) => new SheetView(info, r.grids[i])), styles: r.styles, date1904: r.date1904 };
const chart = (title: string): SheetObject => {
  const found = r.sheets[0].objects.find((o) => o.title === title);
  if (!found) throw new Error(`no chart titled ${title}`);
  return found;
};
const plot = (title: string): ChartPlot => plotChart(chart(title).chart!, src);
const LOOK: ChartLook = { palette: [], background: '#fff', text: '#111', muted: '#555', gridline: '#ddd', axis: '#ccc', font: 'Calibri' };
const config = (title: string) => chartConfig(plot(title), LOOK, english.text) as ChartConfiguration;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const datasets = (c: ChartConfiguration) => c.data.datasets as any[];

describe('the chart gallery', () => {
  it('draws every kind but the surface chart, which is kept as it is', () => {
    const drawn = r.sheets[0].objects.filter((o) => o.kind === 'chart').map((o) => [o.title, o.chart?.unsupported ?? o.chart?.groups[0].kind]);
    expect(drawn).toEqual([
      ['Skills (radar)', 'radar'],
      ['Skills (filled radar)', 'radar'],
      ['Price, margin and units (bubble)', 'bubble'],
      ['High, low, close (stock)', 'stock'],
      ['Open, high, low, close (stock)', 'stock'],
      ['Sales (3-D column)', 'bar'],
      ['Sales (3-D line)', 'line'],
      ['Sales (3-D area)', 'area'],
      ['Costs (3-D pie)', 'pie'],
      ['Costs (pie of pie)', 'pie'],
      ['Sales (surface, kept as it is)', 'surfaceChart'],
    ]);
  });
});

describe('radar charts', () => {
  it('draw lines with markers, or filled areas, round the categories', () => {
    const lines = config('Skills (radar)');
    expect(lines.type).toBe('radar');
    expect(lines.data.labels).toEqual(['Speed', 'Strength', 'Stamina', 'Skill', 'Teamwork']);
    expect(datasets(lines).map((d) => [d.label, d.data, d.fill, d.pointRadius > 0])).toEqual([
      ['Ana', [8, 5, 7, 9, 6], false, true],
      ['Ben', [6, 9, 7, 5, 8], false, true],
    ]);
    expect(datasets(config('Skills (filled radar)')).every((d) => d.fill === true && d.pointRadius === 0)).toBe(true);
  });
});

describe('bubble charts', () => {
  it('size bubbles by area, the largest at the full size, and say each point with its size', () => {
    const c = config('Price, margin and units (bubble)');
    expect(c.type).toBe('bubble');
    const points = datasets(c)[0].data as { x: number; y: number; r: number }[];
    expect(points.map((p) => [p.x, p.y])).toEqual([
      [12.5, 0.32],
      [30, 0.41],
      [7.25, 0.18],
      [18, 0.27],
    ]);
    // 2,600 units is the largest; 450 is about 0.42 of it across, as its area is 0.17 of it.
    expect(points[2].r).toBe(30);
    expect(points[1].r).toBeCloseTo(30 * Math.sqrt(450 / 2600), 6);
    expect(describeChart(plot('Price, margin and units (bubble)'), english)).toContain('Products: x 12.5, y 0.32, size 1200');
  });
});

describe('stock charts', () => {
  it('draw a line from each day’s high to its low, the close marked, and leave the lines out of the legend', () => {
    const c = config('High, low, close (stock)');
    const [hiLow, ...series] = datasets(c);
    expect(hiLow.data[0]).toEqual([40.8, 42.9]);
    expect(series.map((s) => [s.label, s.pointRadius > 0])).toEqual([
      ['High', false],
      ['Low', false],
      ['Close', true],
    ]);
    const filter = (c.options!.plugins!.legend!.labels!.filter as (i: LegendItem) => boolean);
    expect([filter({ text: '' } as LegendItem), filter({ text: 'Close' } as LegendItem)]).toEqual([false, true]);
  });

  it('draw open to close bars, light where the price rose and dark where it fell', () => {
    const c = config('Open, high, low, close (stock)');
    const bars = datasets(c)[1];
    expect(bars.data.slice(0, 2)).toEqual([
      [41.2, 42.5],
      [42.5, 41.9],
    ]);
    expect(bars.backgroundColor.slice(0, 2)).toEqual(['#fff', '#111']);
    expect(plot('Open, high, low, close (stock)').categories.slice(0, 2)).toEqual(['1 Oct', '2 Oct']);
  });
});

describe('changing a kind Insert Chart does not make', () => {
  it('keeps the chart its kind when only its title changes', () => {
    const radar = chart('Skills (radar)');
    expect(chartKindOf(radar.chart)).toBeUndefined();
    const s = open('chart-gallery.xlsx');
    s.apply({ type: 'layout', sheet: 0, op: { kind: 'editChart', key: objectKey(radar), edit: { title: 'Players' }, title: 'Chart Title' } });
    const changed = s.model.sheets[0].info.objects.find((o) => objectKey(o) === objectKey(radar))!;
    expect([changed.chart?.groups[0].kind, changed.title, english.text(changed.chartType!)]).toEqual(['radar', 'Players', 'Radar']);
  });
});
