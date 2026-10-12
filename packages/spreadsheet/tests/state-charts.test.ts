import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { ChartData, ChartSeries, ChartSpec } from '../src/core/charts';
import { WorkbookSession } from '../src/core/session';
import { english, msg } from '../src/i18n';
import { chartSheets, describeChart, formatValue, plotChart, type ChartSource } from '../src/state/charts';
import { SheetView } from '../src/state/sheetView';
import { sheetIndex } from './helpers';
import { num, PLAIN, sheetInfo, str, style } from './state-fixtures';

/** tests/corpus/charts.xlsx as the editor holds it: the worker's session, and the views the screen reads. */
function workbook() {
  const session = WorkbookSession.open(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'charts.xlsx'))), 'charts.xlsx', 'en-US', { recalc: 'on' });
  session.ensureEngine();
  const r = session.loadResult();
  const src: ChartSource = { sheets: r.sheets.map((info, i) => new SheetView(info, r.grids[i])), styles: r.styles, date1904: r.date1904 };
  const chart = (title: string): ChartSpec => {
    const spec = r.sheets[sheetIndex(session, 'Charts')].objects.find((o) => o.title === title)?.chart;
    if (!spec) throw new Error(`no chart titled ${title}`);
    return spec;
  };
  /** Types into a cell, and hands the views what changed, as the store does. */
  const type = (sheet: string, row: number, col: number, text: string) => {
    for (const u of session.apply({ type: 'input', sheet: sheetIndex(session, sheet), row, col, text }).updates) src.sheets[u.sheet].apply(u);
  };
  return { session, src, chart, type };
}

const series = (extra: Partial<ChartSeries> = {}): ChartSeries => ({ values: { cache: [] }, line: false, markers: false, smooth: false, dashed: false, ...extra });
const spec = (extra: Partial<ChartSpec> = {}): ChartSpec => ({
  groups: [{ kind: 'bar', horizontal: false, grouping: 'clustered', varyColors: false, secondary: false, gapWidth: 150, series: [series()] }],
  axes: {},
  blanks: 'gap',
  visibleOnly: true,
  accents: [],
  ...extra,
});

describe('plotting a chart from the cells as they are now', () => {
  it('reads names, categories and values from the cells, in their formats', () => {
    const { src, chart } = workbook();
    const plot = plotChart(chart('Quarterly sales'), src);
    expect(plot.title).toBe('Quarterly sales');
    expect(plot.categories).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
    expect(plot.groups[0].series.map((s) => [s.name, s.values, s.format, s.index])).toEqual([
      ['North', [120, 150, 130, 170], '"$"#,##0', 0],
      ['South', [90, 110, 100, 120], '"$"#,##0', 1],
      ['East', [60, 80, 70, 90], '"$"#,##0', 2],
    ]);
    expect(plot.axes.category).toMatchObject({ title: 'Quarter', format: 'General' });
    expect(plot.axes.value).toMatchObject({ title: 'Units', min: 0, max: 200, format: '"$"#,##0' });
    expect(plot.legend).toEqual({ position: 'bottom', deleted: [] });
  });

  it('follows edits, and formulas the edits recalculate', () => {
    const { src, chart, type } = workbook();
    type('Data', 1, 1, '500');
    type('Data', 0, 2, 'Southwest');
    type('Data', 2, 0, 'Second');
    const plot = plotChart(chart('Quarterly sales'), src);
    expect(plot.groups[0].series[0].values[0]).toBe(500);
    expect(plot.groups[0].series[1].name).toBe('Southwest');
    expect(plot.categories[1]).toBe('Second');
    // A title from a cell follows it too.
    type('Data', 0, 3, 'Eastern');
    expect(plotChart(chart('East'), src).title).toBe('Eastern');
  });

  it('plots text as zero and leaves out empty cells and errors, as Excel does', () => {
    const { src, chart, type } = workbook();
    type('Data', 1, 1, 'n/a');
    type('Data', 2, 1, '');
    type('Data', 3, 1, '=1/0');
    expect(plotChart(chart('Quarterly sales'), src).groups[0].series[0].values).toEqual([0, null, null, 170]);
  });

  it('leaves out hidden rows and shows dates as their cells do', () => {
    const { src, chart } = workbook();
    const plot = plotChart(chart('Visits by month'), src);
    expect(plot.categories).toEqual(['Jan 2026', 'Feb 2026', 'Mar 2026', 'May 2026', 'Jun 2026']);
    expect(plot.groups[0].series[0].values).toEqual([410, 380, 520, 610, 590]);
    expect(plot.blanks).toBe('span');
    // Unless the chart asks for hidden cells too.
    expect(plotChart({ ...chart('Visits by month'), visibleOnly: false }, src).categories).toHaveLength(6);
  });

  it('places scatter points by their x values, read through a defined name', () => {
    const { src, chart } = workbook();
    const [s] = plotChart(chart('Sales by temperature'), src).groups[0].series;
    expect(s.x).toEqual([12, 15, 18, 22, 26, 29]);
    expect(s.values).toEqual([210, 260, 300, 390, 420, 480]);
  });

  it('reads another sheet, and a secondary axis takes the format of its own series', () => {
    const { src, chart } = workbook();
    expect(plotChart(chart('Units by product'), src).categories).toEqual(['Widget', 'Gadget', 'Gizmo']);
    const combo = plotChart(chart('Sales and margin'), src);
    expect(combo.axes.secondary).toMatchObject({ title: 'Margin', format: '0%' });
    expect(combo.groups[1].series[0]).toMatchObject({ name: 'Margin', values: [0.12, 0.15, 0.1, 0.18], format: '0%', index: 1 });
    expect(plotChart(chart('Share by quarter'), src).axes.value?.format).toBe('0%');
  });

  it('lists the sheets a chart reads, to draw it again when they change', () => {
    const { session, chart } = workbook();
    expect(chartSheets(chart('Quarterly sales'))).toEqual([sheetIndex(session, 'Data')]);
    expect(chartSheets(chart('Units by product'))).toEqual([sheetIndex(session, 'Sales Data')]);
    expect(chartSheets(spec())).toEqual([]);
  });
});

describe('plotting what Excel saved', () => {
  const src: ChartSource = { sheets: [new SheetView(sheetInfo(), null), new SheetView(sheetInfo({ kind: 'chartsheet' }), null)], styles: [PLAIN], date1904: false };
  const saved = (cache: ChartData['cache'], extra: Partial<ChartData> = {}): ChartData => ({ cache, ...extra });

  it('uses the saved values when a reference can’t be read, formatting numbers as saved', () => {
    const plot = plotChart(
      spec({
        title: saved(['Saved title']),
        groups: [
          {
            ...spec().groups[0],
            series: [
              series({
                name: saved(['Saved']),
                categories: saved([46023, 46054, null], { format: 'mmm yyyy' }),
                values: saved([1, 'text', null], { formula: '[1]Other!$A$1:$A$3' }),
              }),
              // A chart sheet's "cells" can't be read.
              series({ values: saved([5], { ref: { sheet: 1, rect: { r1: 0, c1: 0, r2: 0, c2: 0 } } }) }),
            ],
          },
        ],
      }),
      src,
    );
    expect(plot.title).toBe('Saved title');
    expect(plot.categories).toEqual(['Jan 2026', 'Feb 2026', '']);
    expect(plot.groups[0].series.map((s) => [s.name, s.values, s.format])).toEqual([
      ['Saved', [1, 0, null], 'General'],
      [msg('charts.seriesName', { n: 2 }), [5, null, null], 'General'],
    ]);
    expect(english.text(plot.groups[0].series[1].name)).toBe('Series2');
  });

  it('numbers categories that aren’t given, and plots text x values in order', () => {
    const plot = plotChart(
      spec({ groups: [{ ...spec().groups[0], kind: 'scatter', series: [series({ categories: saved(['a', 'b']), values: saved([3, 4, 5]) }), series({ values: saved([1]) })] }] }),
      src,
    );
    expect(plot.categories).toEqual(['a', 'b', '3']);
    expect(plot.groups[0].series.map((s) => s.x)).toEqual([
      [1, 2, 3],
      [1, 2, 3],
    ]);
    expect(plotChart(spec({ groups: [{ ...spec().groups[0], series: [series({ values: saved([1, 2]) })] }] }), src).categories).toEqual(['1', '2']);
  });

  it('reads a whole column only as far as the sheet is used', () => {
    const grid = new SheetView(sheetInfo({ maxRow: 2, maxCol: 0 }), null);
    for (const [row, v] of [[0, 1], [1, 2], [2, 3]]) grid.apply({ sheet: 0, row, col: 0, t: num(v).t, n: v, style: 0, flags: 0 });
    grid.apply({ sheet: 0, row: 0, col: 1, t: str('x').t, s: 'Label', style: 1, flags: 0 });
    const plot = plotChart(
      spec({
        groups: [{ ...spec().groups[0], series: [series({ name: { cache: [], ref: { sheet: 0, rect: { r1: 0, c1: 1, r2: 0, c2: 3 } } }, values: { cache: [], ref: { sheet: 0, rect: { r1: 0, c1: 0, r2: 1_048_575, c2: 0 } } } })] }],
      }),
      { sheets: [grid], styles: [PLAIN, style({ numFmt: '@' })], date1904: false },
    );
    expect(plot.groups[0].series[0].values).toEqual([1, 2, 3]);
    expect(plot.groups[0].series[0].name).toBe('Label');
  });
});

describe('a chart in words', () => {
  it('names each series and its points, and the axes', () => {
    const { src, chart } = workbook();
    expect(describeChart(plotChart(chart('Quarterly sales'), src), english)).toBe(
      '3 series. North: Q1 $120, Q2 $150, Q3 $130, and Q4 $170. South: Q1 $90, Q2 $110, Q3 $100, and Q4 $120. ' +
        'East: Q1 $60, Q2 $80, Q3 $70, and Q4 $90. Horizontal axis: Quarter. Vertical axis: Units.',
    );
    expect(describeChart(plotChart(chart('Ice cream sales'), src), english)).toBe(
      '1 series. Ice cream: x 12, y 210, x 15, y 260, x 18, y 300, x 22, y 390, x 26, y 420, and x 29, y 480. Horizontal axis: Temperature.',
    );
    expect(describeChart(plotChart(chart('Sales and margin'), src), english)).toContain('Margin: Q1 12%, Q2 15%, Q3 10%, and Q4 18%. Secondary axis: Margin.');
  });

  it('puts a bar chart’s categories up the side, and says only what a pie shows', () => {
    const plot = plotChart(
      spec({
        groups: [{ ...spec().groups[0], horizontal: true, series: [series({ values: { cache: [1, null] } })] }],
        axes: { category: { title: { cache: ['Region'] }, reversed: false, hidden: false, gridlines: false }, value: { title: { cache: ['Sales'] }, reversed: false, hidden: false, gridlines: true } },
      }),
      { sheets: [], styles: [PLAIN], date1904: false },
    );
    expect(describeChart(plot, english)).toBe('1 series. Series1: 1 1 and 2 blank. Horizontal axis: Sales. Vertical axis: Region.');
    const pie = plotChart(spec({ groups: [{ ...spec().groups[0], kind: 'pie', series: [series({ values: { cache: [1] } }), series({ values: { cache: [2] } })] }] }), {
      sheets: [],
      styles: [PLAIN],
      date1904: false,
    });
    expect(describeChart(pie, english)).toBe('1 series. Series1: 1 1.');
  });

  it('says how many more points a long series has, and when there is nothing to plot', () => {
    const plot = plotChart(spec({ groups: [{ ...spec().groups[0], series: [series({ values: { cache: Array.from({ length: 30 }, (_, i) => i) } })] }] }), {
      sheets: [],
      styles: [PLAIN],
      date1904: false,
    });
    expect(describeChart(plot, english)).toMatch(/24 23, and 6 more\.$/);
    expect(describeChart(plotChart(spec(), { sheets: [], styles: [PLAIN], date1904: false }), english)).toBe('There is no data to plot.');
    expect(describeChart(plotChart(spec({ groups: [] }), { sheets: [], styles: [PLAIN], date1904: false }), english)).toBe('There is no data to plot.');
  });

  it('formats values as their cells would', () => {
    expect(formatValue(1234.5, '#,##0.00')).toBe('1,234.50');
    expect(formatValue(0.125, '0%')).toBe('13%');
    expect(formatValue(0, 'yyyy-mm-dd', true)).toBe('1904-01-01');
  });
});
