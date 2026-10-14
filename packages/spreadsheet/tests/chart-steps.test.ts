import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { parseChartEx } from '../src/core/charts';
import { WorkbookSession } from '../src/core/session';
import { SheetView } from '../src/state/sheetView';
import { parseXml } from '../src/core/xml';
import { english } from '../src/i18n';
import { plotChart } from '../src/state/charts';
import { chartConfig, type ChartLook } from '../src/ui/chartConfig';

// Office's theme: lt1, dk1, lt2, dk2, then the accents.
const ctx = { theme: ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'], sheetIndex: new Map(), names: [] };
const LOOK: ChartLook = { palette: [], background: '#fff', text: '#111', muted: '#555', gridline: '#ddd', axis: '#ccc', font: 'Calibri' };
const src = { sheets: [], styles: [], date1904: false };

/** A chartEx part as Excel writes it, with the points it saved. */
const chartEx = (layout: string, labels: string[], values: number[], extra = '') =>
  parseXml(
    '<cx:chartSpace><cx:chartData><cx:data id="0">' +
      `<cx:strDim type="cat"><cx:f>Data!$A$2:$A$${labels.length + 1}</cx:f><cx:lvl ptCount="${labels.length}">${labels.map((l, i) => `<cx:pt idx="${i}">${l}</cx:pt>`).join('')}</cx:lvl></cx:strDim>` +
      `<cx:numDim type="val"><cx:f>Data!$B$2:$B$${values.length + 1}</cx:f><cx:lvl ptCount="${values.length}" formatCode="General">${values.map((v, i) => `<cx:pt idx="${i}">${v}</cx:pt>`).join('')}</cx:lvl></cx:numDim>` +
      '</cx:data></cx:chartData><cx:chart><cx:title><cx:tx><cx:txData><cx:v>Cash</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion>' +
      `<cx:series layoutId="${layout}"><cx:tx><cx:txData><cx:f>Data!$B$1</cx:f><cx:v>Amount</cx:v></cx:txData></cx:tx><cx:dataId val="0"/>${extra}</cx:series>` +
      '</cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea><cx:legend pos="t"/></cx:chart></cx:chartSpace>',
  );
type Loose = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any

describe("Excel 2016's waterfall and funnel charts", () => {
  it('read a waterfall: its points, totals, title, legend and gap', () => {
    const spec = parseChartEx(chartEx('waterfall', ['Start', 'Sales', 'Costs', 'End'], [100, 50, -30, 120], '<cx:layoutPr><cx:subtotals><cx:idx val="0"/><cx:idx val="3"/></cx:subtotals></cx:layoutPr>'), ctx);
    expect(spec.unsupported).toBeUndefined();
    expect(spec.title).toEqual({ cache: ['Cash'] });
    expect(spec.legend?.position).toBe('top');
    expect(spec.groups[0]).toMatchObject({ kind: 'waterfall', gapWidth: 50, subtotals: [0, 3] });
    expect(spec.groups[0].series[0]).toMatchObject({ name: { formula: 'Data!$B$1', cache: ['Amount'] }, values: { cache: [100, 50, -30, 120] }, categories: { cache: ['Start', 'Sales', 'Costs', 'End'] } });
  });

  it('draw a waterfall: steps from the running total, totals from zero, coloured by kind', () => {
    const spec = parseChartEx(chartEx('waterfall', ['Start', 'Sales', 'Costs', 'End'], [100, 50, -30, 120], '<cx:layoutPr><cx:subtotals><cx:idx val="0"/><cx:idx val="3"/></cx:subtotals></cx:layoutPr>'), ctx);
    const config = chartConfig(plotChart(spec, src), LOOK, english.text);
    const sets = config.data.datasets as Loose[];
    expect(sets.map((d) => d.label)).toEqual(['Increase', 'Decrease', 'Total']);
    expect(sets.map((d) => d.data)).toEqual([
      [null, [100, 150], null, null],
      [null, null, [150, 120], null],
      [[0, 100], null, null, [0, 120]],
    ]);
    expect(sets.map((d) => d.backgroundColor[0])).toEqual(['#4472C4', '#ED7D31', '#A5A5A5']);
    expect((config.options as Loose).indexAxis).toBe('x');
  });

  it('draw a funnel: centred bars, the first at the top, with no value axis', () => {
    const spec = parseChartEx(chartEx('funnel', ['Leads', 'Calls', 'Deals'], [300, 120, 40]), ctx);
    expect(spec.groups[0]).toMatchObject({ kind: 'funnel', horizontal: true });
    const config = chartConfig(plotChart(spec, src), LOOK, english.text);
    const [set] = config.data.datasets as Loose[];
    expect(set.data).toEqual([[-150, 150], [-60, 60], [-20, 20]]);
    const options = config.options as Loose;
    expect([options.indexAxis, options.scales.x.display, options.scales.y.reverse]).toEqual(['y', false, false]);
  });

  it('leave the other newer kinds as placeholders', () => {
    expect(parseChartEx(chartEx('treemap', ['A'], [1]), ctx).unsupported).toBe('treemap');
    expect(parseChartEx(chartEx('clusteredColumn', ['A'], [1]), ctx).unsupported).toBe('clusteredColumn');
  });
});

describe('a waterfall and a funnel made in Excel for the web', () => {
  // tests/corpus/waterfall-funnel.xlsx: Step/Amount in A1:B6 (Start 100, Sales 60, Refunds -25, Costs -40, End 95, no totals) and Stage/People in D1:E5.
  const s = WorkbookSession.open(new Uint8Array(readFileSync(new URL('./corpus/waterfall-funnel.xlsx', import.meta.url))), 'w.xlsx');
  const r = s.loadResult();
  const src = { sheets: r.sheets.map((info, i) => new SheetView(info, r.grids[i])), styles: r.styles, date1904: r.date1904 };
  const [waterfall, funnel] = r.sheets[0].objects.map((o) => o.chart!);

  it('are read through the hidden names Excel refers to its data by, with no saved points', () => {
    expect(waterfall.groups[0]).toMatchObject({ kind: 'waterfall', gapWidth: 50 });
    expect(waterfall.groups[0].subtotals).toBeUndefined();
    expect(waterfall.groups[0].series[0].values).toMatchObject({ formula: '_xlchart.v1.2', ref: { sheet: 0, rect: { r1: 1, c1: 1, r2: 5, c2: 1 } }, cache: [] });
    expect(funnel.groups[0]).toMatchObject({ kind: 'funnel' });
    expect(funnel.groups[0].gapWidth).toBeCloseTo(6);
    expect(funnel.axes.value).toBeUndefined();
  });

  it('are drawn from the cells, with the labels and legend the file asks for', () => {
    const w = plotChart(waterfall, src);
    expect([w.categories, w.groups[0].series[0].values, w.groups[0].series[0].name]).toEqual([['Start', 'Sales', 'Refunds', 'Costs', 'End'], [100, 60, -25, -40, 95], 'Amount']);
    const sets = chartConfig(w, LOOK, english.text).data.datasets as Loose[];
    expect(sets.map((d) => d.label)).toEqual(['Increase', 'Decrease', 'Total']);
    expect(sets[0].wbeLabels.texts[0]).toBe('Amount, Start, 100');
    const f = chartConfig(plotChart(funnel, src), LOOK, english.text).data.datasets as Loose[];
    expect(f[0].data).toEqual([[-250, 250], [-110, 110], [-45, 45], [-17.5, 17.5]]);
    expect(f[0].wbeLabels).toEqual({ texts: ['500', '220', '90', '35'], position: 'ctr' });
  });
});
