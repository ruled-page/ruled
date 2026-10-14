import { describe, expect, it } from 'vitest';
import { parseChart, parseChartEx } from '../src/core/charts';
import { parseXml } from '../src/core/xml';
import { english } from '../src/i18n';
import { plotChart } from '../src/state/charts';
import { chartConfig, type ChartLook } from '../src/ui/chartConfig';

const ctx = { theme: ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'], sheetIndex: new Map(), names: [] };
const LOOK: ChartLook = { palette: [], background: '#fff', text: '#111', muted: '#555', gridline: '#ddd', axis: '#ccc', font: 'Calibri' };
const src = { sheets: [], styles: [], date1904: false };
type Loose = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any

const cats = '<c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South</c:v></c:pt><c:pt idx="2"><c:v>East</c:v></c:pt></c:strCache></c:strRef></c:cat>';
const vals = '<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>30</c:v></c:pt><c:pt idx="2"><c:v>60</c:v></c:pt></c:numCache></c:numRef></c:val>';
const chart = (group: string, labels: string, groupLabels = '') =>
  parseChart(
    parseXml(`<c:chartSpace><c:chart><c:plotArea><c:${group}>${group === 'barChart' ? '<c:barDir val="col"/>' : ''}<c:ser><c:tx><c:v>Sales</c:v></c:tx>${labels}${cats}${vals}</c:ser>${groupLabels}</c:${group}></c:plotArea></c:chart></c:chartSpace>`),
    ctx,
  );
const datasets = (spec: ReturnType<typeof parseChart>) => {
  const config = chartConfig(plotChart(spec, src), LOOK, english.text);
  return { sets: config.data.datasets as Loose[], plugins: config.plugins ?? [] };
};

describe('data labels', () => {
  it("are read from a series, with their place, number format and deleted points, the series' setting winning over its chart's", () => {
    const spec = chart(
      'barChart',
      '<c:dLbls><c:dLbl><c:idx val="1"/><c:delete val="1"/></c:dLbl><c:numFmt formatCode="0.0" sourceLinked="0"/><c:dLblPos val="inEnd"/><c:showVal val="1"/><c:showCatName val="1"/></c:dLbls>',
      '<c:dLbls><c:showSerName val="1"/></c:dLbls>',
    );
    expect(spec.groups[0].series[0].labels).toEqual({ value: true, category: true, series: false, percent: false, position: 'inEnd', format: '0.0', hidden: [1] });
    const { sets, plugins } = datasets(spec);
    expect(sets[0].wbeLabels).toEqual({ texts: ['North, 10.0', undefined, 'East, 60.0'], position: 'inEnd' });
    expect(plugins).toHaveLength(1);
  });

  it("are taken from the chart where the series sets none, and a pie's show each slice's share", () => {
    expect(chart('barChart', '', '<c:dLbls><c:showSerName val="1"/></c:dLbls>').groups[0].series[0].labels).toMatchObject({ series: true });
    const pie = chart('pieChart', '<c:dLbls><c:showPercent val="1"/><c:showCatName val="1"/></c:dLbls>');
    expect(datasets(pie).sets[0].wbeLabels).toEqual({ texts: ['North, 10%', 'South, 30%', 'East, 60%'], position: 'bestFit' });
  });

  it('are none where nothing is shown, or they are deleted', () => {
    for (const labels of ['', '<c:dLbls><c:showVal val="0"/></c:dLbls>', '<c:dLbls><c:delete val="1"/></c:dLbls>']) {
      const spec = chart('lineChart', labels);
      expect(spec.groups[0].series[0].labels).toBeUndefined();
      expect(datasets(spec).plugins).toEqual([]);
    }
  });

  it("show a waterfall's values at each step's end, and a funnel's in the middle of its bars", () => {
    const ex = (layout: string, labels: string) =>
      parseChartEx(
        parseXml(
          '<cx:chartSpace><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:lvl ptCount="2"><cx:pt idx="0">A</cx:pt><cx:pt idx="1">B</cx:pt></cx:lvl></cx:strDim>' +
            '<cx:numDim type="val"><cx:lvl ptCount="2" formatCode="General"><cx:pt idx="0">100</cx:pt><cx:pt idx="1">-25</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>' +
            `<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="${layout}"><cx:dataId val="0"/>${labels}</cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`,
        ),
        ctx,
      );
    const waterfall = ex('waterfall', '<cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>');
    expect(waterfall.groups[0].series[0].labels).toEqual({ value: true, category: false, series: false, percent: false, position: 'outEnd' });
    expect(datasets(waterfall).sets.map((d) => d.wbeLabels.texts)).toEqual([['100', '-25'], ['100', '-25']]);
    const funnel = ex('funnel', '<cx:dataLabels/>');
    expect(datasets(funnel).sets[0].wbeLabels).toEqual({ texts: ['100', '-25'], position: 'ctr' });
    expect(ex('funnel', '').groups[0].series[0].labels).toBeUndefined();
  });
});
