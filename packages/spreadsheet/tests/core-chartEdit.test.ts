import { describe, expect, it } from 'vitest';
import { chartKindOf, editChartXml, seriesRefsOf } from '../src/core/chartEdit';
import { parseChart, type ChartContext } from '../src/core/charts';
import type { NewChartSeries } from '../src/core/chartWriter';
import { objectKey } from '../src/core/objectKey';
import { WorkbookSession } from '../src/core/session';
import type { ChartEdit } from '../src/core/chartEdit';
import { parseXml } from '../src/core/xml';
import { english } from '../src/i18n';
import { open as openSample, parts } from './helpers';

const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const CTX: ChartContext = { theme: [], sheetIndex: new Map([['s', 0]]), names: [] };

/** A column chart as Excel saves one: a formatted title, a green series, a legend on the right. */
const CHART =
  `<c:chartSpace ${C}><c:chart>` +
  '<c:title><c:tx><c:rich><a:bodyPr rot="0"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr><a:r><a:rPr lang="en-US" b="1"/><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
  '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>S!$B$1</c:f></c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="217346"/></a:solidFill></c:spPr>' +
  '<c:dPt><c:idx val="1"/></c:dPt><c:cat><c:strRef><c:f>S!$A$2:$A$3</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>S!$B$2:$B$3</c:f></c:numRef></c:val></c:ser>' +
  '<c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:crossAx val="1"/></c:valAx></c:plotArea>' +
  '<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>';

const spec = (xml: string) => parseChart(parseXml(xml), CTX);
const series = (n: number): NewChartSeries => ({
  name: { ref: `S!$${'BCD'[n]}$1`, text: `Q${n + 1}` },
  categories: { ref: 'S!$A$2:$A$3', points: ['North', 'South'], numeric: false, format: 'General' },
  values: { ref: `S!$${'BCD'[n]}$2:$${'BCD'[n]}$3`, points: [n + 1, n + 2], format: 'General' },
});

describe('changing a chart part', () => {
  it("sets the title in the old title's formatting, or takes it away, and adds one to a chart without", () => {
    const titled = editChartXml(CHART, { kind: 'column', rebuild: false, title: 'Revenue' });
    expect(titled).toContain('<c:tx><c:rich><a:bodyPr rot="0"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400"/></a:pPr><a:r><a:rPr lang="en-US" b="1"/><a:t>Revenue</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>');
    expect(spec(titled).title?.cache).toEqual(['Revenue']);
    const none = editChartXml(CHART, { kind: 'column', rebuild: false, title: '' });
    expect(none).toContain('<c:chart><c:autoTitleDeleted val="1"/><c:plotArea>');
    expect(spec(none).title).toBeUndefined();
    expect(spec(editChartXml(none, { kind: 'column', rebuild: false, title: 'Back' })).title?.cache).toEqual(['Back']);
  });

  it('moves the legend, takes it away, and puts one back after the plot area', () => {
    expect(spec(editChartXml(CHART, { kind: 'column', rebuild: false, legend: 'top' })).legend?.position).toBe('top');
    const gone = editChartXml(CHART, { kind: 'column', rebuild: false, legend: null });
    expect(spec(gone).legend).toBeUndefined();
    expect(editChartXml(gone, { kind: 'column', rebuild: false, legend: 'bottom' })).toContain('</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>');
  });

  it("sets series in place, each keeping its look; a new one takes the last one's without its points' colours", () => {
    const two = editChartXml(CHART, { kind: 'column', rebuild: false, series: [series(0), series(1)] });
    const s = spec(two).groups[0].series;
    expect(s.map((x) => [x.name?.formula, x.values.formula, x.values.cache, x.color])).toEqual([
      ['S!$B$1', 'S!$B$2:$B$3', [1, 2], '#217346'],
      ['S!$C$1', 'S!$C$2:$C$3', [2, 3], '#217346'],
    ]);
    expect(two.match(/<c:dPt>/g)).toHaveLength(1);
    expect(two).toContain('<c:idx val="1"/><c:order val="1"/>');
    // Down to one series again, which keeps its own name written in when given one.
    const one = editChartXml(two, { kind: 'column', rebuild: false, series: [{ ...series(2), name: { text: 'Plan' } }] });
    expect(spec(one).groups[0].series.map((x) => [x.name?.cache, x.values.formula])).toEqual([[['Plan'], 'S!$D$2:$D$3']]);
  });

  it('makes a new plot area for another type, keeping the title and legend', () => {
    const line = editChartXml(CHART, { kind: 'line', rebuild: true, series: [series(0)] });
    const s = spec(line);
    expect([chartKindOf(s), s.title?.cache, s.legend?.position]).toEqual(['line', ['Sales'], 'right']);
    expect(chartKindOf(spec(editChartXml(CHART, { kind: 'pie', rebuild: true, series: [series(0)] })))).toBe('pie');
    expect(seriesRefsOf(s)).toEqual([{ name: 'S!$B$1', values: 'S!$B$2:$B$3', categories: 'S!$A$2:$A$3' }]);
  });

  it('adds a name and categories to a series without them, and keeps a legend that had no position', () => {
    const plain = CHART.replace(/<c:tx><c:strRef><c:f>S!\$B\$1<\/c:f><\/c:strRef><\/c:tx>/, '').replace(/<c:cat>.*?<\/c:cat>/, '').replace('<c:legendPos val="r"/>', '');
    const named = editChartXml(plain, { kind: 'column', rebuild: false, series: [series(1)], legend: 'left' });
    const s = spec(named);
    expect([s.groups[0].series[0].name?.formula, s.groups[0].series[0].categories?.formula, s.legend?.position]).toEqual(['S!$C$1', 'S!$A$2:$A$3', 'left']);
    expect(spec(editChartXml(CHART.replace(/<c:legend>.*?<\/c:legend>/, '<c:legend/>'), { kind: 'column', rebuild: false, legend: 'bottom' })).legend?.position).toBe('bottom');
    // Nothing to set series into, or no chart at all: left as it is, but for the prefixes declared.
    const withR = (x: string) => x.replace('/2006/main"', '/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
    const empty = `<c:chartSpace ${C}><c:chart><c:plotArea/></c:chart></c:chartSpace>`;
    expect(editChartXml(empty, { kind: 'column', rebuild: false, series: [series(0)] })).toBe(withR(empty));
    expect(editChartXml(`<c:chartSpace ${C}/>`, { kind: 'column', rebuild: false, title: 'x', legend: null })).toBe(withR(`<c:chartSpace ${C}/>`));
  });

  it('declares the prefixes it writes with where the file does not', () => {
    const bare = '<chartSpace xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart"><chart><plotArea/></chart></chartSpace>';
    const out = editChartXml(bare, { kind: 'column', rebuild: false, legend: 'left' });
    expect(out).toMatch(/^<chartSpace xmlns="[^"]+" xmlns:c="[^"]+" xmlns:a="[^"]+" xmlns:r="[^"]+">/);
    expect(out).toContain('<plotArea/><c:legend><c:legendPos val="l"/>');
  });
});

const edit = (s: WorkbookSession, key: string, e: ChartEdit) => s.apply({ type: 'layout', sheet: 0, op: { kind: 'editChart', key, edit: e, title: 'Chart Title' } });

describe("a file's chart changed", () => {
  it('is saved over its own part, keeping what it was not asked to change', () => {
    const s = openSample('kitchen-sink.xlsx');
    const chart = s.model.sheets[0].info.objects.find((o) => o.kind === 'chart')!;
    const part = chart.chartPart!;
    const before = parts(s.save())[part];
    const done = edit(s, objectKey(chart), { title: 'Takings by region', legend: 'top' });
    expect(done.rejected).toBeUndefined();
    const saved = s.save();
    const after = parts(saved)[part];
    expect(after).toContain('<a:t>Takings by region</a:t>');
    expect(after).toContain('<a:srgbClr val="217346"/>');
    expect(after.length).toBeGreaterThan(before.length - 50);
    const t = WorkbookSession.open(saved, 'k.xlsx', 'en-US');
    expect(t.model.sheets[0].info.objects.find((o) => o.chartPart === part)?.title).toBe('Takings by region');
    // Undone, the part is as the file had it.
    s.apply({ type: 'restore', cells: [], layout: done.layout!.undo });
    expect(parts(s.save())[part]).toBe(before);
  });

  it('takes new data from a range down the rows, or from references, and refuses what is not a range', () => {
    const s = openSample('kitchen-sink.xlsx');
    const chart = s.model.sheets[0].info.objects.find((o) => o.kind === 'chart')!;
    const key = objectKey(chart);
    edit(s, key, { range: { ref: 'Summary!$A$4:$C$6', byColumns: false } });
    const rows = s.model.sheets[0].info.objects.find((o) => objectKey(o) === key)!.chart!.groups[0].series;
    expect(rows.map((r) => [r.name?.cache, r.values.formula])).toEqual([
      [['North'], 'Summary!$B$5:$C$5'],
      [['South'], 'Summary!$B$6:$C$6'],
    ]);
    edit(s, key, { kind: 'bar', series: [{ name: 'Totals', values: '=Summary!$F$5:$F$8', categories: 'Summary!$A$5:$A$8' }] });
    const bar = s.model.sheets[0].info.objects.find((o) => objectKey(o) === key)!;
    expect([chartKindOf(bar.chart), english.text(bar.chartType!), bar.chart!.groups[0].series[0].values.cache]).toEqual(['bar', 'Bar', [36359.5, 42651.25, 42312.25, 36602.25]]);
    expect(english.text(edit(s, key, { series: [{ values: 'nowhere' }] }).rejected!)).toBe("That isn't a range on a sheet: nowhere");
  });
});
