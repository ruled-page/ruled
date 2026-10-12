import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { drawingColor, parseChart, type ChartContext, type ChartSpec } from '../src/core/charts';
import { WorkbookSession } from '../src/core/session';
import type { SheetObject } from '../src/core/types';
import { parseXml } from '../src/core/xml';
import { sheetIndex } from './helpers';

/** tests/corpus/charts.xlsx, made by scripts/gen_charts.py with XlsxWriter. */
const session = WorkbookSession.open(new Uint8Array(readFileSync(join(__dirname, 'corpus', 'charts.xlsx'))), 'charts.xlsx', 'en-US');
const objects = (sheet: string): SheetObject[] => session.model.sheets[sheetIndex(session, sheet)].info.objects;
const DATA = sheetIndex(session, 'Data');
const SALES = sheetIndex(session, 'Sales Data');

/** The chart titled `title` on the Charts sheet. */
function chart(title: string): ChartSpec {
  const found = objects('Charts').find((o) => o.title === title)?.chart;
  if (!found) throw new Error(`no chart titled ${title}`);
  return found;
}

const rect = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });
/** XlsxWriter writes the Office 2007 theme. */
const OFFICE_2007_ACCENTS = ['#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646'];

describe('reading charts made by XlsxWriter', () => {
  it('reads a clustered column chart: series from cells, saved values, colours, axes and legend', () => {
    const c = chart('Quarterly sales');
    expect(c.unsupported).toBeUndefined();
    expect(c.title).toEqual({ cache: ['Quarterly sales'] });
    expect(c.groups).toHaveLength(1);
    const [g] = c.groups;
    expect(g).toMatchObject({ kind: 'bar', horizontal: false, grouping: 'clustered', varyColors: false, secondary: false, gapWidth: 150 });
    expect(g.series.map((s) => s.name?.ref)).toEqual([1, 2, 3].map((col) => ({ sheet: DATA, rect: rect(0, col, 0, col) })));
    expect(g.series.map((s) => s.name?.cache)).toEqual([['North'], ['South'], ['East']]);
    expect(g.series[0].categories).toEqual({ formula: 'Data!$A$2:$A$5', ref: { sheet: DATA, rect: rect(1, 0, 4, 0) }, cache: ['Q1', 'Q2', 'Q3', 'Q4'], format: undefined });
    expect(g.series[0].values).toEqual({ formula: 'Data!$B$2:$B$5', ref: { sheet: DATA, rect: rect(1, 1, 4, 1) }, cache: [120, 150, 130, 170], format: 'General' });
    expect(g.series.map((s) => s.color)).toEqual(['#1F77B4', undefined, undefined]);
    expect(c.axes.category).toMatchObject({ title: { cache: ['Quarter'] }, hidden: false, gridlines: false, reversed: false });
    expect(c.axes.value).toMatchObject({ title: { cache: ['Units'] }, min: 0, max: 200, majorUnit: 50, gridlines: true, format: undefined });
    expect(c.axes.secondary).toBeUndefined();
    expect(c.legend).toEqual({ position: 'bottom', deleted: [] });
    expect(c).toMatchObject({ blanks: 'zero', visibleOnly: true, accents: OFFICE_2007_ACCENTS });
  });

  it('reads how bars are stacked and which way they run', () => {
    expect(chart('Stacked columns').groups[0]).toMatchObject({ kind: 'bar', grouping: 'stacked', horizontal: false });
    expect(chart('Share by quarter').groups[0]).toMatchObject({ kind: 'bar', grouping: 'percentStacked' });
    expect(chart('Horizontal bars').groups[0]).toMatchObject({ kind: 'bar', grouping: 'clustered', horizontal: true });
    expect(chart('Horizontal bars').legend?.position).toBe('top');
    expect(chart('Stacked bars').groups[0]).toMatchObject({ grouping: 'stacked', horizontal: true });
  });

  it('reads lines: their colour, width, dashes, markers and smoothing', () => {
    const [north, south] = chart('Sales trend').groups[0].series;
    expect(chart('Sales trend').groups[0]).toMatchObject({ kind: 'line', grouping: 'standard' });
    expect(north).toMatchObject({ color: '#C00000', lineWidth: 3.3, line: true, markers: true, smooth: false, dashed: false });
    expect(south).toMatchObject({ color: undefined, line: true, markers: false, smooth: true, dashed: true });
    expect(chart('Sales trend').legend?.position).toBe('right');
  });

  it('reads how empty cells are plotted, and dates as categories', () => {
    const visits = chart('Visits by month');
    expect(visits.blanks).toBe('span');
    expect(visits.visibleOnly).toBe(true);
    expect(visits.groups[0].series[0].categories?.ref).toEqual({ sheet: DATA, rect: rect(10, 0, 15, 0) });
    expect(visits.groups[0].series[0].categories?.cache[0]).toBe(46023);
    expect(visits.groups[0].series[0].name).toEqual({ cache: ['Visits'] });
  });

  it('reads areas', () => {
    expect(chart('Area').groups[0]).toMatchObject({ kind: 'area', grouping: 'standard' });
    expect(chart('Stacked area').groups[0]).toMatchObject({ kind: 'area', grouping: 'stacked' });
  });

  it('reads pies and doughnuts: slice colours, rotation and the hole', () => {
    const pie = chart('North by quarter');
    expect(pie.groups[0]).toMatchObject({ kind: 'pie', varyColors: true, firstSliceAngle: 90, holeSize: undefined });
    expect(pie.groups[0].series[0].pointColors).toEqual({ 0: '#FF0000', 2: '#00B050' });
    const doughnut = chart('Units by product');
    expect(doughnut.groups[0]).toMatchObject({ kind: 'doughnut', holeSize: 40, firstSliceAngle: 0 });
    // A sheet whose name needs quoting.
    expect(doughnut.groups[0].series[0].values.ref).toEqual({ sheet: SALES, rect: rect(1, 1, 3, 1) });
    expect(doughnut.legend?.position).toBe('left');
  });

  it('reads scatter charts, with x values from a defined name', () => {
    const markers = chart('Ice cream sales');
    expect(markers.groups[0]).toMatchObject({ kind: 'scatter' });
    expect(markers.groups[0].series[0]).toMatchObject({ line: false, markers: true });
    expect(markers.groups[0].series[0].categories?.ref).toEqual({ sheet: DATA, rect: rect(1, 7, 6, 7) });
    expect(markers.axes.category).toMatchObject({ min: 10, title: { cache: ['Temperature'] } });
    const lines = chart('Sales by temperature');
    expect(lines.groups[0].series[0]).toMatchObject({ line: true, markers: true });
    // XlsxWriter saves no values for a name; the name leads to the cells.
    expect(lines.groups[0].series[0].categories).toEqual({ formula: 'Temperatures', ref: { sheet: DATA, rect: rect(1, 7, 6, 7) }, cache: [], format: undefined });
  });

  it('reads a combination with a secondary axis', () => {
    const c = chart('Sales and margin');
    expect(c.unsupported).toBeUndefined();
    expect(c.groups.map((g) => [g.kind, g.secondary])).toEqual([
      ['bar', false],
      ['line', true],
    ]);
    expect(c.axes.secondary).toMatchObject({ title: { cache: ['Margin'] }, format: '0%', gridlines: false });
    expect(c.axes.value?.format).toBeUndefined();
  });

  it('reads a radar chart, as one of the kinds it draws', () => {
    const radar = chart('Radar');
    expect(radar.unsupported).toBeUndefined();
    expect(radar.groups[0]).toMatchObject({ kind: 'radar', filled: false });
  });

  it('titles an untitled single-series chart with the series name, as Excel does', () => {
    const untitled = objects('Charts').find((o) => o.kind === 'chart' && !o.title);
    expect(untitled).toBeUndefined();
    const east = objects('Charts').find((o) => o.title === 'East')!;
    expect(east.chart?.title?.ref).toEqual({ sheet: DATA, rect: rect(0, 3, 0, 3) });
  });

  it('reads the chart on a chart sheet', () => {
    const [trend] = objects('Trend');
    expect(trend.chart?.groups[0]).toMatchObject({ kind: 'line' });
    expect(trend.chart?.groups[0].series).toHaveLength(3);
  });
});

const C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const CTX: ChartContext = {
  theme: ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'],
  sheetIndex: new Map([
    ['data', 0],
    ['other sheet', 1],
  ]),
  names: [
    { name: 'Wide', formula: 'Data!$A$1:$C$1' },
    { name: 'Local', formula: "'Other Sheet'!$B$2:$B$3", localSheet: 1 },
    { name: 'Local', formula: 'Data!$Z$1' },
    { name: 'Loop', formula: 'Wide' },
  ],
};
const parse = (body: string, after = '') => parseChart(parseXml(`<c:chartSpace ${C}><c:chart>${body}</c:chart>${after}</c:chartSpace>`), CTX);
const plot = (groups: string, axes = '') => `<c:plotArea><c:layout/>${groups}${axes}</c:plotArea>`;
const ser = (inner: string) => `<c:ser><c:idx val="0"/><c:order val="0"/>${inner}</c:ser>`;
const numRef = (f: string, values: (number | string)[] = []) =>
  `<c:numRef><c:f>${f}</c:f><c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache></c:numRef>`;
const color = (inner: string) => drawingColor(parseXml(inner), CTX.theme);

describe('reading chart parts', () => {
  it('turns DrawingML colours into CSS colours', () => {
    expect(color('<a:srgbClr val="1f77b4"/>')).toBe('#1F77B4');
    expect(color('<a:schemeClr val="accent2"/>')).toBe('#ED7D31');
    expect(color('<a:schemeClr val="tx1"/>')).toBe('#000000');
    // Lighter 40% and darker 25%, as Excel's colour picker writes them.
    expect(color('<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>')).toBe('#8FAADC');
    expect(color('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>')).toBe('#2F5597');
    expect(color('<a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr>')).toBe('#D9D9D9');
    expect(color('<a:srgbClr val="808080"><a:tint val="50000"/></a:srgbClr>')).toMatch(/^#[0-9A-F]{6}$/);
    expect(color('<a:srgbClr val="808080"><a:shade val="50000"/></a:srgbClr>')).toBe('#404040');
    expect(color('<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>')).toBe('rgba(255, 0, 0, 0.5)');
    expect(color('<a:sysClr val="windowText" lastClr="000000"/>')).toBe('#000000');
    expect(color('<a:sysClr val="window"/>')).toBe('#FFFFFF');
    expect(color('<a:prstClr val="red"/>')).toBe('#FF0000');
    expect(color('<a:schemeClr val="phClr"/>')).toBeUndefined();
    expect(color('<a:srgbClr val="nonsense"/>')).toBeUndefined();
    expect(color('<a:hslClr hue="0" sat="0" lum="0"/>')).toBeUndefined();
    expect(drawingColor(undefined, CTX.theme)).toBeUndefined();
  });

  it('colours series and slices from the theme, and leaves gradients to the default', () => {
    const c = parse(
      plot(
        `<c:barChart><c:barDir val="col"/>${ser(
          `<c:spPr><a:solidFill><a:schemeClr val="accent3"/></a:solidFill></c:spPr>` +
            `<c:dPt><c:idx val="1"/><c:spPr><a:gradFill/></c:spPr></c:dPt><c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></c:spPr></c:dPt>` +
            `<c:val>${numRef('Data!$B$1:$B$3', [1, 2, 3])}</c:val>`,
        )}<c:ser><c:spPr><a:gradFill/></c:spPr><c:val>${numRef('Data!$C$1:$C$3')}</c:val></c:ser></c:barChart>`,
      ),
    );
    expect(c.groups[0].series.map((s) => [s.color, s.pointColors])).toEqual([
      ['#A5A5A5', { 2: '#70AD47' }],
      [undefined, undefined],
    ]);
    // A bar chart with no direction stands up, as Excel's default.
    expect(c.groups[0]).toMatchObject({ horizontal: false, grouping: 'clustered' });
  });

  it('reads values written into the chart, and several levels of labels from what Excel saved', () => {
    const c = parse(
      plot(
        `<c:lineChart>${ser(
          '<c:tx><c:v>Written in</c:v></c:tx>' +
            '<c:cat><c:multiLvlStrRef><c:f>Data!$A$1:$B$3</c:f><c:multiLvlStrCache><c:ptCount val="3"/><c:lvl><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="2"><c:v>c</c:v></c:pt></c:lvl><c:lvl><c:pt idx="0"><c:v>Group</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>' +
            '<c:val><c:numLit><c:formatCode>0%</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>0.5</c:v></c:pt><c:pt idx="1"><c:v>oops</c:v></c:pt><c:pt idx="7"><c:v>9</c:v></c:pt></c:numLit></c:val>',
        )}<c:ser><c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>x</c:v></c:pt></c:strLit></c:cat><c:val><c:numRef><c:f>Data!$C$1</c:f></c:numRef></c:val></c:ser></c:lineChart>`,
      ),
    );
    const [first, second] = c.groups[0].series;
    expect(first.name).toEqual({ cache: ['Written in'] });
    expect(first.categories).toEqual({ formula: 'Data!$A$1:$B$3', ref: undefined, cache: ['a', null, 'c'], format: undefined });
    // A point past the count it declares is ignored; text where a number belongs is no number.
    expect(first.values).toEqual({ cache: [0.5, null, null], format: '0%' });
    expect(second.categories).toEqual({ cache: ['x'], format: undefined });
    expect(second.values).toEqual({ formula: 'Data!$C$1', ref: { sheet: 0, rect: rect(0, 2, 0, 2) }, cache: [], format: undefined });
    expect(second.name).toBeUndefined();
    // Lines have markers unless the series or the chart turns them off.
    expect([first.markers, second.markers]).toEqual([true, true]);
  });

  it('follows references through names, and keeps saved values for ones it cannot follow', () => {
    const refs = [
      'Wide',
      'Data!Local',
      "'Other Sheet'!Local",
      'Loop',
      'Missing',
      '[1]Data!$A$1:$A$3',
      '(Data!$A$1,Data!$A$3)',
      'Nowhere!$A$1:$A$3',
      '$A$1:$A$3',
      'Data!$B:$B',
      'Data!$2:$2',
    ];
    const c = parse(plot(`<c:barChart>${refs.map((f) => `<c:ser><c:val>${numRef(f, [1])}</c:val></c:ser>`).join('')}</c:barChart>`));
    expect(c.groups[0].series.map((s) => s.values.ref)).toEqual([
      { sheet: 0, rect: rect(0, 0, 0, 2) },
      { sheet: 0, rect: rect(0, 25, 0, 25) },
      { sheet: 1, rect: rect(1, 1, 2, 1) },
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      { sheet: 0, rect: rect(0, 1, 1_048_575, 1) },
      { sheet: 0, rect: rect(1, 0, 1, 16_383) },
    ]);
    expect(c.groups[0].series.every((s) => s.values.cache[0] === 1)).toBe(true);
  });

  it('reads titles as rich text or from a cell, and axis settings', () => {
    const c = parse(
      '<c:title><c:tx><c:rich><a:p><a:r><a:t>Two </a:t></a:r><a:r><a:t>runs</a:t></a:r></a:p><a:p><a:r><a:t>and lines</a:t></a:r></a:p></c:rich></c:tx></c:title>' +
        plot(
          `<c:areaChart><c:grouping val="percentStacked"/>${ser(`<c:val>${numRef('Data!$B$1:$B$2', [1, 2])}</c:val>`)}<c:axId val="1"/><c:axId val="2"/></c:areaChart>`,
          '<c:dateAx><c:axId val="1"/><c:scaling><c:orientation val="maxMin"/></c:scaling><c:delete val="1"/><c:title><c:tx><c:strRef><c:f>Data!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>When</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title></c:dateAx>' +
            '<c:valAx><c:axId val="2"/><c:scaling><c:logBase val="10"/><c:max val="x"/></c:scaling><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorGridlines/></c:valAx>',
        ) +
        '<c:legend><c:legendPos val="tr"/><c:legendEntry><c:idx val="1"/><c:delete val="1"/></c:legendEntry><c:legendEntry><c:idx val="2"/><c:delete val="0"/></c:legendEntry></c:legend>' +
        '<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/>',
    );
    expect(c.title).toEqual({ cache: ['Two runs\nand lines'] });
    expect(c.axes.category).toMatchObject({ reversed: true, hidden: true, title: { formula: 'Data!$A$1', cache: ['When'] } });
    expect(c.axes.value).toMatchObject({ logBase: 10, max: undefined, format: '#,##0', gridlines: true });
    expect(c.legend).toEqual({ position: 'right', deleted: [1] });
    expect(c).toMatchObject({ visibleOnly: false, blanks: 'gap' });
    expect(c.groups[0].grouping).toBe('percentStacked');
  });

  it('places legends where the file says', () => {
    const legend = (pos: string) => parse(plot(`<c:pieChart>${ser('')}</c:pieChart>`) + `<c:legend><c:legendPos val="${pos}"/></c:legend>`).legend?.position;
    expect(['t', 'b', 'l', 'r'].map(legend)).toEqual(['top', 'bottom', 'left', 'right']);
    expect(parse(plot(`<c:pieChart>${ser('')}</c:pieChart>`) + '<c:legend/>').legend?.position).toBe('right');
  });

  it('leaves the title out when Excel would, and an empty rich title is none', () => {
    const one = plot(`<c:barChart>${ser(`<c:tx><c:v>Only</c:v></c:tx><c:val>${numRef('Data!$B$1', [1])}</c:val>`)}</c:barChart>`);
    expect(parse(one).title).toEqual({ cache: ['Only'] });
    expect(parse('<c:autoTitleDeleted val="1"/>' + one).title).toBeUndefined();
    expect(parse('<c:title><c:tx><c:rich><a:p/></c:rich></c:tx></c:title><c:autoTitleDeleted val="1"/>' + one).title).toBeUndefined();
    expect(parse('<c:title/>' + one).title).toEqual({ cache: ['Only'] });
  });

  it('draws 3-D charts flat, and refuses what it cannot draw', () => {
    const bar3d = parse(plot(`<c:bar3DChart><c:barDir val="bar"/><c:grouping val="standard"/>${ser('')}</c:bar3DChart>`));
    expect(bar3d.unsupported).toBeUndefined();
    expect(bar3d.groups[0]).toMatchObject({ kind: 'bar', horizontal: true, grouping: 'clustered' });
    expect(parse(plot(`<c:pie3DChart>${ser('')}</c:pie3DChart>`)).groups[0].kind).toBe('pie');
    expect(parse(plot(`<c:barChart>${ser('')}</c:barChart><c:pieChart>${ser('')}</c:pieChart>`)).unsupported).toBe('combination');
    expect(parse(plot(`<c:barChart><c:barDir val="col"/>${ser('')}</c:barChart><c:barChart><c:barDir val="bar"/>${ser('')}</c:barChart>`)).unsupported).toBe('combination');
    expect(parse(plot(`<c:barChart>${ser('')}</c:barChart><c:surfaceChart/>`)).unsupported).toBe('surfaceChart');
    expect(parse(plot('<c:barChart/>')).unsupported).toBe('empty');
    expect(parse(plot('')).unsupported).toBe('empty');
    expect(parse('').unsupported).toBe('empty');
  });

  it('puts series in their plotting order, and reads scatter lines and smoothing', () => {
    const c = parse(
      plot(
        '<c:scatterChart><c:scatterStyle val="smoothMarker"/>' +
          '<c:ser><c:idx val="0"/><c:order val="1"/><c:tx><c:v>Second</c:v></c:tx><c:marker><c:symbol val="none"/></c:marker><c:smooth val="1"/></c:ser>' +
          '<c:ser><c:idx val="1"/><c:order val="0"/><c:tx><c:v>First</c:v></c:tx><c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="solid"/></a:ln></c:spPr><c:smooth val="0"/></c:ser>' +
          '</c:scatterChart>',
      ),
    );
    const [first, second] = c.groups[0].series;
    expect(first).toMatchObject({ name: { cache: ['First'] }, color: '#00FF00', lineWidth: 1.3, dashed: false, smooth: false, markers: true, line: true });
    expect(second).toMatchObject({ name: { cache: ['Second'] }, markers: false, smooth: true, line: true });
  });

  it('caps the points it reads from a part, and takes a series with no values as empty', () => {
    const c = parse(plot(`<c:barChart>${ser('<c:val><c:numLit><c:ptCount val="999999999"/></c:numLit></c:val>')}<c:ser><c:val/></c:ser></c:barChart>`));
    expect(c.groups[0].series[0].values.cache).toHaveLength(10_000);
    expect(c.groups[0].series[1].values).toEqual({ cache: [] });
  });
});
