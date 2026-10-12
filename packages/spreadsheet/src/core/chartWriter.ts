/**
 * Charts made in the editor, as Excel's Insert Chart makes them: what to
 * plot from a range of cells, and the chart part (xl/charts/chartN.xml) and
 * drawing anchor that put it on a sheet.
 */
import { quoteSheetName, rangeToA1 } from '@sitepen/formula-engine/refs';
import type { Anchor, Rect } from './types';
import { escapeXml } from './xml';

/** The kinds of chart Insert Chart offers. */
export type NewChartKind = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';

export const NEW_CHART_KINDS: NewChartKind[] = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter'];

/** A cell as Insert Chart sees it: its value, and whether its format makes it a date. */
export interface ChartCell {
  value: string | number | boolean | null;
  date: boolean;
  format: string;
  /** As the cell shows it. */
  text: string;
}

export interface NewChartSeries {
  name?: { ref: string; text: string };
  /** Labels along the category axis, or a scatter chart's x values. */
  categories?: { ref: string; points: (string | number | null)[]; numeric: boolean; format: string };
  values: { ref: string; points: (number | null)[]; format: string };
}

export interface NewChart {
  kind: NewChartKind;
  title?: string;
  legend: boolean;
  series: NewChartSeries[];
}

const isLabel = (c: ChartCell) => typeof c.value === 'string' && c.value !== '';
const isEmpty = (c: ChartCell) => c.value === null || c.value === '';

/**
 * What Insert Chart plots from a range, as Excel works it out: a first row
 * of text names the series and a first column of text or dates gives the
 * categories (both, when the top-left cell is empty); series run down the
 * columns unless the data is wider than it is tall. A scatter chart takes its
 * first column of numbers as the x values. A chart of several series gets
 * `placeholder` as its title, as Excel gives it "Chart Title".
 */
export function chartFromRange(rect: Rect, at: (row: number, col: number) => ChartCell, sheet: string, kind: NewChartKind, placeholder: string): NewChart {
  const rows = rect.r2 - rect.r1 + 1;
  const cols = rect.c2 - rect.c1 + 1;
  const corner = at(rect.r1, rect.c1);
  const span = (r1: number, c1: number, r2: number, c2: number) => {
    const out: ChartCell[] = [];
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.push(at(r, c));
    return out;
  };
  const filled = (cells: ChartCell[]) => cells.filter((c) => !isEmpty(c));
  // A single column's heading is its first cell; otherwise the first row past the labels.
  const topRow = filled(span(rect.r1, cols > 1 ? rect.c1 + 1 : rect.c1, rect.r1, rect.c2));
  const leftCol = filled(span(rect.r1 + 1, rect.c1, rect.r2, rect.c1));
  const blankCorner = isEmpty(corner) && rows > 1 && cols > 1;
  let header = rows > 1 && (blankCorner || (topRow.length > 0 && topRow.every(isLabel) && (cols === 1 || isLabel(corner) || isEmpty(corner))));
  let labels = cols > 1 && (blankCorner || (leftCol.length > 0 && leftCol.every((c) => isLabel(c) || c.date)));
  // A lone column (or row) of text is labels for nothing: plot what there is.
  if (header && rows - 1 < 1) header = false;
  if (labels && cols - 1 < 1) labels = false;

  const ref = (r1: number, c1: number, r2: number, c2: number) => `${quoteSheetName(sheet)}!${absolute(rangeToA1({ r1, c1, r2, c2 }))}`;
  const data = { r1: rect.r1 + (header ? 1 : 0), c1: rect.c1 + (labels ? 1 : 0), r2: rect.r2, c2: rect.c2 };
  // Series down the columns, unless there are more columns of data than rows.
  const byCols = data.r2 - data.r1 + 1 >= data.c2 - data.c1 + 1;
  const lines: { cells: ChartCell[]; ref: string; name?: { ref: string; text: string } }[] = [];
  const along: { cells: ChartCell[]; ref: string } | undefined = labels && byCols
    ? { cells: span(data.r1, rect.c1, data.r2, rect.c1), ref: ref(data.r1, rect.c1, data.r2, rect.c1) }
    : header && !byCols
      ? { cells: span(rect.r1, data.c1, rect.r1, data.c2), ref: ref(rect.r1, data.c1, rect.r1, data.c2) }
      : undefined;
  if (byCols) {
    for (let c = data.c1; c <= data.c2; c++) {
      const name = header ? at(rect.r1, c) : undefined;
      lines.push({ cells: span(data.r1, c, data.r2, c), ref: ref(data.r1, c, data.r2, c), ...(name && { name: { ref: ref(rect.r1, c, rect.r1, c), text: name.text } }) });
    }
  } else {
    for (let r = data.r1; r <= data.r2; r++) {
      const name = labels ? at(r, rect.c1) : undefined;
      lines.push({ cells: span(r, data.c1, r, data.c2), ref: ref(r, data.c1, r, data.c2), ...(name && { name: { ref: ref(r, rect.c1, r, rect.c1), text: name.text } }) });
    }
  }
  let categories: NewChartSeries['categories'] = along && {
    ref: along.ref,
    points: along.cells.map((c) => (c.date && typeof c.value === 'number' ? c.value : c.text || null)),
    numeric: along.cells.every((c) => c.date || isEmpty(c)),
    format: along.cells.find((c) => c.date)?.format ?? 'General',
  };
  // A scatter chart's x values are its first series of numbers, when it has more than one.
  if (kind === 'scatter' && !categories && lines.length > 1) {
    const x = lines.shift()!;
    categories = { ref: x.ref, points: x.cells.map((c) => (typeof c.value === 'number' ? c.value : null)), numeric: true, format: x.cells.find((c) => typeof c.value === 'number')?.format ?? 'General' };
  }
  const series = lines.map((l): NewChartSeries => ({
    ...(l.name && { name: l.name }),
    ...(categories && { categories }),
    values: { ref: l.ref, points: l.cells.map((c) => (typeof c.value === 'number' ? c.value : null)), format: l.cells.find((c) => typeof c.value === 'number')?.format ?? 'General' },
  }));
  const round = kind === 'pie' || kind === 'doughnut';
  return { kind, ...(series.length > 1 && !round && { title: placeholder }), legend: series.length > 1 || round, series };
}

/** A1:B2 as $A$1:$B$2. */
const absolute = (a1: string) => a1.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const num = (n: number) => (Number.isFinite(n) ? String(n) : '0');

function strCache(points: (string | number | null)[]): string {
  const pts = points.map((p, i) => (p === null ? '' : `<c:pt idx="${i}"><c:v>${escapeXml(String(p))}</c:v></c:pt>`)).join('');
  return `<c:strCache><c:ptCount val="${points.length}"/>${pts}</c:strCache>`;
}

function numCache(points: (string | number | null)[], format: string): string {
  const pts = points.map((p, i) => (typeof p === 'number' ? `<c:pt idx="${i}"><c:v>${num(p)}</c:v></c:pt>` : '')).join('');
  return `<c:numCache><c:formatCode>${escapeXml(format)}</c:formatCode><c:ptCount val="${points.length}"/>${pts}</c:numCache>`;
}

/** A series' parts, in the order the chart schema has them for its kind. */
function seriesXml(s: NewChartSeries, i: number, kind: NewChartKind): string {
  const tx = s.name ? `<c:tx><c:strRef><c:f>${escapeXml(s.name.ref)}</c:f>${strCache([s.name.text])}</c:strRef></c:tx>` : '';
  const cat = s.categories
    ? s.categories.numeric
      ? `<c:numRef><c:f>${escapeXml(s.categories.ref)}</c:f>${numCache(s.categories.points, s.categories.format)}</c:numRef>`
      : `<c:strRef><c:f>${escapeXml(s.categories.ref)}</c:f>${strCache(s.categories.points)}</c:strRef>`
    : '';
  const val = `<c:numRef><c:f>${escapeXml(s.values.ref)}</c:f>${numCache(s.values.points, s.values.format)}</c:numRef>`;
  const head = `<c:idx val="${i}"/><c:order val="${i}"/>${tx}`;
  switch (kind) {
    case 'column':
    case 'bar':
      return `<c:ser>${head}<c:invertIfNegative val="0"/>${cat && `<c:cat>${cat}</c:cat>`}<c:val>${val}</c:val></c:ser>`;
    case 'line':
      return `<c:ser>${head}<c:marker><c:symbol val="none"/></c:marker>${cat && `<c:cat>${cat}</c:cat>`}<c:val>${val}</c:val><c:smooth val="0"/></c:ser>`;
    case 'area':
    case 'pie':
    case 'doughnut':
      return `<c:ser>${head}${cat && `<c:cat>${cat}</c:cat>`}<c:val>${val}</c:val></c:ser>`;
    case 'scatter':
      return `<c:ser>${head}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>${cat && `<c:xVal>${cat}</c:xVal>`}<c:yVal>${val}</c:yVal><c:smooth val="0"/></c:ser>`;
  }
}

/** An axis, as Excel's default chart has it: labels beside it, no tick marks, gridlines on the value axis. */
function axisXml(tag: 'catAx' | 'valAx', id: number, cross: number, pos: 'b' | 'l', gridlines: boolean, format: string, between: 'between' | 'midCat' = 'between'): string {
  const head = `<c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${gridlines ? '<c:majorGridlines/>' : ''}`;
  const ticks = `<c:numFmt formatCode="${escapeXml(format)}" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/>`;
  return tag === 'catAx'
    ? `<c:catAx>${head}${ticks}<c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
    : `<c:valAx>${head}${ticks}<c:crossBetween val="${between}"/></c:valAx>`;
}

/** The chart part for a new chart, as Excel writes a default one of its kind. */
export function chartPartXml(chart: NewChart): string {
  const series = chart.series.map((s, i) => seriesXml(s, i, chart.kind)).join('');
  const valueFormat = chart.series[0]?.values.format ?? 'General';
  const catFormat = chart.series[0]?.categories?.format ?? 'General';
  const axes = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  let plot: string;
  switch (chart.kind) {
    case 'column':
    case 'bar': {
      const dir = chart.kind === 'column' ? 'col' : 'bar';
      plot =
        `<c:barChart><c:barDir val="${dir}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${series}<c:gapWidth val="219"/><c:overlap val="-27"/>${axes}</c:barChart>` +
        axisXml('catAx', 500000001, 500000002, chart.kind === 'column' ? 'b' : 'l', false, catFormat) +
        axisXml('valAx', 500000002, 500000001, chart.kind === 'column' ? 'l' : 'b', true, valueFormat);
      break;
    }
    case 'line':
    case 'area': {
      const tag = chart.kind === 'line' ? 'lineChart' : 'areaChart';
      const marker = chart.kind === 'line' ? '<c:marker val="1"/>' : '';
      plot =
        `<c:${tag}><c:grouping val="standard"/><c:varyColors val="0"/>${series}${marker}${axes}</c:${tag}>` +
        axisXml('catAx', 500000001, 500000002, 'b', false, catFormat) +
        axisXml('valAx', 500000002, 500000001, 'l', true, valueFormat);
      break;
    }
    case 'pie':
      plot = `<c:pieChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/></c:pieChart>`;
      break;
    case 'doughnut':
      plot = `<c:doughnutChart><c:varyColors val="1"/>${series}<c:firstSliceAng val="0"/><c:holeSize val="75"/></c:doughnutChart>`;
      break;
    case 'scatter':
      plot =
        `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${series}${axes}</c:scatterChart>` +
        axisXml('valAx', 500000001, 500000002, 'b', false, catFormat, 'midCat') +
        axisXml('valAx', 500000002, 500000001, 'l', true, valueFormat, 'midCat');
      break;
  }
  const title = chart.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : `<c:autoTitleDeleted val="${chart.series.length === 1 ? 0 : 1}"/>`;
  const legend = chart.legend ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : '';
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
    `<c:chartSpace xmlns:c="${C_NS}" xmlns:a="${A_NS}" xmlns:r="${R_NS}"><c:roundedCorners val="0"/>` +
    `<c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    '</c:chartSpace>'
  );
}

const EMU_PER_PX = 9525;

function cellAnchor(p: string, tag: 'from' | 'to', a: Anchor): string {
  return `<${p}${tag}><${p}col>${a.col}</${p}col><${p}colOff>${a.colOffPx * EMU_PER_PX}</${p}colOff><${p}row>${a.row}</${p}row><${p}rowOff>${a.rowOffPx * EMU_PER_PX}</${p}rowOff></${p}${tag}>`;
}

/** An anchor's from or to, written in the drawing's own prefix. */
export const anchorXml = cellAnchor;

/** The drawing anchor that puts a chart on its sheet between two cells, pointing at its part by relationship id. */
export function chartAnchorXml(p: string, from: Anchor, to: Anchor, id: number, name: string, rel: string): string {
  return (
    `<${p}twoCellAnchor editAs="oneCell">${cellAnchor(p, 'from', from)}${cellAnchor(p, 'to', to)}` +
    `<${p}graphicFrame macro=""><${p}nvGraphicFramePr><${p}cNvPr id="${id}" name="${escapeXml(name)}"/><${p}cNvGraphicFramePr/></${p}nvGraphicFramePr>` +
    `<${p}xfrm><a:off xmlns:a="${A_NS}" x="0" y="0"/><a:ext xmlns:a="${A_NS}" cx="0" cy="0"/></${p}xfrm>` +
    `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${C_NS}"><c:chart xmlns:c="${C_NS}" xmlns:r="${R_NS}" r:id="${rel}"/></a:graphicData></a:graphic>` +
    `</${p}graphicFrame><${p}clientData/></${p}twoCellAnchor>`
  );
}
