/**
 * Changes to a chart part, as Excel's Change Chart Type, Select Data and
 * chart elements make them. A title, legend or series changed in place
 * keeps everything else about the chart, its formatting included; another
 * type of chart gets a new plot area, its series in the theme's colours.
 */
import type { ChartSpec } from './charts';
import { plotAreaXml, seriesCategoriesXml, seriesTxXml, seriesValuesXml, type NewChartKind, type NewChartSeries, type SeriesRefs } from './chartWriter';
import { rootTag, siblings, type Sibling } from './hyperlinks';
import { escapeXml, findTagEnd, firstElement } from './xml';

export type LegendPosition = 'top' | 'bottom' | 'left' | 'right';

/** What Excel's chart dialogs change: the type, the title (empty for none), the legend (null for none) and the data. */
export interface ChartEdit {
  kind?: NewChartKind;
  title?: string;
  legend?: LegendPosition | null;
  /** New data from a range, its series down the columns or along the rows. */
  range?: { ref: string; byColumns: boolean };
  /** Each series' name, values and categories, as references. */
  series?: SeriesRefs[];
}

/** The chart's type, as Insert Chart names it; undefined for kinds and combinations it doesn't make. */
export function chartKindOf(spec: ChartSpec | undefined): NewChartKind | undefined {
  const g = spec?.groups[0];
  if (!g || spec.unsupported || spec.groups.length > 1) return undefined;
  return g.kind === 'bar' ? (g.horizontal ? 'bar' : 'column') : g.kind;
}

/** A chart's series as references, as Select Data shows them. */
export function seriesRefsOf(spec: ChartSpec | undefined): SeriesRefs[] {
  return (spec?.groups ?? []).flatMap((g) =>
    g.series.map((s) => ({
      ...(s.name && { name: s.name.formula ?? String(s.name.cache[0] ?? '') }),
      values: s.values.formula ?? '',
      ...(s.categories?.formula && { categories: s.categories.formula }),
    })),
  );
}

const C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The c:, a: and r: prefixes the new parts are written with, declared on the root where the file hasn't. */
function withNamespaces(xml: string): string {
  const root = rootTag(xml);
  if (!root) return xml;
  const add = [
    ['c', C_NS],
    ['a', A_NS],
    ['r', R_NS],
  ]
    .filter(([p]) => root.attrs[`xmlns:${p}`] === undefined)
    .map(([p, ns]) => ` xmlns:${p}="${ns}"`)
    .join('');
  if (!add) return xml;
  const at = root.empty ? root.end - 1 : root.end;
  return xml.slice(0, at) + add + xml.slice(at);
}

/** An element's direct children. */
function childrenOf(xml: string, el: { start: number; end: number }): Sibling[] {
  const open = findTagEnd(xml, el.start);
  if (xml.charCodeAt(open - 1) === 47) return [];
  return siblings(xml, open + 1).filter((c) => c.end <= el.end);
}

const replace = (xml: string, at: { start: number; end: number }, text: string) => xml.slice(0, at.start) + text + xml.slice(at.end);
const insert = (xml: string, at: number, text: string) => xml.slice(0, at) + text + xml.slice(at);
/** Just inside an element, before its first child. */
const inside = (xml: string, el: { start: number }) => findTagEnd(xml, el.start) + 1;
const all = (xml: string) => ({ start: 0, end: xml.length });

/** The first match of an element in drawing XML, whole: <a:bodyPr …/> or <a:bodyPr>…</a:bodyPr>. */
const piece = (xml: string, name: string) => new RegExp(`<a:${name}\\b[^>]*?(?:/>|>[\\s\\S]*?</a:${name}>)`).exec(xml)?.[0];

/** A title's text, in a run that keeps the old title's text box, paragraph and run formatting. */
function titleTx(text: string, old = ''): string {
  return (
    `<c:tx><c:rich>${piece(old, 'bodyPr') ?? '<a:bodyPr/>'}${piece(old, 'lstStyle') ?? '<a:lstStyle/>'}` +
    `<a:p>${piece(old, 'pPr') ?? ''}<a:r>${piece(old, 'rPr') ?? ''}<a:t>${escapeXml(text)}</a:t></a:r></a:p></c:rich></c:tx>`
  );
}

/** The chart's title set to `text`, or taken off when it is empty, with Excel's own title off too. */
function setTitle(xml: string, text: string): string {
  const chart = firstElement(xml, 'chart');
  if (!chart) return xml;
  const kids = childrenOf(xml, chart);
  const title = kids.find((k) => k.name === 'title');
  const deleted = kids.find((k) => k.name === 'autoTitleDeleted');
  let next = '';
  if (text && title) {
    const t = xml.slice(title.start, title.end);
    const tx = childrenOf(t, all(t)).find((k) => k.name === 'tx');
    next = tx ? replace(t, tx, titleTx(text, t.slice(tx.start, tx.end))) : insert(t, inside(t, all(t)), titleTx(text));
  } else if (text) {
    next = `<c:title>${titleTx(text)}<c:overlay val="0"/></c:title>`;
  }
  // The title and its flag are the chart's first children, side by side: both are written afresh where they were.
  const from = (title ?? deleted)?.start ?? inside(xml, chart);
  const to = (deleted ?? title)?.end ?? from;
  return xml.slice(0, from) + next + `<c:autoTitleDeleted val="${text ? 0 : 1}"/>` + xml.slice(to);
}

const LEGEND_POS: Record<LegendPosition, string> = { top: 't', bottom: 'b', left: 'l', right: 'r' };

/** The legend placed at `pos`, added after the plot area if there was none; or taken away. */
function setLegend(xml: string, pos: LegendPosition | null): string {
  const chart = firstElement(xml, 'chart');
  if (!chart) return xml;
  const kids = childrenOf(xml, chart);
  const legend = kids.find((k) => k.name === 'legend');
  const plot = kids.find((k) => k.name === 'plotArea');
  if (pos === null) return legend ? replace(xml, legend, '') : xml;
  const fresh = `<c:legend><c:legendPos val="${LEGEND_POS[pos]}"/><c:overlay val="0"/></c:legend>`;
  if (!legend) return plot ? insert(xml, plot.end, fresh) : xml;
  const at = childrenOf(xml, legend).find((k) => k.name === 'legendPos');
  if (at) return replace(xml, at, `<c:legendPos val="${LEGEND_POS[pos]}"/>`);
  if (xml.charCodeAt(findTagEnd(xml, legend.start) - 1) === 47) return replace(xml, legend, fresh);
  return insert(xml, inside(xml, legend), `<c:legendPos val="${LEGEND_POS[pos]}"/>`);
}

/** A series element with its name, categories and values set; `idx` renumbers it. */
function patchSeries(ser: string, s: NewChartSeries, idx?: number): string {
  let out = ser;
  const kids = () => childrenOf(out, all(out));
  const scatter = kids().some((k) => k.name === 'xVal' || k.name === 'yVal');
  // From the end of the series back: values, categories, name, then its number.
  const val = kids().find((k) => k.name === (scatter ? 'yVal' : 'val'));
  if (val) out = replace(out, val, seriesValuesXml(scatter ? 'yVal' : 'val', s.values));
  const cat = kids().find((k) => k.name === (scatter ? 'xVal' : 'cat'));
  const catXml = s.categories ? seriesCategoriesXml(scatter ? 'xVal' : 'cat', s.categories) : '';
  if (cat) out = replace(out, cat, catXml);
  else if (catXml) {
    const before = kids().find((k) => k.name === (scatter ? 'yVal' : 'val'));
    if (before) out = insert(out, before.start, catXml);
  }
  const tx = kids().find((k) => k.name === 'tx');
  const txXml = s.name ? seriesTxXml(s.name) : '';
  if (tx) out = replace(out, tx, txXml);
  else if (txXml) {
    const order = kids().find((k) => k.name === 'order') ?? kids().find((k) => k.name === 'idx');
    if (order) out = insert(out, order.end, txXml);
  }
  if (idx !== undefined) {
    for (const name of ['order', 'idx']) {
      const el = kids().find((k) => k.name === name);
      if (el) out = replace(out, el, `<c:${name} val="${idx}"/>`);
    }
  }
  return out;
}

/** The chart's series set to `series` in place: each keeps its own formatting, new ones take the last one's, and those left over go. */
function setSeries(xml: string, series: NewChartSeries[]): string {
  const plot = firstElement(xml, 'plotArea');
  if (!plot) return xml;
  const groups = childrenOf(xml, plot).filter((g) => g.name.endsWith('Chart'));
  const sers = groups.flatMap((g) => childrenOf(xml, g).filter((c) => c.name === 'ser'));
  if (!sers.length) return xml;
  let next = 1 + Math.max(...sers.map((s) => Number(/<(?:\w+:)?idx val="(\d+)"/.exec(xml.slice(s.start, s.end))?.[1] ?? 0)));
  const last = xml.slice(sers[sers.length - 1].start, sers[sers.length - 1].end);
  // A new series takes the last one's look, but not its points' own colours.
  const template = childrenOf(last, all(last))
    .filter((k) => k.name === 'dPt')
    .reverse()
    .reduce((t, k) => replace(t, k, ''), last);
  const texts = series.map((s, i) => (i < sers.length ? patchSeries(xml.slice(sers[i].start, sers[i].end), s) : patchSeries(template, s, next++)));
  const extra = texts.slice(sers.length).join('');
  let out = xml;
  for (let i = sers.length - 1; i >= 0; i--) out = replace(out, sers[i], (texts[i] ?? '') + (i === sers.length - 1 ? extra : ''));
  return out;
}

/**
 * A chart part changed as Excel's chart dialogs change it. With `rebuild`
 * (another type of chart), the plot area is made afresh for `kind` with
 * `series`; otherwise new series are set in place.
 */
export function editChartXml(
  xml: string,
  edit: { kind: NewChartKind; rebuild: boolean; series?: NewChartSeries[]; title?: string; legend?: LegendPosition | null },
): string {
  let out = withNamespaces(xml);
  if (edit.rebuild && edit.series) {
    const plot = firstElement(out, 'plotArea');
    if (plot) out = replace(out, plot, plotAreaXml({ kind: edit.kind, series: edit.series }));
  } else if (edit.series) {
    out = setSeries(out, edit.series);
  }
  if (edit.legend !== undefined) out = setLegend(out, edit.legend);
  if (edit.title !== undefined) out = setTitle(out, edit.title);
  return out;
}
