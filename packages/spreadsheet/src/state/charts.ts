import type { ChartAxis, ChartData, ChartGroup, ChartSeries, ChartSpec } from '../core/charts';
import { VT, type DisplayStyle } from '../core/types';
import type { Translator } from '../i18n';
import { msg, type Text } from '../i18n/message';
import { displayText, type CellRecord } from './format';
import type { SheetView } from './sheetView';

/** The workbook as charts read it: its sheets as they are now. */
export interface ChartSource {
  sheets: SheetView[];
  styles: DisplayStyle[];
  date1904: boolean;
}

export interface PlotSeries {
  /** The series' name, or Excel's "Series1" for one without. */
  name: Text;
  /** A value per point; null where there is nothing to plot. */
  values: (number | null)[];
  /** A scatter chart's x value for each point. */
  x?: (number | null)[];
  /** Number format the values are shown in. */
  format: string;
  /** Its position among all the chart's series, which picks its default colour. */
  index: number;
  style: ChartSeries;
}

export interface PlotGroup extends Omit<ChartGroup, 'series'> {
  series: PlotSeries[];
}

export interface PlotAxis extends Omit<ChartAxis, 'title' | 'format'> {
  title?: string;
  /** Number format of its labels. */
  format: string;
}

/** A chart with its data read from the cells as they are now, ready to draw and describe. */
export interface ChartPlot {
  title?: string;
  /** Category labels as their cells show them; for a pie, its slices. */
  categories: string[];
  groups: PlotGroup[];
  axes: { category?: PlotAxis; value?: PlotAxis; secondary?: PlotAxis };
  legend?: ChartSpec['legend'];
  blanks: ChartSpec['blanks'];
  accents: string[];
}

/** Points read from a reference at most, so a whole column doesn't plot a million rows. */
const MAX_POINTS = 10_000;

interface Cell {
  rec: CellRecord | null;
  style: DisplayStyle;
  hidden: boolean;
}

/** The cells a reference covers, row by row, within the sheet's used range; undefined when they can't be read. */
function cellsOf(data: ChartData | undefined, src: ChartSource): Cell[] | undefined {
  const ref = data?.ref;
  const view = ref ? src.sheets[ref.sheet] : undefined;
  if (!ref || !view || view.info.kind !== 'worksheet') return undefined;
  const { maxRow, maxCol } = view.extent();
  const r = ref.rect;
  const r2 = Math.min(r.r2, Math.max(maxRow, r.r1));
  const c2 = Math.min(r.c2, Math.max(maxCol, r.c1));
  const hiddenRows = new Set(view.info.hiddenRows);
  const hiddenCols = new Set(view.info.hiddenCols);
  const out: Cell[] = [];
  for (let row = r.r1; row <= r2 && out.length < MAX_POINTS; row++) {
    for (let col = r.c1; col <= c2 && out.length < MAX_POINTS; col++) {
      const rec = view.get(row, col);
      out.push({ rec, style: src.styles[view.styleAt(row, col, rec)] ?? src.styles[0], hidden: hiddenRows.has(row) || hiddenCols.has(col) });
    }
  }
  return out;
}

const GENERAL = { numFmt: 'General' } as DisplayStyle;

/** What Excel plots for a cell: its number; nothing for an empty cell or an error; zero for text. */
function plotted(rec: CellRecord | null): number | null {
  if (!rec || rec.t === VT.None || rec.t === VT.Error) return null;
  return rec.t === VT.Number ? (rec.n ?? 0) : 0;
}

function label(rec: CellRecord | null, style: DisplayStyle, date1904: boolean): string {
  return rec ? displayText(rec, style, date1904).text : '';
}

/** A saved point as text: numbers in the cache's own format. */
function cachedLabel(v: string | number | null, format: string | undefined, date1904: boolean): string {
  if (v === null) return '';
  if (typeof v === 'string') return v;
  return label({ t: VT.Number, n: v, style: 0, flags: 0 }, { ...GENERAL, numFmt: format || 'General' }, date1904);
}

/** Text from a name or title: its cells joined with spaces, as Excel joins them, or what was saved. */
function textOf(data: ChartData | undefined, src: ChartSource): string | undefined {
  if (!data) return undefined;
  const cells = cellsOf(data, src);
  if (cells) return cells.map((c) => label(c.rec, c.style, src.date1904)).filter(Boolean).join(' ');
  return data.cache.map((v) => cachedLabel(v, data.format, src.date1904)).join(' ');
}

interface Read {
  values: (number | null)[];
  labels: string[];
  hidden: boolean[];
  format: string;
  /** Every point is a number or empty: as x values, they place the points. */
  numeric: boolean;
}

/** A series' values (or categories) as they are now; Excel's saved values when the cells can't be read. */
function read(data: ChartData | undefined, src: ChartSource): Read | undefined {
  if (!data) return undefined;
  const cells = cellsOf(data, src);
  if (cells) {
    return {
      values: cells.map((c) => plotted(c.rec)),
      labels: cells.map((c) => label(c.rec, c.style, src.date1904)),
      hidden: cells.map((c) => c.hidden),
      format: cells.find((c) => c.rec?.t === VT.Number)?.style.numFmt ?? cells[0]?.style.numFmt ?? 'General',
      numeric: cells.every((c) => !c.rec || c.rec.t === VT.Number || c.rec.t === VT.None),
    };
  }
  return {
    values: data.cache.map((v) => (typeof v === 'number' ? v : v === null ? null : 0)),
    labels: data.cache.map((v) => cachedLabel(v, data.format, src.date1904)),
    hidden: data.cache.map(() => false),
    format: data.format ?? 'General',
    numeric: data.cache.every((v) => typeof v !== 'string'),
  };
}

/**
 * A chart's data as it is now: each series read from its cells (or from the
 * values Excel saved, where they can't be read), points in hidden rows and
 * columns left out as Excel leaves them out, and labels as the cells show them.
 */
export function plotChart(spec: ChartSpec, src: ChartSource): ChartPlot {
  let index = 0;
  const groups = spec.groups.map((g) => ({
    group: g,
    series: g.series.map((s) => ({ style: s, index: index++, values: read(s.values, src), categories: read(s.categories, src) })),
  }));
  const first = groups.find((g) => g.series.length)?.series[0];
  const scatter = spec.groups[0]?.kind === 'scatter';
  // Points are laid out by the first series, whose categories label the axis.
  const count = Math.max(0, ...groups.flatMap((g) => g.series.map((s) => s.values?.values.length ?? 0)));
  const keep = Array.from({ length: count }, (_, i) => !spec.visibleOnly || !(first?.values?.hidden[i] || first?.categories?.hidden[i]));
  const pick = <T>(list: T[], fill: (i: number) => T): T[] => keep.flatMap((k, i) => (k ? [i < list.length ? list[i] : fill(i)] : []));
  const categories = pick(first?.categories?.labels ?? [], (i) => String(i + 1));

  const plotGroups: PlotGroup[] = groups.map(({ group, series }) => ({
    ...group,
    series: series.map((s): PlotSeries => ({
      name: textOf(s.style.name, src) || msg('charts.seriesName', { n: s.index + 1 }),
      values: pick(s.values?.values ?? [], () => null),
      // X values that aren't all numbers are plotted as 1, 2, 3…, as Excel does.
      x: scatter ? pick(s.categories?.numeric ? s.categories.values : [], (i) => i + 1) : undefined,
      format: s.values?.format ?? 'General',
      index: s.index,
      style: s.style,
    })),
  }));

  const formatFor = (axis: ChartAxis | undefined, secondary: boolean, x = false) => {
    if (axis?.format) return axis.format;
    const g = plotGroups.find((p) => p.secondary === secondary && p.series.length);
    if (g?.grouping === 'percentStacked') return '0%';
    if (x) return groups[0]?.series[0]?.categories?.format ?? 'General';
    return g?.series[0].format ?? 'General';
  };
  const axis = (a: ChartAxis | undefined, format: string): PlotAxis | undefined =>
    a && { ...a, title: textOf(a.title, src) || undefined, format };

  return {
    title: textOf(spec.title, src) || undefined,
    categories,
    groups: plotGroups,
    axes: {
      category: axis(spec.axes.category, formatFor(spec.axes.category, false, true)),
      value: axis(spec.axes.value, formatFor(spec.axes.value, false)),
      secondary: axis(spec.axes.secondary, formatFor(spec.axes.secondary, true)),
    },
    legend: spec.legend,
    blanks: spec.blanks,
    accents: spec.accents,
  };
}

/** The sheets a chart reads, so it is drawn again when one of them changes. */
export function chartSheets(spec: ChartSpec): number[] {
  const refs: (ChartData | undefined)[] = [spec.title, spec.axes.category?.title, spec.axes.value?.title, spec.axes.secondary?.title];
  for (const g of spec.groups) for (const s of g.series) refs.push(s.name, s.categories, s.values);
  return [...new Set(refs.flatMap((d) => (d?.ref ? [d.ref.sheet] : [])))].sort((a, b) => a - b);
}

/** A value as a cell with the series' format would show it. */
export function formatValue(n: number, format: string, date1904 = false): string {
  return label({ t: VT.Number, n, style: 0, flags: 0 }, { ...GENERAL, numFmt: format }, date1904);
}

/** Points listed per series in the text alternative; longer series end with how many more there are. */
const LISTED = 24;

/**
 * The chart in words, for screen readers: each series with its points
 * (category and value, as the cells show them) and the axes' titles.
 */
export function describeChart(plot: ChartPlot, t: Translator, date1904 = false): string {
  const parts: string[] = [];
  const series = plot.groups.flatMap((g) => (g.kind === 'pie' ? g.series.slice(0, 1) : g.series));
  if (!series.length || series.every((s) => !s.values.length)) return t('charts.noData');
  parts.push(t('charts.seriesCount', { count: series.length }));
  for (const s of series) {
    const points: string[] = [];
    s.values.slice(0, LISTED).forEach((v, i) => {
      const value = v === null ? t('charts.blank') : formatValue(v, s.format, date1904);
      if (s.x) points.push(t('charts.xy', { x: formatValue(s.x[i] ?? i + 1, plot.axes.category?.format ?? 'General', date1904), y: value }));
      else points.push(t('charts.point', { label: plot.categories[i] ?? String(i + 1), value }));
    });
    if (s.values.length > LISTED) points.push(t('charts.more', { count: s.values.length - LISTED }));
    parts.push(t('charts.series', { name: t.text(s.name), points: t.list(points) }));
  }
  // Bar charts (Excel's sideways ones) have their categories up the side.
  const sideways = plot.groups.some((g) => g.horizontal);
  const { category, value, secondary } = plot.axes;
  const across = sideways ? value : category;
  const up = sideways ? category : value;
  if (across?.title) parts.push(t('charts.horizontalAxis', { title: across.title }));
  if (up?.title) parts.push(t('charts.verticalAxis', { title: up.title }));
  if (secondary?.title) parts.push(t('charts.secondaryAxis', { title: secondary.title }));
  return parts.join(' ');
}
