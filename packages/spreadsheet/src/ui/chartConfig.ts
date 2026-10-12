import type { ChartConfiguration, ChartDataset, LegendItem, ScaleOptions } from 'chart.js';
import type { Text } from '../i18n/message';
import { formatValue, type ChartPlot, type PlotAxis, type PlotGroup, type PlotSeries } from '../state/charts';

/**
 * The editor's look as charts take it, read from its theme tokens: colours
 * and font for everything the workbook doesn't colour itself.
 */
export interface ChartLook {
  /** Colours for series the workbook doesn't colour; empty for the workbook theme's accents, as Excel uses. */
  palette: string[];
  background: string;
  text: string;
  muted: string;
  gridline: string;
  axis: string;
  font: string;
}

/** The --wbe-chart-colors token: colours separated by spaces or commas, or "auto". */
export function parsePalette(value: string): string[] {
  const v = value.trim();
  if (!v || v === 'auto') return [];
  // Commas inside rgb() and the like belong to the colour.
  return v.split(/[\s,]+(?![^(]*\))/).filter(Boolean);
}

/** Office's accents, for a chart with neither a palette nor a theme. */
const OFFICE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
/** Past the palette's end, colours repeat darker, then lighter, as Excel's do. */
const CYCLES = [0, -0.4, 0.3, -0.2, 0.5];

function mix(color: string, amount: number): string {
  const m = /^#([0-9a-f]{6})$/i.exec(color);
  if (!m || !amount) return color;
  const target = amount < 0 ? 0 : 255;
  const a = Math.abs(amount);
  return '#' + [0, 2, 4].map((i) => Math.round(parseInt(m[1].slice(i, i + 2), 16) * (1 - a) + target * a).toString(16).padStart(2, '0')).join('');
}

export function paletteColor(palette: string[], i: number): string {
  return mix(palette[i % palette.length], CYCLES[Math.floor(i / palette.length) % CYCLES.length]);
}

/**
 * Where Excel's automatic value axis starts or ends: at zero, unless the data
 * sit in a band narrower than a sixth of their size, when it leaves half the
 * band's height again below them (above, for negative data).
 */
export function axisBounds(values: (number | null)[]): { suggestedMin?: number; suggestedMax?: number; beginAtZero: boolean } {
  let min = Infinity;
  let max = -Infinity;
  for (const v of values) {
    if (v === null || !Number.isFinite(v)) continue;
    min = Math.min(min, v);
    max = Math.max(max, v);
  }
  if (min === Infinity || (min < 0 && max > 0) || max === 0) return { beginAtZero: true };
  const band = max - min;
  if (min >= 0) return band / max >= 1 / 6 ? { beginAtZero: true } : { beginAtZero: false, suggestedMin: min - band / 2 };
  return band / -min >= 1 / 6 ? { beginAtZero: true } : { beginAtZero: false, suggestedMax: max + band / 2 };
}

/** Each value as its share of the category's total, for 100% stacked charts. */
function shares(group: PlotGroup): (number | null)[][] {
  const n = Math.max(0, ...group.series.map((s) => s.values.length));
  const totals = Array.from({ length: n }, (_, i) => group.series.reduce((sum, s) => sum + Math.abs(s.values[i] ?? 0), 0));
  return group.series.map((s) => s.values.map((v, i) => (v === null ? null : totals[i] ? v / totals[i] : 0)));
}

/** Builds the Chart.js configuration for a chart's data, in the editor's look. */
export function chartConfig(plot: ChartPlot, look: ChartLook, text: (t: Text) => string, date1904 = false): ChartConfiguration {
  const palette = look.palette.length ? look.palette : plot.accents.length ? plot.accents : OFFICE;
  const font = { family: look.font, size: 12 };
  const color = (s: PlotSeries) => s.style.color ?? paletteColor(palette, s.index);
  const titleOf = (title: string | undefined) => ({ display: !!title, text: title ?? '', color: look.text, font: { ...font, size: 13 } });
  const ticks = (format: string) => ({ color: look.muted, font, callback: (v: number | string) => formatValue(Number(v), format, date1904) });
  const deleted = new Set(plot.legend?.deleted ?? []);
  const valueScale = (
    axis: PlotAxis | undefined,
    position: 'left' | 'right' | 'top' | 'bottom',
    values: (number | null)[],
    percentStacked: boolean,
    format: string,
    secondary = false,
  ): ScaleOptions =>
    ({
      type: axis?.logBase ? 'logarithmic' : 'linear',
      position,
      reverse: !!axis?.reversed,
      min: axis?.min ?? (percentStacked && values.every((v) => (v ?? 0) >= 0) ? 0 : undefined),
      max: axis?.max ?? (percentStacked ? 1 : undefined),
      ...(axis?.logBase ? {} : axisBounds(values)),
      // Excel leaves a little room above the data before the next round number.
      grace: axis?.max === undefined && !percentStacked ? '5%' : 0,
      display: !axis?.hidden,
      grid: { display: axis ? axis.gridlines : !secondary, color: look.gridline, drawOnChartArea: !secondary },
      border: { display: false },
      ticks: { ...ticks(format), stepSize: axis?.majorUnit },
      title: titleOf(axis?.title),
    }) as ScaleOptions;

  const groups = plot.groups.filter((g) => g.series.length);
  const kind = groups[0]?.kind ?? 'bar';
  const round = kind === 'pie' || kind === 'doughnut';
  const options: ChartConfiguration['options'] = {
    responsive: true,
    maintainAspectRatio: false,
    // Redrawn as cells change; nothing to animate, and nothing to point at through the overlay.
    animation: false,
    events: [],
    layout: { padding: 8 },
    plugins: {
      title: { display: !!plot.title, text: plot.title ?? '', color: look.text, font: { ...font, size: 16, weight: 'normal' }, padding: { top: 2, bottom: 8 } },
      legend: {
        display: !!plot.legend,
        position: plot.legend?.position ?? 'right',
        labels: {
          color: look.text,
          font,
          boxWidth: 12,
          // Each series' own sample: a line for lines, a square for bars and areas. Slices are squares.
          usePointStyle: !round,
          boxHeight: round ? 10 : 8,
          // In the order Excel lists them, whatever order they are drawn in.
          sort: (a: LegendItem, b: LegendItem) => (round ? (a.index ?? 0) - (b.index ?? 0) : (a.datasetIndex ?? 0) - (b.datasetIndex ?? 0)),
          filter: (item: LegendItem) => !deleted.has((round ? item.index : item.datasetIndex) ?? -1),
        },
      },
      tooltip: { enabled: false },
    },
  };

  if (round) {
    const g = groups[0];
    // A pie shows its first series; a doughnut has a ring for each.
    const series = g.kind === 'pie' ? g.series.slice(0, 1) : g.series;
    return {
      type: g.kind as 'pie' | 'doughnut',
      data: {
        labels: plot.categories,
        datasets: series.map((s) => ({
          label: text(s.name),
          // Excel draws a negative value's slice at its size.
          data: s.values.map((v) => Math.abs(v ?? 0)),
          backgroundColor: s.values.map((_, i) => s.style.pointColors?.[i] ?? (g.varyColors ? paletteColor(palette, i) : color(s))),
          borderColor: look.background,
          borderWidth: 1,
        })),
      },
      options: { ...options, rotation: g.firstSliceAngle ?? 0, cutout: g.kind === 'doughnut' ? `${g.holeSize ?? 50}%` : 0 } as ChartConfiguration['options'],
    };
  }

  const lineStyle = (s: PlotSeries) => {
    const c = color(s);
    return {
      borderColor: c,
      backgroundColor: c,
      borderWidth: s.style.line ? (s.style.lineWidth ?? 2.5) : 0,
      showLine: s.style.line,
      borderDash: s.style.dashed ? [6, 4] : [],
      pointRadius: s.style.markers ? 3.5 : 0,
      pointBackgroundColor: c,
      pointBorderColor: c,
      pointStyle: s.style.markers ? 'circle' : 'line',
      tension: s.style.smooth ? 0.4 : 0,
      spanGaps: plot.blanks === 'span',
    };
  };

  if (kind === 'scatter') {
    const x = plot.axes.category;
    const y = plot.axes.value;
    const xs = groups.flatMap((g) => g.series.flatMap((s) => s.x ?? []));
    const ys = groups.flatMap((g) => g.series.flatMap((s) => s.values));
    return {
      type: 'scatter',
      data: {
        datasets: groups.flatMap((g) =>
          g.series.map((s) => ({
            label: text(s.name),
            data: s.values.map((v, i) => ({ x: s.x?.[i] ?? i + 1, y: v === null && plot.blanks === 'zero' ? 0 : v })),
            ...lineStyle(s),
          })),
        ) as ChartDataset[],
      },
      options: {
        ...options,
        scales: {
          x: valueScale(x, 'bottom', xs, false, x?.format ?? 'General'),
          y: valueScale(y, 'left', ys, false, y?.format ?? 'General'),
        },
      } as ChartConfiguration['options'],
    };
  }

  // Bars, lines and areas, alone or combined; Excel's bar chart runs sideways.
  const sideways = groups.some((g) => g.kind === 'bar' && g.horizontal);
  const index = sideways ? 'y' : 'x';
  const value = sideways ? 'x' : 'y';
  const second = sideways ? 'x2' : 'y2';
  const plotted: Record<string, (number | null)[]> = { [value]: [], [second]: [] };
  const datasets = groups.flatMap((g, gi) => {
    const stacked = g.grouping === 'stacked' || g.grouping === 'percentStacked';
    const values = g.grouping === 'percentStacked' ? shares(g) : g.series.map((s) => s.values);
    const axes = { [`${index}AxisID`]: index, [`${value}AxisID`]: g.secondary ? second : value };
    // Stacks start from zero, so the axis does too.
    plotted[g.secondary ? second : value].push(...(stacked ? [0] : values.flat()));
    return g.series.map((s, si) => {
      const data = values[si].map((v) => (v === null && plot.blanks === 'zero' ? 0 : v));
      const stack = stacked ? `g${gi}` : undefined;
      if (g.kind === 'bar') {
        const slots = stacked ? 1 : g.series.length;
        const c = color(s);
        const single = groups.length === 1 && g.series.length === 1;
        return {
          type: 'bar' as const,
          label: text(s.name),
          data,
          backgroundColor: data.map((_, i) => s.style.pointColors?.[i] ?? (g.varyColors && single ? paletteColor(palette, i) : c)),
          borderWidth: 0,
          pointStyle: 'rect',
          stack,
          // Excel's gap width is the space between clusters as a share of one bar.
          categoryPercentage: slots / (slots + g.gapWidth / 100),
          barPercentage: 1,
          ...axes,
        };
      }
      const area = g.kind === 'area';
      return {
        type: 'line' as const,
        label: text(s.name),
        data,
        ...lineStyle(s),
        ...(area ? { borderWidth: 0, showLine: true, pointRadius: 0, pointStyle: 'rect', tension: 0, fill: stacked && si > 0 ? '-1' : 'origin' } : { fill: false }),
        stack,
        // Lines go over bars, and later areas over earlier ones; stacked series keep their order.
        order: stacked ? 0 : area ? -si : -1,
        ...axes,
      };
    });
  });

  const percent = (secondary: boolean) => groups.some((g) => g.secondary === secondary && g.grouping === 'percentStacked');
  const category = plot.axes.category;
  const onlyAreas = groups.every((g) => g.kind === 'area');
  const scales: Record<string, ScaleOptions> = {
    [index]: {
      type: 'category',
      position: sideways ? 'left' : 'bottom',
      // Excel lists a bar chart's first category at the bottom.
      reverse: sideways ? !category?.reversed : !!category?.reversed,
      // Bars and line points sit between the ticks; areas run edge to edge.
      offset: !onlyAreas,
      display: !category?.hidden,
      grid: { display: !!category?.gridlines, color: look.gridline },
      border: { color: look.axis },
      ticks: { color: look.muted, font },
      title: titleOf(category?.title),
    } as ScaleOptions,
    [value]: valueScale(plot.axes.value, sideways ? 'bottom' : 'left', plotted[value], percent(false), plot.axes.value?.format ?? 'General'),
  };
  if (groups.some((g) => g.secondary)) {
    scales[second] = valueScale(plot.axes.secondary, sideways ? 'top' : 'right', plotted[second], percent(true), plot.axes.secondary?.format ?? 'General', true);
  }
  return {
    type: groups.some((g) => g.kind === 'bar') ? 'bar' : 'line',
    data: { labels: plot.categories, datasets: datasets as ChartDataset[] },
    options: { ...options, indexAxis: index, scales } as ChartConfiguration['options'],
  };
}
