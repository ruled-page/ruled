import type { ChartConfiguration, ChartDataset, LegendItem, ScaleOptions } from 'chart.js';
import { msg, type Text } from '../i18n/message';
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

  const categoryScale = (): ScaleOptions => {
    const category = plot.axes.category;
    return {
      type: 'category',
      offset: true,
      reverse: !!category?.reversed,
      display: !category?.hidden,
      grid: { display: !!category?.gridlines, color: look.gridline },
      border: { color: look.axis },
      ticks: { color: look.muted, font },
      title: titleOf(category?.title),
    } as ScaleOptions;
  };

  if (kind === 'radar') {
    const g = groups[0];
    const value = plot.axes.value;
    return {
      type: 'radar',
      data: {
        labels: plot.categories,
        datasets: g.series.map((s) => ({
          label: text(s.name),
          data: s.values.map((v) => (v === null && plot.blanks === 'zero' ? 0 : v)),
          ...lineStyle(s),
          // A filled radar's series are solid areas, later ones over earlier ones.
          ...(g.filled ? { fill: true, showLine: true, borderWidth: 0, pointRadius: 0, pointStyle: 'rect' } : { fill: false, showLine: s.style.line, borderWidth: s.style.line ? (s.style.lineWidth ?? 2) : 0 }),
        })) as ChartDataset[],
      },
      options: {
        ...options,
        scales: {
          r: {
            ...axisBounds(g.series.flatMap((s) => s.values)),
            min: value?.min,
            max: value?.max,
            grid: { color: look.gridline },
            angleLines: { color: look.gridline },
            pointLabels: { color: look.muted, font },
            ticks: { ...ticks(value?.format ?? 'General'), backdropColor: 'transparent', stepSize: value?.majorUnit, display: !value?.hidden },
          },
        },
      } as ChartConfiguration['options'],
    };
  }

  if (kind === 'bubble') {
    const g = groups[0];
    const sizes = g.series.flatMap((s) => s.sizes ?? []).filter((v): v is number => v !== null && v > 0);
    const biggest = Math.max(Number.MIN_VALUE, ...sizes);
    // Excel's largest bubble is about a quarter of the plot across at 100%; bubbleScale shrinks or grows it.
    const largest = (30 * (g.bubbles?.scale ?? 100)) / 100;
    const radius = (v: number | null | undefined) => (v === null || v === undefined || !(v > 0) ? 0 : g.bubbles?.area === false ? (largest * v) / biggest : largest * Math.sqrt(v / biggest));
    const x = plot.axes.category;
    const y = plot.axes.value;
    return {
      type: 'bubble',
      data: {
        datasets: g.series.map((s) => {
          const c = color(s);
          return {
            label: text(s.name),
            data: s.values.flatMap((v, i) => (v === null ? [] : [{ x: s.x?.[i] ?? i + 1, y: v, r: radius(s.sizes?.[i]) }])),
            backgroundColor: c,
            borderColor: look.background,
            borderWidth: 1,
          };
        }) as ChartDataset[],
      },
      options: {
        ...options,
        scales: {
          x: valueScale(x, 'bottom', g.series.flatMap((s) => s.x ?? []), false, x?.format ?? 'General'),
          y: valueScale(y, 'left', g.series.flatMap((s) => s.values), false, y?.format ?? 'General'),
        },
      } as ChartConfiguration['options'],
    };
  }

  if (kind === 'stock') {
    const g = groups[0];
    const n = plot.categories.length;
    const known = (i: number) => g.series.map((s) => s.values[i]).filter((v): v is number => v !== null && v !== undefined);
    const datasets: ChartDataset[] = [];
    // A line from each day's highest value to its lowest.
    if (g.stock?.hiLow) {
      datasets.push({
        type: 'bar',
        label: '',
        data: Array.from({ length: n }, (_, i) => {
          const v = known(i);
          return v.length ? [Math.min(...v), Math.max(...v)] : null;
        }),
        backgroundColor: look.text,
        barThickness: 1.5,
        grouped: false,
        order: 2,
      } as ChartDataset);
    }
    // A bar from the first series (open) to the last (close): up days light, down days dark.
    const upDown = g.stock?.upDown;
    if (upDown && g.series.length > 1) {
      const first = g.series[0].values;
      const last = g.series[g.series.length - 1].values;
      const bars = Array.from({ length: n }, (_, i) => (first[i] === null || last[i] === null || first[i] === undefined || last[i] === undefined ? null : [first[i]!, last[i]!]));
      datasets.push({
        type: 'bar',
        label: '',
        data: bars,
        backgroundColor: bars.map((b) => (b && b[1] >= b[0] ? (upDown.up ?? look.background) : (upDown.down ?? look.text))),
        borderColor: look.text,
        borderWidth: 1,
        grouped: false,
        categoryPercentage: 1 / (1 + g.gapWidth / 100),
        barPercentage: 1,
        order: 1,
      } as ChartDataset);
    }
    datasets.push(...(g.series.map((s) => ({ type: 'line', label: text(s.name), data: s.values, ...lineStyle(s), order: 0 })) as ChartDataset[]));
    const legend = options.plugins!.legend!;
    return {
      type: 'bar',
      data: { labels: plot.categories, datasets },
      options: {
        ...options,
        // The lines and bars aren't series of their own, so the legend leaves them out.
        plugins: { ...options.plugins, legend: { ...legend, labels: { ...legend.labels, filter: (item: LegendItem) => !!item.text } } },
        scales: {
          x: categoryScale(),
          y: valueScale(plot.axes.value, 'left', g.series.flatMap((s) => s.values), false, plot.axes.value?.format ?? 'General'),
        },
      } as ChartConfiguration['options'],
    };
  }

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

  if (kind === 'waterfall' || kind === 'funnel') return stepsConfig(plot, groups[0], palette, look, options, text, valueScale);

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

/**
 * Excel 2016's waterfall and funnel charts, as floating bars. A waterfall's
 * bars step up or down from the running total, or stand from zero for the
 * points that are totals, in the theme's first three accents for increases,
 * decreases and totals, as Excel colours them. A funnel's bars are centred,
 * the first at the top, with no value axis.
 */
function stepsConfig(
  plot: ChartPlot,
  g: PlotGroup,
  palette: string[],
  look: ChartLook,
  options: ChartConfiguration['options'],
  text: (t: Text) => string,
  valueScale: (axis: PlotAxis | undefined, position: 'left' | 'bottom', values: (number | null)[], percentStacked: boolean, format: string) => ScaleOptions,
): ChartConfiguration {
  const s = g.series[0];
  const font = { family: look.font, size: 12 };
  const funnel = g.kind === 'funnel';
  const bar = { type: 'bar' as const, borderWidth: 0, grouped: false, categoryPercentage: 1 / (1 + g.gapWidth / 100), barPercentage: 1, pointStyle: 'rect' };
  const category = plot.axes.category;
  const categoryScale = {
    type: 'category',
    position: funnel ? 'left' : 'bottom',
    reverse: !!category?.reversed,
    display: !category?.hidden,
    grid: { display: false },
    border: { color: look.axis },
    ticks: { color: look.muted, font },
    title: { display: !!category?.title, text: category?.title ?? '', color: look.text, font: { ...font, size: 13 } },
  } as ScaleOptions;
  let datasets: ChartDataset[];
  let values: number[];
  if (funnel) {
    const widths = s.values.map((v) => (v === null ? null : Math.abs(v)));
    values = widths.flatMap((w) => (w === null ? [] : [w / 2, -w / 2]));
    const color = s.style.color ?? palette[0];
    datasets = [
      {
        ...bar,
        label: text(s.name),
        data: widths.map((w) => (w === null ? null : [-w / 2, w / 2])),
        backgroundColor: widths.map((_, i) => s.style.pointColors?.[i] ?? color),
      } as ChartDataset,
    ];
  } else {
    const totals = new Set(g.subtotals ?? []);
    let run = 0;
    const steps = s.values.map((v, i) => {
      if (v === null) return null;
      if (totals.has(i)) {
        run = v;
        return { from: 0, to: v, kind: 2 };
      }
      const from = run;
      run += v;
      return { from, to: run, kind: v >= 0 ? 0 : 1 };
    });
    values = [0, ...steps.flatMap((p) => (p ? [p.from, p.to] : []))];
    const names = [msg('charts.waterfall.increase'), msg('charts.waterfall.decrease'), msg('charts.waterfall.total')];
    datasets = names
      .map((name, k) => ({
        ...bar,
        label: text(name),
        data: steps.map((p) => (p && p.kind === k ? [p.from, p.to] : null)),
        backgroundColor: steps.map((_, i) => s.style.pointColors?.[i] ?? palette[k % palette.length]),
      }))
      .filter((d) => d.data.some((v) => v !== null)) as ChartDataset[];
  }
  const value = funnel ? ({ type: 'linear', display: false, grid: { display: false } } as ScaleOptions) : valueScale(plot.axes.value, 'left', values, false, plot.axes.value?.format ?? s.format);
  return {
    type: 'bar',
    data: { labels: plot.categories, datasets },
    options: {
      ...options,
      indexAxis: funnel ? 'y' : 'x',
      plugins: { ...options?.plugins, legend: { ...options?.plugins?.legend, display: !!plot.legend } },
      scales: funnel ? { y: categoryScale, x: value } : { x: categoryScale, y: value },
    } as ChartConfiguration['options'],
  };
}
