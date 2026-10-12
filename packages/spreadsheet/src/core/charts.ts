/**
 * Chart definitions from DrawingML chart parts (xl/charts/chartN.xml): the
 * kind of chart, its series and where their data comes from, its titles,
 * axes and legend, and the colours the file gives its series. The editor
 * draws charts from these, reading each series' cells as they are now and
 * falling back to the values Excel saved with the chart where a reference
 * can't be followed.
 */
import { parseA1Ref } from '@borgar/fx';
import { MAX_COLS, MAX_ROWS } from '@sitepen/formula-engine/refs';
import { tintColor } from './styles';
import type { Rect } from './types';
import { attr, child, children, descendants, type XmlNode } from './xml';

export type ChartKind = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'bubble' | 'stock';
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked' | 'standard';

/** A series' name, categories or values: where they come from, and what Excel saved. */
export interface ChartData {
  /** The reference as the file writes it, such as "Data!$B$2:$B$5"; missing for values written into the chart. */
  formula?: string;
  /** The reference as one range on one sheet of this workbook, when it is that. */
  ref?: { sheet: number; rect: Rect };
  /** What Excel saved with the chart, point by point; points it didn't save are null. */
  cache: (string | number | null)[];
  /** Number format of the saved values, such as "General" or "mmm yyyy". */
  format?: string;
}

export interface ChartSeries {
  name?: ChartData;
  /** Labels along the category axis, or a scatter or bubble chart's x values. */
  categories?: ChartData;
  values: ChartData;
  /** A bubble chart's bubble sizes. */
  sizes?: ChartData;
  /** The colour the file gives the series: its fill, or its line for lines and scatter. */
  color?: string;
  /** Colours the file gives single points (a pie's slices), by point. */
  pointColors?: Record<number, string>;
  /** A line joins the points (lines and scatter). */
  line: boolean;
  /** Each point has a marker (lines and scatter). */
  markers: boolean;
  smooth: boolean;
  dashed: boolean;
  /** Line width in pixels, when the file sets one. */
  lineWidth?: number;
}

/** One kind of chart in the plot area; a combination chart has several. */
export interface ChartGroup {
  kind: ChartKind;
  /** Bars run sideways (Excel's bar chart) rather than up (its column chart). */
  horizontal: boolean;
  grouping: ChartGrouping;
  /** Each point gets its own colour, as a pie's slices do. */
  varyColors: boolean;
  /** Plotted against the secondary value axis. */
  secondary: boolean;
  series: ChartSeries[];
  /** Space between clusters of bars, as a percentage of a bar's width. */
  gapWidth: number;
  /** A doughnut's hole, as a percentage of its size. */
  holeSize?: number;
  /** Where a pie's first slice starts, in degrees clockwise from the top. */
  firstSliceAngle?: number;
  /** A radar's areas are filled, rather than drawn as lines. */
  filled?: boolean;
  /** A bubble chart's largest bubble, as a percentage of Excel's default size, and whether sizes give a bubble's area or its width. */
  bubbles?: { scale: number; area: boolean };
  /** A stock chart's lines from each category's highest value to its lowest, and its bars from its first series to its last (up and down colours). */
  stock?: { hiLow: boolean; upDown?: { up?: string; down?: string } };
}

export interface ChartAxis {
  title?: ChartData;
  min?: number;
  max?: number;
  majorUnit?: number;
  logBase?: number;
  /** Runs from maximum to minimum. */
  reversed: boolean;
  /** Deleted in Excel, so not drawn. */
  hidden: boolean;
  gridlines: boolean;
  /** Number format of its labels; missing when they take the format of the cells plotted. */
  format?: string;
}

export interface ChartSpec {
  /** The chart's title, or its only series' name, which Excel shows when the chart has no title of its own. */
  title?: ChartData;
  groups: ChartGroup[];
  /** The category axis (a scatter chart's x axis), the value axis, and a secondary value axis. */
  axes: { category?: ChartAxis; value?: ChartAxis; secondary?: ChartAxis };
  legend?: { position: 'top' | 'bottom' | 'left' | 'right'; deleted: number[] };
  /** How empty cells are plotted. */
  blanks: 'gap' | 'zero' | 'span';
  /** Cells in hidden rows and columns are left out, as Excel does unless told otherwise. */
  visibleOnly: boolean;
  /** The workbook theme's accent colours, which Excel gives series the file doesn't colour. */
  accents: string[];
  /** Why the editor can't draw it: the kind of chart, as the file names it, or "combination". */
  unsupported?: string;
}

/** What chart references are resolved against. */
export interface ChartContext {
  /** Theme colours in SpreadsheetML theme-index order (lt1, dk1, lt2, dk2, accents, links). */
  theme: string[];
  /** Lower-cased sheet name → sheet index. */
  sheetIndex: Map<string, number>;
  names: { name: string; formula: string; localSheet?: number }[];
}

const KINDS: Record<string, ChartKind> = {
  barChart: 'bar',
  bar3DChart: 'bar',
  lineChart: 'line',
  line3DChart: 'line',
  areaChart: 'area',
  area3DChart: 'area',
  pieChart: 'pie',
  pie3DChart: 'pie',
  doughnutChart: 'doughnut',
  scatterChart: 'scatter',
  radarChart: 'radar',
  bubbleChart: 'bubble',
  stockChart: 'stock',
  // Pie of pie and bar of pie: drawn as the one pie, without the second plot.
  ofPieChart: 'pie',
};

/** Points read from one chart part at most; the cache can claim any count. */
const MAX_POINTS = 10_000;
const EMU_PER_PX = 9525;

const val = (node: XmlNode | undefined) => attr(node, 'val');
const num = (node: XmlNode | undefined) => {
  const v = val(node);
  if (v === undefined || v.trim() === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
};
/** A CT_Boolean: present without a value means true. */
const flag = (node: XmlNode | undefined, fallback: boolean) => (node ? !['0', 'false'].includes(val(node) ?? '1') : fallback);

/** Theme colour names as DrawingML writes them → SpreadsheetML theme index. */
const SCHEME: Record<string, number> = {
  lt1: 0, bg1: 0, dk1: 1, tx1: 1, lt2: 2, bg2: 2, dk2: 3, tx2: 3,
  accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11,
};
const PRESET: Record<string, string> = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080' };

const hex2 = (n: number) => Math.round(Math.min(255, Math.max(0, n))).toString(16).padStart(2, '0').toUpperCase();

/** Lightness scaled and offset (lumMod, lumOff) in HSL, as DrawingML adjusts theme colours. */
function luminance(hex: string, mod: number, off: number): string {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  const d = max - min;
  const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
  let h = 0;
  if (d) h = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  const l2 = Math.min(1, Math.max(0, l * mod + off));
  const c = (1 - Math.abs(2 * l2 - 1)) * s;
  const x = c * (1 - Math.abs((h % 2) - 1));
  const m = l2 - c / 2;
  const sector = Math.floor((h + 6) % 6);
  const [r1, g1, b1] = [[c, x, 0], [x, c, 0], [0, c, x], [0, x, c], [x, 0, c], [c, 0, x]][sector];
  return hex2((r1 + m) * 255) + hex2((g1 + m) * 255) + hex2((b1 + m) * 255);
}

/** The colour a DrawingML colour element (srgbClr, schemeClr…) stands for, as "#RRGGBB" or rgba(). */
export function drawingColor(node: XmlNode | undefined, theme: string[]): string | undefined {
  if (!node) return undefined;
  let hex: string | undefined;
  if (node.name === 'srgbClr') hex = val(node);
  else if (node.name === 'schemeClr') hex = theme[SCHEME[val(node) ?? ''] ?? -1];
  else if (node.name === 'sysClr') hex = attr(node, 'lastClr') ?? (val(node) === 'window' ? 'FFFFFF' : '000000');
  else if (node.name === 'prstClr') hex = PRESET[val(node) ?? ''];
  if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return undefined;
  hex = hex.toUpperCase();
  const mod = (name: string) => num(child(node, name));
  const lumMod = mod('lumMod');
  const lumOff = mod('lumOff');
  if (lumMod !== undefined || lumOff !== undefined) hex = luminance(hex, (lumMod ?? 100_000) / 100_000, (lumOff ?? 0) / 100_000);
  // A tint keeps that much of the colour and adds white; a shade keeps that much and adds black.
  const tint = mod('tint');
  if (tint !== undefined) hex = tintColor(hex, 1 - tint / 100_000).slice(1);
  const shade = mod('shade');
  if (shade !== undefined) hex = tintColor(hex, shade / 100_000 - 1).slice(1);
  const alpha = mod('alpha');
  if (alpha !== undefined && alpha < 100_000) {
    const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex!.slice(i, i + 2), 16));
    return `rgba(${r}, ${g}, ${b}, ${Math.round(alpha / 1000) / 100})`;
  }
  return '#' + hex;
}

const COLOR_ELEMENTS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr']);

/** The colour of a solid fill, or undefined for no fill or a fill this doesn't read (gradients, patterns). */
function solidFill(node: XmlNode | undefined, theme: string[]): string | undefined {
  const fill = child(node, 'solidFill');
  return drawingColor(fill?.children.find((c) => COLOR_ELEMENTS.has(c.name)), theme);
}

/** A reference resolved to one range on one sheet: directly, or through a defined name. */
export function resolveRef(formula: string, ctx: ChartContext, depth = 0): ChartData['ref'] {
  const ref = parseA1Ref(formula.trim().replace(/^=/, ''), { allowTernary: false });
  if (!ref || (ref.context?.length ?? 0) > 1) return undefined;
  const sheetName = ref.context?.[0];
  const sheet = sheetName === undefined ? undefined : ctx.sheetIndex.get(sheetName.toLowerCase());
  if ('name' in ref) {
    if (depth) return undefined;
    const lower = ref.name.toLowerCase();
    const named =
      (sheet !== undefined ? ctx.names.find((n) => n.name.toLowerCase() === lower && n.localSheet === sheet) : undefined) ??
      ctx.names.find((n) => n.name.toLowerCase() === lower && n.localSheet === undefined);
    return named ? resolveRef(named.formula, ctx, depth + 1) : undefined;
  }
  // Charts always name the sheet; a range without one has nothing to be on.
  if (sheet === undefined) return undefined;
  const r = ref.range;
  return {
    sheet,
    rect: {
      r1: r.top ?? 0,
      c1: r.left ?? 0,
      r2: r.bottom ?? (r.top === null ? MAX_ROWS - 1 : r.top),
      c2: r.right ?? (r.left === null ? MAX_COLS - 1 : r.left),
    },
  };
}

/** Saved points: <c:pt idx="n"><c:v>…</c:v></c:pt>, laid out by index. */
function points(cache: XmlNode | undefined, numeric: boolean): (string | number | null)[] {
  if (!cache) return [];
  const pts = children(cache, 'pt');
  const declared = num(child(cache, 'ptCount'));
  const last = pts.reduce((m, p) => Math.max(m, Number(attr(p, 'idx') ?? -1)), -1);
  const count = Math.min(MAX_POINTS, declared ?? last + 1);
  const out: (string | number | null)[] = new Array(count).fill(null);
  for (const p of pts) {
    const i = Number(attr(p, 'idx'));
    if (!(i >= 0 && i < count)) continue;
    const text = child(p, 'v')?.text ?? '';
    if (numeric) {
      const n = Number(text);
      out[i] = text.trim() !== '' && Number.isFinite(n) ? n : null;
    } else {
      out[i] = text;
    }
  }
  return out;
}

/** A name, categories or values element: a reference with its saved values, or values written in. */
function readData(node: XmlNode | undefined, ctx: ChartContext): ChartData | undefined {
  if (!node) return undefined;
  const literal = child(node, 'v');
  if (literal) return { cache: [literal.text] };
  for (const c of node.children) {
    switch (c.name) {
      case 'numRef':
      case 'strRef':
      case 'multiLvlStrRef': {
        const formula = child(c, 'f')?.text.trim() || undefined;
        const numeric = c.name === 'numRef';
        const cache = child(c, numeric ? 'numCache' : c.name === 'strRef' ? 'strCache' : 'multiLvlStrCache');
        // Several levels of labels: the cache's first level is the one next to the axis.
        const level = c.name === 'multiLvlStrRef' ? child(cache, 'lvl') : cache;
        return {
          formula,
          // Only single-level references are read live; the rest keep Excel's labels.
          ref: formula && c.name !== 'multiLvlStrRef' ? resolveRef(formula, ctx) : undefined,
          cache: points(level, numeric),
          format: child(cache, 'formatCode')?.text || undefined,
        };
      }
      case 'numLit':
      case 'strLit':
        return { cache: points(c, c.name === 'numLit'), format: child(c, 'formatCode')?.text || undefined };
    }
  }
  return undefined;
}

/** A title's text: rich text runs, or a reference to a cell. */
function readTitle(title: XmlNode | undefined, ctx: ChartContext): ChartData | undefined {
  const tx = child(title, 'tx');
  if (!tx) return undefined;
  const rich = child(tx, 'rich');
  if (rich) {
    const text = children(rich, 'p')
      .map((p) => descendants(p, 't').map((t) => t.text).join(''))
      .join('\n')
      .trim();
    return text ? { cache: [text] } : undefined;
  }
  return readData(tx, ctx);
}

function readAxis(node: XmlNode | undefined, ctx: ChartContext): ChartAxis | undefined {
  if (!node) return undefined;
  const scaling = child(node, 'scaling');
  const fmt = child(node, 'numFmt');
  const linked = ['1', 'true'].includes(attr(fmt, 'sourceLinked') ?? '0');
  return {
    title: readTitle(child(node, 'title'), ctx),
    min: num(child(scaling, 'min')),
    max: num(child(scaling, 'max')),
    majorUnit: num(child(node, 'majorUnit')),
    logBase: num(child(scaling, 'logBase')),
    reversed: val(child(scaling, 'orientation')) === 'maxMin',
    hidden: flag(child(node, 'delete'), false),
    gridlines: !!child(node, 'majorGridlines'),
    format: fmt && !linked ? attr(fmt, 'formatCode') : undefined,
  };
}

function readSeries(ser: XmlNode, kind: ChartKind, group: XmlNode, ctx: ChartContext): ChartSeries {
  const spPr = child(ser, 'spPr');
  const ln = child(spPr, 'ln');
  const lineColor = solidFill(ln, ctx.theme);
  const fillColor = solidFill(spPr, ctx.theme);
  const lined = kind === 'line' || kind === 'scatter' || kind === 'radar' || kind === 'stock';
  const xy = kind === 'scatter' || kind === 'bubble';
  const symbol = val(child(child(ser, 'marker'), 'symbol'));
  const pointColors: Record<number, string> = {};
  for (const dPt of children(ser, 'dPt')) {
    const idx = num(child(dPt, 'idx'));
    const color = solidFill(child(dPt, 'spPr'), ctx.theme);
    if (idx !== undefined && color) pointColors[idx] = color;
  }
  const width = Number(attr(ln, 'w'));
  const dash = val(child(ln, 'prstDash'));
  return {
    name: readData(child(ser, 'tx'), ctx),
    categories: readData(child(ser, xy ? 'xVal' : 'cat'), ctx),
    values: readData(child(ser, xy ? 'yVal' : 'val'), ctx) ?? { cache: [] },
    ...(kind === 'bubble' && { sizes: readData(child(ser, 'bubbleSize'), ctx) }),
    color: lined ? (lineColor ?? fillColor) : (fillColor ?? lineColor),
    pointColors: Object.keys(pointColors).length ? pointColors : undefined,
    // Excel goes by each series' line and markers, whatever the scatter style says.
    line: lined && !child(ln, 'noFill'),
    markers:
      lined &&
      symbol !== 'none' &&
      (kind === 'scatter' || (kind === 'radar' ? val(child(group, 'radarStyle')) === 'marker' || !!symbol : kind === 'stock' ? !!symbol : flag(child(group, 'marker'), true))),
    smooth: lined && flag(child(ser, 'smooth'), false),
    dashed: dash !== undefined && dash !== 'solid',
    lineWidth: width > 0 ? Math.max(1, Math.round((width / EMU_PER_PX) * 10) / 10) : undefined,
  };
}

/**
 * Reads a chart part. Kinds of chart the editor doesn't draw, and
 * combinations it can't (a pie with anything else, bars both ways), are
 * marked `unsupported`, keeping what is known about them.
 */
export function parseChart(root: XmlNode, ctx: ChartContext): ChartSpec {
  const chart = child(root, 'chart');
  const plot = child(chart, 'plotArea');
  const axes = new Map<string, XmlNode>();
  for (const c of plot?.children ?? []) {
    if (/^(catAx|valAx|dateAx|serAx)$/.test(c.name)) {
      const id = val(child(c, 'axId'));
      if (id) axes.set(id, c);
    }
  }

  let unsupported: string | undefined;
  const groups: ChartGroup[] = [];
  let primaryValueAxis: string | undefined;
  let categoryAxis: XmlNode | undefined;
  let valueAxis: XmlNode | undefined;
  let secondaryAxis: XmlNode | undefined;
  for (const g of plot?.children ?? []) {
    if (!g.name.endsWith('Chart')) continue;
    const kind = KINDS[g.name];
    if (!kind) {
      unsupported ??= g.name;
      continue;
    }
    const ids = children(g, 'axId').map((a) => val(a) ?? '');
    const valueId = ids[1];
    const secondary = primaryValueAxis !== undefined && valueId !== undefined && valueId !== primaryValueAxis;
    if (!groups.length) {
      primaryValueAxis = valueId;
      categoryAxis = axes.get(ids[0]);
      valueAxis = axes.get(valueId);
    } else if (secondary) {
      secondaryAxis ??= axes.get(valueId);
    }
    const grouping = (val(child(g, 'grouping')) ?? (kind === 'bar' ? 'clustered' : 'standard')) as ChartGrouping;
    groups.push({
      kind,
      horizontal: kind === 'bar' && val(child(g, 'barDir')) === 'bar',
      grouping: kind === 'bar' && grouping === 'standard' ? 'clustered' : grouping,
      varyColors: flag(child(g, 'varyColors'), kind === 'pie' || kind === 'doughnut'),
      secondary,
      series: children(g, 'ser')
        .sort((a, b) => (num(child(a, 'order')) ?? 0) - (num(child(b, 'order')) ?? 0))
        .map((s) => readSeries(s, kind, g, ctx)),
      gapWidth: num(child(g, 'gapWidth')) ?? 150,
      holeSize: kind === 'doughnut' ? (num(child(g, 'holeSize')) ?? 50) : undefined,
      firstSliceAngle: kind === 'pie' || kind === 'doughnut' ? (num(child(g, 'firstSliceAng')) ?? 0) : undefined,
      ...(kind === 'radar' && { filled: val(child(g, 'radarStyle')) === 'filled' }),
      ...(kind === 'bubble' && { bubbles: { scale: num(child(g, 'bubbleScale')) ?? 100, area: val(child(g, 'sizeRepresents')) !== 'w' } }),
      ...(kind === 'stock' && {
        stock: {
          hiLow: !!child(g, 'hiLowLines'),
          ...(child(g, 'upDownBars') && {
            upDown: { up: solidFill(child(child(child(g, 'upDownBars'), 'upBars'), 'spPr'), ctx.theme), down: solidFill(child(child(child(g, 'upDownBars'), 'downBars'), 'spPr'), ctx.theme) },
          }),
        },
      }),
    });
  }

  if (!unsupported) {
    const kinds = new Set(groups.map((g) => g.kind));
    const cartesian = [...kinds].every((k) => k === 'bar' || k === 'line' || k === 'area');
    const directions = new Set(groups.filter((g) => g.kind === 'bar').map((g) => g.horizontal));
    if (!groups.length) unsupported = 'empty';
    else if (groups.length > 1 && (!cartesian || directions.size > 1)) unsupported = 'combination';
    else if (!groups.some((g) => g.series.length)) unsupported = 'empty';
  }

  const series = groups.flatMap((g) => g.series);
  const legend = child(chart, 'legend');
  const position = val(child(legend, 'legendPos')) ?? 'r';
  const blanks = val(child(chart, 'dispBlanksAs'));
  return {
    title:
      readTitle(child(chart, 'title'), ctx) ??
      (series.length === 1 && !flag(child(chart, 'autoTitleDeleted'), false) ? series[0].name : undefined),
    groups,
    axes: {
      category: readAxis(categoryAxis, ctx),
      value: readAxis(valueAxis, ctx),
      secondary: readAxis(secondaryAxis, ctx),
    },
    legend: legend
      ? {
          position: position === 't' ? 'top' : position === 'b' ? 'bottom' : position === 'l' ? 'left' : 'right',
          deleted: children(legend, 'legendEntry')
            .filter((e) => flag(child(e, 'delete'), false))
            .map((e) => num(child(e, 'idx')) ?? -1),
        }
      : undefined,
    blanks: blanks === 'gap' || blanks === 'span' ? blanks : 'zero',
    visibleOnly: flag(child(chart, 'plotVisOnly'), true),
    accents: ctx.theme.slice(4, 10).map((c) => '#' + c.toUpperCase()),
    unsupported,
  };
}
