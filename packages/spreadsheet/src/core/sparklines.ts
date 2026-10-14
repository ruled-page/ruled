import { parseRange } from '@ruled/formula/refs';
import type { Rect, SparkCell, SparklineGroup } from './types';
import { attr, child, children, descendants, type XmlNode } from './xml';

/** Excel 2010's sparkline groups, from the sheet's extensions. `color` resolves a colour element, theme colours included. */
export function parseSparklines(meta: XmlNode, color: (node: XmlNode | undefined) => string | undefined): SparklineGroup[] {
  const groups: SparklineGroup[] = [];
  for (const g of descendants(meta, 'sparklineGroup')) {
    const flag = (name: string) => /^(1|true)$/i.test(attr(g, name) ?? '');
    const num = (name: string) => (attr(g, name) === undefined ? undefined : Number(attr(g, name)));
    const axis = (name: string) => (attr(g, name) ?? 'individual') as SparklineGroup['minAxis'];
    const type = attr(g, 'type');
    const lines: SparklineGroup['lines'] = [];
    for (const s of children(child(g, 'sparklines'), 'sparkline')) {
      const at = parseRange(child(s, 'sqref')?.text.trim() ?? '');
      const source = child(s, 'f')?.text.trim();
      if (at && source) lines.push({ row: at.r1, col: at.c1, source });
    }
    if (!lines.length) continue;
    groups.push({
      type: type === 'column' || type === 'stacked' ? type : 'line',
      colors: {
        series: color(child(g, 'colorSeries')) ?? '#376092',
        negative: color(child(g, 'colorNegative')) ?? '#D00000',
        axis: color(child(g, 'colorAxis')) ?? '#000000',
        markers: color(child(g, 'colorMarkers')) ?? '#D00000',
        first: color(child(g, 'colorFirst')) ?? '#D00000',
        last: color(child(g, 'colorLast')) ?? '#D00000',
        high: color(child(g, 'colorHigh')) ?? '#D00000',
        low: color(child(g, 'colorLow')) ?? '#D00000',
      },
      show: { markers: flag('markers'), high: flag('high'), low: flag('low'), first: flag('first'), last: flag('last'), negative: flag('negative'), axis: flag('displayXAxis') },
      lineWeight: num('lineWeight') ?? 0.75,
      empty: (attr(g, 'displayEmptyCellsAs') ?? 'zero') as SparklineGroup['empty'],
      minAxis: axis('minAxisType'),
      maxAxis: axis('maxAxisType'),
      manualMin: num('manualMin'),
      manualMax: num('manualMax'),
      displayHidden: flag('displayHidden'),
      rightToLeft: flag('rightToLeft'),
      lines,
    });
  }
  return groups;
}

/** A sparkline's source: its sheet's name (undefined for its own) and range. */
export function sparklineSource(source: string): { sheet?: string; rect: Rect } | undefined {
  const bang = source.lastIndexOf('!');
  const rect = parseRange((bang < 0 ? source : source.slice(bang + 1)).replace(/\$/g, ''));
  if (!rect) return undefined;
  if (bang < 0) return { rect };
  const name = source.slice(0, bang);
  return { sheet: /^'.*'$/.test(name) ? name.slice(1, -1).replace(/''/g, "'") : name, rect };
}

/**
 * What each of a sheet's sparklines draws: its points, from the values its
 * source holds now (null where a cell is empty or isn't a number), and the
 * range they are drawn against, its own or its group's.
 */
export function evaluateSparklines(
  groups: SparklineGroup[],
  points: (source: string, displayHidden: boolean) => (number | null)[] | undefined,
): SparkCell[] {
  const out: SparkCell[] = [];
  groups.forEach((g, group) => {
    const cells: SparkCell[] = [];
    for (const line of g.lines) {
      const values = points(line.source, g.displayHidden);
      if (!values) continue;
      const shown = (g.empty === 'zero' ? values.map((v) => v ?? 0) : values).filter((v): v is number => v !== null);
      cells.push({ row: line.row, col: line.col, group, values, min: Math.min(...shown), max: Math.max(...shown) });
    }
    const low = Math.min(...cells.map((c) => c.min));
    const high = Math.max(...cells.map((c) => c.max));
    for (const c of cells) {
      if (g.minAxis === 'group') c.min = low;
      else if (g.minAxis === 'custom' && g.manualMin !== undefined) c.min = g.manualMin;
      if (g.maxAxis === 'group') c.max = high;
      else if (g.maxAxis === 'custom' && g.manualMax !== undefined) c.max = g.manualMax;
    }
    out.push(...cells);
  });
  return out;
}
