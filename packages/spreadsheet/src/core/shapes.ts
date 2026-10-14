import { drawingColor } from './charts';
import type { ShapeSpec, ShapeText } from './types';
import { attr, child, children, type XmlNode } from './xml';

/** The preset shapes the editor draws; others show as placeholders. */
export const DRAWN_SHAPES = new Set([
  'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle', 'diamond', 'parallelogram', 'trapezoid', 'pentagon', 'hexagon', 'octagon',
  'homePlate', 'chevron', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'star5', 'plus',
  'line', 'straightConnector1', 'flowChartProcess', 'flowChartAlternateProcess', 'flowChartDecision', 'flowChartTerminator',
]);

const EMU_PER_PX = 9525;
const COLOR_ELEMENTS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr']);
/** Line widths of Office's theme line styles, by a style reference's index, in pixels. */
const THEME_LINE_PX = [0, 6350 / EMU_PER_PX, 12700 / EMU_PER_PX, 19050 / EMU_PER_PX];

const colorIn = (node: XmlNode | undefined, theme: string[], placeholder?: string) => {
  const el = node?.children.find((c) => COLOR_ELEMENTS.has(c.name));
  // phClr in a theme style stands for the colour the reference gives.
  if (el?.name === 'schemeClr' && attr(el, 'val') === 'phClr') return placeholder;
  return drawingColor(el, theme);
};

/**
 * A fill as drawn: a solid fill's colour, a gradient's first stop or a
 * pattern's foreground, null for none, undefined where the element sets
 * none (so a style reference's applies).
 */
function fillOf(node: XmlNode | undefined, theme: string[]): string | null | undefined {
  if (!node) return undefined;
  if (child(node, 'noFill')) return null;
  const solid = child(node, 'solidFill');
  if (solid) return colorIn(solid, theme) ?? null;
  const grad = child(node, 'gradFill');
  if (grad) return colorIn(children(child(grad, 'gsLst'), 'gs')[0], theme) ?? null;
  const patt = child(node, 'pattFill');
  if (patt) return colorIn(child(patt, 'fgClr'), theme) ?? null;
  return undefined;
}

/** A shape (<sp>) or connector (<cxnSp>) the editor draws, read with the workbook's theme colours; undefined for kinds it doesn't. */
export function parseShape(sp: XmlNode, theme: string[]): ShapeSpec | undefined {
  const spPr = child(sp, 'spPr');
  const geom = child(spPr, 'prstGeom');
  const geometry = attr(geom, 'prst');
  if (!geometry || !DRAWN_SHAPES.has(geometry)) return undefined;
  const style = child(sp, 'style');
  const ref = (name: string) => {
    const r = child(style, name);
    return r && { idx: attr(r, 'idx') ?? '0', color: colorIn(r, theme) };
  };
  const out: ShapeSpec = { geometry };

  const adjust: Record<string, number> = {};
  for (const gd of children(child(geom, 'avLst'), 'gd')) {
    const m = /^val\s+(-?\d+)$/.exec(attr(gd, 'fmla') ?? '');
    if (m) adjust[attr(gd, 'name') ?? ''] = Number(m[1]);
  }
  if (Object.keys(adjust).length) out.adjust = adjust;

  const fill = fillOf(spPr, theme);
  const fillRef = ref('fillRef');
  const styled = fill === undefined && fillRef && fillRef.idx !== '0' ? fillRef.color : undefined;
  if (fill ?? styled) out.fill = (fill ?? styled)!;

  const ln = child(spPr, 'ln');
  const lnRef = ref('lnRef');
  const lineFill = fillOf(ln, theme);
  const lineColor = lineFill === undefined ? (lnRef && lnRef.idx !== '0' ? lnRef.color : undefined) : lineFill;
  if (lineColor) {
    const w = attr(ln, 'w');
    const width = w !== undefined ? Number(w) / EMU_PER_PX : (THEME_LINE_PX[Number(lnRef?.idx ?? 0)] ?? 1);
    const dash = attr(child(ln, 'prstDash'), 'val');
    out.line = { color: lineColor, width: Math.max(0.5, width), ...(dash && dash !== 'solid' && { dash }) };
    const end = (name: string) => {
      const type = attr(child(ln, name), 'type');
      return type && type !== 'none' ? (type === 'oval' ? 'oval' : 'arrow') : undefined;
    };
    const head = end('headEnd');
    const tail = end('tailEnd');
    if (head) out.line.head = head;
    if (tail) out.line.tail = tail;
  }

  const xfrm = child(spPr, 'xfrm');
  const rot = Number(attr(xfrm, 'rot') ?? 0) / 60_000;
  if (rot) out.rotate = rot;
  if (attr(xfrm, 'flipH') === '1') out.flipH = true;
  if (attr(xfrm, 'flipV') === '1') out.flipV = true;

  const text = textOf(child(sp, 'txBody'), theme, ref('fontRef')?.color);
  if (text) out.text = text;
  return out;
}

/** The text in a shape, as drawn: its paragraphs and their runs, where it sits, and how far in from the edges. */
function textOf(body: XmlNode | undefined, theme: string[], styleColor: string | undefined): ShapeText | undefined {
  if (!body) return undefined;
  const pr = child(body, 'bodyPr');
  const inset = (name: string, dflt: number) => Number(attr(pr, name) ?? dflt) / EMU_PER_PX;
  const defaults = child(child(child(body, 'lstStyle'), 'lvl1pPr'), 'defRPr');
  const paragraphs: ShapeText['paragraphs'] = [];
  for (const p of children(body, 'p')) {
    const pPr = child(p, 'pPr');
    const pDefaults = child(pPr, 'defRPr');
    const runs: ShapeText['paragraphs'][number]['runs'] = [];
    for (const r of p.children) {
      if (r.name === 'br') {
        runs.push({ text: '\n' });
        continue;
      }
      if (r.name !== 'r' && r.name !== 'fld') continue;
      const rPr = child(r, 'rPr');
      const prop = (name: string) => attr(rPr, name) ?? attr(pDefaults, name) ?? attr(defaults, name);
      const sz = prop('sz');
      const typeface = attr(child(rPr, 'latin') ?? child(pDefaults, 'latin') ?? child(defaults, 'latin'), 'typeface');
      // Text with no colour of its own or from the shape's style is the theme's dark text colour.
      const color = fillOf(rPr, theme) ?? fillOf(pDefaults, theme) ?? fillOf(defaults, theme) ?? styleColor ?? (theme[1] ? `#${theme[1]}` : '#000000');
      runs.push({
        text: child(r, 't')?.text ?? '',
        ...(sz && { size: (Number(sz) / 100) * (4 / 3) }),
        ...(prop('b') === '1' && { bold: true }),
        ...(prop('i') === '1' && { italic: true }),
        ...(prop('u') && prop('u') !== 'none' && { underline: true }),
        ...(color && { color }),
        // A theme font (+mn-lt, +mj-lt) is drawn in the workbook's own.
        ...(typeface && !typeface.startsWith('+') && { font: typeface }),
      });
    }
    const end = child(p, 'endParaRPr');
    const algn = attr(pPr, 'algn');
    paragraphs.push({
      align: algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'justify' : 'left',
      runs,
      ...(!runs.length && attr(end, 'sz') && { size: (Number(attr(end, 'sz')) / 100) * (4 / 3) }),
    });
  }
  if (!paragraphs.some((p) => p.runs.some((r) => r.text.trim()))) return undefined;
  const anchor = attr(pr, 'anchor');
  const vert = attr(pr, 'vert');
  return {
    paragraphs,
    anchor: anchor === 'ctr' ? 'middle' : anchor === 'b' ? 'bottom' : 'top',
    insets: [inset('lIns', 91440), inset('tIns', 45720), inset('rIns', 91440), inset('bIns', 45720)],
    ...(attr(pr, 'wrap') === 'none' && { nowrap: true }),
    ...(vert && vert !== 'horz' && { vertical: true }),
  };
}
