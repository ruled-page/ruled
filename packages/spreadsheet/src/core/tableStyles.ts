import { tintColor } from './styles';
import type { BorderEdge, DiffStyle, Rect, TableElement, TableInfo, TableLook } from './types';

const WHITE = '#FFFFFF';
const BLACK = '#000000';
const thin = (color: string): BorderEdge => ({ width: 1, pattern: 'solid', color });
const medium = (color: string): BorderEdge => ({ width: 2, pattern: 'solid', color });
const double = (color: string): BorderEdge => ({ width: 3, pattern: 'double', color });
const box = (e: BorderEdge) => ({ top: e, bottom: e, left: e, right: e });

/**
 * Excel's built-in table styles, which files name but don't define. Each is
 * drawn from its family's pattern (TableStyleLight1–21, Medium1–28,
 * Dark1–11) in its theme colour. This is close to Excel, not exact.
 */
export function builtinTableLook(name: string, theme: string[]): TableLook | undefined {
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/i.exec(name);
  if (!m) return undefined;
  const family = m[1].toLowerCase();
  const n = Number(m[2]);
  // Each band of seven uses dark 1 then accents 1–6; Dark 8–11 pair accents up.
  const slot = family === 'dark' && n > 7 ? [0, 1, 3, 5][n - 8] : (n - 1) % 7;
  if (slot === undefined) return undefined;
  const band = Math.floor((n - 1) / 7);
  const plain = slot === 0;
  const a = '#' + (plain ? theme[1] : theme[3 + slot]);
  const tint = (t: number) => tintColor(a, plain && t > 0 ? Math.min(0.95, t + 0.05) : t);
  const stripe = tint(0.8);
  const bold: DiffStyle = { bold: true };
  const header: DiffStyle = { fill: a, color: WHITE, bold: true };
  let e: TableLook['elements'];

  if (family === 'light' && band === 0) {
    e = {
      wholeTable: { color: plain ? undefined : tint(-0.25), border: { top: thin(a), bottom: thin(a) } },
      headerRow: { bold: true, border: { bottom: thin(a) } },
      totalRow: { bold: true, border: { top: thin(a) } },
      firstRowStripe: { fill: stripe },
      firstColumnStripe: { fill: stripe },
      firstColumn: bold,
      lastColumn: bold,
    };
  } else if (family === 'light' && band === 1) {
    e = {
      wholeTable: { border: box(thin(a)) },
      headerRow: header,
      totalRow: { bold: true, border: { top: double(a) } },
      firstRowStripe: { border: { top: thin(a), bottom: thin(a) } },
      firstColumnStripe: { border: { left: thin(a), right: thin(a) } },
      firstColumn: bold,
      lastColumn: bold,
    };
  } else if (family === 'light') {
    e = {
      wholeTable: { border: { ...box(thin(a)), horizontal: thin(a), vertical: thin(a) } },
      headerRow: { bold: true, border: { bottom: thin(a) } },
      totalRow: { bold: true, border: { top: double(a) } },
      firstRowStripe: { fill: stripe },
      firstColumnStripe: { fill: stripe },
      firstColumn: bold,
      lastColumn: bold,
    };
  } else if (family === 'medium' && band === 0) {
    e = {
      wholeTable: { border: { ...box(thin(tint(0.4))), horizontal: thin(tint(0.4)) } },
      headerRow: header,
      totalRow: { bold: true, border: { top: double(a) } },
      firstRowStripe: { fill: stripe },
      firstColumnStripe: { fill: stripe },
      firstColumn: bold,
      lastColumn: bold,
    };
  } else if (family === 'medium' && band === 1) {
    e = {
      wholeTable: { fill: stripe, color: BLACK, border: { horizontal: thin(WHITE), vertical: thin(WHITE) } },
      headerRow: { ...header, border: { bottom: medium(WHITE) } },
      totalRow: { ...header, border: { top: medium(WHITE) } },
      firstColumn: header,
      lastColumn: header,
      firstRowStripe: { fill: tint(0.6) },
      firstColumnStripe: { fill: tint(0.6) },
    };
  } else if (family === 'medium' && band === 2) {
    e = {
      wholeTable: { border: { top: thin(BLACK), bottom: thin(BLACK) } },
      headerRow: { ...header, border: { bottom: medium(BLACK) } },
      totalRow: { bold: true, border: { top: double(BLACK) } },
      firstRowStripe: { fill: '#D9D9D9' },
      firstColumnStripe: { fill: '#D9D9D9' },
      firstColumn: header,
      lastColumn: header,
    };
  } else if (family === 'medium') {
    e = {
      wholeTable: { fill: stripe, border: { ...box(thin(tint(0.4))), horizontal: thin(tint(0.4)), vertical: thin(tint(0.4)) } },
      headerRow: bold,
      totalRow: { bold: true, border: { top: double(a) } },
      firstRowStripe: { fill: tint(0.6) },
      firstColumnStripe: { fill: tint(0.6) },
      firstColumn: bold,
      lastColumn: bold,
    };
  } else if (band === 0) {
    const body = plain ? tintColor(BLACK, 0.45) : tint(-0.25);
    const deep = plain ? tintColor(BLACK, 0.25) : tint(-0.5);
    e = {
      wholeTable: { fill: body, color: WHITE },
      headerRow: { fill: BLACK, color: WHITE, bold: true, border: { bottom: medium(WHITE) } },
      totalRow: { fill: deep, color: WHITE, bold: true, border: { top: double(WHITE) } },
      firstColumn: { fill: deep, color: WHITE, bold: true, border: { right: medium(WHITE) } },
      lastColumn: { fill: deep, color: WHITE, bold: true, border: { left: medium(WHITE) } },
      firstRowStripe: { fill: deep },
      firstColumnStripe: { fill: deep },
    };
  } else {
    e = {
      wholeTable: { fill: stripe },
      headerRow: { fill: BLACK, color: WHITE, bold: true },
      totalRow: { bold: true, border: { top: double(BLACK) } },
      firstRowStripe: { fill: tint(0.6) },
      firstColumnStripe: { fill: tint(0.6) },
      firstColumn: bold,
      lastColumn: bold,
    };
  }
  return { elements: e, rowStripe: 1, colStripe: 1 };
}

const formatCache = new WeakMap<TableInfo, Map<string, DiffStyle | undefined>>();

/**
 * How a table's style formats one of its cells. Elements are layered as
 * Excel layers them, from the whole table up to the corner cells, and each
 * element's borders apply to the edge of the area it covers.
 */
export function tableCellFormat(t: TableInfo, row: number, col: number): DiffStyle | undefined {
  const look = t.look;
  const { r1, c1, r2, c2 } = t.rect;
  if (!look || row < r1 || row > r2 || col < c1 || col > c2) return undefined;
  const header = t.headerRows > 0 && row === r1;
  const total = t.totalsRows > 0 && row === r2;
  const top = r1 + (t.headerRows > 0 ? 1 : 0);
  const bottom = r2 - (t.totalsRows > 0 ? 1 : 0);
  const layers: [TableElement, Rect][] = [['wholeTable', t.rect]];
  if (t.showColumnStripes && !header && !total) {
    const k = Math.floor((col - c1) / look.colStripe);
    const from = c1 + k * look.colStripe;
    layers.push([k % 2 ? 'secondColumnStripe' : 'firstColumnStripe', { r1: top, r2: bottom, c1: from, c2: Math.min(c2, from + look.colStripe - 1) }]);
  }
  if (t.showRowStripes && !header && !total) {
    const k = Math.floor((row - top) / look.rowStripe);
    const from = top + k * look.rowStripe;
    layers.push([k % 2 ? 'secondRowStripe' : 'firstRowStripe', { r1: from, r2: Math.min(bottom, from + look.rowStripe - 1), c1, c2 }]);
  }
  if (t.showLastColumn && col === c2) layers.push(['lastColumn', { r1, r2, c1: c2, c2 }]);
  if (t.showFirstColumn && col === c1) layers.push(['firstColumn', { r1, r2, c1, c2: c1 }]);
  if (header) layers.push(['headerRow', { r1, r2: r1, c1, c2 }]);
  if (total) layers.push(['totalRow', { r1: r2, r2, c1, c2 }]);
  if (header && col === c1) layers.push(['firstHeaderCell', { r1, r2: r1, c1, c2: c1 }]);
  if (header && col === c2) layers.push(['lastHeaderCell', { r1, r2: r1, c1: c2, c2 }]);
  if (total && col === c1) layers.push(['firstTotalCell', { r1: r2, r2, c1, c2: c1 }]);
  if (total && col === c2) layers.push(['lastTotalCell', { r1: r2, r2, c1: c2, c2 }]);

  // Cells with the same layers and the same position within each area share a format.
  const used = layers.filter(([el]) => look.elements[el]);
  const key = used.map(([el, r]) => `${el}${+(row === r.r1)}${+(row === r.r2)}${+(col === r.c1)}${+(col === r.c2)}`).join('|');
  let cache = formatCache.get(t);
  if (!cache) formatCache.set(t, (cache = new Map()));
  if (cache.has(key)) return cache.get(key);

  let out: DiffStyle | undefined;
  for (const [el, r] of used) {
    const s = look.elements[el]!;
    out ??= {};
    for (const k of ['fill', 'color', 'bold', 'italic', 'underline', 'strike'] as const) {
      if (s[k] !== undefined) (out as Record<string, unknown>)[k] = s[k];
    }
    const b = s.border;
    if (!b) continue;
    const edges = {
      top: row === r.r1 ? b.top : b.horizontal,
      bottom: row === r.r2 ? b.bottom : b.horizontal,
      left: col === r.c1 ? b.left : b.vertical,
      right: col === r.c2 ? b.right : b.vertical,
    };
    for (const [side, edge] of Object.entries(edges)) {
      if (edge) out.border = { ...out.border, [side]: edge };
    }
  }
  cache.set(key, out);
  return out;
}
