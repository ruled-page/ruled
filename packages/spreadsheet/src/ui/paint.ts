import type { BorderEdge, DisplayStyle } from '../core/types';
import { CELL_PADDING, textTop } from '../state/overflow';

/**
 * Draws a cell's left or top edge, starting on the boundary pixel (where the
 * grid line sits) and growing into the cell for wider borders. The context is
 * restored afterwards, so what is drawn next keeps its own fill style. `auto`
 * is the colour for a border in Excel's automatic colour.
 */
export function drawEdge(ctx: CanvasRenderingContext2D, side: 'left' | 'top', x: number, y: number, length: number, edge: BorderEdge, auto: string): void {
  ctx.save();
  try {
    paintEdge(ctx, side, x, y, length, edge, auto);
  } finally {
    ctx.restore();
  }
}

function paintEdge(ctx: CanvasRenderingContext2D, side: 'left' | 'top', x: number, y: number, length: number, edge: BorderEdge, auto: string): void {
  ctx.fillStyle = edge.color ?? auto;
  const line = (offset: number, thickness: number, from = 0, len = length) => {
    if (side === 'left') ctx.fillRect(x + offset, y + from, thickness, len);
    else ctx.fillRect(x + from, y + offset, len, thickness);
  };
  if (edge.pattern === 'solid') return line(0, edge.width);
  if (edge.pattern === 'double') {
    line(0, 1);
    line(2, 1);
    return;
  }
  // Each style's run of dashes and gaps, in multiples of the line's width; a hair's dots are as fine as can be.
  const w = edge.width;
  const dash = 3 * w;
  const gap = w + 1;
  const runs: Record<Exclude<BorderEdge['pattern'], 'solid' | 'double'>, number[]> = {
    hair: [1, 1],
    dotted: [w, w],
    dashed: [dash, gap],
    dashDot: [dash, gap, w, gap],
    dashDotDot: [dash, gap, w, gap, w, gap],
    // Excel slants its dashes; upright here.
    slantDashDot: [dash, gap, w, gap],
  };
  const run = runs[edge.pattern];
  for (let p = 0, i = 0; p < length; p += run[i], i = (i + 1) % run.length) {
    if (i % 2 === 0) line(0, edge.pattern === 'hair' ? 1 : w, p, Math.min(run[i], length - p));
  }
}

/** A conditional-format data bar, behind the cell's text. */
export function drawDataBar(
  ctx: CanvasRenderingContext2D,
  rect: { x: number; y: number; width: number; height: number },
  bar: { from: number; to: number; color: string; gradient: boolean },
): void {
  const left = rect.x + 2;
  const width = rect.width - 4;
  const x = left + width * bar.from;
  const w = width * (bar.to - bar.from);
  if (w <= 0) return;
  const y = rect.y + 2;
  const h = rect.height - 3;
  ctx.save();
  if (bar.gradient) {
    const g = ctx.createLinearGradient(x, 0, x + w, 0);
    g.addColorStop(0, bar.color);
    g.addColorStop(1, '#FFFFFF');
    ctx.fillStyle = g;
  } else {
    ctx.fillStyle = bar.color;
  }
  ctx.fillRect(x, y, w, h);
  ctx.restore();
}

const RED = '#E0301E';
const AMBER = '#F5B800';
const GREEN = '#1E9E4A';
const GREY = '#8A8A8A';

/** Colour for icon `index` (0 = lowest) of an n-icon traffic-light style set. */
function ramp(index: number, n: number): string {
  if (index === 0) return RED;
  if (index === n - 1) return GREEN;
  return AMBER;
}

/**
 * Draws one icon of an Excel icon set in a size×size box at (x, y). Index 0
 * is the icon for the lowest values. Shapes are simplified versions of
 * Excel's, in the same colours.
 */
export function drawIcon(ctx: CanvasRenderingContext2D, set: string, index: number, x: number, y: number, size = 14): void {
  const n = Number(set[0]) || 3;
  const cx = x + size / 2;
  const cy = y + size / 2;
  const r = size / 2 - 1;
  ctx.save();
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  const circle = (fill: string) => {
    ctx.beginPath();
    ctx.arc(cx, cy, r, 0, Math.PI * 2);
    ctx.fillStyle = fill;
    ctx.fill();
  };
  const name = set.slice(1);
  if (name.startsWith('Arrows')) {
    // Down, down-right, right, up-right, up, spread over the set.
    const angles = n === 3 ? [90, 0, -90] : n === 4 ? [90, 45, -45, -90] : [90, 45, 0, -45, -90];
    const color = name === 'ArrowsGray' ? GREY : n === 3 ? ramp(index, n) : index === 0 ? RED : index === n - 1 ? GREEN : AMBER;
    ctx.translate(cx, cy);
    ctx.rotate(((angles[index] ?? 0) * Math.PI) / 180);
    ctx.strokeStyle = color;
    ctx.fillStyle = color;
    ctx.lineWidth = 2.4;
    ctx.beginPath();
    ctx.moveTo(-r + 1, 0);
    ctx.lineTo(r - 3, 0);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(r, 0);
    ctx.lineTo(r - 5, -4.5);
    ctx.lineTo(r - 5, 4.5);
    ctx.closePath();
    ctx.fill();
  } else if (name.startsWith('TrafficLights') || name === 'RedToBlack') {
    const colors = name === 'RedToBlack' ? ['#1A1A1A', '#7F7F7F', '#F4A6A6', RED] : n === 4 ? ['#1A1A1A', RED, AMBER, GREEN] : [RED, AMBER, GREEN];
    if (set === '3TrafficLights2') {
      ctx.fillStyle = '#3A3A3A';
      ctx.fillRect(x, y, size, size);
    }
    circle(colors[index] ?? GREY);
  } else if (name === 'Signs') {
    ctx.fillStyle = ramp(index, 3);
    ctx.beginPath();
    if (index === 0) {
      ctx.moveTo(cx, y);
      ctx.lineTo(x + size, cy);
      ctx.lineTo(cx, y + size);
      ctx.lineTo(x, cy);
    } else if (index === 1) {
      ctx.moveTo(cx, y + 1);
      ctx.lineTo(x + size, y + size - 1);
      ctx.lineTo(x, y + size - 1);
    } else {
      ctx.arc(cx, cy, r, 0, Math.PI * 2);
    }
    ctx.closePath();
    ctx.fill();
  } else if (name.startsWith('Symbols')) {
    const color = ramp(index, 3);
    const ringed = name === 'Symbols';
    if (ringed) circle(color);
    ctx.strokeStyle = ringed ? '#FFFFFF' : color;
    ctx.lineWidth = 2;
    ctx.beginPath();
    if (index === 0) {
      ctx.moveTo(cx - 3, cy - 3);
      ctx.lineTo(cx + 3, cy + 3);
      ctx.moveTo(cx + 3, cy - 3);
      ctx.lineTo(cx - 3, cy + 3);
    } else if (index === 1) {
      ctx.moveTo(cx, cy - 4);
      ctx.lineTo(cx, cy + 1);
      ctx.moveTo(cx, cy + 3.5);
      ctx.lineTo(cx, cy + 3.6);
    } else {
      ctx.moveTo(cx - 3.5, cy);
      ctx.lineTo(cx - 1, cy + 3);
      ctx.lineTo(cx + 4, cy - 3);
    }
    ctx.stroke();
  } else if (name === 'Flags') {
    ctx.strokeStyle = '#3A3A3A';
    ctx.lineWidth = 1.2;
    ctx.beginPath();
    ctx.moveTo(x + 3, y + 1);
    ctx.lineTo(x + 3, y + size);
    ctx.stroke();
    ctx.fillStyle = ramp(index, 3);
    ctx.beginPath();
    ctx.moveTo(x + 3.5, y + 1);
    ctx.lineTo(x + size - 1, y + 4);
    ctx.lineTo(x + 3.5, y + 8);
    ctx.closePath();
    ctx.fill();
  } else if (name === 'Stars') {
    ctx.beginPath();
    for (let i = 0; i < 10; i++) {
      const a = -Math.PI / 2 + (i * Math.PI) / 5;
      const rr = i % 2 ? r * 0.45 : r;
      ctx.lineTo(cx + rr * Math.cos(a), cy + rr * Math.sin(a));
    }
    ctx.closePath();
    ctx.strokeStyle = '#E2A700';
    ctx.lineWidth = 1;
    ctx.stroke();
    if (index > 0) {
      ctx.save();
      ctx.clip();
      ctx.fillStyle = '#F5C400';
      ctx.fillRect(x, y, index === 1 ? size / 2 : size, size);
      ctx.restore();
    }
  } else if (name === 'Triangles') {
    ctx.fillStyle = ramp(index, 3);
    ctx.beginPath();
    if (index === 1) {
      ctx.fillRect(x + 2, cy - 1.5, size - 4, 3);
    } else {
      const up = index === 2;
      ctx.moveTo(cx, up ? y + 2 : y + size - 2);
      ctx.lineTo(x + size - 1, up ? y + size - 3 : y + 3);
      ctx.lineTo(x + 1, up ? y + size - 3 : y + 3);
      ctx.closePath();
      ctx.fill();
    }
  } else if (name === 'Rating') {
    // Signal bars, one more filled for each step up.
    const bars = n - 1;
    const bw = (size - 2) / bars - 1;
    for (let i = 0; i < bars; i++) {
      const bh = ((i + 1) / bars) * (size - 2);
      ctx.fillStyle = i < index ? '#2F6EBA' : '#D0D7E2';
      ctx.fillRect(x + 1 + i * (bw + 1), y + size - 1 - bh, bw, bh);
    }
  } else if (name === 'Quarters') {
    circle('#FFFFFF');
    ctx.strokeStyle = '#3A3A3A';
    ctx.lineWidth = 1;
    ctx.stroke();
    if (index > 0) {
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, -Math.PI / 2, -Math.PI / 2 + (index / 4) * Math.PI * 2);
      ctx.closePath();
      ctx.fillStyle = '#3A3A3A';
      ctx.fill();
    }
  } else if (name === 'Boxes') {
    const cell = (size - 3) / 2;
    for (let i = 0; i < 4; i++) {
      ctx.fillStyle = i < index ? '#2F6EBA' : '#D0D7E2';
      ctx.fillRect(x + 1 + (i % 2) * (cell + 1), y + 1 + Math.floor(i / 2) * (cell + 1), cell, cell);
    }
  } else {
    circle(ramp(index, n));
  }
  ctx.restore();
}

/** Where a heading's filter button sits in its cell: a square at the right, as in Excel. */
export function filterButtonBox(cell: { x: number; y: number; width: number; height: number }): { x: number; y: number; size: number } {
  const size = Math.max(8, Math.min(16, cell.height - 4, cell.width - 4));
  return { x: cell.x + cell.width - size - 2, y: cell.y + Math.round((cell.height - size) / 2), size };
}

/**
 * A heading's filter button: a framed arrow, or for a column that is
 * filtered, a funnel, as Excel draws them.
 */
export function drawFilterButton(
  ctx: CanvasRenderingContext2D,
  cell: { x: number; y: number; width: number; height: number },
  filtered: boolean,
  colors: { bg: string; line: string; arrow: string; filtered: string },
): void {
  const b = filterButtonBox(cell);
  ctx.save();
  try {
    ctx.fillStyle = colors.line;
    ctx.fillRect(b.x, b.y, b.size, b.size);
    ctx.fillStyle = colors.bg;
    ctx.fillRect(b.x + 1, b.y + 1, b.size - 2, b.size - 2);
    const cx = b.x + b.size / 2;
    const cy = b.y + b.size / 2;
    ctx.beginPath();
    if (filtered) {
      ctx.moveTo(cx - 4, cy - 4);
      ctx.lineTo(cx + 4, cy - 4);
      ctx.lineTo(cx + 1, cy);
      ctx.lineTo(cx + 1, cy + 4);
      ctx.lineTo(cx - 1, cy + 3);
      ctx.lineTo(cx - 1, cy);
      ctx.closePath();
      ctx.fillStyle = colors.filtered;
      ctx.fill();
    } else {
      ctx.moveTo(cx - 3.5, cy - 1.5);
      ctx.lineTo(cx, cy + 2);
      ctx.lineTo(cx + 3.5, cy - 1.5);
      ctx.strokeStyle = colors.arrow;
      ctx.lineWidth = 1.5;
      ctx.stroke();
    }
  } finally {
    ctx.restore();
  }
}

/** Space between a checkbox and its cell's top or bottom edge, when aligned there. */
const CHECKBOX_PAD = 3;

/**
 * Where a cell's checkbox sits: as big as its font, placed by the cell's
 * alignment, centred across unless set otherwise, as Excel centres TRUE and
 * FALSE, and at the bottom unless set.
 */
export function checkboxBox(
  cell: { x: number; y: number; width: number; height: number },
  style: Pick<DisplayStyle, 'fontSizePx' | 'hAlign' | 'vAlign'>,
): { x: number; y: number; size: number } {
  const size = Math.round(Math.max(8, Math.min(style.fontSizePx * 0.9, cell.height - 4, cell.width - 4)));
  const x =
    style.hAlign === 'left' ? cell.x + CELL_PADDING : style.hAlign === 'right' ? cell.x + cell.width - CELL_PADDING - size : cell.x + Math.round((cell.width - size) / 2);
  return { x: Math.round(x), y: cell.y + Math.round(textTop(style.vAlign, cell.height, size, CHECKBOX_PAD)), size };
}

/** Excel 365's checkbox in a cell: a box in the font's colour, filled with a tick in the cell's own colour once ticked. */
export function drawCheckbox(ctx: CanvasRenderingContext2D, box: { x: number; y: number; size: number }, ticked: boolean, color: string, ground: string): void {
  const { x, y, size } = box;
  ctx.save();
  ctx.beginPath();
  ctx.roundRect(x + 0.5, y + 0.5, size - 1, size - 1, Math.max(2, size * 0.18));
  if (ticked) {
    ctx.fillStyle = color;
    ctx.fill();
    ctx.beginPath();
    ctx.moveTo(x + size * 0.25, y + size * 0.52);
    ctx.lineTo(x + size * 0.43, y + size * 0.7);
    ctx.lineTo(x + size * 0.76, y + size * 0.31);
    ctx.strokeStyle = ground;
    ctx.lineWidth = Math.max(1.5, size / 8);
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    ctx.stroke();
  } else {
    ctx.strokeStyle = color;
    ctx.lineWidth = 1;
    ctx.stroke();
  }
  ctx.restore();
}
