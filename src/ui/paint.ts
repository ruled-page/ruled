import type { BorderEdge } from '../core/types';

export const GRID_LINE = '#E1E3E3';
/** The line Excel draws where frozen columns meet the scrolling ones. */
export const FREEZE_LINE = '#B7BCC2';

/**
 * Draws a cell's left or top edge, starting on the boundary pixel (where the
 * grid line sits) and growing into the cell for wider borders. The context is
 * restored afterwards: Glide caches the fill style between cells.
 */
export function drawEdge(ctx: CanvasRenderingContext2D, side: 'left' | 'top', x: number, y: number, length: number, edge: BorderEdge): void {
  ctx.save();
  try {
    paintEdge(ctx, side, x, y, length, edge);
  } finally {
    ctx.restore();
  }
}

function paintEdge(ctx: CanvasRenderingContext2D, side: 'left' | 'top', x: number, y: number, length: number, edge: BorderEdge): void {
  ctx.fillStyle = edge.color;
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
  const on = edge.pattern === 'dotted' ? edge.width : 3 * edge.width;
  const off = edge.pattern === 'dotted' ? edge.width : edge.width + 1;
  for (let p = 0; p < length; p += on + off) line(0, edge.width, p, Math.min(on, length - p));
}

export function drawGridLine(ctx: CanvasRenderingContext2D, side: 'left' | 'top', x: number, y: number, length: number, color = GRID_LINE): void {
  const fill = ctx.fillStyle;
  ctx.fillStyle = color;
  if (side === 'left') ctx.fillRect(x, y, 1, length);
  else ctx.fillRect(x, y, length, 1);
  ctx.fillStyle = fill;
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
