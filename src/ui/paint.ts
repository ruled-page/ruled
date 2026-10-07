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
