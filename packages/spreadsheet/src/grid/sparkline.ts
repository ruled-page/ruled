import type { SparkCell, SparklineGroup } from '../core/types';
import type { Box } from './contract';

/** Room left between a sparkline and its cell's edges. */
const PAD = 3;

/**
 * Draws a sparkline in its cell as Excel does: a line through its points, a
 * column for each, or win/loss columns up or down from the middle; points
 * Excel's options pick out (high, low, first, last, negative) in their own
 * colours, and the axis at zero when asked for.
 */
export function drawSparkline(ctx: CanvasRenderingContext2D, box: Box, spark: SparkCell, g: SparklineGroup): void {
  let values = g.empty === 'zero' ? spark.values.map((v) => v ?? 0) : spark.values;
  if (g.rightToLeft) values = [...values].reverse();
  const n = values.length;
  if (!n || values.every((v) => v === null)) return;
  const line = g.type === 'line';
  const weight = Math.max(1, (g.lineWeight * 4) / 3);
  const inset = PAD + (line && (g.show.markers || g.show.high || g.show.low || g.show.first || g.show.last || g.show.negative) ? weight : 0);
  const left = box.x + inset;
  const width = box.width - 2 * inset;
  const top = box.y + inset;
  const height = box.height - 2 * inset;
  if (width <= 0 || height <= 0) return;

  const { min, max } = spark;
  const y = (v: number) => (max === min ? top + height / 2 : top + ((max - Math.min(max, Math.max(min, v))) / (max - min)) * height);
  const zero = g.type === 'stacked' ? top + height / 2 : y(Math.min(max, Math.max(min, 0)));
  const shown = values.filter((v): v is number => v !== null);
  const high = Math.max(...shown);
  const low = Math.min(...shown);
  const first = values.findIndex((v) => v !== null);
  const last = values.length - 1 - [...values].reverse().findIndex((v) => v !== null);
  /** A point's colour where an option picks it out, the latest option winning as in Excel. */
  const special = (v: number, i: number): string | undefined => {
    let c: string | undefined;
    if (g.show.negative && v < 0) c = g.colors.negative;
    if (g.show.low && v === low) c = g.colors.low;
    if (g.show.high && v === high) c = g.colors.high;
    if (g.show.first && i === first) c = g.colors.first;
    if (g.show.last && i === last) c = g.colors.last;
    return c;
  };

  ctx.save();
  ctx.beginPath();
  ctx.rect(box.x, box.y, box.width, box.height);
  ctx.clip();
  if (g.show.axis && (g.type === 'stacked' || (min <= 0 && max >= 0))) {
    ctx.fillStyle = g.colors.axis;
    ctx.fillRect(box.x + PAD, Math.round(zero), box.width - 2 * PAD, 1);
  }

  if (line) {
    const x = (i: number) => (n === 1 ? left + width / 2 : left + (i * width) / (n - 1));
    ctx.strokeStyle = g.colors.series;
    ctx.lineWidth = weight;
    ctx.lineJoin = 'round';
    ctx.lineCap = 'round';
    ctx.beginPath();
    let pen = false;
    values.forEach((v, i) => {
      if (v === null) {
        if (g.empty === 'gap') pen = false;
        return;
      }
      if (pen) ctx.lineTo(x(i), y(v));
      else ctx.moveTo(x(i), y(v));
      pen = true;
    });
    ctx.stroke();
    // Small dots, as Excel's: a little wider than the line.
    const r = Math.max(1.25, weight * 0.9);
    values.forEach((v, i) => {
      if (v === null) return;
      const color = special(v, i) ?? (g.show.markers ? g.colors.markers : undefined);
      if (!color) return;
      ctx.fillStyle = color;
      ctx.beginPath();
      ctx.arc(x(i), y(v), r, 0, Math.PI * 2);
      ctx.fill();
    });
  } else {
    const slot = width / n;
    const bar = Math.max(1, slot * 0.8);
    values.forEach((v, i) => {
      if (v === null || (g.type === 'stacked' && v === 0)) return;
      const x = left + i * slot + (slot - bar) / 2;
      let a: number;
      let b: number;
      if (g.type === 'stacked') [a, b] = v > 0 ? [top, zero - 1] : [zero + 1, top + height];
      else [a, b] = [Math.min(y(v), zero), Math.max(y(v), zero)];
      ctx.fillStyle = special(v, i) ?? g.colors.series;
      ctx.fillRect(x, a, bar, Math.max(1, b - a));
    });
  }
  ctx.restore();
}
