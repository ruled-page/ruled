import { describe, expect, it } from 'vitest';
import type { SparkCell, SparklineGroup } from '../src/core/types';
import { drawSparkline } from '../src/grid/sparkline';
import { Recorder } from './canvasRecorder';

const BOX = { x: 0, y: 0, width: 106, height: 26 };

const group = (over: Partial<SparklineGroup> = {}): SparklineGroup => ({
  type: 'line',
  colors: { series: 'S', negative: 'N', axis: 'A', markers: 'M', first: 'F', last: 'L', high: 'H', low: 'W' },
  show: { markers: false, high: false, low: false, first: false, last: false, negative: false, axis: false },
  lineWeight: 0.75,
  empty: 'gap',
  minAxis: 'individual',
  maxAxis: 'individual',
  displayHidden: false,
  rightToLeft: false,
  lines: [],
  ...over,
});

const spark = (values: (number | null)[], min = Math.min(...values.filter((v) => v !== null)), max = Math.max(...values.filter((v) => v !== null))): SparkCell => ({
  row: 0,
  col: 0,
  group: 0,
  values,
  min,
  max,
});

function draw(values: (number | null)[], g: Partial<SparklineGroup>, s = spark(values)): Recorder {
  const r = new Recorder();
  drawSparkline(r.ctx, BOX, s, group(g));
  return r;
}

describe('sparklines', () => {
  it('draws a line through the points, clipped to the cell, its highest at the top', () => {
    const r = draw([1, 3, 2], {});
    expect(r.args('rect')).toEqual([[0, 0, 106, 26]]);
    expect(r.args('clip')).toHaveLength(1);
    expect(r.args('moveTo')).toEqual([[3, 23]]);
    expect(r.args('lineTo')).toEqual([
      [53, 3],
      [103, 13],
    ]);
    expect(r.paints()).toEqual(['stroke S']);
  });

  it('breaks the line at an empty cell, draws through it as zero, or carries it across', () => {
    expect(draw([1, null, 2], { empty: 'gap' }).args('moveTo')).toHaveLength(2);
    expect(draw([1, null, 2], { empty: 'span' }).args('moveTo')).toHaveLength(1);
    const zero = draw([1, null, 2], { empty: 'zero' }, spark([1, 0, 2]));
    expect(zero.args('lineTo')).toHaveLength(2);
  });

  it('marks points in the colours Excel’s options pick, the later option winning', () => {
    const r = draw([-1, 4, 2, 0], { show: { markers: true, high: true, low: true, first: true, last: true, negative: true, axis: false } });
    // First (also low and negative), high, a plain marker, last.
    expect(r.paints()).toEqual(['stroke S', 'fill F', 'fill H', 'fill M', 'fill L']);
  });

  it('marks only the points an option picks when markers are off', () => {
    expect(draw([1, 5, 2], { show: { markers: false, high: true, low: false, first: false, last: false, negative: false, axis: false } }).paints()).toEqual(['stroke S', 'fill H']);
  });

  it('draws one point in the middle, and nothing for no values or a cell too small', () => {
    expect(draw([7], {}).args('moveTo')).toEqual([[53, 13]]);
    expect(draw([null, null], {}).calls).toEqual([]);
    expect(draw([], {}).calls).toEqual([]);
    const tiny = new Recorder();
    drawSparkline(tiny.ctx, { x: 0, y: 0, width: 5, height: 5 }, spark([1, 2]), group());
    expect(tiny.calls).toEqual([]);
  });

  it('runs right to left when the group says so', () => {
    expect(draw([1, 3], { rightToLeft: true }).args('moveTo')).toEqual([[3, 3]]);
  });

  it('draws columns up or down from the axis at zero, with the axis when asked for', () => {
    const r = draw([2, -2], { type: 'column', show: { markers: false, high: false, low: false, first: false, last: false, negative: true, axis: true } });
    expect(r.rects()).toEqual([
      [3, 13, 106 - 6, 1, 'A'],
      [8, 3, 40, 10, 'S'],
      [58, 13, 40, 10, 'N'],
    ]);
  });

  it('draws win/loss columns from the middle, skipping zeros', () => {
    const r = draw([5, 0, -1], { type: 'stacked', show: { markers: false, high: false, low: false, first: false, last: false, negative: false, axis: true } });
    const rects = r.rects();
    expect(rects[0]).toEqual([3, 13, 100, 1, 'A']);
    expect(rects.slice(1).map(([x, y, , h]) => [Math.round(x as number), y, h])).toEqual([
      [6, 3, 9],
      [73, 14, 9],
    ]);
  });
});
