export interface Gradient {
  from: number[];
  stops: [number, string][];
  addColorStop(offset: number, color: string): void;
}

interface Call {
  op: string;
  args: unknown[];
  /** The fill style for fillRect and fill, the stroke style for stroke. */
  style?: unknown;
}

const STATE = ['fillStyle', 'strokeStyle', 'lineWidth', 'lineCap', 'lineJoin', 'font', 'textAlign', 'textBaseline'] as const;

/**
 * A canvas context that records what is drawn, and keeps save/restore state
 * as a real one does. Text measures half its font size a character, with
 * 0.75 of the size above the baseline and 0.25 below.
 */
export class Recorder {
  calls: Call[] = [];
  fillStyle: unknown = '#000000';
  strokeStyle: unknown = '#000000';
  lineWidth = 1;
  lineCap = 'butt';
  lineJoin = 'miter';
  font = '10px sans-serif';
  textAlign = 'start';
  textBaseline = 'alphabetic';
  stack: Record<string, unknown>[] = [];
  /** Makes fillRect throw, to check the context is restored anyway. */
  failFill = false;

  private record(op: string, args: unknown[], style?: unknown) {
    this.calls.push(style === undefined ? { op, args } : { op, args, style });
  }
  save() {
    this.stack.push(Object.fromEntries(STATE.map((k) => [k, this[k]])));
    this.record('save', []);
  }
  restore() {
    Object.assign(this, this.stack.pop());
    this.record('restore', []);
  }
  fillRect(...args: number[]) {
    if (this.failFill) throw new Error('canvas lost');
    this.record('fillRect', args, this.fillStyle);
  }
  fill(...a: unknown[]) {
    this.record('fill', a, this.fillStyle);
  }
  stroke() {
    this.record('stroke', [], this.strokeStyle);
  }
  createLinearGradient(...from: number[]): Gradient {
    return { from, stops: [], addColorStop(offset, color) { this.stops.push([offset, color]); } };
  }
  beginPath = (...a: unknown[]) => this.record('beginPath', a);
  closePath = (...a: unknown[]) => this.record('closePath', a);
  moveTo = (...a: unknown[]) => this.record('moveTo', a);
  lineTo = (...a: unknown[]) => this.record('lineTo', a);
  arc = (...a: unknown[]) => this.record('arc', a);
  roundRect = (...a: unknown[]) => this.record('roundRect', a);
  translate = (...a: unknown[]) => this.record('translate', a);
  rotate = (...a: unknown[]) => this.record('rotate', a);
  clip = (...a: unknown[]) => this.record('clip', a);
  rect = (...a: unknown[]) => this.record('rect', a);
  setLineDash = (...a: unknown[]) => this.record('setLineDash', a);
  strokeRect(...args: number[]) {
    this.record('strokeRect', args, this.strokeStyle);
  }
  setTransform = (...a: unknown[]) => this.record('setTransform', a);
  fillText(text: string, x: number, y: number) {
    this.calls.push({ op: 'fillText', args: [text, x, y], style: `${this.fillStyle} ${this.textAlign} ${this.font}` });
  }
  measureText(text: string) {
    const size = Number(/(\d+(?:\.\d+)?)px/.exec(this.font)?.[1] ?? 10);
    return { width: text.length * size * 0.5, actualBoundingBoxAscent: size * 0.75, actualBoundingBoxDescent: size * 0.25 };
  }

  get ctx(): CanvasRenderingContext2D {
    return this as unknown as CanvasRenderingContext2D;
  }
  rects() {
    return this.calls.filter((c) => c.op === 'fillRect').map((c) => [...(c.args as number[]), c.style]);
  }
  /** Colours of shapes filled or stroked, in order. */
  paints() {
    return this.calls.filter((c) => c.op === 'fill' || c.op === 'stroke').map((c) => `${c.op} ${c.style}`);
  }
  args(op: string) {
    return this.calls.filter((c) => c.op === op).map((c) => c.args);
  }
  /** Text drawn: what, where, and in what colour, alignment and font. */
  texts() {
    return this.calls.filter((c) => c.op === 'fillText').map((c) => [...c.args, c.style] as [string, number, number, string]);
  }
}

