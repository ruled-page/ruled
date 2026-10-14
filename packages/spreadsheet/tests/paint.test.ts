import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import type { BorderEdge } from '../src/core/types';
import { REF_COLORS } from '../src/core/formulaAssist';
import { DEFAULT_PALETTE, DEFAULT_REF_COLORS, readGridPalette, readRefColors } from '../src/ui/gridPalette';
import { Recorder, type Gradient } from './canvasRecorder';
import { checkboxBox, drawCheckbox, drawDataBar, drawEdge, drawFilterButton, drawIcon, filterButtonBox } from '../src/ui/paint';

const RED = '#E0301E';
const AMBER = '#F5B800';
const GREEN = '#1E9E4A';
const GREY = '#8A8A8A';
const BLUE = '#2F6EBA';
const OFF = '#D0D7E2';

describe('cell borders', () => {
  const edge = (pattern: BorderEdge['pattern'], width: BorderEdge['width'], color?: string): BorderEdge => ({ pattern, width, color });

  it('draws a solid border its full width, growing into the cell', () => {
    const r = new Recorder();
    drawEdge(r.ctx, 'left', 10, 20, 30, edge('solid', 2, '#C00000'), '#000000');
    drawEdge(r.ctx, 'top', 10, 20, 30, edge('solid', 3, '#C00000'), '#000000');
    expect(r.rects()).toEqual([
      [10, 20, 2, 30, '#C00000'],
      [10, 20, 30, 3, '#C00000'],
    ]);
  });

  it('draws a double border as two hairlines a pixel apart', () => {
    const r = new Recorder();
    drawEdge(r.ctx, 'left', 10, 20, 30, edge('double', 3, '#0000FF'), '#000000');
    drawEdge(r.ctx, 'top', 10, 20, 30, edge('double', 3, '#0000FF'), '#000000');
    expect(r.rects()).toEqual([
      [10, 20, 1, 30, '#0000FF'],
      [12, 20, 1, 30, '#0000FF'],
      [10, 20, 30, 1, '#0000FF'],
      [10, 22, 30, 1, '#0000FF'],
    ]);
  });

  it('draws dashes three times as long as the line is wide, cutting the last one short', () => {
    const r = new Recorder();
    drawEdge(r.ctx, 'left', 10, 20, 12, edge('dashed', 1), '#000000');
    // 3 on, 2 off.
    expect(r.rects()).toEqual([
      [10, 20, 1, 3, '#000000'],
      [10, 25, 1, 3, '#000000'],
      [10, 30, 1, 2, '#000000'],
    ]);
    const wide = new Recorder();
    drawEdge(wide.ctx, 'top', 10, 20, 20, edge('dashed', 2), '#000000');
    // 6 on, 3 off.
    expect(wide.rects()).toEqual([
      [10, 20, 6, 2, '#000000'],
      [19, 20, 6, 2, '#000000'],
      [28, 20, 2, 2, '#000000'],
    ]);
  });

  it('draws dots as long as the line is wide, with gaps the same', () => {
    const r = new Recorder();
    drawEdge(r.ctx, 'top', 0, 0, 5, edge('dotted', 1), '#000000');
    expect(r.rects().map(([x, , w]) => [x, w])).toEqual([
      [0, 1],
      [2, 1],
      [4, 1],
    ]);
  });

  it("draws Excel's dash-dot styles and its hair's finest dots", () => {
    const runs = (pattern: BorderEdge['pattern'], width: BorderEdge['width'], length: number) => {
      const r = new Recorder();
      drawEdge(r.ctx, 'top', 0, 0, length, edge(pattern, width), '#000000');
      return r.rects().map(([x, , w, h]) => [x, w, h]);
    };
    // Dash, gap, dot, gap: 3, 2, 1, 2.
    expect(runs('dashDot', 1, 12)).toEqual([
      [0, 3, 1],
      [5, 1, 1],
      [8, 3, 1],
    ]);
    expect(runs('dashDotDot', 1, 14).map(([x]) => x)).toEqual([0, 5, 8, 11]);
    expect(runs('slantDashDot', 2, 12)).toEqual([
      [0, 6, 2],
      [9, 2, 2],
    ]);
    expect(runs('hair', 2, 4)).toEqual([
      [0, 1, 1],
      [2, 1, 1],
    ]);
  });

  it("uses the theme's colour for borders in Excel's automatic colour", () => {
    const r = new Recorder();
    drawEdge(r.ctx, 'left', 0, 0, 10, edge('solid', 1), '#E8EAED');
    expect(r.rects()).toEqual([[0, 0, 1, 10, '#E8EAED']]);
  });

  it('leaves the context as it found it', () => {
    const r = new Recorder();
    r.fillStyle = '#123456';
    drawEdge(r.ctx, 'left', 0, 0, 10, edge('dashed', 1, '#C00000'), '#000000');
    expect(r.fillStyle).toBe('#123456');
    expect(r.stack).toHaveLength(0);
    expect(r.calls[0].op).toBe('save');
    expect(r.calls.at(-1)?.op).toBe('restore');
  });

  it('restores the context even when drawing fails', () => {
    const r = new Recorder();
    r.fillStyle = '#123456';
    r.failFill = true;
    expect(() => drawEdge(r.ctx, 'top', 0, 0, 10, edge('solid', 1, '#C00000'), '#000000')).toThrow('canvas lost');
    expect(r.fillStyle).toBe('#123456');
    expect(r.stack).toHaveLength(0);
  });
});

describe('data bars', () => {
  const cell = { x: 100, y: 40, width: 104, height: 23 };

  it('fills the share of the cell the value reaches, inside a small margin', () => {
    const r = new Recorder();
    drawDataBar(r.ctx, cell, { from: 0, to: 0.5, color: '#638EC6', gradient: false });
    expect(r.rects()).toEqual([[102, 42, 50, 20, '#638EC6']]);
  });

  it('starts bars for negative values partway along, at the axis', () => {
    const r = new Recorder();
    drawDataBar(r.ctx, cell, { from: 0.25, to: 0.5, color: '#FF0000', gradient: false });
    expect(r.rects()).toEqual([[127, 42, 25, 20, '#FF0000']]);
  });

  it('fades a gradient bar from its colour to white across its own length', () => {
    const r = new Recorder();
    drawDataBar(r.ctx, cell, { from: 0, to: 1, color: '#63BE7B', gradient: true });
    const [[x, y, w, h, fill]] = r.rects() as [number, number, number, number, Gradient][];
    expect([x, y, w, h]).toEqual([102, 42, 100, 20]);
    expect(fill.from).toEqual([102, 0, 202, 0]);
    expect(fill.stops).toEqual([
      [0, '#63BE7B'],
      [1, '#FFFFFF'],
    ]);
  });

  it('draws nothing for an empty bar, and leaves the context as it was', () => {
    const r = new Recorder();
    r.fillStyle = '#ABCDEF';
    drawDataBar(r.ctx, cell, { from: 0.5, to: 0.5, color: '#638EC6', gradient: true });
    drawDataBar(r.ctx, cell, { from: 0.6, to: 0.5, color: '#638EC6', gradient: true });
    expect(r.calls).toEqual([]);
    drawDataBar(r.ctx, cell, { from: 0, to: 0.2, color: '#638EC6', gradient: true });
    expect(r.fillStyle).toBe('#ABCDEF');
    expect(r.stack).toHaveLength(0);
  });
});

describe('icon sets', () => {
  const SETS: Record<string, number> = {
    '3Arrows': 3,
    '3ArrowsGray': 3,
    '3Flags': 3,
    '3TrafficLights1': 3,
    '3TrafficLights2': 3,
    '3Signs': 3,
    '3Symbols': 3,
    '3Symbols2': 3,
    '3Stars': 3,
    '3Triangles': 3,
    '4Arrows': 4,
    '4ArrowsGray': 4,
    '4RedToBlack': 4,
    '4Rating': 4,
    '4TrafficLights': 4,
    '5Arrows': 5,
    '5ArrowsGray': 5,
    '5Rating': 5,
    '5Quarters': 5,
    '5Boxes': 5,
  };

  const icon = (set: string, index: number, size?: number) => {
    const r = new Recorder();
    r.fillStyle = '#ABCDEF';
    drawIcon(r.ctx, set, index, 10, 20, size);
    return r;
  };

  it('draws every icon of every set, each one different, and restores the context', () => {
    for (const [set, n] of Object.entries(SETS)) {
      const drawings = new Set<string>();
      for (let i = 0; i < n; i++) {
        const r = icon(set, i);
        expect(r.calls.some((c) => c.op === 'fill' || c.op === 'stroke' || c.op === 'fillRect'), `${set} ${i}`).toBe(true);
        expect(r.stack, `${set} ${i}`).toHaveLength(0);
        expect(r.fillStyle, `${set} ${i}`).toBe('#ABCDEF');
        drawings.add(JSON.stringify(r.calls));
      }
      expect(drawings.size, set).toBe(n);
    }
  });

  it('colours traffic lights from red to green', () => {
    expect([0, 1, 2].map((i) => icon('3TrafficLights1', i).paints())).toEqual([[`fill ${RED}`], [`fill ${AMBER}`], [`fill ${GREEN}`]]);
    expect([0, 1, 2, 3].map((i) => icon('4TrafficLights', i).paints()[0])).toEqual([`fill #1A1A1A`, `fill ${RED}`, `fill ${AMBER}`, `fill ${GREEN}`]);
    expect([0, 3].map((i) => icon('4RedToBlack', i).paints()[0])).toEqual([`fill #1A1A1A`, `fill ${RED}`]);
    // The rimmed set sits on a dark square.
    expect(icon('3TrafficLights2', 0).rects()).toEqual([[10, 20, 14, 14, '#3A3A3A']]);
    expect(icon('3TrafficLights1', 0).rects()).toEqual([]);
  });

  it('points arrows from down for the lowest values to up for the highest', () => {
    const turn = (set: string, i: number) => (icon(set, i).args('rotate')[0][0] as number) * (180 / Math.PI);
    expect([0, 1, 2].map((i) => turn('3Arrows', i))).toEqual([90, 0, -90]);
    expect([0, 1, 2, 3].map((i) => turn('4Arrows', i))).toEqual([90, 45, -45, -90]);
    expect([0, 1, 2, 3, 4].map((i) => turn('5Arrows', i))).toEqual([90, 45, 0, -45, -90]);
    expect(icon('3Arrows', 0).args('translate')).toEqual([[17, 27]]);
  });

  it('colours arrows red, amber and green, or all grey', () => {
    expect([0, 1, 2].map((i) => icon('3Arrows', i).paints())).toEqual([0, 1, 2].map((i) => [`stroke ${[RED, AMBER, GREEN][i]}`, `fill ${[RED, AMBER, GREEN][i]}`]));
    expect([0, 1, 2, 3, 4].map((i) => icon('5Arrows', i).paints()[0])).toEqual([RED, AMBER, AMBER, AMBER, GREEN].map((c) => `stroke ${c}`));
    for (const set of ['3ArrowsGray', '5ArrowsGray']) expect(icon(set, 0).paints(), set).toEqual([`stroke ${GREY}`, `fill ${GREY}`]);
  });

  it('draws signs as a diamond, a triangle and a circle', () => {
    expect(icon('3Signs', 0).args('lineTo')).toHaveLength(3);
    expect(icon('3Signs', 1).args('lineTo')).toHaveLength(2);
    expect(icon('3Signs', 2).args('arc')).toHaveLength(1);
    expect([0, 1, 2].map((i) => icon('3Signs', i).paints())).toEqual([[`fill ${RED}`], [`fill ${AMBER}`], [`fill ${GREEN}`]]);
  });

  it('draws symbols on a coloured circle, or in colour on their own', () => {
    expect(icon('3Symbols', 0).paints()).toEqual([`fill ${RED}`, 'stroke #FFFFFF']);
    expect(icon('3Symbols2', 2).paints()).toEqual([`stroke ${GREEN}`]);
  });

  it('fills stars by none, half or all', () => {
    expect(icon('3Stars', 0).rects()).toEqual([]);
    expect(icon('3Stars', 1).rects()).toEqual([[10, 20, 7, 14, '#F5C400']]);
    expect(icon('3Stars', 2).rects()).toEqual([[10, 20, 14, 14, '#F5C400']]);
  });

  it('points triangles down, draws a dash, then points up', () => {
    expect(icon('3Triangles', 0).args('moveTo')).toEqual([[17, 32]]);
    expect(icon('3Triangles', 1).rects()).toEqual([[12, 25.5, 10, 3, AMBER]]);
    expect(icon('3Triangles', 2).args('moveTo')).toEqual([[17, 22]]);
  });

  it('fills one more rating bar, box or quarter for each step up', () => {
    const lit = (set: string, i: number) => icon(set, i).rects().map((r) => (r[4] === BLUE ? 1 : 0)).join('');
    expect([0, 1, 2, 3].map((i) => lit('4Rating', i))).toEqual(['000', '100', '110', '111']);
    expect(lit('5Rating', 4)).toBe('1111');
    expect([0, 2, 4].map((i) => lit('5Boxes', i))).toEqual(['0000', '1100', '1111']);
    expect(icon('5Boxes', 0).rects().every((r) => r[4] === OFF)).toBe(true);
    // An empty quarter circle is outlined only; a full one ends where it started, a turn later.
    expect(icon('5Quarters', 0).args('arc')).toHaveLength(1);
    expect(icon('5Quarters', 4).args('arc')[1]).toEqual([17, 27, 6, -Math.PI / 2, -Math.PI / 2 + Math.PI * 2]);
  });

  it('flies flags in red, amber and green from a dark pole', () => {
    expect([0, 1, 2].map((i) => icon('3Flags', i).paints())).toEqual([RED, AMBER, GREEN].map((c) => ['stroke #3A3A3A', `fill ${c}`]));
  });

  it('falls back to a coloured circle for unknown sets, and grey past the end of a set', () => {
    expect(icon('3Unknown', 2).paints()).toEqual([`fill ${GREEN}`]);
    expect(icon('Unknown', 2).paints()).toEqual([`fill ${GREEN}`]);
    expect(icon('5Unknown', 2).paints()).toEqual([`fill ${AMBER}`]);
    expect(icon('3TrafficLights1', 3).paints()).toEqual([`fill ${GREY}`]);
    expect(icon('3Arrows', 7).args('rotate')).toEqual([[0]]);
  });

  it('scales to the size asked for', () => {
    expect(icon('3TrafficLights1', 0, 20).args('arc')).toEqual([[20, 30, 9, 0, Math.PI * 2]]);
  });
});

describe('the grid palette', () => {
  const original = globalThis.getComputedStyle;
  afterEach(() => {
    globalThis.getComputedStyle = original;
  });

  /** A stylesheet that sets these custom properties, and records which were read. */
  const stylesheet = (props: Record<string, string>) => {
    const read: string[] = [];
    globalThis.getComputedStyle = (() => ({
      getPropertyValue: (name: string) => {
        read.push(name);
        return props[name] ?? '';
      },
    })) as unknown as typeof getComputedStyle;
    return read;
  };
  const el = {} as Element;

  it("uses the stylesheet's colours and fonts, falling back for any it leaves out", () => {
    stylesheet({
      '--wbe-grid-bg': '  #101010 ',
      '--wbe-grid-header-selected-bg': '#123456',
      '--wbe-grid-font': 'Inter, sans-serif',
      '--wbe-grid-line': '   ',
    });
    const p = readGridPalette(el);
    expect(p).toEqual({ ...DEFAULT_PALETTE, bg: '#101010', headerSelectedBg: '#123456', font: 'Inter, sans-serif' });
    expect(DEFAULT_PALETTE.bg).toBe('#ffffff');
  });

  it("takes the selection's text colour from the brand's", () => {
    stylesheet({ '--wbe-on-accent': '#000000', '--wbe-grid-on-selection': '#FF00FF' });
    expect(readGridPalette(el).onSelection).toBe('#000000');
  });

  it('reads one custom property for each entry, named after it', () => {
    const read = stylesheet({});
    readGridPalette(el);
    expect(read).toHaveLength(Object.keys(DEFAULT_PALETTE).length);
    expect(read).toEqual(expect.arrayContaining(['--wbe-grid-header-selected-text', '--wbe-grid-text-on-fill', '--wbe-grid-cut-fill', '--wbe-grid-header-font-weight', '--wbe-on-accent']));
  });

  it('uses the defaults with no element, or outside a browser', () => {
    stylesheet({ '--wbe-grid-bg': '#101010' });
    expect(readGridPalette(null)).toBe(DEFAULT_PALETTE);
    delete (globalThis as { getComputedStyle?: unknown }).getComputedStyle;
    expect(readGridPalette(el)).toBe(DEFAULT_PALETTE);
  });

  it('has the same defaults as the stylesheet', () => {
    const css = readFileSync(join(__dirname, '..', 'src', 'editor', 'editor.css'), 'utf8');
    const declared = (name: string) => new RegExp(`${name}:\\s*([^;]+);`).exec(css)?.[1].trim();
    for (const [key, value] of Object.entries(DEFAULT_PALETTE)) {
      const name = key === 'onSelection' ? '--wbe-on-accent' : `--wbe-grid-${key.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)}`;
      expect(declared(name), name).toBe(value);
    }
  });

  it('reads the colours of references in a formula being edited, falling back for any the stylesheet leaves out', () => {
    stylesheet({ '--wbe-formula-ref2': ' #ff0000 ' });
    expect(readRefColors(el)).toEqual([DEFAULT_REF_COLORS[0], '#ff0000', ...DEFAULT_REF_COLORS.slice(2)]);
    expect(DEFAULT_REF_COLORS).toHaveLength(REF_COLORS);
    expect(readRefColors(null)).toBe(DEFAULT_REF_COLORS);
    const css = readFileSync(join(__dirname, '..', 'src', 'editor', 'editor.css'), 'utf8');
    DEFAULT_REF_COLORS.forEach((value, i) => expect(new RegExp(`--wbe-formula-ref${i + 1}:\\s*([^;]+);`).exec(css)?.[1]).toBe(value));
    delete (globalThis as { getComputedStyle?: unknown }).getComputedStyle;
    expect(readRefColors(el)).toBe(DEFAULT_REF_COLORS);
  });
});

describe('filter buttons', () => {
  const colors = { bg: '#FFFFFF', line: '#CCCCCC', arrow: '#444444', filtered: '#2F6EBA' };
  const cell = { x: 100, y: 20, width: 64, height: 20 };

  it('sit at the right of the heading, shrinking to fit a small cell', () => {
    expect(filterButtonBox(cell)).toEqual({ x: 146, y: 22, size: 16 });
    expect(filterButtonBox({ x: 0, y: 0, width: 10, height: 30 })).toEqual({ x: 0, y: 11, size: 8 });
  });

  it('draw a framed arrow, or a funnel on a filtered column, and restore the context', () => {
    const plain = new Recorder();
    drawFilterButton(plain.ctx, cell, false, colors);
    expect(plain.rects()).toEqual([
      [146, 22, 16, 16, '#CCCCCC'],
      [147, 23, 14, 14, '#FFFFFF'],
    ]);
    expect(plain.paints()).toEqual(['stroke #444444']);
    const filtered = new Recorder();
    drawFilterButton(filtered.ctx, cell, true, colors);
    expect(filtered.paints()).toEqual(['fill #2F6EBA']);
    expect(filtered.stack).toHaveLength(0);
    expect(filtered.fillStyle).toBe('#000000');
  });
});

describe('checkboxes', () => {
  const cell = { x: 100, y: 20, width: 64, height: 20 };
  const font = { fontSizePx: 14.7 };

  it('are as big as their font, centred across and at the bottom unless the cell is aligned otherwise', () => {
    expect(checkboxBox(cell, font)).toEqual({ x: 126, y: 24, size: 13 });
    expect(checkboxBox(cell, { ...font, hAlign: 'left', vAlign: 'top' })).toEqual({ x: 105, y: 23, size: 13 });
    expect(checkboxBox(cell, { ...font, hAlign: 'right', vAlign: 'center' })).toEqual({ x: 147, y: 24, size: 13 });
    // A short row makes them smaller, never under 8 pixels.
    expect(checkboxBox({ ...cell, height: 6 }, font).size).toBe(8);
  });

  it('draw an empty box, or a filled one with a tick in the cell\'s colour, and restore the context', () => {
    const box = { x: 125, y: 24, size: 13 };
    const empty = new Recorder();
    drawCheckbox(empty.ctx, box, false, '#222222', '#FFFFFF');
    expect(empty.paints()).toEqual(['stroke #222222']);
    const ticked = new Recorder();
    drawCheckbox(ticked.ctx, box, true, '#222222', '#FFFFFF');
    expect(ticked.paints()).toEqual(['fill #222222', 'stroke #FFFFFF']);
    expect(ticked.stack).toHaveLength(0);
    expect(ticked.strokeStyle).toBe('#000000');
  });
});
