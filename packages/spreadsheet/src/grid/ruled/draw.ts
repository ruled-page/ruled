import { cellKey, colToLetters } from '@ruled/formula/refs';
import { CellFlag, VT, type BorderEdge, type DisplayStyle, type MergeRange } from '../../core/types';
import { wrapLines } from '../../state/autofit';
import { effectiveStyle } from '../../state/cellStyle';
import { checkboxState } from '../../state/checkbox';
import { alignment, displayText } from '../../state/format';
import { CELL_PADDING, excelTextWidth, fitNumber, indentPx, overflowTextX, rowOverflows, textTop, type Overflow } from '../../state/overflow';
import type { SheetView } from '../../state/sheetView';
import type { Selection } from '../../state/store';
import type { GridPalette } from '../../ui/gridPalette';
import { drawEdge } from '../../ui/paint';
import { withAlpha } from '../color';
import type { Box } from '../contract';
import type { Mark } from '../marks';
import type { Layout, Run } from './layout';

/** Space above and below text in a cell, and between wrapped lines, as a share of the font's height. */
const V_PAD = 2;
const LINE_HEIGHT = 1.25;
/** Text whose height a font's lines are measured by: capitals, ascenders, descenders and digits. */
const METRIC_TEXT = 'ABCi09jgqpy';

/** What the grid is drawing. */
export interface PaintInput {
  view: SheetView;
  layout: Layout;
  styles: DisplayStyle[];
  palette: GridPalette;
  date1904: boolean;
  locale: string;
  /** Numbers in full, rather than rounded as their formats show them. */
  fullNumbers: boolean;
  /** Changed cells are shown in a fill of their own. */
  highlightEdits: boolean;
  /** A schema's label cells, filled to set them apart. */
  isLabel?: (row: number, col: number) => boolean;
  selection?: Selection;
  /** Ranges marked besides the selection (see sheetMarks). */
  marks?: readonly Mark[];
}

/** How a cell is drawn. */
interface Look {
  /** Its own style, borders and all. */
  style: DisplayStyle;
  /** The style its text is drawn in: its own, or smaller for a number that wouldn't fit. */
  font: DisplayStyle;
  fill?: string;
  text: string;
  color: string;
  align: 'left' | 'right' | 'center';
}

/**
 * Draws a sheet on a canvas, keeping what it works out about each cell
 * (its text, colours and the text running into its neighbours) until it is
 * told the cells changed.
 */
export class Painter {
  private readonly looks = new Map<number, Look>();
  private readonly overflowsByRow = new Map<number, Overflow[]>();
  private readonly wraps = new Map<string, string[]>();
  private readonly metrics = new Map<string, { ascent: number; descent: number }>();
  private palette?: GridPalette;

  constructor(private readonly ctx: CanvasRenderingContext2D) {}

  /** Forgets what it worked out: for when cells, styles, column widths or the palette change. */
  reset(): void {
    this.looks.clear();
    this.overflowsByRow.clear();
    this.wraps.clear();
  }

  /** The CSS font a cell's text is drawn in. */
  font(style: DisplayStyle): string {
    return `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${style.fontSizePx}px "${style.fontFamily}", ${this.palette?.font ?? 'sans-serif'}`;
  }

  /** How wide text is in a cell's font, as drawn here. */
  measure(text: string, style: DisplayStyle): number {
    this.ctx.font = this.font(style);
    return this.ctx.measureText(text).width;
  }

  /** How far a font's text reaches above and below its baseline. */
  private lineMetrics(font: string): { ascent: number; descent: number } {
    let m = this.metrics.get(font);
    if (!m) {
      this.ctx.font = font;
      const t = this.ctx.measureText(METRIC_TEXT);
      this.metrics.set(font, (m = { ascent: t.actualBoundingBoxAscent, descent: t.actualBoundingBoxDescent }));
    }
    return m;
  }

  draw(input: PaintInput): void {
    const { ctx } = this;
    const { layout, palette } = input;
    if (this.palette !== palette) {
      this.palette = palette;
      this.reset();
    }
    const { width, height } = layout.spec;
    ctx.save();
    ctx.fillStyle = palette.bg;
    ctx.fillRect(0, 0, width, height);
    for (const q of layout.quadrants()) this.drawCells(input, q.rows, q.cols, q.clip);
    this.drawHeadings(input);
    ctx.fillStyle = palette.freezeLine;
    if (layout.frozenColRun.last >= 0) this.fill(layout.box(layout.bodyX - 1, layout.spec.letters, 1, height - layout.spec.letters));
    if (layout.frozenRowRun.last >= 0) this.fill(layout.box(layout.spec.numbers, layout.bodyY - 1, width - layout.spec.numbers, 1));
    ctx.restore();
  }

  private fill(b: Box): void {
    this.ctx.fillRect(b.x, b.y, b.width, b.height);
  }

  private clip(b: Box): void {
    this.ctx.beginPath();
    this.ctx.rect(b.x, b.y, b.width, b.height);
    this.ctx.clip();
  }

  /** How a cell is drawn: for a merge, its top-left cell across the whole merge. */
  private look(input: PaintInput, row: number, col: number, merge?: MergeRange): Look {
    const key = cellKey(row, col);
    let look = this.looks.get(key);
    if (look) return look;
    if (this.looks.size > 50_000) this.looks.clear();
    const { view, styles, palette } = input;
    const rec = view.get(row, col);
    const blocked = view.blockedAt(row, col);
    const locked = !!blocked && blocked.kind !== 'tableHeader';
    const label = !!input.isLabel?.(row, col);
    if (!rec) {
      const style = effectiveStyle(view, styles, row, col, view.styleAt(row, col, null));
      look = { style, font: style, fill: style.fill ?? (locked ? palette.lockedFill : label ? palette.labelFill : undefined), text: '', color: palette.text, align: 'left' };
    } else {
      const style = effectiveStyle(view, styles, row, col, rec.style, palette.link);
      let { text, color } = displayText(rec, style, input.date1904, input.fullNumbers, input.locale);
      // A checkbox is drawn in place of TRUE or FALSE; a data bar may show without its value.
      if (checkboxState(rec, style) !== undefined || view.cfAt(row, col)?.hideValue) text = '';
      let font = style;
      if (rec.t === VT.Number && !style.wrap) {
        // Numbers never spill into neighbours; Excel shortens them or shows ####.
        let room = -CELL_PADDING * 2;
        for (let c = merge?.c1 ?? col; c <= (merge?.c2 ?? col); c++) room += view.colWidth(c);
        if (text.length * style.fontSizePx * 0.62 > room) {
          // Whether it fits is Excel's call, made with Calibri's widths; where
          // the browser's stand-in font is wider, the number is drawn smaller.
          text = fitNumber(text, rec.n ?? 0, /^general$/i.test(style.numFmt), room, (s) => excelTextWidth(s, style));
          const drawn = text.startsWith('#') ? 0 : this.measure(text, style);
          if (drawn > room) font = { ...style, fontSizePx: Math.max(6, Math.floor(((style.fontSizePx * room) / drawn) * 10) / 10) };
        }
      }
      // A changed cell shows its own fill, if it has one; text with no colour of its own on a fill from the file is drawn as Excel draws it.
      const edited = !label && input.highlightEdits && !!(rec.flags & CellFlag.Edited);
      const fill = label && !style.fill ? palette.labelFill : edited && !style.fill ? palette.editedFill : (style.fill ?? (locked ? palette.lockedFill : undefined));
      look = { style, font, fill, text, color: color ?? style.color ?? (style.fill && !edited ? palette.textOnFill : palette.text), align: alignment(rec, style) };
    }
    this.looks.set(key, look);
    return look;
  }

  /** A cell's text as drawn, for the table screen readers read: a number shown as #### in full. */
  cellText(input: PaintInput, row: number, col: number): string {
    const m = input.view.mergeAt(row, col);
    if (m && (m.r1 !== row || m.c1 !== col)) return '';
    const text = this.look(input, row, col, m).text;
    if (!/^#+$/.test(text)) return text;
    const rec = input.view.get(row, col);
    const style = this.look(input, row, col, m).style;
    return rec ? displayText(rec, style, input.date1904, input.fullNumbers, input.locale).text : text;
  }

  /** Text running out of a row's cells into empty neighbours. */
  private overflows(input: PaintInput, row: number): Overflow[] {
    let list = this.overflowsByRow.get(row);
    if (!list) {
      const { view, styles, palette } = input;
      if (this.overflowsByRow.size > 2_000) this.overflowsByRow.clear();
      list = rowOverflows(view, row, (col, rec) => effectiveStyle(view, styles, row, col, rec.style, palette.link), input.date1904, (s, style) => this.measure(s, style));
      this.overflowsByRow.set(row, list);
    }
    return list;
  }

  /** A merge's box: from its first line in view to its last. */
  private mergeBox(layout: Layout, view: SheetView, m: MergeRange): Box {
    const [r1, r2] = view.lineSpan('rows', m.r1, m.r2)!;
    const [c1, c2] = view.lineSpan('cols', m.c1, m.c2)!;
    const x = layout.x(c1);
    const y = layout.y(r1);
    return layout.box(x, y, layout.x(c2) + layout.spec.cols.size(c2) - x, layout.y(r2) + layout.spec.rows.size(r2) - y);
  }

  /** One part of the cells' area: fills, then grid lines, text and borders. */
  private drawCells(input: PaintInput, rows: Run, cols: Run, clip: Box): void {
    const { ctx } = this;
    const { view, layout, palette } = input;
    const sheetRow = (g: number) => view.visibleRows[g];
    const sheetCol = (g: number) => view.visibleCols[g];
    ctx.save();
    this.clip(clip);

    const merges = new Set<MergeRange>();
    for (let gr = rows.first; gr <= rows.last; gr++) {
      for (let gc = cols.first; gc <= cols.last; gc++) {
        const m = view.mergeAt(sheetRow(gr), sheetCol(gc));
        if (m) merges.add(m);
        const look = m ? undefined : this.look(input, sheetRow(gr), sheetCol(gc));
        if (look?.fill) {
          ctx.fillStyle = look.fill;
          this.fill(layout.cell(gr, gc));
        }
      }
    }
    for (const m of merges) {
      const look = this.look(input, m.r1, m.c1, m);
      if (!look.fill) continue;
      ctx.fillStyle = look.fill;
      this.fill(this.mergeBox(layout, view, m));
    }

    // Grid lines: each cell's left and top, but not inside a merge or under text running across.
    if (view.info.showGridLines) {
      ctx.fillStyle = palette.line;
      for (let gr = rows.first; gr <= rows.last; gr++) {
        const row = sheetRow(gr);
        const overflows = this.overflows(input, row);
        for (let gc = cols.first; gc <= cols.last; gc++) {
          const col = sheetCol(gc);
          const m = view.mergeAt(row, col);
          const cell = layout.cell(gr, gc);
          if ((!m || col === m.c1) && !overflows.some((o) => o.from < col && col <= o.to)) this.fill(layout.box(layout.x(gc), cell.y, 1, cell.height));
          if (!m || row === m.r1) this.fill({ ...cell, height: 1 });
        }
      }
    }

    // Text: each cell's in its own box, a merge's across it, and text running into neighbours once, across them.
    for (let gr = rows.first; gr <= rows.last; gr++) {
      const row = sheetRow(gr);
      const overflows = this.overflows(input, row);
      for (let gc = cols.first; gc <= cols.last; gc++) {
        const col = sheetCol(gc);
        if (view.mergeAt(row, col) || overflows.some((o) => o.col === col)) continue;
        const look = this.look(input, row, col);
        if (look.text) this.drawText(input, layout.cell(gr, gc), look);
      }
      for (const o of overflows) {
        if (view.gridCol(o.from) <= cols.last && view.gridCol(o.to) >= cols.first) this.drawOverflow(input, gr, o);
      }
    }
    for (const m of merges) {
      const look = this.look(input, m.r1, m.c1, m);
      if (look.text) this.drawText(input, this.mergeBox(layout, view, m), look);
    }

    // Borders: each cell's left and top, its own or its neighbour's right and bottom.
    for (let gr = rows.first; gr <= rows.last; gr++) {
      const row = sheetRow(gr);
      for (let gc = cols.first; gc <= cols.last; gc++) {
        const col = sheetCol(gc);
        const m = view.mergeAt(row, col);
        const own = this.look(input, m?.r1 ?? row, m?.c1 ?? col, m).style.border;
        const cell = layout.cell(gr, gc);
        if (!m || col === m.c1) {
          const edge = own?.left ?? (gc > 0 ? this.look(input, row, sheetCol(gc - 1)).style.border?.right : undefined);
          if (edge) this.edge('left', layout.x(gc), cell.y, cell.height, edge, input);
        }
        if (!m || row === m.r1) {
          const edge = own?.top ?? (gr > 0 ? this.look(input, sheetRow(gr - 1), col).style.border?.bottom : undefined);
          if (edge) this.edge('top', layout.x(gc), cell.y, cell.width, edge, input);
        }
      }
    }
    this.drawSelection(input);
    ctx.restore();
  }

  /** The box a range takes on screen, from its first line in view to its last; undefined when all of it is hidden. */
  private rangeBox(input: PaintInput, r: { r1: number; r2: number; c1: number; c2: number }): Box | undefined {
    const { view, layout } = input;
    const span = view.gridSpan(r);
    if (!span) return undefined;
    const x = layout.x(span.x1);
    const y = layout.y(span.y1);
    return layout.box(x, y, layout.x(span.x2) + layout.spec.cols.size(span.x2) - x, layout.y(span.y2) + layout.spec.rows.size(span.y2) - y);
  }

  /**
   * The selection, tinted but for its active cell and outlined, as Excel
   * shows it: whole rows and columns in full, past the cells used. Then the
   * other marks, each filled and outlined in its colour.
   */
  private drawSelection(input: PaintInput): void {
    const { ctx } = this;
    const { selection: sel, view, palette } = input;
    if (sel) {
      const lastRow = view.visibleRows[view.rowCount - 1] ?? 0;
      const lastCol = view.visibleCols[view.colCount - 1] ?? 0;
      const whole = (r: Selection['rects'][number]) =>
        sel.lines === 'all' ? { r1: 0, r2: lastRow, c1: 0, c2: lastCol } : sel.lines === 'cols' ? { ...r, r1: 0, r2: lastRow } : sel.lines === 'rows' ? { ...r, c1: 0, c2: lastCol } : r;
      const boxes = sel.rects.map((r) => this.rangeBox(input, whole(r)));
      const active = this.rangeBox(input, view.mergeAt(sel.row, sel.col) ?? { r1: sel.row, r2: sel.row, c1: sel.col, c2: sel.col });
      ctx.fillStyle = palette.selectionFill;
      for (const b of boxes) {
        if (!b) continue;
        ctx.beginPath();
        ctx.rect(b.x, b.y, b.width, b.height);
        if (active) ctx.rect(active.x, active.y, active.width, active.height);
        ctx.fill('evenodd');
      }
      ctx.fillStyle = palette.selection;
      const first = boxes[0] ?? active;
      if (first) this.outline(first, 2);
    }
    for (const m of input.marks ?? []) {
      const b = this.rangeBox(input, m.rect);
      if (!b) continue;
      ctx.fillStyle = m.color;
      this.fill(b);
      ctx.save();
      ctx.strokeStyle = withAlpha(m.color, 1);
      ctx.lineWidth = 1;
      ctx.setLineDash(m.style === 'dashed' ? [4, 3] : []);
      ctx.strokeRect(b.x + 0.5, b.y + 0.5, b.width - 1, b.height - 1);
      ctx.restore();
    }
  }

  /** A line `width` pixels wide round a box, on its edges, in the fill style. */
  private outline(b: Box, width: number): void {
    const h = width / 2;
    this.ctx.fillRect(b.x - h, b.y - h, b.width + width, width);
    this.ctx.fillRect(b.x - h, b.y + b.height - h, b.width + width, width);
    this.ctx.fillRect(b.x - h, b.y - h, width, b.height + width);
    this.ctx.fillRect(b.x + b.width - h, b.y - h, width, b.height + width);
  }

  /** A border along a cell's start edge (from its start, measured from the start edge) or its top. */
  private edge(side: 'left' | 'top', x: number, y: number, length: number, edge: BorderEdge, input: PaintInput): void {
    const { layout, palette } = input;
    const at = side === 'left' ? layout.box(x, y, edge.width, length) : layout.box(x, y, length, edge.width);
    drawEdge(this.ctx, side, at.x, at.y, length, edge, palette.borderAuto);
  }

  /** Text in a box, cut at its edges: on one line, or wrapped, where Excel's alignment puts it. */
  private drawText(input: PaintInput, box: Box, look: Look): void {
    const { ctx } = this;
    const s = look.font;
    const font = this.font(s);
    const { ascent, descent } = this.lineMetrics(font);
    const em = ascent + descent;
    const indent = indentPx(s, look.align);
    const lines = s.wrap ? this.wrap(look.text, box.width - CELL_PADDING * 2 - indent, s, font) : [look.text];
    const step = LINE_HEIGHT * em;
    const top = box.y + textTop(s.vAlign, box.height, em + step * (lines.length - 1), V_PAD);
    // The start edge is the right one when the sheet reads right to left.
    const align = input.layout.spec.rtl && look.align !== 'center' ? (look.align === 'left' ? 'right' : 'left') : look.align;
    const x = align === 'left' ? box.x + CELL_PADDING + indent : align === 'right' ? box.x + box.width - CELL_PADDING - indent : box.x + box.width / 2;
    ctx.save();
    this.clip(box);
    ctx.font = font;
    ctx.fillStyle = look.color;
    ctx.textAlign = align;
    ctx.textBaseline = 'alphabetic';
    lines.forEach((line, i) => {
      const y = top + i * step;
      if (y > box.y + box.height || y + em < box.y) return;
      ctx.fillText(line, x, y + ascent);
      this.decorate(line, x, align, y + em / 2, s);
    });
    ctx.restore();
  }

  /** Strikethrough and underline, across text drawn from `x` with its middle at `mid`. */
  private decorate(text: string, x: number, align: 'left' | 'right' | 'center', mid: number, s: DisplayStyle): void {
    if (!s.strike && !s.underline) return;
    const width = this.measure(text, s);
    const left = align === 'right' ? x - width : align === 'center' ? x - width / 2 : x;
    const thickness = s.bold ? 1.5 : 1;
    if (s.strike) this.ctx.fillRect(left, mid, width, thickness);
    if (s.underline) this.ctx.fillRect(left, mid + s.fontSizePx * 0.45, width, thickness);
  }

  /** Text running out of its cell into the empty ones beside it, drawn once across them all. */
  private drawOverflow(input: PaintInput, gr: number, o: Overflow): void {
    const { ctx } = this;
    const { view, layout, palette } = input;
    const font = this.font(o.style);
    const { ascent, descent } = this.lineMetrics(font);
    const em = ascent + descent;
    const own = view.gridCol(o.col);
    const from = view.gridCol(o.from);
    const to = view.gridCol(o.to);
    const y = layout.y(gr);
    const height = layout.spec.rows.size(gr);
    const startX = layout.x(own) + overflowTextX(o, view.colWidth(o.col));
    const span = layout.box(layout.x(from), y, layout.x(to) + layout.spec.cols.size(to) - layout.x(from), height);
    const at = layout.box(startX, y, o.width, height);
    const top = y + textTop(o.style.vAlign, height, em, V_PAD);
    ctx.save();
    this.clip(span);
    ctx.font = font;
    ctx.fillStyle = o.color ?? o.style.color ?? (o.style.fill ? palette.textOnFill : palette.text);
    ctx.textAlign = 'left';
    ctx.textBaseline = 'alphabetic';
    ctx.fillText(o.text, at.x, top + ascent);
    this.decorate(o.text, at.x, 'left', top + em / 2, o.style);
    ctx.restore();
  }

  /** Wrapped text's lines, kept while the same text is drawn at the same width. */
  private wrap(text: string, width: number, style: DisplayStyle, font: string): string[] {
    const key = `${font}\u0000${width}\u0000${text}`;
    let lines = this.wraps.get(key);
    if (!lines) {
      if (this.wraps.size > 2_000) this.wraps.clear();
      lines = wrapLines(text, width, (s) => this.measure(s, style));
      this.wraps.set(key, lines);
    }
    return lines;
  }

  /** The column letters and row numbers, and the corner where they meet. */
  private drawHeadings(input: PaintInput): void {
    const { ctx } = this;
    const { view, layout, palette } = input;
    const { width, height, numbers, letters } = layout.spec;
    const font = `${palette.headerFontWeight} ${palette.headerFontSize} ${palette.headerFont}`;
    const { ascent, descent } = this.lineMetrics(font);
    const sel = input.selection;
    /** Whether a heading's line is in the selection, or selected whole. */
    const state = (axis: 'rows' | 'cols', line: number): 'line' | 'range' | undefined => {
      if (!sel) return undefined;
      const inside = sel.lines === 'all' || sel.rects.some((r) => (axis === 'rows' ? line >= r.r1 && line <= r.r2 : line >= r.c1 && line <= r.c2));
      if (!inside) return undefined;
      return sel.lines === 'all' || sel.lines === axis ? 'line' : 'range';
    };
    const label = (text: string, b: Box, color: string, selected?: 'line' | 'range') => {
      if (selected) {
        ctx.fillStyle = selected === 'line' ? palette.selection : palette.headerSelectedBg;
        this.fill(b);
      }
      ctx.fillStyle = selected === 'line' ? palette.headerSelectedText : color;
      ctx.fillText(text, b.x + b.width / 2, b.y + (b.height - ascent - descent) / 2 + ascent);
    };
    ctx.save();
    ctx.font = font;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'alphabetic';
    ctx.fillStyle = palette.headerBg;
    this.fill(layout.box(0, 0, width, letters));
    this.fill(layout.box(0, letters, numbers, height - letters));

    const strips = [
      [layout.frozenColRun, numbers, layout.bodyX],
      [layout.colRun, layout.bodyX, width],
    ] as const;
    for (const [run, start, end] of strips) {
      if (run.first > run.last || end <= start) continue;
      ctx.save();
      this.clip(layout.box(start, 0, end - start, letters));
      for (let gc = run.first; gc <= run.last; gc++) {
        const b = layout.colHeading(gc);
        label(colToLetters(view.visibleCols[gc]), b, palette.headerText, state('cols', view.visibleCols[gc]));
        ctx.fillStyle = palette.line;
        this.fill(layout.box(layout.x(gc), 0, 1, letters));
      }
      ctx.restore();
    }
    const bands = [
      [layout.frozenRowRun, letters, layout.bodyY],
      [layout.rowRun, layout.bodyY, height],
    ] as const;
    for (const [run, top, bottom] of bands) {
      if (run.first > run.last || bottom <= top) continue;
      ctx.save();
      this.clip(layout.box(0, top, numbers, bottom - top));
      for (let gr = run.first; gr <= run.last; gr++) {
        const b = layout.rowHeading(gr);
        label(String(view.visibleRows[gr] + 1), b, palette.rowHeaderText, state('rows', view.visibleRows[gr]));
        ctx.fillStyle = palette.line;
        this.fill({ ...b, height: 1 });
      }
      ctx.restore();
    }
    // The line under the letters, and the one after the numbers.
    ctx.fillStyle = palette.headerLine;
    this.fill(layout.box(0, letters - 1, width, 1));
    ctx.fillStyle = palette.line;
    this.fill(layout.box(numbers - 1, letters, 1, height - letters));
    ctx.restore();
  }
}
