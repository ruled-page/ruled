import {
  AllCellRenderers,
  CompactSelection,
  DataEditorCore,
  GridCellKind,
  ImageWindowLoaderImpl,
  getMiddleCenterBias,
  sprites,
  type DataEditorProps,
  type DataEditorRef,
  type DrawCellCallback,
  type DrawHeaderCallback,
  type FillPatternEventArgs,
  type GridCell,
  type GridColumn,
  type GridKeyEventArgs,
  type GridMouseEventArgs,
  type GridSelection,
  type CellClickedEventArgs,
  type HeaderClickedEventArgs,
  type Item,
  type Rectangle,
  type TextCell,
  type Theme,
} from '@glideapps/glide-data-grid';
import '@glideapps/glide-data-grid/dist/index.css';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type MouseEvent as ReactMouseEvent, type RefObject } from 'react';
import { colToLetters, inRect } from '../core/refs';
import { CellFlag, VT, type CellNote, type DisplayStyle, type Rect } from '../core/types';
import { clipboardNumbers, parseClipboardText } from '../state/clipboard';
import { effectiveStyle } from '../state/cellStyle';
import { alignment, displayText, editText } from '../state/format';
import { CELL_PADDING, excelTextWidth, fitNumber, overflowTextX, rowOverflows, type Overflow } from '../state/overflow';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { EnumEditor, enumOptionsOf, type EnumCell } from './EnumEditor';
import { NoteTip } from './NoteTip';
import { ObjectsOverlay } from './ObjectsOverlay';
import { GLIDE_CLIP_REGION, GLIDE_SCROLLER } from './glideDom';
import { DEFAULT_PALETTE, glideTheme, readGridPalette, type GridPalette } from './gridPalette';
import { drawDataBar, drawEdge, drawGridLine, drawIcon } from './paint';
import { useT } from '../i18n/react';

/**
 * Glide's cell renderers, with text cells read to screen readers as they are
 * shown (6) rather than as they are edited (=B2*C2), which the formula bar
 * gives instead, followed by any problem with them.
 */
const RENDERERS = AllCellRenderers.map((r) =>
  r.kind === GridCellKind.Text ? { ...r, getAccessibilityString: (c: SpokenCell) => c.spokenText ?? c.displayData } : r,
) as typeof AllCellRenderers;

/**
 * A cell with what a screen reader should say for it, when that is more than
 * its value: its problems too, which the grid only marks in colour.
 */
type SpokenCell = TextCell & { spokenText?: string };

/** Everything the grid draws with, from the theme's palette. */
function gridStyles(p: GridPalette) {
  const fontFor = (style: DisplayStyle) => `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${style.fontSizePx}px "${style.fontFamily}", ${p.font}`;
  let measureCtx: CanvasRenderingContext2D | null = null;
  const measure = (text: string, style: DisplayStyle) => {
    measureCtx ??= document.createElement('canvas').getContext('2d');
    if (!measureCtx) return 0;
    measureCtx.font = fontFor(style);
    return measureCtx.measureText(text).width;
  };
  const cache = new Map<string, Partial<Theme>>();
  /** A cell's theme override: its fill (or the edited or locked fill), text colour and font. */
  const cellTheme = (style: DisplayStyle, color: string | undefined, edited: boolean, locked = false): Partial<Theme> | undefined => {
    // A changed cell shows its own fill, if it has one, and is marked in its corner instead.
    const fill = edited && !style.fill ? p.editedFill : (style.fill ?? (locked ? p.lockedFill : undefined));
    // Text with no colour of its own on a fill from the file is drawn as Excel draws it, whatever the theme's text colour.
    const textColor = color ?? style.color ?? (style.fill && !edited ? p.textOnFill : undefined);
    const plain = !fill && !textColor && !style.bold && !style.italic && Math.abs(style.fontSizePx - 14.7) < 0.2 && /^Calibri$/i.test(style.fontFamily);
    if (plain) return undefined;
    const key = `${fill}|${textColor}|${style.bold}|${style.italic}|${style.fontSizePx}|${style.fontFamily}`;
    let t = cache.get(key);
    if (!t) {
      // Glide spreads overrides over the base theme, so unset keys must be absent, not undefined.
      t = {
        baseFontStyle: `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${style.fontSizePx}px`,
        fontFamily: `"${style.fontFamily}", ${p.font}`,
      };
      if (fill) t.bgCell = fill;
      if (textColor) t.textDark = textColor;
      cache.set(key, t);
    }
    return t;
  };
  /** Strikethrough and underline, which Glide doesn't draw. */
  const decorate = (ctx: CanvasRenderingContext2D, rect: Rectangle, text: string, align: 'left' | 'right' | 'center', style: DisplayStyle) => {
    const width = measure(text, style);
    const x = align === 'right' ? rect.x + rect.width - CELL_PADDING - width : align === 'center' ? rect.x + (rect.width - width) / 2 : rect.x + CELL_PADDING;
    const mid = rect.y + rect.height / 2;
    ctx.save();
    ctx.fillStyle = style.color ?? p.text;
    const thickness = style.bold ? 1.5 : 1;
    if (style.strike) ctx.fillRect(x, mid, width, thickness);
    if (style.underline) ctx.fillRect(x, mid + style.fontSizePx * 0.45, width, thickness);
    ctx.restore();
  };
  return {
    theme: glideTheme(p),
    rowHeader: { bgCell: p.headerBg, textDark: p.rowHeaderText, baseFontStyle: p.headerFontSize, fontFamily: p.headerFont } as Partial<Theme>,
    fontFor,
    measure,
    cellTheme,
    decorate,
  };
}

interface Props {
  store: WorkbookStore;
  view: SheetView;
  /** Changes whenever cell data or styles change, so the grid redraws. */
  dataVersion: number;
  /** Changes whenever schema issues change, so markers redraw. */
  schemaVersion: number;
  selection: Selection;
  /** Changes whenever a column is resized; shared by both panes. */
  colVersion: number;
  onColumnResized: () => void;
  /**
   * The rows this grid shows, as grid rows of the sheet view: frozen rows are
   * a separate pane above the scrolling one. Defaults to every row.
   */
  rowStart?: number;
  rowCount?: number;
  /** Whether this pane shows the column letters. */
  header?: boolean;
  gridRef?: RefObject<DataEditorRef | null>;
  onKeyDown?: (e: GridKeyEventArgs) => void;
  className?: string;
  style?: React.CSSProperties;
}

type Highlights = NonNullable<DataEditorProps['highlightRegions']>[number][];

export function SheetGrid({
  store,
  view,
  dataVersion,
  schemaVersion,
  selection,
  colVersion,
  onColumnResized,
  rowStart = 0,
  rowCount,
  header = true,
  gridRef,
  onKeyDown,
  className,
  style: hostStyle,
}: Props) {
  const t = useT();
  const ownRef = useRef<DataEditorRef>(null);
  const ref = gridRef ?? ownRef;
  const [images] = useState(() => new ImageWindowLoaderImpl());
  const host = useRef<HTMLDivElement>(null);
  const rows = rowCount ?? view.rowCount - rowStart;
  // The grid's colours and fonts come from the theme's custom properties, read once laid out.
  const [palette, setPalette] = useState<GridPalette>(DEFAULT_PALETTE);
  const themeVersion = store.themeVersion;
  useLayoutEffect(() => {
    const next = readGridPalette(host.current);
    setPalette((prev) => (JSON.stringify(prev) === JSON.stringify(next) ? prev : next));
  }, [themeVersion]);
  const g = useMemo(() => gridStyles(palette), [palette]);
  const [region, setRegion] = useState(0);
  const [hoverNote, setHoverNote] = useState<{ note: CellNote; x: number; y: number } | null>(null);
  const styles = store.styles;
  const date1904 = store.date1904;
  const locale = store.locale;
  const highlight = store.highlightEdits;
  const numbers = store.numbers;
  const schema = store.schema;
  const sheetIndex = view.info.index;
  const rowHeaderWidth = Math.max(40, String(view.visibleRows[view.rowCount - 1] + 1).length * 8 + 16);

  const layoutVersion = view.layoutVersion;
  const frozen = useMemo(() => {
    let n = 0;
    for (let c = 0; c < view.info.frozenCols; c++) if (view.gridCol(c) >= 0) n++;
    return n;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [view, layoutVersion]);

  const columns = useMemo<GridColumn[]>(
    () => [
      { title: '', id: '#', width: rowHeaderWidth },
      ...Array.from(view.visibleCols, (c) => ({ title: colToLetters(c), id: String(c), width: view.colWidth(c) })),
    ],
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, view.colCount, colVersion, rowHeaderWidth],
  );

  const getCellContent = useCallback(
    ([gc, gr]: Item): GridCell => {
      const row = view.visibleRows[gr + rowStart];
      if (gc === 0) {
        return {
          kind: GridCellKind.Text,
          data: '',
          displayData: String(row + 1),
          allowOverlay: false,
          readonly: true,
          contentAlign: 'center',
          themeOverride: g.rowHeader,
          cursor: 'default',
        };
      }
      let col = view.visibleCols[gc - 1];
      let span: [number, number] | undefined;
      let showContent = true;
      const merge = view.mergeAt(row, col);
      if (merge) {
        let start = view.gridCol(merge.c1) + 1;
        let end = view.gridCol(merge.c2) + 1;
        for (let c = merge.c2; end <= 0 && c >= merge.c1; c--) end = view.gridCol(c) + 1;
        // Glide clips a span's text to its frozen part, so a merge that
        // crosses the freeze line is split there and its text drawn in the
        // scrolling part.
        const firstScrolling = 1 + frozen;
        if (start > 0 && start < firstScrolling && end >= firstScrolling) {
          if (gc < firstScrolling) {
            end = firstScrolling - 1;
            showContent = false;
          } else {
            start = firstScrolling;
          }
        }
        if (start > 0 && end > start) span = [start, end];
        showContent = showContent && row === merge.r1;
        col = merge.c1;
      }
      const rec = showContent ? view.get(merge ? merge.r1 : row, col) : null;
      const style = rec ? effectiveStyle(view, styles, merge ? merge.r1 : row, col, rec.style) : styles[0];
      const readOnly = view.readOnlyReason(row, view.visibleCols[gc - 1]);
      const blocked = view.blockedAt(row, view.visibleCols[gc - 1]);
      const locked = blocked !== undefined && blocked.kind !== 'tableHeader';
      const label = !!schema?.isLabel(sheetIndex, row, view.visibleCols[gc - 1]);
      // Enum fields edit with a dropdown; clicking the selected cell opens it.
      const field = schema && !readOnly ? schema.fieldAt(sheetIndex, row, view.visibleCols[gc - 1])?.field : undefined;
      let options: readonly string[] | undefined = field?.type === 'enum' && field.values?.length ? field.values : undefined;
      if (!options && !readOnly && view.info.listValidations.length) {
        // Excel's own list validations get the same dropdown.
        const dv = store.listAt(sheetIndex, row, view.visibleCols[gc - 1]);
        if (dv?.inCellDropdown) options = store.listOptions(dv, row, view.visibleCols[gc - 1]);
        if (!options?.length) options = undefined;
      }
      const issues = schema?.cellIssues(sheetIndex, row, view.visibleCols[gc - 1]) ?? [];
      const spokenNote = issues.length
        ? issues.map((i) => t(i.severity === 'error' ? 'grid.spokenError' : 'grid.spokenWarning', { message: t.text(i.message) })).join(' ')
        : undefined;
      const asEnum = (cell: EnumCell): EnumCell & SpokenCell => {
        const spoken = spokenNote ? { ...cell, spokenText: t('grid.spokenCell', { value: cell.displayData || t('grid.empty'), note: spokenNote }) } : cell;
        return options ? { ...spoken, enumOptions: options, activationBehaviorOverride: 'second-click' } : spoken;
      };
      if (!rec) {
        const baseStyle = merge ? effectiveStyle(view, styles, merge.r1, merge.c1) : effectiveStyle(view, styles, row, col, view.styleAt(row, col, null));
        const bg = baseStyle?.fill ?? (locked ? palette.lockedFill : label ? palette.labelFill : undefined);
        return asEnum({
          kind: GridCellKind.Text,
          data: '',
          displayData: '',
          allowOverlay: !readOnly,
          span,
          themeOverride: bg ? { bgCell: bg } : undefined,
        });
      }
      let { text, color } = displayText(rec, style, date1904, numbers.display === 'full');
      let drawStyle = style;
      if (rec.t === VT.Number && !style.wrap) {
        // Numbers never spill into neighbours; Excel shortens them or shows ####.
        let room = 0;
        for (let c = merge?.c1 ?? col; c <= (merge?.c2 ?? col); c++) room += view.colWidth(c);
        room -= CELL_PADDING * 2;
        if (text.length * style.fontSizePx * 0.62 > room) {
          // Whether it fits is Excel's call, made with Calibri's widths; where
          // the browser's stand-in font is wider, the number is drawn smaller.
          text = fitNumber(text, rec.n ?? 0, /^general$/i.test(style.numFmt), room, (s) => excelTextWidth(s, style));
          const drawn = text.startsWith('#') ? 0 : g.measure(text, style);
          if (drawn > room) drawStyle = { ...style, fontSizePx: Math.max(6, Math.floor(((style.fontSizePx * room) / drawn) * 10) / 10) };
        }
      }
      return asEnum({
        kind: GridCellKind.Text,
        data: editText(rec, style, date1904, locale),
        displayData: text,
        allowOverlay: !readOnly,
        contentAlign: alignment(rec, style),
        allowWrapping: style.wrap,
        span,
        themeOverride:
          label && !style.fill
            ? { ...g.cellTheme(drawStyle, color, false, false), bgCell: palette.labelFill }
            : g.cellTheme(drawStyle, color, highlight && !!(rec.flags & CellFlag.Edited), locked),
      });
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, dataVersion, styles, date1904, locale, highlight, frozen, numbers, schema, schemaVersion, rowStart, g, palette, t],
  );

  // Text spilling into empty neighbours, worked out per row as rows are drawn.
  const overflowCache = useMemo(
    () => new Map<number, Overflow[]>(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, dataVersion, styles, date1904, colVersion],
  );
  const overflowsFor = useCallback(
    (row: number) => {
      let list = overflowCache.get(row);
      if (!list) {
        list = rowOverflows(view, row, (col, rec) => effectiveStyle(view, styles, row, col, rec.style), date1904, g.measure);
        overflowCache.set(row, list);
      }
      return list;
    },
    [overflowCache, view, styles, date1904, g],
  );

  /** Pixels from the left edge of sheet column `from` to the left edge of `to`, over visible columns. */
  const offsetBetween = useCallback(
    (from: number, to: number) => {
      const a = view.gridCol(from);
      const b = view.gridCol(to);
      let px = 0;
      for (let g = Math.min(a, b); g < Math.max(a, b); g++) px += view.colWidth(view.visibleCols[g]);
      return a <= b ? px : -px;
    },
    [view],
  );

  const drawCell = useCallback<DrawCellCallback>(
    (args, drawContent) => {
      const { ctx, rect } = args;
      const row = view.visibleRows[args.row + rowStart];
      const gridLines = view.info.showGridLines;
      if (args.col === 0) {
        drawContent();
        drawGridLine(ctx, 'top', rect.x, rect.y, rect.width, palette.line);
        drawGridLine(ctx, 'left', rect.x + rect.width - 1, rect.y, rect.height, palette.line);
        return;
      }
      const col = view.visibleCols[args.col - 1];

      // Overflowing text is drawn here in full, in its own cell and in each neighbour it covers.
      const merge = view.mergeAt(row, col);
      const style = effectiveStyle(view, styles, merge?.r1 ?? row, merge?.c1 ?? col);
      const overflows = overflowsFor(row);
      const own = overflows.find((o) => o.col === col);
      const through = own ?? overflows.find((o) => col >= o.from && col <= o.to);

      // Conditional formatting's bars and icons sit behind the text.
      const cf = view.cfAt(merge?.r1 ?? row, merge?.c1 ?? col);
      if (cf?.bar) drawDataBar(ctx, rect, cf.bar);
      if (cf?.icon) drawIcon(ctx, cf.icon.set, cf.icon.index, rect.x + 3, rect.y + (rect.height - 14) / 2);
      if (!own && !cf?.hideValue) {
        drawContent();
        if ((style.strike || style.underline) && args.cell.kind === GridCellKind.Text && args.cell.displayData) {
          g.decorate(ctx, rect, args.cell.displayData, args.cell.contentAlign ?? 'left', style);
        }
      }
      if (through) {
        const font = g.fontFor(through.style);
        ctx.save();
        ctx.font = font;
        ctx.fillStyle = through.color ?? through.style.color ?? (through.style.fill ? palette.textOnFill : palette.text);
        ctx.textAlign = 'left';
        const x = rect.x - offsetBetween(through.col, col) + overflowTextX(through, view.colWidth(through.col));
        ctx.fillText(through.text, x, rect.y + rect.height / 2 + getMiddleCenterBias(ctx, font));
        ctx.restore();
      }

      // Left and top edges belong to this cell: a border from either side, else a grid line.
      const gc = view.gridCol(col);
      const gr = view.gridRow(row);
      if (!merge || col === merge.c1) {
        const leftCol = gc > 0 ? view.visibleCols[gc - 1] : -1;
        const edge = style.border?.left ?? (leftCol >= 0 ? effectiveStyle(view, styles, row, leftCol).border?.right : undefined);
        if (edge) drawEdge(ctx, 'left', rect.x, rect.y, rect.height, edge, palette.borderAuto);
        else if (gridLines && !overflows.some((o) => o.from < col && col <= o.to)) drawGridLine(ctx, 'left', rect.x, rect.y, rect.height, palette.line);
      }
      if (!merge || row === merge.r1) {
        const upRow = gr > 0 ? view.visibleRows[gr - 1] : -1;
        const edge = style.border?.top ?? (upRow >= 0 ? effectiveStyle(view, styles, upRow, col).border?.bottom : undefined);
        if (edge) drawEdge(ctx, 'top', rect.x, rect.y, rect.width, edge, palette.borderAuto);
        else if (gridLines) drawGridLine(ctx, 'top', rect.x, rect.y, rect.width, palette.line);
      }
      if (frozen > 0 && args.col === frozen) drawGridLine(ctx, 'left', rect.x + rect.width - 1, rect.y, rect.height, palette.freezeLine);

      if (args.highlighted && enumOptionsOf(args.cell)) {
        // Dropdown chevron on selected enum cells, like Excel's list validation.
        const cx = rect.x + rect.width - 10;
        const cy = rect.y + rect.height / 2 + 1;
        ctx.save();
        ctx.beginPath();
        ctx.moveTo(cx - 3.5, cy - 2);
        ctx.lineTo(cx, cy + 1.5);
        ctx.lineTo(cx + 3.5, cy - 2);
        ctx.strokeStyle = palette.dropdownArrow;
        ctx.lineWidth = 1.5;
        ctx.stroke();
        ctx.restore();
      }
      const rec = view.get(row, col);
      if (rec && store.highlightEdits && rec.flags & CellFlag.Edited && styles[rec.style]?.fill) {
        const bottom = rect.y + rect.height - 1;
        ctx.save();
        ctx.beginPath();
        ctx.moveTo(rect.x + 1, bottom);
        ctx.lineTo(rect.x + 7, bottom);
        ctx.lineTo(rect.x + 1, bottom - 6);
        ctx.closePath();
        ctx.fillStyle = palette.editedMark;
        ctx.fill();
        ctx.restore();
      }
      if (rec && rec.flags & (CellFlag.Unsupported | CellFlag.Mismatch | CellFlag.Stale)) {
        ctx.save();
        ctx.beginPath();
        ctx.moveTo(rect.x + 1, rect.y + 1);
        ctx.lineTo(rect.x + 8, rect.y + 1);
        ctx.lineTo(rect.x + 1, rect.y + 8);
        ctx.closePath();
        ctx.fillStyle = rec.flags & CellFlag.Stale ? palette.staleMark : palette.unsupportedMark;
        ctx.fill();
        ctx.restore();
      }
      if (view.noteAt(row, col)) {
        const right = rect.x + rect.width - 1;
        ctx.save();
        ctx.beginPath();
        ctx.moveTo(right, rect.y + 1);
        ctx.lineTo(right - 6, rect.y + 1);
        ctx.lineTo(right, rect.y + 7);
        ctx.closePath();
        ctx.fillStyle = palette.noteMark;
        ctx.fill();
        ctx.restore();
      }
      if (!schema) return;
      const issues = schema.cellIssues(sheetIndex, row, col);
      const right = rect.x + rect.width - 1;
      ctx.save();
      if (issues.length) {
        ctx.beginPath();
        ctx.moveTo(right, rect.y + 1);
        ctx.lineTo(right - 8, rect.y + 1);
        ctx.lineTo(right, rect.y + 9);
        ctx.closePath();
        ctx.fillStyle = issues.some((i) => i.severity === 'error') ? palette.errorMark : palette.warningMark;
        ctx.fill();
      } else if (schema.isPending(sheetIndex, row, col)) {
        ctx.beginPath();
        ctx.arc(right - 4, rect.y + 5, 2.5, 0, Math.PI * 2);
        ctx.fillStyle = palette.pendingMark;
        ctx.fill();
      }
      ctx.restore();
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, schema, schemaVersion, sheetIndex, styles, overflowsFor, offsetBetween, frozen, rowStart, g, palette],
  );

  const drawHeader = useCallback<DrawHeaderCallback>(
    (args, drawContent) => {
      drawContent();
    if (args.columnIndex > 0) drawGridLine(args.ctx, 'left', args.rect.x, args.rect.y, args.rect.height, palette.line);
    },
    [palette],
  );

  // Grid selection <-> sheet selection. Grid column 0 is the row-number column.
  const toSheetRect = useCallback(
    (r: Rectangle): Rect => {
      const x1 = Math.max(1, r.x);
      const x2 = Math.max(x1, r.x + r.width - 1);
      return {
        r1: view.visibleRows[r.y + rowStart],
        r2: view.visibleRows[Math.min(rowStart + rows - 1, rowStart + r.y + r.height - 1)],
        c1: view.visibleCols[x1 - 1],
        c2: view.visibleCols[Math.min(view.colCount, x2) - 1],
      };
    },
    [view, rowStart, rows],
  );

  const cutRange = store.cutRange;
  // The selection, clipped to this pane. When the active cell is in the other
  // pane, this one shows its part of the selection as a highlight instead.
  const { gridSelection, highlights } = useMemo(() => {
    const last = rowStart + rows - 1;
    const at = (row: number) => Math.max(0, view.gridRow(row));
    const clip = (r: Rect): Rectangle | undefined => {
      const y1 = Math.max(rowStart, at(r.r1));
      const y2 = Math.min(last, Math.max(at(r.r1), at(r.r2)));
      if (y1 > y2) return undefined;
      const x1 = Math.max(0, view.gridCol(r.c1)) + 1;
      const x2 = Math.max(x1, view.gridCol(r.c2) + 1);
      return { x: x1, y: y1 - rowStart, width: x2 - x1 + 1, height: y2 - y1 + 1 };
    };
    const rects = selection.rects.map(clip).filter((r): r is Rectangle => !!r);
    const gr = at(selection.row);
    const gc = Math.max(0, view.gridCol(selection.col)) + 1;
    const empty = { columns: CompactSelection.empty(), rows: CompactSelection.empty() };
    // A cut waiting to be pasted gets a dashed outline.
    const cut = cutRange && cutRange.sheet === view.info.index ? clip(cutRange.rect) : undefined;
    const marks: Highlights = cut ? [{ color: palette.cutFill, range: cut, style: 'dashed' }] : [];
    if (gr < rowStart || gr > last) {
      marks.push(...rects.map((range) => ({ color: palette.selectionFill, range, style: 'solid-outline' as const })));
      return { gridSelection: { current: undefined, ...empty } as GridSelection, highlights: marks.length ? marks : undefined };
    }
    const cell = { x: gc, y: gr - rowStart, width: 1, height: 1 };
    const [first, ...rest] = rects;
    return {
      gridSelection: { current: { cell: [gc, gr - rowStart], range: first ?? cell, rangeStack: rest }, ...empty } as GridSelection,
      highlights: marks.length ? marks : undefined,
    };
  }, [view, selection, rowStart, rows, cutRange, palette]);

  const onSelectionChange = useCallback(
    (sel: GridSelection) => {
      if (sel.columns.length > 0) {
        const cols = sel.columns.toArray();
        const c1 = view.visibleCols[Math.max(1, Math.min(...cols)) - 1];
        const c2 = view.visibleCols[Math.max(1, Math.max(...cols)) - 1];
        const r2 = Math.max(view.info.maxRow, 0);
        store.setSelection({ row: 0, col: c1, rects: [{ r1: 0, c1, r2, c2 }] });
        return;
      }
      if (!sel.current) return;
      const [gc, gr] = sel.current.cell;
      if (gc === 0) {
        // Clicking a row number selects the whole row.
        const picked = [sel.current.range.y, sel.current.range.y + sel.current.range.height - 1].map((y) => view.visibleRows[y + rowStart]);
        store.setSelection({
          row: picked[0],
          col: view.visibleCols[0],
          rects: [{ r1: picked[0], r2: picked[1], c1: 0, c2: Math.max(view.info.maxCol, 0) }],
        });
        return;
      }
      store.setSelection({
        row: view.visibleRows[gr + rowStart],
        col: view.visibleCols[gc - 1],
        rects: [sel.current.range, ...sel.current.rangeStack].map(toSheetRect),
      });
    },
    [store, view, toSheetRect, rowStart],
  );

  // The editor writes the clipboard itself, the text and an escaped HTML
  // table, rather than letting the grid write cell text into HTML unescaped.
  // It remembers what was copied, so pasting it back carries formulas. A cut
  // only marks the cells: like Excel, pasting it moves them, and references follow.
  useEffect(() => {
    const inEditorField = () => {
      const active = document.activeElement;
      return !!active && /^(INPUT|TEXTAREA)$/.test(active.tagName) && !active.closest(`.${GLIDE_CLIP_REGION}, .${GLIDE_SCROLLER}`);
    };
    const write = (e: ClipboardEvent, cut: boolean) => {
      if (inEditorField() || !host.current?.contains(document.activeElement)) return;
      const first = store.selection.rects[0];
      if (!first || !e.clipboardData) return;
      e.preventDefault();
      e.clipboardData.setData('text/plain', store.selectionText(first));
      e.clipboardData.setData('text/html', store.selectionHtml(first));
      store.noteCopy(first, cut);
    };
    const onCopy = (e: ClipboardEvent) => write(e, false);
    const onCut = (e: ClipboardEvent) => write(e, true);
    const onKey = (e: KeyboardEvent) => {
      if (!host.current?.contains(document.activeElement)) return;
      if (e.key === 'Escape' && store.cutRange) store.cancelCut();
      // The menu key or Shift+F10 opens the right-click menu at the active cell.
      if (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10')) {
        e.preventDefault();
        const sel = store.selection;
        const v = store.view;
        const at = v && ref.current?.getBounds(v.gridCol(sel.col) + 1, v.gridRow(sel.row) - rowStart);
        if (at) store.set('menu', { x: at.x + Math.min(at.width, 24), y: at.y + at.height, kind: 'cell' });
      }
    };
    window.addEventListener('copy', onCopy, true);
    window.addEventListener('cut', onCut, true);
    window.addEventListener('keydown', onKey);
    return () => {
      window.removeEventListener('copy', onCopy, true);
      window.removeEventListener('cut', onCut, true);
      window.removeEventListener('keydown', onKey);
    };
  }, [store, ref, rowStart]);

  // Scroll to cells that find, the name box or the change list point at.
  // Glide itself scrolls a grid that is showing to a cell selected from
  // outside it, so this only scrolls when Glide won't: to the cell already
  // selected, or in a grid that has just appeared for another sheet, once it
  // has a size. Scrolling as well as Glide would scroll twice as far.
  const reveal = store.revealSeq;
  const shown = useRef<{ sheet: number; cell: string }>(undefined);
  useEffect(() => {
    if (!reveal) return;
    const gr = view.gridRow(selection.row) - rowStart;
    const gc = view.gridCol(selection.col) + 1;
    if (gr >= 0 && gr < rows && gc > 0) {
      const prev = shown.current;
      if (prev?.sheet !== view.info.index || prev.cell === `${gc}:${gr}`) {
        let tries = 20;
        const scroll = () => {
          const b = ref.current?.getBounds(gc, gr);
          if (b?.width && b.height) ref.current?.scrollTo(gc, gr, 'both', 40, 40);
          else if (--tries > 0) requestAnimationFrame(scroll);
        };
        scroll();
      }
      if (store.revealFocus) {
        store.revealFocus = false;
        // Once React has handed the grid its new active cell. A grid that has
        // just appeared, on another sheet, may not take the focus at first in
        // Safari, so it is tried again for a few frames.
        let tries = 10;
        const focus = () => {
          ref.current?.focus();
          if (!host.current?.contains(document.activeElement) && --tries > 0) requestAnimationFrame(focus);
        };
        setTimeout(focus, 0);
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reveal]);
  // After the effect above, so that it sees the cell selected before.
  useEffect(() => {
    shown.current = { sheet: view.info.index, cell: `${view.gridCol(selection.col) + 1}:${view.gridRow(selection.row) - rowStart}` };
  }, [view, selection, rowStart]);

  const onCellContextMenu = useCallback(
    ([gc, gr]: Item, e: CellClickedEventArgs) => {
      e.preventDefault();
      const row = view.visibleRows[gr + rowStart];
      const col = gc === 0 ? view.visibleCols[0] : view.visibleCols[gc - 1];
      const sel = store.selection;
      const inside = sel.rects.some((r) => inRect(r, row, gc === 0 ? r.c1 : col));
      // Right-clicking outside the selection selects what was clicked first, as Excel does.
      if (!inside) {
        const rect = gc === 0 ? { r1: row, r2: row, c1: 0, c2: Math.max(view.info.maxCol, 0) } : { r1: row, r2: row, c1: col, c2: col };
        store.setSelection({ row, col, rects: [rect] });
      }
      store.set('menu', { x: e.bounds.x + e.localEventX, y: e.bounds.y + e.localEventY, kind: gc === 0 ? 'row' : 'cell' });
    },
    [store, view, rowStart],
  );

  const onHeaderContextMenu = useCallback(
    (gc: number, e: HeaderClickedEventArgs) => {
      e.preventDefault();
      if (gc === 0) return;
      const col = view.visibleCols[gc - 1];
      const sel = store.selection;
      if (!sel.rects.some((r) => col >= r.c1 && col <= r.c2 && r.r1 === 0)) {
        store.setSelection({ row: 0, col, rects: [{ r1: 0, r2: Math.max(view.info.maxRow, 0), c1: col, c2: col }] });
      }
      store.set('menu', { x: e.bounds.x + e.localEventX, y: e.bounds.y + e.localEventY, kind: 'column' });
    },
    [store, view],
  );

  // Glide reads pastes through navigator.clipboard.read(), which makes
  // browsers ask for clipboard permission. The paste event already carries
  // the data, so handle it here instead.
  useEffect(() => {
    const onPaste = (e: ClipboardEvent) => {
      if (!host.current?.contains(document.activeElement)) return;
      const text = e.clipboardData?.getData('text/plain');
      if (!text) return;
      e.preventDefault();
      const values = parseClipboardText(text);
      // Excel's HTML has the exact values of the numbers its text shows rounded.
      const html = e.clipboardData?.getData('text/html');
      const exact = html ? clipboardNumbers(html) : undefined;
      const sel = store.selection;
      const rect = sel.rects[0] ?? { r1: sel.row, c1: sel.col, r2: sel.row, c2: sel.col };
      void store.paste({ row: rect.r1, col: rect.c1 }, values, rect, false, exact);
    };
    window.addEventListener('paste', onPaste, true);
    return () => window.removeEventListener('paste', onPaste, true);
  }, [store]);

  const onDelete = useCallback(
    (sel: GridSelection) => {
      const rects: Rect[] = [];
      if (sel.columns.length > 0) {
        const cols = sel.columns.toArray().filter((c) => c > 0);
        if (cols.length) {
          rects.push({ r1: 0, r2: Math.max(view.info.maxRow, 0), c1: view.visibleCols[Math.min(...cols) - 1], c2: view.visibleCols[Math.max(...cols) - 1] });
        }
      } else if (sel.current) {
        rects.push(...[sel.current.range, ...sel.current.rangeStack].map(toSheetRect));
      }
      if (rects.length) void store.clear(rects);
      return false;
    },
    [store, view, toSheetRect],
  );

  const onFillPattern = useCallback(
    (e: FillPatternEventArgs) => {
      e.preventDefault();
      void store.fill(toSheetRect(e.patternSource), toSheetRect(e.fillDestination));
    },
    [store, toSheetRect],
  );

  const onCellEdited = useCallback(
    ([gc, gr]: Item, value: { data?: unknown }) => {
      if (gc === 0) return;
      void store.input(view.visibleRows[gr + rowStart], view.visibleCols[gc - 1], String(value.data ?? ''));
    },
    [store, view, rowStart],
  );

  const onCellActivated = useCallback(
    ([gc, gr]: Item) => {
      if (gc === 0) return;
      const reason = view.readOnlyReason(view.visibleRows[gr + rowStart], view.visibleCols[gc - 1]);
      if (reason) store.toast('warn', reason);
    },
    [store, view, rowStart],
  );

  // Columns keep the width dragged to, as in Excel, except on a protected sheet.
  const resizable = view.info.kind === 'worksheet' && !view.info.protected;
  const onColumnResize = useCallback(
    (_c: GridColumn, size: number, index: number) => {
      if (index === 0 || !resizable) return;
      view.colWidths.set(view.visibleCols[index - 1], size);
      onColumnResized();
    },
    [view, onColumnResized, resizable],
  );
  // Rows are resized by dragging the bottom edge of their number, which Glide
  // doesn't offer: a press there is taken before Glide sees it.
  const visible = useRef({ y: 0, height: 0 });
  const [rowDrag, setRowDrag] = useState<{ row: number; from: number; px: number; y: number } | null>(null);
  const rowEdgeAt = (x: number, y: number): { row: number; edge: number } | undefined => {
    const grid = ref.current;
    if (!grid || !resizable) return undefined;
    for (let gr = visible.current.y; gr <= visible.current.y + visible.current.height; gr++) {
      const b = grid.getBounds(0, gr);
      if (!b || x < b.x || x > b.x + b.width) return undefined;
      const edge = b.y + b.height;
      if (Math.abs(y - edge) <= 3 && gr + rowStart < view.rowCount) return { row: view.visibleRows[gr + rowStart], edge };
    }
    return undefined;
  };
  const onHostMouseMove = (e: ReactMouseEvent) => {
    if (!rowDrag && host.current) host.current.style.cursor = rowEdgeAt(e.clientX, e.clientY) ? 'row-resize' : '';
  };
  const onHostMouseDown = (e: ReactMouseEvent) => {
    const at = e.button === 0 ? rowEdgeAt(e.clientX, e.clientY) : undefined;
    if (!at) return;
    e.preventDefault();
    e.stopPropagation();
    const start = { row: at.row, from: e.clientY, px: view.rowHeight(at.row), y: at.edge };
    setRowDrag(start);
    const move = (m: MouseEvent) => setRowDrag({ ...start, y: start.y + Math.max(-start.px, m.clientY - start.from) });
    const up = (m: MouseEvent) => {
      window.removeEventListener('mousemove', move, true);
      window.removeEventListener('mouseup', up, true);
      // The click that ends the drag isn't one on the grid.
      window.addEventListener('click', (c) => c.stopPropagation(), { capture: true, once: true });
      setRowDrag(null);
      const px = Math.max(0, Math.round(start.px + m.clientY - start.from));
      if (px !== start.px) void store.resizedRow(start.row, px);
    };
    window.addEventListener('mousemove', move, true);
    window.addEventListener('mouseup', up, true);
  };

  const onColumnResizeEnd = useCallback(
    (_c: GridColumn, size: number, index: number) => {
      if (index > 0 && resizable) void store.resizedColumn(view.visibleCols[index - 1], size);
    },
    [store, view, resizable],
  );

  const onItemHovered = useCallback(
    (args: GridMouseEventArgs) => {
      const [gc, gr] = args.location;
      const note = args.kind === 'cell' && gc > 0 ? view.noteAt(view.visibleRows[gr + rowStart], view.visibleCols[gc - 1]) : undefined;
      if (!note || args.kind !== 'cell') {
        setHoverNote(null);
        return;
      }
      const { x, y, width } = args.bounds;
      setHoverNote((prev) => (prev?.note === note ? prev : { note, x: x + width + 4, y }));
    },
    [view, rowStart],
  );

  const rowHeight = useCallback(
    (gr: number) => view.rowHeight(view.visibleRows[gr + rowStart]),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, rowStart, layoutVersion],
  );
  const provideEditor = useCallback(
    (cell: GridCell) => (enumOptionsOf(cell) ? { editor: EnumEditor, disablePadding: true } : undefined),
    [],
  );

  return (
    <div className={`wbe-grid-host ${className ?? ''}`} ref={host} style={hostStyle} onMouseMoveCapture={onHostMouseMove} onMouseDownCapture={onHostMouseDown}>
      <DataEditorCore
        ref={ref}
        key={view.info.index}
        renderers={RENDERERS}
        headerIcons={sprites}
        imageWindowLoader={images}
        columns={columns}
        rows={rows}
        getCellContent={getCellContent}
        drawCell={drawCell}
        drawHeader={drawHeader}
        gridSelection={gridSelection}
        onGridSelectionChange={onSelectionChange}
        onCellEdited={onCellEdited}
        onCellActivated={onCellActivated}
        keybindings={{ copy: false, paste: false, cut: false }}
        onCellContextMenu={onCellContextMenu}
        onHeaderContextMenu={onHeaderContextMenu}
        provideEditor={provideEditor}
        onDelete={onDelete}
        fillHandle
        allowedFillDirections="orthogonal"
        onFillPattern={onFillPattern}
        onColumnResize={onColumnResize}
        onColumnResizeEnd={onColumnResizeEnd}
        getCellsForSelection
        rangeSelect="rect"
        columnSelect="multi"
        rowSelect="none"
        rowMarkers="none"
        freezeColumns={1 + frozen}
        rowHeight={rowHeight}
        headerHeight={header ? 24 : 0}
        highlightRegions={highlights}
        onKeyDown={onKeyDown}
        smoothScrollX
        smoothScrollY
        cellActivationBehavior="double-click"
        editOnType
        theme={g.theme}
        width="100%"
        height="100%"
        minColumnWidth={18}
        maxColumnWidth={1200}
        onVisibleRegionChanged={(r) => {
          visible.current = { y: r.y, height: r.height };
          setRegion((n) => n + 1);
        }}
        onItemHovered={onItemHovered}
      />
      {hoverNote && <NoteTip {...hoverNote} />}
      {rowDrag && host.current && <div className="wbe-row-guide" style={{ top: rowDrag.y - host.current.getBoundingClientRect().top }} />}
      <ObjectsOverlay store={store} view={view} gridRef={ref} tick={region + colVersion} rowStart={rowStart} rowCount={rows} top={header ? 24 : 0} />
    </div>
  );
}
