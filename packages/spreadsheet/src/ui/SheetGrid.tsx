import {
  AllCellRenderers,
  CompactSelection,
  DataEditorCore,
  GridCellKind,
  ImageWindowLoaderImpl,
  getMiddleCenterBias,
  sprites,
  withAlpha,
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
import { linkAddress } from '../core/links';
import { colToLetters, inRect, rangeToA1, toA1 } from '@sitepen/formula-engine/refs';
import { CellFlag, VT, type CellNote, type DisplayStyle, type Rect } from '../core/types';
import { clipboardNumbers, parseClipboardText } from '../state/clipboard';
import { clipboardFormats } from '../state/clipboardFormats';
import { effectiveStyle } from '../state/cellStyle';
import { alignment, displayText, editText } from '../state/format';
import { cellLink, type CellLink } from '../state/links';
import { CELL_PADDING, excelTextWidth, fitNumber, indentPx, overflowTextX, rowOverflows, textTop, verticalShift, type Overflow } from '../state/overflow';
import { wrapLines } from '../state/autofit';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { textCellEditor } from './CellEditor';
import { EnumEditor, enumOptionsOf, type EnumCell } from './EnumEditor';
import { noteSummary } from './FormulaBar';
import { APPLE } from './keys';
import { LinkTip } from './LinkTip';
import { NoteTip } from './NoteTip';
import { ObjectsOverlay } from './ObjectsOverlay';
import { GLIDE_CLIP_REGION, GLIDE_SCROLLER } from './glideDom';
import { DEFAULT_PALETTE, DEFAULT_REF_COLORS, glideTheme, readGridPalette, readRefColors, type GridPalette } from './gridPalette';
import { drawDataBar, drawEdge, drawFilterButton, drawGridLine, drawIcon, filterButtonBox } from './paint';
import { useT } from '../i18n/react';

/**
 * Glide's cell renderers, with text cells in the table for screen readers
 * as they are shown (6) rather than as they are edited (=B2*C2), which the
 * formula bar gives instead, and in full where the grid draws #### or fewer
 * digits.
 */
const RENDERERS = AllCellRenderers.map((r) =>
  r.kind === GridCellKind.Text ? { ...r, getAccessibilityString: (c: SpokenCell) => c.shownText ?? c.displayData } : r,
) as typeof AllCellRenderers;

/**
 * A cell with its text in full, and what a screen reader says for it: where
 * it is, what it shows, and what the grid only marks in colour, such as a
 * formula out of date, a note or a problem. That is worked out only when read.
 */
type SpokenCell = TextCell & { shownText?: string; spoken?: () => string };

/**
 * What Glide's table for screen readers is told (see vite.glide.ts): its
 * name, and the sheet's own row and column numbers for its grid rows and
 * columns, so that rows after hidden ones or below frozen ones are numbered
 * as the sheet numbers them.
 */
interface SheetAccess {
  label: string;
  rowCount: number;
  colCount: number;
  rowIndex: (gridRow: number) => number;
  colIndex: (gridCol: number) => number;
  /** What a cell is called: its reference, what it shows and what else is true of it. */
  cellLabel: (cell: GridCell) => string | undefined;
}

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
  /** The height Glide gives a line of text in a font, as it measures it. */
  const ems = new Map<string, number>();
  const emHeight = (style: DisplayStyle) => {
    const font = fontFor(style);
    let em = ems.get(font);
    if (em === undefined) {
      measureCtx ??= document.createElement('canvas').getContext('2d');
      if (!measureCtx) return style.fontSizePx;
      measureCtx.font = font;
      const m = measureCtx.measureText('ABCi09jgqpy');
      ems.set(font, (em = m.actualBoundingBoxAscent + m.actualBoundingBoxDescent));
    }
    return em;
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
  // Wrapped text's lines, kept while the same text is drawn at the same width.
  const wraps = new Map<string, string[]>();
  /**
   * Wrapped text, drawn line by line where Excel's vertical alignment puts it
   * and cut at the cell's edges. Glide's own wrapping stops at the bottom of
   * the cell it centres text in, so text moved up for bottom alignment would
   * lose its last lines.
   */
  const drawWrapped = (
    ctx: CanvasRenderingContext2D,
    rect: Rectangle,
    text: string,
    align: 'left' | 'right' | 'center',
    style: DisplayStyle,
    theme: Pick<Theme, 'lineHeight' | 'cellVerticalPadding' | 'textDark'>,
  ) => {
    const font = fontFor(style);
    const indent = indentPx(style, align);
    const width = rect.width - CELL_PADDING * 2 - indent;
    const key = `${font}\u0000${width}\u0000${text}`;
    let lines = wraps.get(key);
    if (!lines) {
      if (wraps.size > 2000) wraps.clear();
      lines = wrapLines(text, width, (s) => measure(s, style));
      wraps.set(key, lines);
    }
    const em = emHeight(style);
    const step = theme.lineHeight * em;
    const top = rect.y + textTop(style.vAlign, rect.height, em + step * (lines.length - 1), theme.cellVerticalPadding);
    ctx.save();
    ctx.beginPath();
    ctx.rect(rect.x, rect.y, rect.width, rect.height);
    ctx.clip();
    ctx.font = font;
    ctx.fillStyle = theme.textDark;
    ctx.textAlign = align;
    ctx.textBaseline = 'middle';
    const bias = getMiddleCenterBias(ctx, font);
    const x = align === 'right' ? rect.x + rect.width - CELL_PADDING - indent : align === 'center' ? rect.x + rect.width / 2 : rect.x + CELL_PADDING + indent;
    const thickness = style.bold ? 1.5 : 1;
    lines.forEach((line, i) => {
      const mid = top + i * step + em / 2;
      if (mid + em < rect.y || mid - em > rect.y + rect.height) return;
      ctx.fillText(line, x, mid + bias);
      if (!style.strike && !style.underline) return;
      const w = measure(line, style);
      const left = align === 'right' ? x - w : align === 'center' ? x - w / 2 : x;
      if (style.strike) ctx.fillRect(left, mid, w, thickness);
      if (style.underline) ctx.fillRect(left, mid + style.fontSizePx * 0.45, w, thickness);
    });
    ctx.restore();
  };
  return {
    theme: glideTheme(p),
    rowHeader: { bgCell: p.headerBg, textDark: p.rowHeaderText, baseFontStyle: p.headerFontSize, fontFamily: p.headerFont } as Partial<Theme>,
    fontFor,
    measure,
    emHeight,
    cellTheme,
    decorate,
    drawWrapped,
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
  /** Told the cells in view whenever they change, as the grid scrolls or resizes. */
  onRegion?: (r: Rectangle) => void;
  className?: string;
  style?: React.CSSProperties;
}

type Highlights = NonNullable<DataEditorProps['highlightRegions']>[number][];


/**
 * Stops the click that follows a mouse button being let go, so a press or
 * drag that ends on the grid isn't also taken as a click on a cell. If no
 * click follows (the button was let go over something else), nothing later
 * is stopped.
 */
function swallowNextClick(): void {
  const stop = (c: MouseEvent) => c.stopPropagation();
  window.addEventListener('click', stop, { capture: true, once: true });
  setTimeout(() => window.removeEventListener('click', stop, true), 0);
}
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
  onRegion,
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
  const [refColors, setRefColors] = useState(DEFAULT_REF_COLORS);
  const themeVersion = store.themeVersion;
  useLayoutEffect(() => {
    const next = readGridPalette(host.current);
    setPalette((prev) => (JSON.stringify(prev) === JSON.stringify(next) ? prev : next));
    const colors = readRefColors(host.current);
    setRefColors((prev) => (prev.join() === colors.join() ? prev : colors));
  }, [themeVersion]);
  const g = useMemo(() => gridStyles(palette), [palette]);
  // AutoFit measures text as it is drawn here.
  useEffect(() => {
    store.measureText = g.measure;
  }, [store, g]);
  const [region, setRegion] = useState(0);
  const [hoverNote, setHoverNote] = useState<{ note: CellNote; x: number; y: number } | null>(null);
  const [hoverLink, setHoverLink] = useState<{ link: CellLink; x: number; y: number } | null>(null);
  /** Where the selection being dragged by its border would land, and whether it would be copied. */
  const [cellDrag, setCellDrag] = useState<{ rect: Rect; copy: boolean } | null>(null);
  // Where Shift+dragged cells would go in between others, in the host's pixels.
  const [insertGuide, setInsertGuide] = useState<{ left: number; top: number; width: number; height: number } | null>(null);
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

  /**
   * What a screen reader says for a cell: its reference, what it shows, and
   * then what the grid only shows by colour or marks: a merge, a formula and
   * whether its value is out of date or Excel's, a change, a list, a note
   * or comments, a link and where it goes, and any problem with it.
   */
  const spokenText = useCallback(
    (row: number, col: number, merge: Rect | undefined, shown: string, list: boolean): string => {
      const at = merge ? { row: merge.r1, col: merge.c1 } : { row, col };
      const rec = view.get(at.row, at.col);
      // The cells a merge covers are read as the merge, whose top-left cell holds its value.
      if (merge && !shown && rec) shown = displayText(rec, effectiveStyle(view, styles, at.row, at.col, rec.style), date1904, numbers.display === 'full', locale).text;
      const parts = [toA1(at.row, at.col), shown || t('grid.empty')];
      const link = cellLink(view, at.row, at.col);
      if (link) {
        const address = linkAddress(link);
        parts[1] = address ? t('links.spoken', { value: parts[1], address }) : t('links.spokenUnknown', { value: parts[1] });
      }
      if (merge) parts.push(t('grid.flag.merged', { range: rangeToA1(merge) }));
      if (rec && rec.flags & CellFlag.Formula) parts.push(t('grid.flag.formula'));
      // Where a spilled value comes from, which the grid outlines only while it is selected.
      if (rec?.spill && rec.flags & CellFlag.Spilled) parts.push(t('grid.flag.spilled', { ref: toA1(rec.spill.r1, rec.spill.c1) }));
      if (rec && rec.flags & CellFlag.Stale) parts.push(t('grid.flag.stale'));
      else if (rec && rec.flags & (CellFlag.Unsupported | CellFlag.Mismatch)) parts.push(t('grid.flag.savedValue'));
      if (rec && highlight && rec.flags & CellFlag.Edited) parts.push(t('grid.flag.edited'));
      if (list) parts.push(t('grid.flag.list'));
      const note = view.noteAt(at.row, at.col);
      if (note) parts.push(noteSummary(t, note));
      for (const i of schema?.cellIssues(sheetIndex, row, col) ?? []) {
        parts.push(t(i.severity === 'error' ? 'grid.spokenError' : 'grid.spokenWarning', { message: t.text(i.message) }));
      }
      return parts.join(', ');
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, dataVersion, styles, date1904, locale, numbers, highlight, schema, schemaVersion, sheetIndex, t],
  );

  const sheetName = view.info.name;
  const sheetAccess = useMemo<SheetAccess>(
    () => ({
      label: t(rowCount !== undefined && rowStart === 0 ? 'grid.labelFrozen' : 'grid.label', { name: sheetName }),
      rowCount: (view.visibleRows[view.rowCount - 1] ?? -1) + 1,
      colCount: (view.visibleCols[view.colCount - 1] ?? -1) + 1,
      rowIndex: (gr) => view.visibleRows[gr + rowStart] + 1,
      colIndex: (gc) => view.visibleCols[gc - 1] + 1,
      cellLabel: (cell) => (cell as SpokenCell).spoken?.(),
    }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, layoutVersion, view.rowCount, view.colCount, rowStart, rowCount, sheetName, t],
  );
  const experimental = useMemo(() => ({ sheetAccess }) as unknown as DataEditorProps['experimental'], [sheetAccess]);

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
      const style = rec ? effectiveStyle(view, styles, merge ? merge.r1 : row, col, rec.style, palette.link) : styles[0];
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
      const asEnum = (cell: EnumCell, shown = cell.displayData): EnumCell & SpokenCell => {
        const spoken = { ...cell, shownText: shown, spoken: () => spokenText(row, view.visibleCols[gc - 1], merge, shown, !!options) };
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
      let { text, color } = displayText(rec, style, date1904, numbers.display === 'full', locale);
      const full = text;
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
      return asEnum(
        {
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
        },
        // A number too wide for its cell is said as it is drawn, with fewer digits, but in full where it is drawn as ####.
        /^#+$/.test(text) ? full : text,
      );
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, dataVersion, styles, date1904, locale, highlight, frozen, numbers, schema, schemaVersion, rowStart, g, palette, t, spokenText],
  );

  // Text spilling into empty neighbours, worked out per row as rows are drawn.
  const overflowCache = useMemo(
    () => new Map<number, Overflow[]>(),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, dataVersion, styles, date1904, colVersion, palette],
  );
  const overflowsFor = useCallback(
    (row: number) => {
      let list = overflowCache.get(row);
      if (!list) {
        list = rowOverflows(view, row, (col, rec) => effectiveStyle(view, styles, row, col, rec.style, palette.link), date1904, g.measure);
        overflowCache.set(row, list);
      }
      return list;
    },
    [overflowCache, view, styles, date1904, g, palette],
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
      const style = effectiveStyle(view, styles, merge?.r1 ?? row, merge?.c1 ?? col, undefined, palette.link);
      const overflows = overflowsFor(row);
      const own = overflows.find((o) => o.col === col);
      const through = own ?? overflows.find((o) => col >= o.from && col <= o.to);

      // Conditional formatting's bars and icons sit behind the text.
      const cf = view.cfAt(merge?.r1 ?? row, merge?.c1 ?? col);
      if (cf?.bar) drawDataBar(ctx, rect, cf.bar);
      if (cf?.icon) drawIcon(ctx, cf.icon.set, cf.icon.index, rect.x + 3, rect.y + (rect.height - 14) / 2);
      if (!own && !cf?.hideValue) {
        const text = args.cell.kind === GridCellKind.Text ? args.cell.displayData : '';
        const align = args.cell.kind === GridCellKind.Text ? (args.cell.contentAlign ?? 'left') : 'left';
        if (text && args.cell.kind === GridCellKind.Text && args.cell.allowWrapping) {
          g.drawWrapped(ctx, rect, text, align, style, args.theme);
        } else {
          // Glide centres text; Excel's vertical alignment (bottom unless set) and indent move it.
          let dx = 0;
          let dy = 0;
          let clip = false;
          if (text) {
            dx = align === 'right' ? -indentPx(style, 'right') : indentPx(style, align);
            const em = g.emHeight(style);
            dy = verticalShift(style.vAlign, rect.height, em, args.theme.cellVerticalPadding);
            // Text moved past an edge of its cell is cut there, as Excel cuts it.
            const top = rect.height / 2 - em / 2 + dy;
            clip = top < 0 || top + em > rect.height;
          }
          if (clip) {
            ctx.save();
            ctx.beginPath();
            ctx.rect(rect.x, rect.y, rect.width, rect.height);
            ctx.clip();
          }
          ctx.translate(dx, dy);
          drawContent();
          if ((style.strike || style.underline) && text) g.decorate(ctx, rect, text, align, style);
          ctx.translate(-dx, -dy);
          if (clip) ctx.restore();
        }
      }
      if (through) {
        const font = g.fontFor(through.style);
        ctx.save();
        ctx.font = font;
        ctx.fillStyle = through.color ?? through.style.color ?? (through.style.fill ? palette.textOnFill : palette.text);
        ctx.textAlign = 'left';
        const x = rect.x - offsetBetween(through.col, col) + overflowTextX(through, view.colWidth(through.col));
        const lift = verticalShift(through.style.vAlign, rect.height, g.emHeight(through.style), args.theme.cellVerticalPadding);
        ctx.fillText(through.text, x, rect.y + rect.height / 2 + lift + getMiddleCenterBias(ctx, font));
        // Underlined text, such as a link's, keeps its line as it runs on into the next cells.
        if (through.style.underline || through.style.strike) {
          const width = g.measure(through.text, through.style);
          const mid = rect.y + rect.height / 2 + lift;
          const thickness = through.style.bold ? 1.5 : 1;
          if (through.style.strike) ctx.fillRect(x, mid, width, thickness);
          if (through.style.underline) ctx.fillRect(x, mid + through.style.fontSizePx * 0.45, width, thickness);
        }
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
      // A filter's headings have its buttons, as Excel draws them.
      const filterButton = view.filterButtonAt(row, col);
      if (filterButton) drawFilterButton(ctx, rect, filterButton.filtered, { bg: palette.bg, line: palette.headerLine, arrow: palette.dropdownArrow, filtered: palette.selection });
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
    // eslint-disable-next-line react-hooks/exhaustive-deps -- the filter's buttons follow the filter
    [view, view.filter, schema, schemaVersion, sheetIndex, styles, overflowsFor, offsetBetween, frozen, rowStart, g, palette],
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

  /** The grid columns and rows in view, frozen ones aside. */
  const visible = useRef({ x: 1, y: 0, width: 0, height: 0 });
  /** The sheet row at the top of this pane's view, where Excel puts the active cell of columns selected. */
  const topRow = useCallback(() => view.visibleRows[Math.min(visible.current.y + rowStart, view.rowCount - 1)] ?? 0, [view, rowStart]);

  const cutRange = store.cutRange;
  const formulaRefs = store.formulaRefs;
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
    // Where the selection being dragged by its border would land.
    const landing = cellDrag ? clip(cellDrag.rect) : undefined;
    if (landing) marks.push({ color: palette.selectionFill, range: landing, style: 'dashed' });
    // The range the selected cell's formula spills into, or spills from, outlined as Excel does; dashed while it is blocked.
    const spill = view.get(selection.row, selection.col)?.spill;
    const spillSpan = spill && view.gridSpan(spill);
    if (spillSpan) {
      const y1 = Math.max(rowStart, spillSpan.y1);
      const y2 = Math.min(last, spillSpan.y2);
      const blocked = view.get(spill.r1, spill.c1)?.s === '#SPILL!';
      if (y1 <= y2) marks.push({ color: withAlpha(palette.spillLine, 0), range: { x: spillSpan.x1 + 1, y: y1 - rowStart, width: spillSpan.x2 - spillSpan.x1 + 1, height: y2 - y1 + 1 }, style: blocked ? 'dashed' : 'solid' });
    }
    // The references in a formula being typed, outlined and lightly filled in their colours, as Excel does.
    for (const ref of formulaRefs) {
      const span = ref.sheet === view.info.index ? view.gridSpan(ref.rect) : undefined;
      if (!span) continue;
      const y1 = Math.max(rowStart, span.y1);
      const y2 = Math.min(last, span.y2);
      if (y1 > y2) continue;
      const range = { x: span.x1 + 1, y: y1 - rowStart, width: span.x2 - span.x1 + 1, height: y2 - y1 + 1 };
      marks.push({ color: withAlpha(refColors[ref.color % refColors.length], 0.1), range, style: 'solid' });
    }
    // Whole rows and columns are shown in full, past the used range, with their numbers or letters marked.
    const lines = selection.lines;
    if (lines) {
      let columns = CompactSelection.empty();
      let lineRows = CompactSelection.empty();
      if (lines === 'all') columns = CompactSelection.fromSingleSelection([1, view.colCount + 1]);
      else {
        for (const r of selection.rects) {
          const s = view.lineSpan(lines, lines === 'rows' ? r.r1 : r.c1, lines === 'rows' ? r.r2 : r.c2);
          if (!s) continue;
          if (lines === 'cols') columns = columns.add([s[0] + 1, s[1] + 2]);
          else if (Math.max(rowStart, s[0]) <= Math.min(last, s[1])) lineRows = lineRows.add([Math.max(rowStart, s[0]) - rowStart, Math.min(last, s[1]) - rowStart + 1]);
        }
      }
      const cell = { x: gc, y: gr - rowStart, width: 1, height: 1 };
      const current = gr < rowStart || gr > last ? undefined : { cell: [gc, gr - rowStart] as Item, range: cell, rangeStack: [] };
      return { gridSelection: { current, columns, rows: lineRows } as GridSelection, highlights: marks.length ? marks : undefined };
    }
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
    // eslint-disable-next-line react-hooks/exhaustive-deps -- the spill outline follows the cells' data
  }, [view, dataVersion, selection, rowStart, rows, cutRange, formulaRefs, palette, refColors, cellDrag]);

  // Mouse presses on the letters, numbers and corner are taken before Glide
  // sees them (below); touches on them, and Glide's keys for whole lines, come here.
  const onSelectionChange = useCallback(
    (sel: GridSelection) => {
      /** Runs of grid lines as sheet lines, first to last. */
      const ranges = (picked: CompactSelection, sheetLine: (g: number) => number) => {
        const runs: [number, number][] = [];
        for (const g of picked) {
          const last = runs[runs.length - 1];
          if (last && g === last[1] + 1) last[1] = g;
          else runs.push([g, g]);
        }
        return runs.map(([a, b]): [number, number] => [sheetLine(a), sheetLine(b)]);
      };
      if (sel.columns.length > 0) {
        if (sel.columns.hasIndex(0)) return store.selectAll();
        const runs = ranges(sel.columns, (g) => view.visibleCols[g - 1]);
        store.selectLines('cols', runs, { row: topRow(), col: runs[runs.length - 1][0] });
        return;
      }
      if (sel.rows.length > 0) {
        const runs = ranges(sel.rows, (g) => view.visibleRows[g + rowStart]);
        store.selectLines('rows', runs, { row: runs[runs.length - 1][0], col: view.visibleCols[0] });
        return;
      }
      if (!sel.current) return;
      const [gc, gr] = sel.current.cell;
      if (gc === 0) {
        const picked = [sel.current.range.y, sel.current.range.y + sel.current.range.height - 1].map((y) => view.visibleRows[y + rowStart]);
        store.selectLines('rows', [[picked[0], picked[1]]], { row: picked[0], col: view.visibleCols[0] });
        return;
      }
      store.setSelection({
        row: view.visibleRows[gr + rowStart],
        col: view.visibleCols[gc - 1],
        rects: [sel.current.range, ...sel.current.rangeStack].map(toSheetRect),
      });
    },
    [store, view, toSheetRect, rowStart, topRow],
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
      // Shift+F2, Excel's key for a cell's note, opens its note or comments beside it.
      if (e.shiftKey && e.key === 'F2' && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        const sel = store.selection;
        const v = store.view;
        const at = v && ref.current?.getBounds(v.gridCol(sel.col) + 1, v.gridRow(sel.row) - rowStart);
        if (at) store.openComment(sel.row, sel.col, { x: at.x + at.width + 4, y: at.y });
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
      if (!inside && gc === 0) store.selectLines('rows', [[row, row]], { row, col });
      else if (!inside) store.setSelection({ row, col, rects: [{ r1: row, r2: row, c1: col, c2: col }] });
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
      if (!sel.rects.some((r) => col >= r.c1 && col <= r.c2 && r.r1 === 0)) store.selectLines('cols', [[col, col]], { row: topRow(), col });
      store.set('menu', { x: e.bounds.x + e.localEventX, y: e.bounds.y + e.localEventY, kind: 'column' });
    },
    [store, view, topRow],
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
      // Excel's HTML has the exact values of the numbers its text shows rounded, and each cell's formats.
      const html = e.clipboardData?.getData('text/html');
      const exact = html ? clipboardNumbers(html) : undefined;
      const formats = html ? clipboardFormats(html) : undefined;
      const sel = store.selection;
      const rect = sel.rects[0] ?? { r1: sel.row, c1: sel.col, r2: sel.row, c2: sel.col };
      void store.paste({ row: rect.r1, col: rect.c1 }, values, rect, false, exact, formats);
    };
    window.addEventListener('paste', onPaste, true);
    return () => window.removeEventListener('paste', onPaste, true);
  }, [store]);

  // Delete clears what is selected, whole rows and columns as far as they hold anything.
  const onDelete = useCallback(
    (sel: GridSelection) => {
      const rects = store.selection.lines ? store.selection.rects : sel.current ? [sel.current.range, ...sel.current.rangeStack].map(toSheetRect) : [];
      if (rects.length) void store.clear(rects);
      return false;
    },
    [store, toSheetRect],
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
  // The column whose edge is being dragged. Glide ends a press on an edge
  // that never moved with the narrowest width, so only a drag sets one.
  const resizing = useRef<{ index: number; moved: boolean } | null>(null);
  const onColumnResizeStart = useCallback((_c: GridColumn, _size: number, index: number) => {
    resizing.current = { index, moved: false };
  }, []);
  const onColumnResize = useCallback(
    (_c: GridColumn, size: number, index: number) => {
      // Outside a drag, Glide is fitting the column to the rows in view, after
      // what it takes for a double-click on the edge; the editor fits columns itself (below).
      if (index === 0 || !resizable || !resizing.current) return;
      // Glide repeats a drag for the other selected columns; the store sizes those with this one at the end.
      if (index !== resizing.current.index) return;
      resizing.current.moved = true;
      view.colWidths.set(view.visibleCols[index - 1], size);
      onColumnResized();
    },
    [view, onColumnResized, resizable],
  );
  // Rows are resized by dragging the bottom edge of their number, which Glide
  // doesn't offer: a press there is taken before Glide sees it.
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
  /** The grid columns in view, frozen ones first, the row numbers aside. */
  const columnsInView = () => {
    const out: number[] = [];
    for (let gc = 1; gc <= frozen; gc++) out.push(gc);
    for (let gc = Math.max(frozen + 1, visible.current.x); gc <= Math.min(view.colCount, visible.current.x + visible.current.width); gc++) out.push(gc);
    return out;
  };
  /**
   * The letters, row number or corner under the pointer, if any: not where a
   * column's edge is, which Glide drags to resize it.
   */
  const headingAt = (x: number, y: number): { axis: 'rows' | 'cols'; line: number } | { axis: 'all' } | undefined => {
    const grid = ref.current;
    if (!grid) return undefined;
    for (let gr = visible.current.y; gr <= visible.current.y + visible.current.height && gr + rowStart < view.rowCount; gr++) {
      const b = grid.getBounds(0, gr);
      if (!b || x < b.x || x >= b.x + b.width) break;
      if (y >= b.y && y < b.y + b.height) return { axis: 'rows', line: view.visibleRows[gr + rowStart] };
    }
    const corner = header ? grid.getBounds(0, -1) : undefined;
    if (!corner || y < corner.y || y >= corner.y + corner.height) return undefined;
    if (x >= corner.x && x < corner.x + corner.width) return { axis: 'all' };
    for (const gc of columnsInView()) {
      const b = grid.getBounds(gc, -1);
      if (!b || x < b.x || x >= b.x + b.width) continue;
      return x - b.x <= 5 || b.x + b.width - x <= 5 ? undefined : { axis: 'cols', line: view.visibleCols[gc - 1] };
    }
    return undefined;
  };
  /** The column whose right edge is under the pointer, among the letters, where a drag resizes it. */
  const columnEdgeAt = (x: number, y: number): number | undefined => {
    const grid = ref.current;
    const corner = header ? grid?.getBounds(0, -1) : undefined;
    if (!grid || !corner || y < corner.y || y >= corner.y + corner.height) return undefined;
    for (const gc of columnsInView()) {
      const b = grid.getBounds(gc, -1);
      if (!b || x < b.x - 5 || x > b.x + b.width + 5) continue;
      if (Math.abs(x - (b.x + b.width)) <= 5) return view.visibleCols[gc - 1];
      if (Math.abs(x - b.x) <= 5 && gc > 1) return view.visibleCols[gc - 2];
    }
    return undefined;
  };
  /** The row (or column) the pointer is level with while dragging over the numbers (or letters), the nearest in view past either end. */
  const lineLevelWith = (axis: 'rows' | 'cols', x: number, y: number): number | undefined => {
    const grid = ref.current;
    if (!grid) return undefined;
    const lines = axis === 'rows' ? Array.from({ length: visible.current.height + 1 }, (_, i) => visible.current.y + i).filter((gr) => gr + rowStart < view.rowCount) : columnsInView();
    let nearest: number | undefined;
    for (const g of lines) {
      const b = axis === 'rows' ? grid.getBounds(0, g) : grid.getBounds(g, -1);
      if (!b) continue;
      const [at, from, size] = axis === 'rows' ? [y, b.y, b.height] : [x, b.x, b.width];
      if (nearest === undefined || at >= from) nearest = g;
      if (at < from + size) break;
    }
    if (nearest === undefined) return undefined;
    return axis === 'rows' ? view.visibleRows[nearest + rowStart] : view.visibleCols[nearest - 1];
  };
  /**
   * Selects whole rows or columns from their numbers or letters as Excel
   * does: a click picks one, dragging spans several, Shift+click reaches
   * from the active line, and Ctrl+click (⌘-click on a Mac) adds one. The
   * corner selects the whole sheet.
   */
  const pressHeading = (e: ReactMouseEvent, hit: NonNullable<ReturnType<typeof headingAt>>) => {
    e.preventDefault();
    e.stopPropagation();
    ref.current?.focus();
    // The click that ends the press isn't one on a cell.
    window.addEventListener('mouseup', swallowNextClick, { capture: true, once: true });
    if (hit.axis === 'all') return store.selectAll();
    const { axis, line } = hit;
    const sel = store.selection;
    const same = sel.lines === axis;
    const add = same && (APPLE ? e.metaKey : e.ctrlKey);
    const extend = same && e.shiftKey;
    const anchor = extend ? (axis === 'rows' ? sel.row : sel.col) : line;
    const kept: [number, number][] = add || extend ? sel.rects.map((r) => (axis === 'rows' ? [r.r1, r.r2] : [r.c1, r.c2])) : [];
    if (extend) kept.pop();
    const active = axis === 'rows' ? { row: anchor, col: view.visibleCols[0] ?? 0 } : { row: topRow(), col: anchor };
    const pick = (to: number) => store.selectLines(axis, [...kept, [anchor, to]], active);
    pick(line);
    const move = (m: MouseEvent) => {
      const to = lineLevelWith(axis, m.clientX, m.clientY);
      if (to !== undefined) pick(to);
    };
    const up = () => {
      window.removeEventListener('mousemove', move, true);
      window.removeEventListener('mouseup', up, true);
    };
    window.addEventListener('mousemove', move, true);
    window.addEventListener('mouseup', up, true);
  };

  /**
   * Whether the pointer is on the selection's border, which can be dragged
   * to move its cells, as in Excel: within a few pixels of its edge on this
   * pane, but not on the fill handle at its bottom-right corner.
   */
  const onSelectionEdge = (x: number, y: number): boolean => {
    const grid = ref.current;
    const sel = store.selection;
    const r = sel.rects[0];
    if (!grid || !resizable || sel.rects.length !== 1 || sel.lines === 'all' || !r) return false;
    const last = rowStart + rows - 1;
    const y1 = Math.max(rowStart, view.gridRow(r.r1));
    const y2 = Math.min(last, view.gridRow(r.r2));
    if (y1 > y2) return false;
    const x1 = Math.max(0, view.gridCol(r.c1)) + 1;
    const x2 = Math.max(x1, view.gridCol(r.c2) + 1);
    const a = grid.getBounds(x1, y1 - rowStart);
    const b = grid.getBounds(x2, y2 - rowStart);
    if (!a || !b) return false;
    const [left, top, right, bottom] = [a.x, a.y, b.x + b.width, b.y + b.height];
    const near = (p: number, edge: number) => Math.abs(p - edge) <= 4;
    if (near(x, right) && near(y, bottom)) return false;
    const across = x >= left - 4 && x <= right + 4;
    const down = y >= top - 4 && y <= bottom + 4;
    return (across && (near(y, top) || near(y, bottom))) || (down && (near(x, left) || near(x, right)));
  };
  /**
   * The filter button a cell's heading has, with where it is on screen: under
   * the pointer at `x`, `y`, or, without them, on the cell itself.
   */
  const filterButtonOf = useCallback(
    (row: number, col: number, x?: number, y?: number) => {
      const button = view.filterButtonAt(row, col);
      const grid = ref.current;
      const cell = button && grid ? grid.getBounds(Math.max(0, view.gridCol(col)) + 1, view.gridRow(row) - rowStart) : undefined;
      if (!button || !cell) return undefined;
      const box = filterButtonBox(cell);
      if (x !== undefined && y !== undefined && (x < box.x || x > box.x + box.size || y < box.y || y > box.y + box.size)) return undefined;
      return { col, rect: button.rect, x: cell.x, y: cell.y + cell.height };
    },
    [view, rowStart, ref],
  );
  /** The cell under the pointer, or the nearest in view past either end. */
  const cellUnder = (x: number, y: number) => {
    const row = lineLevelWith('rows', x, y);
    const col = lineLevelWith('cols', x, y);
    return row === undefined || col === undefined ? undefined : { row, col };
  };
  /** Where a block of cells is on this pane, in the host's pixels. */
  const boxOf = (r: Rect) => {
    const grid = ref.current;
    if (!grid || !host.current) return undefined;
    const a = grid.getBounds(Math.max(0, view.gridCol(r.c1)) + 1, view.gridRow(r.r1) - rowStart);
    const b = grid.getBounds(Math.max(0, view.gridCol(r.c2)) + 1, view.gridRow(r.r2) - rowStart);
    if (!a || !b) return undefined;
    const h = host.current.getBoundingClientRect();
    return { left: a.x - h.left, top: a.y - h.top, right: b.x + b.width - h.left, bottom: b.y + b.height - h.top };
  };
  /**
   * Drags the selection by its border: its cells move where it is dropped,
   * or with Ctrl (Option on a Mac) held, are copied there, as in Excel. A
   * dashed outline shows where they'll go. With Shift, they go in between the
   * cells there instead, at the cell edge nearest the pointer, which a line
   * shows. Escape gives up.
   */
  const dragSelection = (e: ReactMouseEvent) => {
    const src = store.selection.rects[0];
    const from = cellUnder(e.clientX, e.clientY);
    if (!src || !from) return;
    e.preventDefault();
    e.stopPropagation();
    // Where in the selection it was picked up, so it keeps that point under the pointer.
    const off = { row: Math.min(Math.max(from.row - src.r1, 0), src.r2 - src.r1), col: Math.min(Math.max(from.col - src.c1, 0), src.c2 - src.c1) };
    const height = src.r2 - src.r1;
    const width = src.c2 - src.c1;
    let drop: Rect | undefined;
    let insert: { row: number; col: number; shift: 'down' | 'right' } | undefined;
    let copy = false;
    let pointer = { x: e.clientX, y: e.clientY };
    const copying = (m: MouseEvent | KeyboardEvent) => (APPLE ? m.altKey : m.ctrlKey);
    const show = () => {
      setCellDrag(drop && !insert ? { rect: drop, copy } : null);
      if (host.current) host.current.style.cursor = copy ? 'copy' : 'move';
    };
    const update = (m: MouseEvent | KeyboardEvent) => {
      const at = cellUnder(pointer.x, pointer.y);
      copy = copying(m);
      insert = undefined;
      setInsertGuide(null);
      if (at) {
        const row = Math.max(0, at.row - off.row);
        const col = Math.max(0, at.col - off.col);
        drop = { r1: row, c1: col, r2: row + height, c2: col + width };
        const cell = m.shiftKey && !copy ? boxOf({ r1: at.row, c1: at.col, r2: at.row, c2: at.col }) : undefined;
        const h = host.current?.getBoundingClientRect();
        if (cell && h) {
          // The edge nearest the pointer, for the cell's shape: top or bottom puts the cells between rows, left or right between columns.
          const x = (pointer.x - h.left - cell.left) / (cell.right - cell.left);
          const y = (pointer.y - h.top - cell.top) / (cell.bottom - cell.top);
          const across = Math.min(x, 1 - x) < Math.min(y, 1 - y);
          insert = across ? { row, col: x < 0.5 ? at.col : at.col + 1, shift: 'right' } : { row: y < 0.5 ? at.row : at.row + 1, col, shift: 'down' };
          const box = across ? boxOf({ r1: row, c1: at.col, r2: row + height, c2: at.col }) : boxOf({ r1: at.row, c1: col, r2: at.row, c2: col + width });
          if (box)
            setInsertGuide(
              across
                ? { left: x < 0.5 ? box.left : box.right, top: box.top, width: 0, height: box.bottom - box.top }
                : { left: box.left, top: y < 0.5 ? box.top : box.bottom, width: box.right - box.left, height: 0 },
            );
        }
      }
      show();
    };
    const move = (m: MouseEvent) => {
      pointer = { x: m.clientX, y: m.clientY };
      update(m);
    };
    const key = (k: KeyboardEvent) => {
      if (k.key === 'Escape') {
        drop = undefined;
        insert = undefined;
        end();
      } else {
        update(k);
      }
    };
    const end = () => {
      window.removeEventListener('mousemove', move, true);
      window.removeEventListener('mouseup', up, true);
      window.removeEventListener('keydown', key, true);
      window.removeEventListener('keyup', key, true);
      setCellDrag(null);
      setInsertGuide(null);
      if (host.current) host.current.style.cursor = '';
    };
    const up = () => {
      // The click that ends the drag isn't one on the grid.
      swallowNextClick();
      const to = drop;
      const between = insert;
      end();
      if (between) void store.dragInsert(between.row, between.col, between.shift);
      else if (to && (to.r1 !== src.r1 || to.c1 !== src.c1)) void store.dragSelection(to.r1, to.c1, copy);
    };
    window.addEventListener('mousemove', move, true);
    window.addEventListener('mouseup', up, true);
    window.addEventListener('keydown', key, true);
    window.addEventListener('keyup', key, true);
  };

  const onHostMouseMove = (e: ReactMouseEvent) => {
    if (rowDrag || cellDrag || !host.current) return;
    host.current.style.cursor = rowEdgeAt(e.clientX, e.clientY) ? 'row-resize' : onSelectionEdge(e.clientX, e.clientY) ? 'move' : '';
  };
  const onHostMouseDown = (e: ReactMouseEvent) => {
    // A filter button opens its column's filter menu, leaving the selection where it is.
    const under = e.button === 0 ? cellUnder(e.clientX, e.clientY) : undefined;
    const button = under && filterButtonOf(under.row, under.col, e.clientX, e.clientY);
    if (button) {
      e.preventDefault();
      e.stopPropagation();
      store.set('filterPicker', button);
      return;
    }
    // A double-click on a column's edge fits it to its widest value, as in Excel.
    const edge = e.button === 0 && e.detail >= 2 && resizable ? columnEdgeAt(e.clientX, e.clientY) : undefined;
    if (edge !== undefined) {
      e.preventDefault();
      e.stopPropagation();
      void store.autofit('cols', edge);
      return;
    }
    const at = e.button === 0 ? rowEdgeAt(e.clientX, e.clientY) : undefined;
    if (!at && e.button === 0 && e.detail < 2 && onSelectionEdge(e.clientX, e.clientY)) {
      dragSelection(e);
      return;
    }
    if (!at) {
      const hit = e.button === 0 ? headingAt(e.clientX, e.clientY) : undefined;
      if (hit) pressHeading(e, hit);
      return;
    }
    // A double-click on the edge fits the row to what it holds, as in Excel.
    if (e.detail >= 2) {
      e.preventDefault();
      e.stopPropagation();
      void store.autofit('rows', at.row);
      return;
    }
    e.preventDefault();
    e.stopPropagation();
    const start = { row: at.row, from: e.clientY, px: view.rowHeight(at.row), y: at.edge };
    setRowDrag(start);
    const move = (m: MouseEvent) => setRowDrag({ ...start, y: start.y + Math.max(-start.px, m.clientY - start.from) });
    const up = (m: MouseEvent) => {
      window.removeEventListener('mousemove', move, true);
      window.removeEventListener('mouseup', up, true);
      // The click that ends the drag isn't one on the grid.
      swallowNextClick();
      setRowDrag(null);
      const px = Math.max(0, Math.round(start.px + m.clientY - start.from));
      if (px !== start.px) void store.resizedRow(start.row, px);
    };
    window.addEventListener('mousemove', move, true);
    window.addEventListener('mouseup', up, true);
  };

  const onColumnResizeEnd = useCallback(
    (_c: GridColumn, size: number, index: number) => {
      const dragged = resizing.current;
      // Glide also ends the other selected columns' resizes; the store sizes those with the one dragged.
      if (!dragged || dragged.index !== index) return;
      resizing.current = null;
      if (dragged.moved && index > 0 && resizable) void store.resizedColumn(view.visibleCols[index - 1], size);
    },
    [store, view, resizable],
  );

  const onItemHovered = useCallback(
    (args: GridMouseEventArgs) => {
      const [gc, gr] = args.location;
      if (args.kind !== 'cell' || gc === 0) {
        setHoverNote(null);
        setHoverLink(null);
        return;
      }
      const row = view.visibleRows[gr + rowStart];
      const col = view.visibleCols[gc - 1];
      const note = view.noteAt(row, col);
      const link = cellLink(view, row, col);
      const { x, y, width, height } = args.bounds;
      setHoverNote((prev) => (!note ? null : prev?.note === note ? prev : { note, x: x + width + 4, y }));
      // Below the cell, as Excel shows where a link goes, so a note beside it stays clear.
      setHoverLink(link ? { link, x, y: y + height + 4 } : null);
    },
    [view, rowStart],
  );

  // Ctrl+Enter (⌘Enter on a Mac) follows the active cell's link, as Ctrl+click does, in place of the grid's own use of it.
  // Alt+Shift+Right and Left group and ungroup the selection, as in Excel.
  const onGridKeyDown = useCallback(
    (e: GridKeyEventArgs) => {
      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey) && !e.shiftKey && !e.altKey) {
        e.cancel();
        e.preventDefault();
        void store.followLink();
        return;
      }
      if ((e.key === 'ArrowRight' || e.key === 'ArrowLeft') && e.altKey && e.shiftKey && !e.ctrlKey && !e.metaKey) {
        e.cancel();
        e.preventDefault();
        void store.groupSelection(e.key === 'ArrowRight' ? 1 : -1);
        return;
      }
      // Excel's keys for whole lines: Shift+Space the selection's rows, Ctrl+Space its columns, Ctrl+A (⌘A) the sheet.
      const sel = store.selection;
      if (e.key === ' ' && !e.altKey && !e.metaKey && e.shiftKey !== e.ctrlKey) {
        e.cancel();
        e.preventDefault();
        const axis = e.shiftKey ? 'rows' : 'cols';
        store.selectLines(axis, sel.rects.map((r) => (axis === 'rows' ? [r.r1, r.r2] : [r.c1, r.c2])), { row: sel.row, col: sel.col });
        return;
      }
      if (e.key.toLowerCase() === 'a' && (APPLE ? e.metaKey : e.ctrlKey) && !e.shiftKey && !e.altKey) {
        e.cancel();
        e.preventDefault();
        store.selectAll();
        return;
      }
      // Alt+Down on a heading with a filter button opens its menu, as in Excel.
      if (e.key === 'ArrowDown' && e.altKey && !e.ctrlKey && !e.metaKey && !e.shiftKey) {
        const button = filterButtonOf(sel.row, sel.col);
        if (button) {
          e.cancel();
          e.preventDefault();
          store.set('filterPicker', button);
          return;
        }
      }
      // Ctrl+Shift+L, or ⇧⌘F on a Mac, turns the filter buttons on or off, as in Excel.
      const letter = e.key.toLowerCase();
      if ((e.ctrlKey || e.metaKey) && e.shiftKey && !e.altKey && (letter === 'l' || (APPLE && letter === 'f'))) {
        e.cancel();
        e.preventDefault();
        void store.toggleFilter();
        return;
      }
      // Ctrl++ (Ctrl+Shift+= on a US keyboard) inserts and Ctrl+- deletes, as in Excel, in place of the browser's zoom.
      const insert = e.key === '+' || (e.key === '=' && e.shiftKey);
      if ((insert || e.key === '-') && (e.ctrlKey || e.metaKey) && !e.altKey) {
        e.cancel();
        e.preventDefault();
        store.changeCells(insert ? 'insert' : 'delete');
        return;
      }
      onKeyDown?.(e);
    },
    [store, onKeyDown, filterButtonOf],
  );

  // Ctrl+click (⌘-click on a Mac) follows a cell's link; a plain click only selects it.
  const onCellClicked = useCallback(
    ([gc, gr]: Item, e: CellClickedEventArgs) => {
      if (gc === 0 || !(APPLE ? e.metaKey : e.ctrlKey) || e.shiftKey) return;
      const row = view.visibleRows[gr + rowStart];
      const col = view.visibleCols[gc - 1];
      if (cellLink(view, row, col)) void store.followLink(row, col);
    },
    [store, view, rowStart],
  );

  const rowHeight = useCallback(
    (gr: number) => view.rowHeight(view.visibleRows[gr + rowStart]),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [view, rowStart, layoutVersion],
  );
  // Lists get a dropdown; text gets the editor that helps with formulas.
  const textEditor = useMemo(() => textCellEditor(store), [store]);
  const provideEditor = useCallback(
    (cell: GridCell) =>
      enumOptionsOf(cell)
        ? { editor: EnumEditor, disablePadding: true }
        : cell.kind === GridCellKind.Text
          ? { editor: textEditor }
          : undefined,
    [textEditor],
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
        onCellClicked={onCellClicked}
        keybindings={{ copy: false, paste: false, cut: false }}
        onCellContextMenu={onCellContextMenu}
        onHeaderContextMenu={onHeaderContextMenu}
        provideEditor={provideEditor}
        onDelete={onDelete}
        fillHandle
        allowedFillDirections="orthogonal"
        onFillPattern={onFillPattern}
        onColumnResizeStart={onColumnResizeStart}
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
        onKeyDown={onGridKeyDown}
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
          visible.current = { x: r.x, y: r.y, width: r.width, height: r.height };
          setRegion((n) => n + 1);
          onRegion?.(r);
          // Near the last row or column shown, more follow, as far as Excel's sheet goes.
          const lastRow = Math.min(rowStart + r.y + r.height, view.rowCount - 1);
          const lastCol = Math.min(r.x + r.width - 1, view.colCount) - 1;
          if (lastRow + 50 >= view.rowCount || lastCol + 3 >= view.colCount) store.scrolledTo(view, view.visibleRows[lastRow] ?? 0, view.visibleCols[lastCol] ?? 0);
        }}
        onItemHovered={onItemHovered}
        experimental={experimental}
      />
      {hoverNote && <NoteTip {...hoverNote} />}
      {hoverLink && <LinkTip {...hoverLink} follows={store.followsLinks} />}
      {rowDrag && host.current && <div className="wbe-row-guide" style={{ top: rowDrag.y - host.current.getBoundingClientRect().top }} />}
      {insertGuide && <div className="wbe-insert-guide" style={insertGuide} />}
      <ObjectsOverlay store={store} view={view} gridRef={ref} tick={region + colVersion} rowStart={rowStart} rowCount={rows} scrollRow={visible.current.y} top={header ? 24 : 0} />
    </div>
  );
}
