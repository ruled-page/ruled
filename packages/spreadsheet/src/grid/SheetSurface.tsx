import { inRect } from '@ruled/formula/refs';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type MouseEvent as ReactMouseEvent, type PointerEvent as ReactPointerEvent } from 'react';
import type { CellNote, Rect } from '../core/types';
import { effectiveStyle } from '../state/cellStyle';
import { editText } from '../state/format';
import { cellLink, type CellLink } from '../state/links';
import { DEFAULT_PALETTE, DEFAULT_REF_COLORS, readGridPalette, readRefColors, type GridPalette } from '../ui/gridPalette';
import { APPLE, commandKey } from '../ui/keys';
import { LinkTip } from '../ui/LinkTip';
import { NoteTip } from '../ui/NoteTip';
import { checkboxBox, filterButtonBox } from '../ui/paint';
import type { Box, EditMove, GridHandle, Pane, SurfaceProps } from './contract';
import { useSheetEvents } from './hooks';
import { editorKey, type FilterButton } from './keys';
import { cellOptions } from './lists';
import { sheetMarks } from './marks';
import { growOverMerges, navigate, navSheet } from './navigation';
import { AccessTable } from './AccessTable';
import { CellEditorBox, type Edit } from './CellEditing';
import { colAxis, rowAxis } from './axis';
import { Painter, type PaintInput } from './draw';
import { contentScroll, Layout, scrollerPosition, scrollLength } from './layout';

/** The height of the column letters. */
const LETTERS = 24;
/** How far a drag past the grid's edge scrolls it each frame. */
const DRAG_STEP = 20;

const span = (a: { row: number; col: number }, b: { row: number; col: number }): Rect => ({
  r1: Math.min(a.row, b.row),
  r2: Math.max(a.row, b.row),
  c1: Math.min(a.col, b.col),
  c2: Math.max(a.col, b.col),
});

/** `line` among sorted `lines`, if it isn't there already. */
const withLine = (lines: number[], line: number) => (lines.includes(line) ? lines : [...lines, line].sort((a, b) => a - b));

/**
 * The grid: the sheet drawn on a canvas, under a scroller whose scrollbars,
 * wheel, trackpad and touch scrolling are the browser's own. Only the cells
 * in view are drawn. Cells are selected, edited, filled and resized with the
 * mouse, a pen, a finger and the keyboard, as in Excel; a table of the cells
 * in view, whose active cell takes the focus, is what screen readers read.
 */
export function SheetSurface({ store, view, dataVersion, schemaVersion, selection, onHandle }: SurfaceProps) {
  const host = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const scroller = useRef<HTMLDivElement>(null);
  const painter = useRef<Painter>(undefined);
  const drawn = useRef<{ painter: Painter; input: PaintInput }>(undefined);
  const [palette, setPalette] = useState<GridPalette>(DEFAULT_PALETTE);
  const [refColors, setRefColors] = useState(DEFAULT_REF_COLORS);
  const themeVersion = store.themeVersion;
  useLayoutEffect(() => {
    const next = readGridPalette(host.current);
    setPalette((prev) => (JSON.stringify(prev) === JSON.stringify(next) ? prev : next));
    const colors = readRefColors(host.current);
    setRefColors((prev) => (prev.join() === colors.join() ? prev : colors));
  }, [themeVersion]);

  const layoutVersion = view.layoutVersion;
  const { rowCount, colCount } = view;
  /** Bumped while a column's edge is dragged, its width changing as it goes. */
  const [sized, setSized] = useState(0);
  const axes = useMemo(
    () => ({ rows: rowAxis(view), cols: colAxis(view) }),
    // eslint-disable-next-line react-hooks/exhaustive-deps -- a view's sizes change in place
    [view, layoutVersion, store.colVersion, rowCount, colCount, sized],
  );
  /** Frozen lines that are shown, not hidden; at least one row scrolls. */
  const frozen = useMemo(() => {
    let rows = 0;
    let cols = 0;
    for (let r = 0; r < view.info.frozenRows; r++) if (view.gridRow(r) >= 0) rows++;
    for (let c = 0; c < view.info.frozenCols; c++) if (view.gridCol(c) >= 0) cols++;
    return { rows: Math.min(rows, rowCount - 1), cols };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [view, layoutVersion, rowCount]);
  const numbers = Math.max(40, String(view.visibleRows[rowCount - 1] + 1).length * 8 + 16);
  /** The sheet reads right to left: column A at the right, the row numbers beside it. */
  const rtl = !!view.info.rightToLeft;
  const contentWidth = numbers + axes.cols.total;
  const contentHeight = LETTERS + axes.rows.total;

  /** The layout as the grid is now scrolled and sized. */
  const layoutNow = useCallback((): Layout | undefined => {
    const sc = scroller.current;
    if (!sc) return undefined;
    const width = sc.clientWidth;
    const height = sc.clientHeight;
    return new Layout({
      width,
      height,
      rows: axes.rows,
      cols: axes.cols,
      frozenRows: frozen.rows,
      frozenCols: frozen.cols,
      numbers,
      letters: LETTERS,
      scrollX: contentScroll(Math.abs(sc.scrollLeft), numbers + axes.cols.total, width),
      scrollY: contentScroll(sc.scrollTop, LETTERS + axes.rows.total, height),
      rtl,
    });
  }, [axes, frozen, numbers, rtl]);

  /** Where cells dragged by the selection's border, or filled by its handle, would go. */
  const [landing, setLanding] = useState<Rect | null>(null);
  const marks = useMemo(
    () => sheetMarks(store, view, selection, palette, refColors, landing ?? undefined),
    // eslint-disable-next-line react-hooks/exhaustive-deps -- marks follow the cut, the formula being typed and the cells
    [store, view, selection, palette, refColors, store.cutRange, store.formulaRefs, dataVersion, landing],
  );

  // The cells the screen readers' table gives: those in view, settled a moment after scrolling stops.
  const [lines, setLines] = useState<{ rows: number[]; cols: number[] }>({ rows: [], cols: [] });
  const tableTimer = useRef(0);
  const settle = useRef<() => void>(() => undefined);
  settle.current = () => {
    const layout = layoutNow();
    if (!layout) return;
    const of = (from: number, to: number, order: Int32Array) => Array.from({ length: Math.max(0, to - from + 1) }, (_, i) => order[from + i]);
    const rows = [...of(layout.frozenRowRun.first, layout.frozenRowRun.last, view.visibleRows), ...of(layout.rowRun.first, layout.rowRun.last, view.visibleRows)];
    const cols = [...of(layout.frozenColRun.first, layout.frozenColRun.last, view.visibleCols), ...of(layout.colRun.first, layout.colRun.last, view.visibleCols)];
    setLines((prev) => (prev.rows.join() === rows.join() && prev.cols.join() === cols.join() ? prev : { rows, cols }));
  };
  useEffect(() => () => clearTimeout(tableTimer.current), []);

  // What the painter worked out about cells goes once they, their styles or the columns change:
  // before this render, so the screen readers' table, which reads the same, is current too.
  const cellsKey = [view, dataVersion, schemaVersion, layoutVersion, store.colVersion, store.highlightEdits, store.showLabels, store.numbers];
  const lastKey = useRef<unknown[]>([]);
  if (painter.current && cellsKey.some((k, i) => k !== lastKey.current[i])) painter.current.reset();
  lastKey.current = cellsKey;

  /** Whether cells can be resized, filled and moved: not on a protected sheet. */
  const editable = view.info.kind === 'worksheet' && !view.info.protected;

  /** What the painter draws from, as the grid now is. */
  const inputFor = (layout: Layout): PaintInput => {
    const schema = store.schema;
    const sheet = view.info.index;
    return {
      view,
      layout,
      styles: store.styles,
      palette,
      date1904: store.date1904,
      locale: store.locale,
      fullNumbers: store.numbers.display === 'full',
      highlightEdits: store.highlightEdits,
      isLabel: schema && store.showLabels ? (row, col) => schema.isLabel(sheet, row, col) : undefined,
      selection: store.selection,
      marks,
      fillHandle: !store.touchMode && !store.selection.lines && store.selection.rects.length === 1 && editable,
      issue: schema
        ? (row, col) => {
            const issues = schema.cellIssues(sheet, row, col);
            if (issues.length) return issues.some((i) => i.severity === 'error') ? 'error' : 'warning';
            return schema.isPending(sheet, row, col) ? 'pending' : undefined;
          }
        : undefined,
      activeList: !!cellOptions(store, view, store.selection.row, store.selection.col),
    };
  };

  const listeners = useRef(new Set<() => void>());
  const draw = useRef<() => void>(() => undefined);
  draw.current = () => {
    const el = canvas.current;
    const layout = layoutNow();
    const ctx = el?.getContext('2d');
    if (!el || !layout || !ctx) return;
    const { width, height } = layout.spec;
    const dpr = window.devicePixelRatio || 1;
    // A sheet read right to left has its scrollbar at the left, and the cells after it.
    const sc = scroller.current;
    el.style.left = rtl && sc ? `${sc.offsetWidth - sc.clientWidth}px` : '';
    if (el.width !== Math.round(width * dpr) || el.height !== Math.round(height * dpr)) {
      el.width = Math.round(width * dpr);
      el.height = Math.round(height * dpr);
      el.style.width = `${width}px`;
      el.style.height = `${height}px`;
    }
    painter.current ??= new Painter(ctx);
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    const input = inputFor(layout);
    painter.current.draw(input);
    drawn.current = { painter: painter.current, input };
    // Near the last row or column shown, more follow, as far as Excel's sheet goes.
    const lastRow = layout.rowRun.last;
    const lastCol = layout.colRun.last;
    if (lastRow + 50 >= rowCount || lastCol + 3 >= colCount) store.scrolledTo(view, view.visibleRows[Math.max(0, lastRow)] ?? 0, view.visibleCols[Math.max(0, lastCol)] ?? 0);
    clearTimeout(tableTimer.current);
    tableTimer.current = window.setTimeout(() => settle.current(), 120);
    for (const fn of listeners.current) fn();
  };
  const frame = useRef(0);
  const redraw = useCallback(() => {
    if (!frame.current) {
      frame.current = requestAnimationFrame(() => {
        frame.current = 0;
        draw.current();
      });
    }
  }, []);
  useEffect(
    () => () => {
      cancelAnimationFrame(frame.current);
      frame.current = 0;
    },
    [],
  );

  // Anything that renders the grid may change what it shows.
  useLayoutEffect(redraw);
  // AutoFit measures text as it is drawn here.
  useEffect(() => {
    store.measureText = (text, style) => painter.current?.measure(text, style) ?? text.length * 7;
  }, [store]);

  useEffect(() => {
    const sc = scroller.current;
    if (!sc) return;
    const watch = typeof ResizeObserver === 'undefined' ? undefined : new ResizeObserver(redraw);
    watch?.observe(sc);
    sc.addEventListener('scroll', redraw, { passive: true });
    // Text is measured again once the sheet's fonts have loaded.
    const fonts = document.fonts;
    const loaded = () => {
      painter.current?.reset();
      redraw();
    };
    fonts?.addEventListener?.('loadingdone', loaded);
    return () => {
      watch?.disconnect();
      sc.removeEventListener('scroll', redraw);
      fonts?.removeEventListener?.('loadingdone', loaded);
    };
  }, [redraw]);

  const handle = useMemo<GridHandle>(() => {
    /** A box on the canvas, in the page's pixels. */
    const onPage = (b: Box): Box => {
      const at = canvas.current?.getBoundingClientRect();
      return { x: b.x + (at?.left ?? 0), y: b.y + (at?.top ?? 0), width: b.width, height: b.height };
    };
    return {
      view() {
        const layout = layoutNow();
        const sc = scroller.current;
        const at = canvas.current?.getBoundingClientRect();
        if (!layout || !sc || !at) return { corner: undefined, panes: [], cols: [], frozenCols: 0, freezeX: 0 };
        const bars = { right: sc.offsetWidth - sc.clientWidth, bottom: sc.offsetHeight - sc.clientHeight };
        const of = (from: number, to: number, order: Int32Array) => Array.from({ length: Math.max(0, to - from + 1) }, (_, i) => order[from + i]);
        const { frozenRowRun: fr, rowRun: rr, frozenColRun: fc, colRun: cr, bodyY, bodyX } = layout;
        const panes: Pane[] = [];
        if (fr.last >= 0) {
          panes.push({
            range: [view.visibleRows[0], view.visibleRows[fr.last]],
            rows: of(fr.first, fr.last, view.visibleRows),
            box: { x: at.left, y: at.top + LETTERS, width: sc.offsetWidth, height: bodyY - LETTERS },
            scrollbars: { right: bars.right, bottom: 0 },
          });
        }
        panes.push({
          range: [view.visibleRows[fr.last + 1], view.visibleRows[rowCount - 1]],
          rows: of(rr.first, rr.last, view.visibleRows),
          box: { x: at.left, y: at.top + bodyY, width: sc.offsetWidth, height: sc.offsetHeight - bodyY },
          scrollbars: bars,
        });
        return {
          corner: onPage(layout.corner()),
          panes,
          cols: [...of(fc.first, fc.last, view.visibleCols), ...of(cr.first, cr.last, view.visibleCols)],
          frozenCols: fc.last + 1,
          freezeX: rtl ? at.left + layout.spec.width - (fc.last >= 0 ? bodyX : numbers) : at.left + (fc.last >= 0 ? bodyX : numbers) - 1,
          rtl,
        };
      },
      cellBox(row, col) {
        const layout = layoutNow();
        const gr = view.gridRow(row);
        const gc = view.gridCol(col);
        return layout && gr >= 0 && gc >= 0 ? onPage(layout.cell(gr, gc)) : undefined;
      },
      headingBox(axis, line) {
        const layout = layoutNow();
        const g = axis === 'rows' ? view.gridRow(line) : view.gridCol(line);
        if (!layout || g < 0) return undefined;
        return onPage(axis === 'rows' ? layout.rowHeading(g) : layout.colHeading(g));
      },
      cellNear(x, y) {
        const layout = layoutNow();
        const at = canvas.current?.getBoundingClientRect();
        if (!layout || !at) return undefined;
        const hit = layout.hit(x - at.left, y - at.top, true);
        return hit.row < 0 || hit.col < 0 ? undefined : { row: view.visibleRows[hit.row], col: view.visibleCols[hit.col] };
      },
      scrollBy(dx, dy) {
        scroller.current?.scrollBy({ left: dx, top: dy });
      },
      subscribe(fn) {
        listeners.current.add(fn);
        return () => listeners.current.delete(fn);
      },
    };
  }, [view, layoutNow, numbers, rowCount, rtl]);
  useEffect(() => {
    onHandle(handle);
    return () => onHandle(undefined);
  }, [handle, onHandle]);

  // Copying, cutting and pasting, and the keys for the cell menu and notes.
  useSheetEvents(store, host, handle.cellBox, () => false);

  /** Scrolls the least that brings a cell into view, unless it is frozen in place. */
  const scrollIntoView = useCallback(
    (row: number, col: number) => {
      const layout = layoutNow();
      const sc = scroller.current;
      const gr = view.gridRow(row);
      const gc = view.gridCol(col);
      if (!layout || !sc || gr < 0 || gc < 0) return;
      const { rows, cols, width, height, scrollX, scrollY } = layout.spec;
      const fit = (start: number, size: number, scroll: number, room: number) => (start < scroll ? start : start + size > scroll + room ? start + size - room : undefined);
      if (gr > layout.frozenRowRun.last) {
        const y = fit(rows.start(gr) - rows.start(layout.frozenRowRun.last + 1), rows.size(gr), scrollY, height - layout.bodyY);
        if (y !== undefined) sc.scrollTop = scrollerPosition(Math.max(0, y), contentHeight, height);
      }
      if (gc > layout.frozenColRun.last) {
        const x = fit(cols.start(gc) - cols.start(layout.frozenColRun.last + 1), cols.size(gc), scrollX, width - layout.bodyX);
        if (x !== undefined) sc.scrollLeft = (rtl ? -1 : 1) * scrollerPosition(Math.max(0, x), contentWidth, width);
      }
    },
    [view, layoutNow, contentHeight, contentWidth, rtl],
  );
  /** A cell to bring into view once the grid has drawn the rows and columns it may have added. */
  const pending = useRef<{ row: number; col: number }>(undefined);
  const wantFocus = useRef(false);
  /** Whether the grid has the focus, or had it when its focused cell went (as when its row is hidden): browsers say nothing then. */
  const hasFocus = useRef(false);
  const activeCell = useRef<HTMLTableCellElement>(null);
  useLayoutEffect(() => {
    if (pending.current) {
      scrollIntoView(pending.current.row, pending.current.col);
      pending.current = undefined;
    }
    // While the grid has the focus, it follows the active cell.
    const el = activeCell.current;
    const now = document.activeElement;
    const kept = hasFocus.current && (!now || now === document.body || !!host.current?.contains(now));
    if (el && !editing.current && (wantFocus.current || kept) && now !== el) el.focus({ preventScroll: true });
    wantFocus.current = false;
  });

  useEffect(() => {
    const el = host.current;
    if (!el) return;
    const into = () => (hasFocus.current = true);
    // A cell taken out of the table (its row hidden, say) loses the focus too, but the grid keeps it.
    const out = (e: FocusEvent) => {
      const from = e.target as Node;
      if (!el.contains(e.relatedTarget as Node | null)) queueMicrotask(() => from.isConnected && (hasFocus.current = false));
    };
    el.addEventListener('focusin', into);
    el.addEventListener('focusout', out);
    return () => {
      el.removeEventListener('focusin', into);
      el.removeEventListener('focusout', out);
    };
  }, []);

  // Find, the name box and the change list point at cells: bring the selection into view, and take the focus when asked.
  const reveal = store.revealSeq;
  useEffect(() => {
    if (!reveal) return;
    pending.current = { row: store.selection.row, col: store.selection.col };
    if (store.revealFocus) {
      store.revealFocus = false;
      wantFocus.current = true;
    }
    redraw();
    // A grid just shown, for another sheet, may not take the focus at first in Safari.
    let tries = 10;
    const focus = () => {
      if (pending.current) scrollIntoView(pending.current.row, pending.current.col);
      pending.current = undefined;
      if (!wantFocus.current) return;
      activeCell.current?.focus({ preventScroll: true });
      if (host.current?.contains(document.activeElement) || --tries <= 0) wantFocus.current = false;
      else requestAnimationFrame(focus);
    };
    requestAnimationFrame(focus);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [reveal]);

  /** Where a heading's filter button opens its menu. */
  const filterButton = useCallback(
    (row: number, col: number): FilterButton | undefined => {
      const button = view.filterButtonAt(row, col);
      const box = button && handle.cellBox(row, col);
      return button && box ? { col, rect: button.rect, x: box.x, y: box.y + box.height } : undefined;
    },
    [view, handle],
  );

  // The cell being edited, if any.
  const [edit, setEdit] = useState<Edit | null>(null);
  const editing = useRef<Edit | null>(null);
  editing.current = edit;
  /** Navigation's view of the sheet, a page being what the grid shows. */
  const navNow = () => {
    const layout = layoutNow();
    const page = (run: { first: number; last: number } | undefined) => (run ? run.last - run.first : 1);
    return navSheet(view, page(layout?.rowRun), page(layout?.colRun));
  };
  /**
   * Edits the active cell (a merge's top-left cell): with `typed`, what
   * was typed in place of what it holds; else what it holds, selected.
   * A cell that can't be edited says why instead.
   */
  const startEdit = (typed?: string) => {
    const sel = store.selection;
    const m = view.mergeAt(sel.row, sel.col);
    const row = m?.r1 ?? sel.row;
    const col = m?.c1 ?? sel.col;
    const reason = view.readOnlyReason(row, col);
    if (reason) {
      store.toast('warn', reason);
      return;
    }
    const rec = view.get(row, col);
    const style = effectiveStyle(view, store.styles, row, col, rec ? rec.style : view.styleAt(row, col, null), palette.link);
    pending.current = { row, col };
    setEdit({
      row,
      col,
      text: typed ?? editText(rec, style, store.date1904, store.locale),
      typed: typed !== undefined,
      selectAll: typed === undefined,
      mode: typed === undefined ? 'edit' : 'enter',
      options: cellOptions(store, view, row, col),
      wrap: style.wrap,
    });
  };
  /** Ends editing: enters the text (unless undefined) and moves, within a selected range as Enter and Tab do there. */
  const finishEdit = (text: string | undefined, move: EditMove) => {
    const e = editing.current;
    if (!e) return;
    editing.current = null;
    setEdit(null);
    wantFocus.current = true;
    if (text !== undefined) void store.input(e.row, e.col, text);
    const [dx, dy] = move;
    if (!dx && !dy) return;
    const sheet = navNow();
    const back = dx < 0 || dy < 0;
    const moved =
      navigate(store.selection, { key: dy ? 'Enter' : 'Tab', shiftKey: back, altKey: false, command: false }, sheet) ??
      navigate(store.selection, { key: dy ? (dy > 0 ? 'ArrowDown' : 'ArrowUp') : dx > 0 ? 'ArrowRight' : 'ArrowLeft', shiftKey: false, altKey: false, command: false }, sheet);
    if (moved) {
      store.setSelection(moved.selection);
      pending.current = moved.reveal;
    }
  };
  // The editor stays over its cell as the grid scrolls.
  const [, scrolled] = useState(0);
  useEffect(() => (edit ? handle.subscribe(() => scrolled((n) => n + 1)) : undefined), [edit, handle]);
  /** The edited cell's box on the page: a merge's whole box. */
  const editBox = (e: Edit): Box | undefined => {
    const m = view.mergeAt(e.row, e.col);
    const a = handle.cellBox(e.row, e.col);
    const b = m ? handle.cellBox(m.r2, m.c2) : a;
    return a && b ? { x: a.x, y: a.y, width: b.x + b.width - a.x, height: b.y + b.height - a.y } : undefined;
  };
  // Text from anywhere but a key (dictation, an emoji picker) starts editing too; an input method's composition is the cell's own.
  useEffect(() => {
    const el = host.current;
    if (!el) return;
    const input = (e: InputEvent) => {
      if (!(e.target as Element).closest?.('.wbe-grid-access') || e.isComposing || e.inputType === 'insertCompositionText') return;
      e.preventDefault();
      if (e.inputType === 'insertText' && e.data) startEditRef.current(e.data);
    };
    el.addEventListener('beforeinput', input);
    return () => el.removeEventListener('beforeinput', input);
  }, []);
  const startEditRef = useRef(startEdit);
  startEditRef.current = startEdit;

  const onKeyDown = (e: ReactKeyboardEvent) => {
    if (editing.current || !(e.target as Element).closest?.('.wbe-grid-access')) return;
    if (editorKey(store, e, filterButton)) {
      e.preventDefault();
      return;
    }
    const moved = navigate(store.selection, { key: e.key, shiftKey: e.shiftKey, altKey: e.altKey, command: commandKey(e) }, navNow());
    if (moved) {
      e.preventDefault();
      if (moved.selection !== store.selection) store.setSelection(moved.selection);
      pending.current = moved.reveal;
      redraw();
      return;
    }
    // Enter and F2 edit what the cell holds; a key that types a character, or Backspace, replaces it, as Excel does.
    const command = e.ctrlKey || e.metaKey;
    const typesText = [...e.key].length === 1 && !command && (!e.altKey || APPLE || e.getModifierState('AltGraph'));
    // Shift+F2 is the note's key, not editing's.
    const opens = (e.key === 'Enter' || (e.key === 'F2' && !e.shiftKey)) && !command && !e.altKey;
    if (opens || typesText || (e.key === 'Backspace' && !APPLE && !command && !e.altKey && !e.shiftKey)) {
      e.preventDefault();
      startEdit(typesText ? e.key : e.key === 'Backspace' ? '' : undefined);
      return;
    }
    // Delete clears what is selected, whole rows and columns as far as they hold anything; so does a Mac's delete key.
    const plain = !e.shiftKey && !e.metaKey && !e.ctrlKey && !e.altKey;
    if (e.key === 'Delete' || (APPLE && e.key === 'Backspace' && plain)) {
      e.preventDefault();
      void store.clear(store.selection.rects);
    }
  };

  /** Where a pointer is on the canvas. */
  const local = (e: { clientX: number; clientY: number }) => {
    const at = canvas.current?.getBoundingClientRect();
    return { x: e.clientX - (at?.left ?? 0), y: e.clientY - (at?.top ?? 0) };
  };

  /**
   * Follows a drag until the button is let go, or Escape gives it up,
   * scrolling the grid while the pointer is past its edge (unless `still`).
   * A key pressed or let go meanwhile counts as a move, for a drag that
   * changes with Ctrl, Alt or Shift.
   */
  const follow = (onMove: (x: number, y: number, e: MouseEvent | KeyboardEvent) => void, onEnd?: (cancelled: boolean) => void, still = false) => {
    let pointer: { x: number; y: number; e: MouseEvent | KeyboardEvent } | undefined;
    let frameId = 0;
    const scroll = () => {
      const at = canvas.current?.getBoundingClientRect();
      const layout = layoutNow();
      if (pointer && at && layout) {
        // From the start edge, which is the right for a sheet read right to left.
        const x = rtl ? at.right - pointer.x : pointer.x - at.left;
        const y = pointer.y - at.top;
        const dx = x > layout.spec.width ? DRAG_STEP : x < layout.bodyX && layout.spec.scrollX > 0 ? -DRAG_STEP : 0;
        const dy = y > layout.spec.height ? DRAG_STEP : y < layout.bodyY && layout.spec.scrollY > 0 ? -DRAG_STEP : 0;
        if (dx || dy) {
          scroller.current?.scrollBy(rtl ? -dx : dx, dy);
          onMove(pointer.x, pointer.y, pointer.e);
        }
      }
      frameId = requestAnimationFrame(scroll);
    };
    if (!still) frameId = requestAnimationFrame(scroll);
    const move = (m: PointerEvent) => {
      pointer = { x: m.clientX, y: m.clientY, e: m };
      onMove(m.clientX, m.clientY, m);
    };
    const key = (k: KeyboardEvent) => {
      if (k.key === 'Escape') {
        k.preventDefault();
        k.stopPropagation();
        end(true);
      } else if (pointer) {
        pointer.e = k;
        onMove(pointer.x, pointer.y, k);
      }
    };
    const end = (cancelled: boolean) => {
      cancelAnimationFrame(frameId);
      window.removeEventListener('pointermove', move, true);
      window.removeEventListener('pointerup', up, true);
      window.removeEventListener('pointercancel', up, true);
      window.removeEventListener('keydown', key, true);
      window.removeEventListener('keyup', key, true);
      onEnd?.(cancelled);
    };
    const up = () => end(false);
    window.addEventListener('pointermove', move, true);
    window.addEventListener('pointerup', up, true);
    window.addEventListener('pointercancel', up, true);
    window.addEventListener('keydown', key, true);
    window.addEventListener('keyup', key, true);
  };

  /** A press on a cell: selects it, Shift reaching from the active cell, Ctrl (⌘) adding a range; a drag spans more. Merges are taken whole. */
  const pressCell = (e: ReactMouseEvent, row: number, col: number, drags = true) => {
    const sel = store.selection;
    const command = APPLE ? e.metaKey : e.ctrlKey;
    const box = handle.cellBox(row, col);
    const inside = (square: { x: number; y: number; size: number }, pad = 0) => {
      // Mirrored across the cell for a sheet read right to left, as it is drawn.
      const b = rtl && box ? { ...square, x: 2 * box.x + box.width - square.x - square.size } : square;
      return e.clientX >= b.x - pad && e.clientX <= b.x + b.size + pad && e.clientY >= b.y - pad && e.clientY <= b.y + b.size + pad;
    };
    // A filter button opens its column's filter menu, leaving the selection where it is.
    const button = view.filterButtonAt(row, col);
    if (button && box && inside(filterButtonBox(box))) {
      store.set('filterPicker', { col, rect: button.rect, x: box.x, y: box.y + box.height });
      return;
    }
    // Clicking a checkbox ticks or unticks it, and selects its cell, as in Excel.
    if (!command && !e.shiftKey && box && store.checkboxAt(row, col) !== undefined && inside(checkboxBox(box, store.styles[view.styleAt(row, col)] ?? store.styles[0]), 2)) void store.toggleCheckboxes({ row, col });
    // A second click on a list cell opens its dropdown.
    const only = sel.rects.length === 1 && sel.rects[0].r1 === sel.rects[0].r2 && sel.rects[0].c1 === sel.rects[0].c2;
    if (!command && !e.shiftKey && only && sel.row === row && sel.col === col && cellOptions(store, view, row, col)) {
      startEdit();
      return;
    }
    // A second tap with a finger edits the cell.
    if (!drags && store.touchMode && !command && !e.shiftKey && only && sel.row === row && sel.col === col) {
      startEdit();
      return;
    }
    // Ctrl+click (⌘-click on a Mac) follows a cell's link; a plain click only selects it.
    if (command && !e.shiftKey && cellLink(view, row, col)) {
      store.setSelection({ row, col, rects: [span({ row, col }, { row, col })] });
      void store.followLink(row, col);
      return;
    }
    const merged = view.mergeAt(row, col);
    const anchor = e.shiftKey ? { row: sel.row, col: sel.col } : merged ? { row: merged.r1, col: merged.c1 } : { row, col };
    const rest = command ? sel.rects : e.shiftKey ? sel.rects.slice(1) : [];
    const select = (to: { row: number; col: number }) => store.setSelection({ row: anchor.row, col: anchor.col, rects: [growOverMerges(span(anchor, to), view.info.merges), ...rest] });
    select({ row, col });
    if (!drags) return;
    follow((x, y) => {
      const to = handle.cellNear(x, y);
      if (to) select(to);
    });
  };

  /**
   * A press on the row numbers or column letters selects whole lines, as
   * Excel does: a click picks one, dragging spans several, Shift+click
   * reaches from the active line, and Ctrl+click (⌘-click on a Mac) adds one.
   */
  const pressHeading = (e: ReactMouseEvent, axis: 'rows' | 'cols', line: number) => {
    const sel = store.selection;
    const same = sel.lines === axis;
    const add = same && (APPLE ? e.metaKey : e.ctrlKey);
    const extend = same && e.shiftKey;
    const anchor = extend ? (axis === 'rows' ? sel.row : sel.col) : line;
    const kept: [number, number][] = add || extend ? sel.rects.map((r) => (axis === 'rows' ? [r.r1, r.r2] : [r.c1, r.c2])) : [];
    if (extend) kept.pop();
    const layout = layoutNow();
    const topRow = view.visibleRows[layout?.rowRun.first ?? 0] ?? 0;
    const active = axis === 'rows' ? { row: anchor, col: view.visibleCols[0] ?? 0 } : { row: topRow, col: anchor };
    const pick = (to: number) => store.selectLines(axis, [...kept, [anchor, to]], active);
    pick(line);
    follow((x, y) => {
      const at = local({ clientX: x, clientY: y });
      const near = layoutNow()?.hit(at.x, at.y, true);
      if (near) pick(axis === 'rows' ? view.visibleRows[near.row] : view.visibleCols[near.col]);
    });
  };

  /** What a press at a point on the canvas takes hold of before it selects: a line's edge, the fill handle, or the selection's border. */
  const grabAt = (x: number, y: number, layout: Layout): { kind: 'col'; col: number } | { kind: 'row'; row: number; edge: number } | { kind: 'fill' } | { kind: 'border' } | undefined => {
    if (!editable) return undefined;
    const { numbers, letters } = layout.spec;
    const runs = (a: { first: number; last: number }, b: { first: number; last: number }) => [a, b].flatMap((r) => Array.from({ length: Math.max(0, r.last - r.first + 1) }, (_, i) => r.first + i));
    // A column's right edge, among the letters; a row's bottom edge, among the numbers.
    if (y < letters && layout.hit(x, y).area === 'letters') {
      for (const gc of runs(layout.frozenColRun, layout.colRun)) {
        const b = layout.colHeading(gc);
        if (Math.abs(x - (rtl ? b.x : b.x + b.width)) <= 5) return { kind: 'col', col: view.visibleCols[gc] };
      }
    }
    if ((rtl ? x > layout.spec.width - numbers : x < numbers) && y >= letters) {
      for (const gr of runs(layout.frozenRowRun, layout.rowRun)) {
        const b = layout.rowHeading(gr);
        if (Math.abs(y - (b.y + b.height)) <= 3) return { kind: 'row', row: view.visibleRows[gr], edge: b.y + b.height };
      }
    }
    const sel = store.selection;
    const r = sel.rects[0];
    if (store.touchMode || sel.rects.length !== 1 || sel.lines === 'all' || !r || layout.hit(x, y).area !== 'cells') return undefined;
    const area = view.gridSpan(r);
    const b = area && layout.span(area.y1, area.y2, area.x1, area.x2);
    if (!b) return undefined;
    const [left, top, right, bottom] = [b.x, b.y, b.x + b.width, b.y + b.height];
    const near = (p: number, edge: number) => Math.abs(p - edge) <= 4;
    if (near(x, rtl ? left : right) && near(y, bottom)) return sel.lines ? undefined : { kind: 'fill' };
    const across = x >= left - 4 && x <= right + 4;
    const down = y >= top - 4 && y <= bottom + 4;
    return (across && (near(y, top) || near(y, bottom))) || (down && (near(x, left) || near(x, right))) ? { kind: 'border' } : undefined;
  };

  /** Drags a column's edge: the column takes the width as it goes, and keeps it once let go; a press alone leaves it. */
  const resizeColumn = (e: ReactMouseEvent, col: number) => {
    const start = e.clientX;
    const had = view.colWidths.get(col);
    const from = view.colWidth(col);
    let px = from;
    follow(
      (x) => {
        px = Math.min(1200, Math.max(18, Math.round(from + (rtl ? start - x : x - start))));
        view.colWidths.set(col, px);
        setSized((n) => n + 1);
      },
      (cancelled) => {
        if (cancelled || px === from) {
          if (had === undefined) view.colWidths.delete(col);
          else view.colWidths.set(col, had);
          setSized((n) => n + 1);
        } else void store.resizedColumn(col, px);
      },
      true,
    );
  };

  // Where a row's edge being dragged would go, in the host's pixels.
  const [rowGuide, setRowGuide] = useState<number | null>(null);
  /** Drags a row's bottom edge: a line shows where it would go, and the row takes that height once let go. */
  const resizeRow = (e: ReactMouseEvent, row: number, edge: number) => {
    const start = e.clientY;
    const from = view.rowHeight(row);
    let px = from;
    setRowGuide(edge);
    follow(
      (_x, y) => {
        px = Math.max(0, Math.round(from + y - start));
        setRowGuide(edge + px - from);
      },
      (cancelled) => {
        setRowGuide(null);
        if (!cancelled && px !== from) void store.resizedRow(row, px);
      },
      true,
    );
  };

  /** Drags the fill handle: the selection's cells continue down, up or across, whichever way the pointer went furthest, as far as it went. */
  const dragFill = () => {
    const sel = store.selection;
    const src = sel.rects[0];
    let dest = src;
    follow(
      (x, y) => {
        const at = handle.cellNear(x, y);
        if (!at) return;
        const down = at.row > src.r2 ? at.row - src.r2 : at.row < src.r1 ? at.row - src.r1 : 0;
        const across = at.col > src.c2 ? at.col - src.c2 : at.col < src.c1 ? at.col - src.c1 : 0;
        dest =
          Math.abs(down) >= Math.abs(across)
            ? { ...src, r1: Math.min(src.r1, at.row), r2: Math.max(src.r2, at.row) }
            : { ...src, c1: Math.min(src.c1, at.col), c2: Math.max(src.c2, at.col) };
        setLanding(dest.r1 === src.r1 && dest.r2 === src.r2 && dest.c1 === src.c1 && dest.c2 === src.c2 ? null : dest);
      },
      (cancelled) => {
        setLanding(null);
        if (cancelled || (dest.r1 === src.r1 && dest.r2 === src.r2 && dest.c1 === src.c1 && dest.c2 === src.c2)) return;
        void store.fill(src, dest).then(() => store.setSelection({ row: sel.row, col: sel.col, rects: [dest] }));
      },
    );
  };

  /** A double-click on the fill handle fills down as far as the data beside the selection goes, as Excel does. */
  const fillDown = () => {
    const sel = store.selection;
    const src = sel.rects[0];
    const filled = (row: number, col: number) => !!view.get(row, col);
    const beside = src.c1 > 0 && filled(src.r2 + 1, src.c1 - 1) ? src.c1 - 1 : filled(src.r2 + 1, src.c2 + 1) ? src.c2 + 1 : -1;
    if (beside < 0) return;
    let last = src.r2 + 1;
    while (filled(last + 1, beside)) last++;
    const dest = { ...src, r2: last };
    void store.fill(src, dest).then(() => store.setSelection({ row: sel.row, col: sel.col, rects: [dest] }));
  };

  // Where Shift+dragged cells would go in between others, in the host's pixels.
  const [insertGuide, setInsertGuide] = useState<Box | null>(null);
  /**
   * Drags the selection by its border: its cells move where it is dropped,
   * or with Ctrl (Option on a Mac) held, are copied there, as in Excel. A
   * dashed outline shows where they'll go. With Shift, they go in between
   * the cells there instead, at the cell edge nearest the pointer, which a
   * line shows. Escape gives up.
   */
  const dragSelection = (e: ReactMouseEvent) => {
    const src = store.selection.rects[0];
    const from = handle.cellNear(e.clientX, e.clientY);
    if (!src || !from) return;
    const off = { row: Math.min(Math.max(from.row - src.r1, 0), src.r2 - src.r1), col: Math.min(Math.max(from.col - src.c1, 0), src.c2 - src.c1) };
    const height = src.r2 - src.r1;
    const width = src.c2 - src.c1;
    let drop: Rect | undefined;
    let insert: { row: number; col: number; shift: 'down' | 'right' } | undefined;
    let copy = false;
    /** A block of cells on the grid, in the host's pixels. */
    const boxOf = (r: Rect) => {
      const a = handle.cellBox(r.r1, r.c1);
      const b = handle.cellBox(r.r2, r.c2);
      const h = host.current?.getBoundingClientRect();
      return a && b && h ? { left: a.x - h.left, top: a.y - h.top, right: b.x + b.width - h.left, bottom: b.y + b.height - h.top } : undefined;
    };
    follow(
      (x, y, m) => {
        const at = handle.cellNear(x, y);
        copy = APPLE ? m.altKey : m.ctrlKey;
        insert = undefined;
        let guide: Box | null = null;
        if (at) {
          const row = Math.max(0, at.row - off.row);
          const col = Math.max(0, at.col - off.col);
          drop = { r1: row, c1: col, r2: row + height, c2: col + width };
          const cell = m.shiftKey && !copy ? boxOf({ r1: at.row, c1: at.col, r2: at.row, c2: at.col }) : undefined;
          const h = host.current?.getBoundingClientRect();
          if (cell && h) {
            // The edge nearest the pointer, for the cell's shape: top or bottom puts the cells between rows, left or right between columns.
            // Across the cell from its start edge: the right one, for a sheet read right to left.
            const fx = rtl ? (cell.right - (x - h.left)) / (cell.right - cell.left) : (x - h.left - cell.left) / (cell.right - cell.left);
            const fy = (y - h.top - cell.top) / (cell.bottom - cell.top);
            const sideways = Math.min(fx, 1 - fx) < Math.min(fy, 1 - fy);
            insert = sideways ? { row, col: fx < 0.5 ? at.col : at.col + 1, shift: 'right' } : { row: fy < 0.5 ? at.row : at.row + 1, col, shift: 'down' };
            const line = sideways ? boxOf({ r1: row, c1: at.col, r2: row + height, c2: at.col }) : boxOf({ r1: at.row, c1: col, r2: at.row, c2: col + width });
            if (line)
              guide = sideways
                ? { x: (fx < 0.5) !== rtl ? line.left : line.right, y: line.top, width: 0, height: line.bottom - line.top }
                : { x: line.left, y: fy < 0.5 ? line.top : line.bottom, width: line.right - line.left, height: 0 };
          }
        }
        setInsertGuide(guide);
        setLanding(drop && !insert ? drop : null);
        if (host.current) host.current.style.cursor = copy ? 'copy' : 'move';
      },
      (cancelled) => {
        setLanding(null);
        setInsertGuide(null);
        if (host.current) host.current.style.cursor = '';
        if (cancelled) return;
        if (insert) void store.dragInsert(insert.row, insert.col, insert.shift);
        else if (drop && (drop.r1 !== src.r1 || drop.c1 !== src.c1)) void store.dragSelection(drop.r1, drop.c1, copy);
      },
    );
  };

  /**
   * A press with a mouse or a pen, or a tap with a finger (whose drags
   * scroll): on a line's edge, the fill handle or the selection's border it
   * takes hold of them; elsewhere it selects.
   */
  const press = (e: ReactMouseEvent<HTMLDivElement>, drags: boolean) => {
    const sc = scroller.current;
    const layout = layoutNow();
    const { x, y } = local(e);
    // Presses on the scrollbars scroll; they are at the left for a sheet read right to left.
    const bar = sc ? sc.offsetWidth - sc.clientWidth : 0;
    if (!sc || !layout || e.button !== 0 || (rtl ? x < bar : x >= sc.clientWidth) || y >= sc.clientHeight) return;
    e.preventDefault();
    store.selectObject(undefined);
    // The grid takes the focus, which then follows the active cell.
    activeCell.current?.focus({ preventScroll: true });
    wantFocus.current = true;
    // A line's edge resizes it, the fill handle fills and the selection's border moves it; a double-click on them is onDoubleClick's.
    const grab = grabAt(x, y, layout);
    if (grab?.kind === 'col') return resizeColumn(e, grab.col);
    if (grab?.kind === 'row') return resizeRow(e, grab.row, grab.edge);
    if (grab?.kind === 'fill') return dragFill();
    if (grab?.kind === 'border') return dragSelection(e);
    const hit = layout.hit(x, y);
    if (hit.area === 'corner') store.selectAll();
    else if (hit.area === 'cells') pressCell(e, view.visibleRows[hit.row], view.visibleCols[hit.col], drags);
    else if (hit.area === 'numbers') pressHeading(e, 'rows', view.visibleRows[hit.row]);
    else pressHeading(e, 'cols', view.visibleCols[hit.col]);
  };

  /**
   * A double-click on a line's edge fits the line to what it holds, on the
   * fill handle fills down beside data, as in Excel; on a cell, it edits it.
   */
  const onDoubleClick = (e: ReactMouseEvent) => {
    const layout = layoutNow();
    if (!layout) return;
    const { x, y } = local(e);
    const grab = grabAt(x, y, layout);
    if (grab?.kind === 'col' || grab?.kind === 'row') void store.autofit(grab.kind === 'col' ? 'cols' : 'rows', grab.kind === 'col' ? grab.col : grab.row);
    else if (grab?.kind === 'fill') fillDown();
    else if (!grab && layout.hit(x, y).area === 'cells') startEdit();
  };

  /** Right-clicking outside the selection selects what was clicked first, as Excel does, then opens the menu for it. */
  const onContextMenu = (e: ReactMouseEvent) => {
    e.preventDefault();
    // A long press has already opened it.
    if (performance.now() - longPressed.current < 1000) return;
    openMenu(e.clientX, e.clientY);
  };
  const openMenu = (clientX: number, clientY: number) => {
    const layout = layoutNow();
    if (!layout) return;
    const { x, y } = local({ clientX, clientY });
    const hit = layout.hit(x, y);
    const sel = store.selection;
    const row = view.visibleRows[hit.row];
    const col = view.visibleCols[hit.col];
    const at = { x: clientX, y: clientY };
    if (hit.area === 'cells') {
      if (!sel.rects.some((r) => inRect(r, row, col))) store.setSelection({ row, col, rects: [span({ row, col }, { row, col })] });
      store.set('menu', { ...at, kind: 'cell' });
    } else if (hit.area === 'numbers') {
      if (!sel.rects.some((r) => row >= r.r1 && row <= r.r2)) store.selectLines('rows', [[row, row]], { row, col: view.visibleCols[0] ?? 0 });
      store.set('menu', { ...at, kind: 'row' });
    } else if (hit.area === 'letters') {
      if (!sel.rects.some((r) => col >= r.c1 && col <= r.c2 && r.r1 === 0)) store.selectLines('cols', [[col, col]], { row: view.visibleRows[layout.rowRun.first] ?? 0, col });
      store.set('menu', { ...at, kind: 'column' });
    }
    activeCell.current?.focus({ preventScroll: true });
  };

  // A cell's note or link, shown while the pointer rests on it.
  const [hoverNote, setHoverNote] = useState<{ note: CellNote; x: number; y: number } | null>(null);
  const [hoverLink, setHoverLink] = useState<{ link: CellLink; x: number; y: number } | null>(null);
  const CURSORS = { col: 'col-resize', row: 'row-resize', fill: 'crosshair', border: 'move' } as const;
  const onMouseMove = (e: ReactMouseEvent) => {
    if (e.buttons) return;
    const layout = layoutNow();
    const { x, y } = local(e);
    const grab = layout && grabAt(x, y, layout);
    if (host.current) host.current.style.cursor = grab ? CURSORS[grab.kind] : '';
    const hit = layout?.hit(x, y);
    const row = hit?.area === 'cells' ? view.visibleRows[hit.row] : -1;
    const col = hit?.area === 'cells' ? view.visibleCols[hit.col] : -1;
    const note = row >= 0 ? view.noteAt(row, col) : undefined;
    const link = row >= 0 ? cellLink(view, row, col) : undefined;
    const box = row >= 0 ? handle.cellBox(row, col) : undefined;
    setHoverNote((prev) => (!note || !box ? null : prev?.note === note ? prev : { note, x: box.x + box.width + 4, y: box.y }));
    // Below the cell, as Excel shows where a link goes, so a note beside it stays clear.
    setHoverLink((prev) => (!link || !box ? null : prev && prev.x === box.x && prev.y === box.y + box.height + 4 ? prev : { link, x: box.x, y: box.y + box.height + 4 }));
  };
  const onMouseLeave = () => {
    setHoverNote(null);
    setHoverLink(null);
  };
  /** Notes whether the grid is being used with a finger, which shows the selection's handles; a touch away from them ends filling by dragging. */
  const onPointerDown = (e: ReactPointerEvent) => {
    const touch = e.pointerType === 'touch';
    if (store.touchMode !== touch) store.set('touchMode', touch);
    if (store.touchFill && !(e.target as Element).closest('.wbe-touch-handle')) store.set('touchFill', false);
  };

  // A finger held still on a cell for half a second opens the cell menu, as a right-click does; iOS doesn't say so itself.
  const longPress = useRef<{ timer: number; x: number; y: number } | undefined>(undefined);
  const longPressed = useRef(0);
  const endLongPress = () => {
    clearTimeout(longPress.current?.timer);
    longPress.current = undefined;
  };
  /** Mouse and pen presses come as pointer events; a finger's taps as the mouse events that follow them, so its drags scroll. */
  const onScrollerPointerDown = (e: ReactPointerEvent<HTMLDivElement>) => {
    if (e.pointerType !== 'touch') return press(e, true);
    const at = { x: e.clientX, y: e.clientY };
    endLongPress();
    longPress.current = {
      ...at,
      timer: window.setTimeout(() => {
        longPress.current = undefined;
        if (store.menu) return;
        longPressed.current = performance.now();
        openMenu(at.x, at.y);
      }, 500),
    };
  };
  const onScrollerPointerMove = (e: ReactPointerEvent) => {
    const held = longPress.current;
    if (held && Math.hypot(e.clientX - held.x, e.clientY - held.y) > 8) endLongPress();
  };
  const onScrollerMouseDown = (e: ReactMouseEvent<HTMLDivElement>) => {
    if (store.touchMode) press(e, false);
  };
  // An Apple Pencil selects, as a mouse does, rather than scrolling as a finger does.
  useEffect(() => {
    const sc = scroller.current;
    if (!sc) return;
    const stylus = (e: TouchEvent) => {
      // Safari says which touches are a stylus's.
      if ([...e.touches].some((t) => (t as Touch & { touchType?: string }).touchType === 'stylus')) e.preventDefault();
    };
    sc.addEventListener('touchstart', stylus, { passive: false });
    return () => sc.removeEventListener('touchstart', stylus);
  }, []);

  // What the screen readers' table needs, kept the same from render to render unless it changed.
  const textOf = useRef<(row: number, col: number) => string>(() => '');
  let tableInput: PaintInput | undefined;
  textOf.current = (row, col) => {
    const layout = drawn.current?.input.layout ?? layoutNow();
    if (!painter.current || !layout) return '';
    tableInput ??= inputFor(layout);
    return painter.current.cellText(tableInput, row, col);
  };
  // A line hidden or shown changes the view in place: layoutVersion says so.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const tableRows = useMemo(() => (view.gridRow(selection.row) >= 0 ? withLine(lines.rows, selection.row) : lines.rows), [view, lines.rows, selection.row, layoutVersion]);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  const tableCols = useMemo(() => (view.gridCol(selection.col) >= 0 ? withLine(lines.cols, selection.col) : lines.cols), [view, lines.cols, selection.col, layoutVersion]);
  const activeBox = useCallback(() => {
    const b = handle.cellBox(store.selection.row, store.selection.col);
    const at = host.current?.getBoundingClientRect();
    return b && at ? { ...b, x: b.x - at.left, y: b.y - at.top } : undefined;
  }, [handle, store]);
  const onComposed = useCallback((text: string) => startEditRef.current(text), []);
  const tableVersion = `${dataVersion}:${schemaVersion}:${layoutVersion}:${store.colVersion}:${store.highlightEdits}:${store.numbers.display}:${store.styles.length}:${palette.text}:${!!painter.current}`;
  return (
    <div
      className="wbe-grid-host"
      ref={host}
      data-testid="wbe-grid"
      onKeyDown={onKeyDown}
      onPointerDownCapture={onPointerDown}
    >
      <canvas className="wbe-grid-canvas" ref={canvas} aria-hidden="true" />
      <AccessTable
        store={store}
        view={view}
        selection={selection}
        textOf={textOf}
        version={tableVersion}
        rows={tableRows}
        cols={tableCols}
        activeRef={activeCell}
        frozenRows={view.info.frozenRows}
        typing={!store.touchMode && !edit}
        activeBox={activeBox}
        onComposed={onComposed}
      />
      {edit && <CellEditorBox store={store} edit={edit} box={editBox(edit)} palette={palette} rtl={rtl} onText={(text) => setEdit((e) => e && { ...e, text })} finish={finishEdit} onMode={() => setEdit((e) => e && { ...e, mode: e.mode === 'edit' ? 'enter' : 'edit' })} />}
      <div
        className="wbe-grid-scroller"
        ref={scroller}
        data-testid="wbe-grid-scroller"
        tabIndex={-1}
        style={rtl ? { direction: 'rtl' } : undefined}
        onPointerDown={onScrollerPointerDown}
        onPointerMove={onScrollerPointerMove}
        onPointerUp={endLongPress}
        onPointerCancel={endLongPress}
        onMouseDown={onScrollerMouseDown}
        onContextMenu={onContextMenu}
        onDoubleClick={onDoubleClick}
        onMouseMove={onMouseMove}
        onMouseLeave={onMouseLeave}
      >
        <div style={{ width: scrollLength(contentWidth), height: scrollLength(contentHeight) }} />
      </div>
      {hoverNote && <NoteTip {...hoverNote} />}
      {hoverLink && <LinkTip {...hoverLink} follows={store.followsLinks} />}
      {rowGuide !== null && <div className="wbe-row-guide" style={{ top: rowGuide - (host.current?.getBoundingClientRect().top ?? 0) + (canvas.current?.getBoundingClientRect().top ?? 0) }} />}
      {insertGuide && <div className="wbe-insert-guide" style={{ left: insertGuide.x, top: insertGuide.y, width: insertGuide.width, height: insertGuide.height }} />}
    </div>
  );
}
