import { inRect } from '@ruled/formula/refs';
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type MouseEvent as ReactMouseEvent, type PointerEvent as ReactPointerEvent } from 'react';
import type { CellNote, Rect } from '../../core/types';
import { effectiveStyle } from '../../state/cellStyle';
import { editText } from '../../state/format';
import { cellLink, type CellLink } from '../../state/links';
import { DEFAULT_PALETTE, DEFAULT_REF_COLORS, readGridPalette, readRefColors, type GridPalette } from '../../ui/gridPalette';
import { APPLE, commandKey } from '../../ui/keys';
import { LinkTip } from '../../ui/LinkTip';
import { NoteTip } from '../../ui/NoteTip';
import type { Box, EditMove, GridHandle, Pane, SurfaceProps } from '../contract';
import { useSheetEvents } from '../hooks';
import { editorKey, type FilterButton } from '../keys';
import { cellOptions } from '../lists';
import { sheetMarks } from '../marks';
import { navigate, navSheet } from '../navigation';
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
 * Ruled's own grid: the sheet drawn on a canvas, under a scroller whose
 * scrollbars, wheel, trackpad and touch scrolling are the browser's own.
 * Only the cells in view are drawn. Cells are selected with the mouse and
 * the keyboard, as in Excel; a table of the cells in view, whose active
 * cell takes the focus, is what screen readers read. Still to come:
 * editing, the fill handle and resizing.
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
  const axes = useMemo(
    () => ({ rows: rowAxis(view), cols: colAxis(view) }),
    // eslint-disable-next-line react-hooks/exhaustive-deps -- a view's sizes change in place
    [view, layoutVersion, store.colVersion, rowCount, colCount],
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
    });
  }, [axes, frozen, numbers]);

  const marks = useMemo(
    () => sheetMarks(store, view, selection, palette, refColors),
    // eslint-disable-next-line react-hooks/exhaustive-deps -- marks follow the cut, the formula being typed and the cells
    [store, view, selection, palette, refColors, store.cutRange, store.formulaRefs, dataVersion],
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
  const cellsKey = [view, dataVersion, schemaVersion, layoutVersion, store.colVersion, store.highlightEdits, store.numbers];
  const lastKey = useRef<unknown[]>([]);
  if (painter.current && cellsKey.some((k, i) => k !== lastKey.current[i])) painter.current.reset();
  lastKey.current = cellsKey;

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
      isLabel: schema ? (row, col) => schema.isLabel(sheet, row, col) : undefined,
      selection: store.selection,
      marks,
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
          freezeX: at.left + (fc.last >= 0 ? bodyX : numbers) - 1,
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
  }, [view, layoutNow, numbers, rowCount]);
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
        if (x !== undefined) sc.scrollLeft = scrollerPosition(Math.max(0, x), contentWidth, width);
      }
    },
    [view, layoutNow, contentHeight, contentWidth],
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

  /** Follows a drag until the button is let go, scrolling the grid while the pointer is past its edge. */
  const follow = (onMove: (x: number, y: number) => void) => {
    let pointer: { x: number; y: number } | undefined;
    let frameId = 0;
    const scroll = () => {
      const at = canvas.current?.getBoundingClientRect();
      const layout = layoutNow();
      if (pointer && at && layout) {
        const x = pointer.x - at.left;
        const y = pointer.y - at.top;
        const dx = x > layout.spec.width ? DRAG_STEP : x < layout.bodyX && layout.spec.scrollX > 0 ? -DRAG_STEP : 0;
        const dy = y > layout.spec.height ? DRAG_STEP : y < layout.bodyY && layout.spec.scrollY > 0 ? -DRAG_STEP : 0;
        if (dx || dy) {
          scroller.current?.scrollBy(dx, dy);
          onMove(pointer.x, pointer.y);
        }
      }
      frameId = requestAnimationFrame(scroll);
    };
    frameId = requestAnimationFrame(scroll);
    const move = (m: MouseEvent) => {
      pointer = { x: m.clientX, y: m.clientY };
      onMove(m.clientX, m.clientY);
    };
    const up = () => {
      cancelAnimationFrame(frameId);
      window.removeEventListener('mousemove', move, true);
      window.removeEventListener('mouseup', up, true);
    };
    window.addEventListener('mousemove', move, true);
    window.addEventListener('mouseup', up, true);
  };

  /** A press on a cell: selects it, Shift reaching from the active cell, Ctrl (⌘) adding a range; a drag spans more. */
  const pressCell = (e: ReactMouseEvent, row: number, col: number) => {
    const sel = store.selection;
    const command = APPLE ? e.metaKey : e.ctrlKey;
    // A second click on a list cell opens its dropdown.
    const only = sel.rects.length === 1 && sel.rects[0].r1 === sel.rects[0].r2 && sel.rects[0].c1 === sel.rects[0].c2;
    if (!command && !e.shiftKey && only && sel.row === row && sel.col === col && cellOptions(store, view, row, col)) {
      startEdit();
      return;
    }
    // Ctrl+click (⌘-click on a Mac) follows a cell's link; a plain click only selects it.
    if (command && !e.shiftKey && cellLink(view, row, col)) {
      store.setSelection({ row, col, rects: [span({ row, col }, { row, col })] });
      void store.followLink(row, col);
      return;
    }
    const anchor = e.shiftKey ? { row: sel.row, col: sel.col } : { row, col };
    const rest = command ? sel.rects : e.shiftKey ? sel.rects.slice(1) : [];
    const select = (to: { row: number; col: number }) => store.setSelection({ row: anchor.row, col: anchor.col, rects: [span(anchor, to), ...rest] });
    select({ row, col });
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
      const near = layoutNow()?.hit(local({ clientX: x, clientY: y }).x, local({ clientX: x, clientY: y }).y, true);
      if (near) pick(axis === 'rows' ? view.visibleRows[near.row] : view.visibleCols[near.col]);
    });
  };

  const onMouseDown = (e: ReactMouseEvent<HTMLDivElement>) => {
    const sc = scroller.current;
    const layout = layoutNow();
    const { x, y } = local(e);
    // Presses on the scrollbars scroll.
    if (!sc || !layout || e.button !== 0 || x >= sc.clientWidth || y >= sc.clientHeight) return;
    e.preventDefault();
    store.selectObject(undefined);
    // The grid takes the focus, which then follows the active cell.
    activeCell.current?.focus({ preventScroll: true });
    wantFocus.current = true;
    const hit = layout.hit(x, y);
    if (hit.area === 'corner') store.selectAll();
    else if (hit.area === 'cells') pressCell(e, view.visibleRows[hit.row], view.visibleCols[hit.col]);
    else if (hit.area === 'numbers') pressHeading(e, 'rows', view.visibleRows[hit.row]);
    else pressHeading(e, 'cols', view.visibleCols[hit.col]);
  };

  /** A double-click on a cell edits it. */
  const onDoubleClick = (e: ReactMouseEvent) => {
    const layout = layoutNow();
    const { x, y } = local(e);
    if (layout?.hit(x, y).area === 'cells') startEdit();
  };

  /** Right-clicking outside the selection selects what was clicked first, as Excel does, then opens the menu for it. */
  const onContextMenu = (e: ReactMouseEvent) => {
    const layout = layoutNow();
    if (!layout) return;
    e.preventDefault();
    const { x, y } = local(e);
    const hit = layout.hit(x, y);
    const sel = store.selection;
    const row = view.visibleRows[hit.row];
    const col = view.visibleCols[hit.col];
    const at = { x: e.clientX, y: e.clientY };
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
  const onMouseMove = (e: ReactMouseEvent) => {
    if (e.buttons) return;
    const layout = layoutNow();
    const { x, y } = local(e);
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

  const tableLayout = drawn.current?.input.layout;
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
        drawn={painter.current && tableLayout ? { painter: painter.current, input: inputFor(tableLayout) } : undefined}
        rows={view.gridRow(selection.row) >= 0 ? withLine(lines.rows, selection.row) : lines.rows}
        cols={view.gridCol(selection.col) >= 0 ? withLine(lines.cols, selection.col) : lines.cols}
        activeRef={activeCell}
        frozenRows={view.info.frozenRows}
        typing={!store.touchMode && !edit}
        activeBox={() => {
          const b = handle.cellBox(store.selection.row, store.selection.col);
          const at = host.current?.getBoundingClientRect();
          return b && at ? { ...b, x: b.x - at.left, y: b.y - at.top } : undefined;
        }}
        onComposed={(text) => startEdit(text)}
      />
      {edit && <CellEditorBox store={store} edit={edit} box={editBox(edit)} palette={palette} onText={(text) => setEdit((e) => e && { ...e, text })} finish={finishEdit} onMode={() => setEdit((e) => e && { ...e, mode: e.mode === 'edit' ? 'enter' : 'edit' })} />}
      <div
        className="wbe-grid-scroller"
        ref={scroller}
        data-testid="wbe-grid-scroller"
        tabIndex={-1}
        onMouseDown={onMouseDown}
        onContextMenu={onContextMenu}
        onDoubleClick={onDoubleClick}
        onMouseMove={onMouseMove}
        onMouseLeave={onMouseLeave}
      >
        <div style={{ width: scrollLength(contentWidth), height: scrollLength(contentHeight) }} />
      </div>
      {hoverNote && <NoteTip {...hoverNote} />}
      {hoverLink && <LinkTip {...hoverLink} follows={store.followsLinks} />}
    </div>
  );
}
