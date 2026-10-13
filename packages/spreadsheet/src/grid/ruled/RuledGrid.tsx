import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { DEFAULT_PALETTE, readGridPalette, type GridPalette } from '../../ui/gridPalette';
import type { Box, GridHandle, Pane, SurfaceProps } from '../contract';
import { colAxis, rowAxis } from './axis';
import { Painter } from './draw';
import { contentScroll, Layout, scrollLength } from './layout';

/** The height of the column letters. */
const LETTERS = 24;

/**
 * Ruled's own grid: the sheet drawn on a canvas, under a scroller whose
 * scrollbars, wheel, trackpad and touch scrolling are the browser's own.
 * Only the cells in view are drawn. Still to come: selection, the
 * keyboard, editing and the table for screen readers.
 */
export function SheetSurface({ store, view, dataVersion, schemaVersion, onHandle }: SurfaceProps) {
  const host = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const scroller = useRef<HTMLDivElement>(null);
  const painter = useRef<Painter>(undefined);
  const [palette, setPalette] = useState<GridPalette>(DEFAULT_PALETTE);
  const themeVersion = store.themeVersion;
  useLayoutEffect(() => {
    const next = readGridPalette(host.current);
    setPalette((prev) => (JSON.stringify(prev) === JSON.stringify(next) ? prev : next));
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

  /** The layout as the grid is now scrolled and sized. */
  const layoutNow = useCallback((): Layout | undefined => {
    const sc = scroller.current;
    if (!sc) return undefined;
    const width = sc.clientWidth;
    const height = sc.clientHeight;
    const contentWidth = numbers + axes.cols.total;
    const contentHeight = LETTERS + axes.rows.total;
    return new Layout({
      width,
      height,
      rows: axes.rows,
      cols: axes.cols,
      frozenRows: frozen.rows,
      frozenCols: frozen.cols,
      numbers,
      letters: LETTERS,
      scrollX: contentScroll(Math.abs(sc.scrollLeft), contentWidth, width),
      scrollY: contentScroll(sc.scrollTop, contentHeight, height),
    });
  }, [axes, frozen, numbers]);

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
    const schema = store.schema;
    const sheet = view.info.index;
    painter.current.draw({
      view,
      layout,
      styles: store.styles,
      palette,
      date1904: store.date1904,
      locale: store.locale,
      fullNumbers: store.numbers.display === 'full',
      highlightEdits: store.highlightEdits,
      isLabel: schema ? (row, col) => schema.isLabel(sheet, row, col) : undefined,
    });
    // Near the last row or column shown, more follow, as far as Excel's sheet goes.
    const lastRow = layout.rowRun.last;
    const lastCol = layout.colRun.last;
    if (lastRow + 50 >= rowCount || lastCol + 3 >= colCount) store.scrolledTo(view, view.visibleRows[Math.max(0, lastRow)] ?? 0, view.visibleCols[Math.max(0, lastCol)] ?? 0);
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

  // What the painter worked out about cells goes once they, their styles or the columns change.
  useLayoutEffect(() => painter.current?.reset(), [view, dataVersion, schemaVersion, layoutVersion, store.colVersion, store.highlightEdits, store.numbers]);
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
        const lines = (from: number, to: number, of: Int32Array) => Array.from({ length: Math.max(0, to - from + 1) }, (_, i) => of[from + i]);
        const { frozenRowRun: fr, rowRun: rr, frozenColRun: fc, colRun: cr, bodyY, bodyX } = layout;
        const panes: Pane[] = [];
        if (fr.last >= 0) {
          panes.push({
            range: [view.visibleRows[0], view.visibleRows[fr.last]],
            rows: lines(fr.first, fr.last, view.visibleRows),
            box: { x: at.left, y: at.top + LETTERS, width: sc.offsetWidth, height: bodyY - LETTERS },
            scrollbars: { right: bars.right, bottom: 0 },
          });
        }
        panes.push({
          range: [view.visibleRows[fr.last + 1], view.visibleRows[rowCount - 1]],
          rows: lines(rr.first, rr.last, view.visibleRows),
          box: { x: at.left, y: at.top + bodyY, width: sc.offsetWidth, height: sc.offsetHeight - bodyY },
          scrollbars: bars,
        });
        return {
          corner: onPage(layout.corner()),
          panes,
          cols: [...lines(fc.first, fc.last, view.visibleCols), ...lines(cr.first, cr.last, view.visibleCols)],
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

  const contentWidth = numbers + axes.cols.total;
  const contentHeight = LETTERS + axes.rows.total;
  return (
    <div className="wbe-grid-host" ref={host} data-testid="wbe-grid">
      <canvas className="wbe-grid-canvas" ref={canvas} aria-hidden="true" />
      <div className="wbe-grid-scroller" ref={scroller} data-testid="wbe-grid-scroller">
        <div style={{ width: scrollLength(contentWidth), height: scrollLength(contentHeight) }} />
      </div>
    </div>
  );
}
