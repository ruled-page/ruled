import type { DataEditorRef, GridKeyEventArgs, Rectangle } from '@glideapps/glide-data-grid';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { SheetView } from '../../state/sheetView';
import type { Box, GridHandle, Pane, SurfaceProps } from '../contract';
import { GLIDE_SCROLLER } from './glideDom';
import { SheetGrid, type GlidePane } from './SheetGrid';

const box = (r: Rectangle | undefined): Box | undefined => r && { x: r.x, y: r.y, width: r.width, height: r.height };

/** The room a grid's scrollbars take; scrollbars drawn over the grid, as macOS draws them, take none of their own. */
function scrollbars(scroller: HTMLElement | null): Pane['scrollbars'] {
  if (!scroller || getComputedStyle(scroller).scrollbarWidth === 'none') return { right: 0, bottom: 0 };
  const bar = (room: number, overflows: boolean) => (room > 0 ? room : overflows ? 10 : 0);
  return {
    right: bar(scroller.offsetWidth - scroller.clientWidth, scroller.scrollHeight > scroller.clientHeight),
    bottom: bar(scroller.offsetHeight - scroller.clientHeight, scroller.scrollWidth > scroller.clientWidth),
  };
}

/** The sheet's handle over its Glide grids: one, or two with the frozen rows in the first. */
function glideHandle(view: SheetView, panes: () => GlidePane[], listeners: Set<() => void>): GridHandle {
  /** The grid holding a sheet row. */
  const paneOf = (row: number) => {
    const g = view.gridRow(row);
    return g < 0 ? undefined : panes().find((p) => g >= p.rowStart && g < p.rowStart + p.rowCount);
  };
  const frozenCols = () => {
    let n = 0;
    for (let c = 0; c < view.info.frozenCols; c++) if (view.gridCol(c) >= 0) n++;
    return n;
  };
  /** Before Glide first says what is in view, as many lines as a screen might hold. */
  const regionOf = (p: GlidePane) => (p.region.width || p.region.height ? p.region : { x: 1, y: 0, width: 40, height: 100 });
  return {
    view() {
      const list = panes();
      const head = list[0];
      const frozen = frozenCols();
      const out = list.map((p): Pane => {
        const el = p.host;
        const rect = el.getBoundingClientRect();
        const top = p.header ? 24 : 0;
        const r = regionOf(p);
        const rows: number[] = [];
        for (let gr = r.y; gr < Math.min(p.rowCount, r.y + r.height); gr++) rows.push(view.visibleRows[p.rowStart + gr]);
        return {
          range: [view.visibleRows[p.rowStart], view.visibleRows[p.rowStart + p.rowCount - 1]],
          rows,
          box: { x: rect.left + el.clientLeft, y: rect.top + el.clientTop + top, width: el.clientWidth, height: el.clientHeight - top },
          scrollbars: scrollbars(el.querySelector<HTMLElement>(`.${GLIDE_SCROLLER}`)),
        };
      });
      const cols: number[] = [];
      for (let gc = 1; gc <= frozen; gc++) cols.push(view.visibleCols[gc - 1]);
      if (head) {
        const r = regionOf(head);
        for (let gc = Math.max(frozen + 1, r.x); gc <= Math.min(view.colCount, r.x + r.width - 1); gc++) cols.push(view.visibleCols[gc - 1]);
      }
      // Glide's bounds take in the line after a cell, which belongs to the next.
      const edge = head?.grid.getBounds(frozen, 0);
      return {
        corner: head?.header ? box(head.grid.getBounds(0, -1)) : undefined,
        panes: out,
        cols,
        frozenCols: frozen,
        freezeX: edge ? edge.x + edge.width - 1 : (out[0]?.box.x ?? 0),
      };
    },
    cellBox(row, col) {
      const p = paneOf(row);
      const gc = view.gridCol(col);
      return p && gc >= 0 ? box(p.grid.getBounds(gc + 1, view.gridRow(row) - p.rowStart)) : undefined;
    },
    headingBox(axis, line) {
      if (axis === 'rows') {
        const p = paneOf(line);
        return p && box(p.grid.getBounds(0, view.gridRow(line) - p.rowStart));
      }
      const head = panes()[0];
      const gc = view.gridCol(line);
      return head?.header && gc >= 0 ? box(head.grid.getBounds(gc + 1, -1)) : undefined;
    },
    cellNear(x, y) {
      const list = panes();
      // The grid the point is level with, or the nearest.
      const p = list.find((q, i) => i === list.length - 1 || y < q.host.getBoundingClientRect().bottom);
      return p?.cellNear(x, y);
    },
    scrollBy(dx, dy) {
      const list = panes();
      list[list.length - 1]?.host.querySelector<HTMLElement>(`.${GLIDE_SCROLLER}`)?.scrollBy({ left: dx, top: dy });
    },
    subscribe(fn) {
      listeners.add(fn);
      return () => listeners.delete(fn);
    },
  };
}

/**
 * The sheet in Glide. Glide can freeze columns but not rows, so a sheet with
 * frozen rows gets two grids: the frozen rows (with the column letters) in a
 * fixed pane, and the rest scrolling beneath, kept in step horizontally.
 */
export function SheetSurface({ onHandle, ...props }: SurfaceProps) {
  const { store, view } = props;
  const [resized, setResized] = useState(0);
  const listeners = useRef(new Set<() => void>());
  const notify = useCallback(() => {
    for (const fn of listeners.current) fn();
  }, []);
  const onColumnResized = useCallback(() => {
    setResized((v) => v + 1);
    notify();
  }, [notify]);
  // Widths change by dragging here, or through the store from the keyboard.
  const colVersion = resized + store.colVersion;
  const top = useRef<GlidePane>(null);
  const bottom = useRef<GlidePane>(null);
  const topGrid = useRef<DataEditorRef>(null);
  const bottomGrid = useRef<DataEditorRef>(null);
  const panes = useRef<HTMLDivElement>(null);

  /** Frozen rows that are visible, i.e. not hidden. */
  const layoutVersion = view.layoutVersion;
  const frozen = useMemo(() => {
    let n = 0;
    for (let r = 0; r < view.info.frozenRows; r++) if (view.gridRow(r) >= 0) n++;
    return Math.min(n, view.rowCount - 1);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [view, layoutVersion]);
  const topHeight = useMemo(() => {
    let h = 24;
    for (let g = 0; g < frozen; g++) h += view.rowHeight(view.visibleRows[g]);
    return h;
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [view, frozen, layoutVersion]);

  const handle = useMemo(() => glideHandle(view, () => [top.current, bottom.current].filter((p): p is GlidePane => !!p), listeners.current), [view]);
  useEffect(() => {
    onHandle(handle);
    return () => onHandle(undefined);
  }, [handle, onHandle]);
  // A pane appearing or going changes what the handle answers.
  useEffect(notify, [notify, frozen]);

  // Scrolling either pane sideways scrolls the other. Scroll events don't
  // bubble, but a capturing listener here still sees each grid's scroller.
  useEffect(() => {
    const el = panes.current;
    if (!frozen || !el) return;
    let lead: EventTarget | null = null;
    const onScroll = (e: Event) => {
      const from = e.target as HTMLElement;
      if (!from.classList?.contains(GLIDE_SCROLLER) || (lead && lead !== from)) return;
      const other = [...el.querySelectorAll<HTMLElement>(`.${GLIDE_SCROLLER}`)].find((s) => s !== from);
      if (!other || other.scrollLeft === from.scrollLeft) return;
      lead = from;
      other.scrollLeft = from.scrollLeft;
      requestAnimationFrame(() => (lead = null));
    };
    el.addEventListener('scroll', onScroll, { capture: true, passive: true });
    return () => el.removeEventListener('scroll', onScroll, { capture: true });
  }, [view, frozen]);

  /** Moves the active cell to another pane with the arrow keys, as if the grid were one. */
  const cross = useCallback(
    (e: GridKeyEventArgs, key: string, fromRow: number, toRow: number, target: typeof topGrid) => {
      const sel = store.selection;
      if (e.key !== key || e.shiftKey || e.ctrlKey || e.metaKey || view.gridRow(sel.row) !== fromRow) return;
      e.cancel();
      e.preventDefault();
      const row = view.visibleRows[toRow];
      store.setSelection({ row, col: sel.col, rects: [{ r1: row, r2: row, c1: sel.col, c2: sel.col }] });
      const gc = view.gridCol(sel.col) + 1;
      if (target === bottomGrid) target.current?.scrollTo(gc, 0, 'vertical');
      // Focus once React has given that grid the new active cell: a grid
      // focused with none picks its own, in the row-number column.
      setTimeout(() => target.current?.focus(), 0);
    },
    [store, view],
  );
  const onTopKey = useCallback((e: GridKeyEventArgs) => cross(e, 'ArrowDown', frozen - 1, frozen, bottomGrid), [cross, frozen]);
  const onBottomKey = useCallback((e: GridKeyEventArgs) => cross(e, 'ArrowUp', frozen, frozen - 1, topGrid), [cross, frozen]);

  return !frozen ? (
    <SheetGrid {...props} gridRef={topGrid} paneRef={top} onRegion={notify} colVersion={colVersion} onColumnResized={onColumnResized} />
  ) : (
    <div className="wbe-grid-panes" ref={panes}>
      <SheetGrid
        {...props}
        className="wbe-pane-top"
        style={{ height: topHeight }}
        rowStart={0}
        rowCount={frozen}
        gridRef={topGrid}
        paneRef={top}
        onKeyDown={onTopKey}
        onRegion={notify}
        colVersion={colVersion}
        onColumnResized={onColumnResized}
      />
      <SheetGrid
        {...props}
        className="wbe-pane-bottom"
        rowStart={frozen}
        header={false}
        gridRef={bottomGrid}
        paneRef={bottom}
        onKeyDown={onBottomKey}
        onRegion={notify}
        colVersion={colVersion}
        onColumnResized={onColumnResized}
      />
    </div>
  );
}
