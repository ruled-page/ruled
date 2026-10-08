import type { DataEditorRef, GridKeyEventArgs, Rectangle } from '@glideapps/glide-data-grid';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { outlineBarSize } from '../state/outline';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { GLIDE_SCROLLER } from './glideDom';
import { lazyModule } from './lazyModule';
import type { OutlinePane } from './OutlineBar';
import { SheetGrid } from './SheetGrid';

/** The outline bars, loaded for the first sheet with groups. */
const outlineBars = lazyModule(() => import('./OutlineBar'));

interface Props {
  store: WorkbookStore;
  view: SheetView;
  dataVersion: number;
  schemaVersion: number;
  selection: Selection;
}

/**
 * The sheet's grid. Glide can freeze columns but not rows, so a sheet with
 * frozen rows gets two grids: the frozen rows (with the column letters) in a
 * fixed pane, and the rest scrolling beneath, kept in step horizontally. A
 * sheet with groups has outline bars beside the row numbers and above the
 * column letters.
 */
export function SheetPanes(props: Props) {
  const { store, view } = props;
  const [resized, setResized] = useState(0);
  const onColumnResized = useCallback(() => setResized((v) => v + 1), []);
  // Widths change by dragging here, or through the store from the keyboard.
  const colVersion = resized + store.colVersion;
  const top = useRef<DataEditorRef>(null);
  const bottom = useRef<DataEditorRef>(null);
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

  // The outline bars take room only on a sheet with groups. The grids are
  // always inside the same elements, so they aren't made again when the first
  // group is made or the last one cleared.
  const barWidth = outlineBarSize(view.info.outline?.rows);
  const barHeight = outlineBarSize(view.info.outline?.cols);
  const Bars = outlineBars.useModule(barWidth > 0 || barHeight > 0);
  const regions = useRef<(Rectangle | undefined)[]>([]);
  const regionListener = useRef<(() => void) | undefined>(undefined);
  const onRegionChange = useCallback((fn: () => void) => {
    regionListener.current = fn;
    return () => {
      if (regionListener.current === fn) regionListener.current = undefined;
    };
  }, []);
  const onTopRegion = useCallback((r: Rectangle) => {
    regions.current[0] = r;
    regionListener.current?.();
  }, []);
  const onBottomRegion = useCallback((r: Rectangle) => {
    regions.current[1] = r;
    regionListener.current?.();
  }, []);
  const rowCount = view.rowCount;
  const outlinePanes = useMemo<OutlinePane[]>(
    () => (frozen ? [{ ref: top, rowStart: 0, rowCount: frozen }, { ref: bottom, rowStart: frozen, rowCount: rowCount - frozen }] : [{ ref: top, rowStart: 0, rowCount }]),
    [frozen, rowCount],
  );

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
    (e: GridKeyEventArgs, key: string, fromRow: number, toRow: number, target: typeof top) => {
      const sel = store.selection;
      if (e.key !== key || e.shiftKey || e.ctrlKey || e.metaKey || view.gridRow(sel.row) !== fromRow) return;
      e.cancel();
      e.preventDefault();
      const row = view.visibleRows[toRow];
      store.setSelection({ row, col: sel.col, rects: [{ r1: row, r2: row, c1: sel.col, c2: sel.col }] });
      const gc = view.gridCol(sel.col) + 1;
      if (target === bottom) target.current?.scrollTo(gc, 0, 'vertical');
      // Focus once React has given that grid the new active cell: a grid
      // focused with none picks its own, in the row-number column.
      setTimeout(() => target.current?.focus(), 0);
    },
    [store, view],
  );
  const onTopKey = useCallback((e: GridKeyEventArgs) => cross(e, 'ArrowDown', frozen - 1, frozen, bottom), [cross, frozen]);
  const onBottomKey = useCallback((e: GridKeyEventArgs) => cross(e, 'ArrowUp', frozen, frozen - 1, top), [cross, frozen]);

  const grids = !frozen ? (
    <SheetGrid {...props} gridRef={top} onRegion={onTopRegion} colVersion={colVersion} onColumnResized={onColumnResized} />
  ) : (
    <div className="wbe-grid-panes" ref={panes}>
      <SheetGrid
        {...props}
        className="wbe-pane-top"
        style={{ height: topHeight }}
        rowStart={0}
        rowCount={frozen}
        gridRef={top}
        onKeyDown={onTopKey}
        onRegion={onTopRegion}
        colVersion={colVersion}
        onColumnResized={onColumnResized}
      />
      <SheetGrid
        {...props}
        className="wbe-pane-bottom"
        rowStart={frozen}
        header={false}
        gridRef={bottom}
        onKeyDown={onBottomKey}
        onRegion={onBottomRegion}
        colVersion={colVersion}
        onColumnResized={onColumnResized}
      />
    </div>
  );
  return (
    <div className="wbe-outline-frame">
      <div className="wbe-outline-body" style={{ left: barWidth, top: barHeight }}>
        {grids}
      </div>
      {Bars && (barWidth > 0 || barHeight > 0) && (
        <Bars.OutlineBars
          store={store}
          view={view}
          panes={outlinePanes}
          regions={regions}
          onRegionChange={onRegionChange}
          width={barWidth}
          height={barHeight}
          colVersion={colVersion}
        />
      )}
    </div>
  );
}
