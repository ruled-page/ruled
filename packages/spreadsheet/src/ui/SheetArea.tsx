import { useRef, useState } from 'react';
import type { GridHandle } from '../grid/contract';
import { SheetSurface } from '../grid/SheetSurface';
import { outlineBarSize } from '../state/outline';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { lazyModule } from './lazyModule';
import { ObjectsOverlay } from './ObjectsOverlay';

/** The outline bars, loaded for the first sheet with groups. */
const outlineBars = lazyModule(() => import('./OutlineBar'));
/** The selection's handles for a finger, loaded once the grid is first touched. */
const touchHandles = lazyModule(() => import('./TouchHandles'));

/** The bands of rows the grid has: the frozen rows, when any are shown, then the rest. */
function paneCount(view: SheetView): number {
  for (let r = 0; r < view.info.frozenRows; r++) if (view.gridRow(r) >= 0) return view.rowCount > 1 ? 2 : 1;
  return 1;
}

interface Props {
  store: WorkbookStore;
  view: SheetView;
  dataVersion: number;
  schemaVersion: number;
  selection: Selection;
}

/**
 * A worksheet: drawn by the grid, with what goes over it (charts, pictures
 * and shapes, and touch handles) and, on a sheet with groups, outline bars
 * beside the row numbers and above the column letters.
 */
export function SheetArea(props: Props) {
  const { store, view } = props;
  const [handle, setHandle] = useState<GridHandle>();
  const body = useRef<HTMLDivElement>(null);

  // The outline bars take room only on a sheet with groups. The grid is
  // always inside the same elements, so it isn't made again when the first
  // group is made or the last one cleared.
  const barWidth = outlineBarSize(view.info.outline?.rows);
  const barHeight = outlineBarSize(view.info.outline?.cols);
  const Bars = outlineBars.useModule(barWidth > 0 || barHeight > 0);
  const Touch = touchHandles.useModule(store.touchMode);
  const panes = paneCount(view);

  return (
    <div className="wbe-outline-frame">
      <div className="wbe-outline-body" ref={body} style={view.info.rightToLeft ? { left: 0, right: barWidth, top: barHeight } : { left: barWidth, top: barHeight }}>
        <SheetSurface {...props} onHandle={setHandle} />
        {handle && Array.from({ length: panes }, (_, i) => <ObjectsOverlay key={i} store={store} view={view} handle={handle} pane={i} colVersion={store.colVersion} />)}
        {handle && store.touchMode && Touch && <Touch.default store={store} host={body} handle={handle} />}
      </div>
      {handle && Bars && (barWidth > 0 || barHeight > 0) && (
        <Bars.OutlineBars store={store} view={view} handle={handle} width={barWidth} height={barHeight} colVersion={store.colVersion} />
      )}
    </div>
  );
}
