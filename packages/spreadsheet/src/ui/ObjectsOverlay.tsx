import { useEffect, useId, useLayoutEffect, useRef, useState, type CSSProperties, type KeyboardEvent as ReactKeyboardEvent, type MouseEvent as ReactMouseEvent, type WheelEvent as ReactWheelEvent } from 'react';
import type { SheetObject } from '../core/types';
import type { SheetView } from '../state/sheetView';
import type { WorkbookStore } from '../state/store';
import { ChartSlot } from './ChartSlot';
import { ContextMenu } from './ContextMenu';
import type { GridHandle } from '../grid/contract';
import { Icon } from './icons';
import type { MessageKey } from '../i18n/message';
import { useT } from '../i18n/react';
import { objectEditable, objectKey, objectSize, rowsHeight } from './objectGeometry';

interface Box {
  key: string;
  /** The object's own key, the same in every pane and as it moves. */
  id: string;
  obj: SheetObject;
  left: number;
  top: number;
  width: number;
  height: number;
  /** How much of its left side, or its right on a sheet read right to left, is under the frozen columns (or the row numbers) and not drawn. */
  clipLeft: number;
  clipRight: number;
  /** The lower part of an object anchored in the frozen rows above; that pane has the object itself. */
  continued: boolean;
  /** It can be moved, resized and deleted. */
  editable: boolean;
}

const LABEL: Record<SheetObject['kind'], MessageKey> = {
  chart: 'objects.chart',
  image: 'objects.picture',
  shape: 'objects.shape',
  pivot: 'objects.pivot',
  control: 'objects.control',
};

/** Where an object is kept but not drawn: a box with its name, so it is clear it exists and will be saved. */
function Placeholder({ box, style }: { box: Box; style: CSSProperties }) {
  const t = useT();
  const { obj } = box;
  const label = obj.title ?? t(LABEL[obj.kind]);
  return (
    <div
      className={`wbe-object-box wbe-object-${obj.kind} ${box.width < 150 || box.height < 56 ? 'wbe-compact' : ''}`}
      style={style}
      title={t('objects.title', { name: label })}
      aria-hidden="true"
    >
      <div className="wbe-object-label">
        <Icon name={obj.kind === 'chart' ? 'chart' : obj.kind === 'image' ? 'image' : 'shape'} />
        <span>
          {obj.title ?? (obj.chartType ? t('objects.typedChart', { type: t.text(obj.chartType) }) : label)}
          <em>{t('objects.kept')}</em>
        </span>
      </div>
    </div>
  );
}

/**
 * A picture, drawn in its box cropped, turned and flipped as Excel draws it;
 * its placeholder while its image loads, and where the browser can't show it.
 */
function Picture({ store, box, style }: { store: WorkbookStore; box: Box; style: CSSProperties }) {
  const pic = box.obj.picture!;
  // Undefined while loading, null where it can't be shown.
  const [url, setUrl] = useState<string | null>();
  useEffect(() => {
    let live = true;
    void store.picture(pic.part).then((u) => live && setUrl(u ?? null));
    return () => {
      live = false;
    };
  }, [store, pic.part]);
  if (!url) return <Placeholder box={box} style={style} />;
  const c = pic.crop ?? { l: 0, t: 0, r: 0, b: 0 };
  // The whole image, of which the box shows the part left after cropping.
  const width = Number(style.width) / Math.max(0.01, 1 - c.l - c.r);
  const height = Number(style.height) / Math.max(0.01, 1 - c.t - c.b);
  const transform = [pic.rotate && `rotate(${pic.rotate}deg)`, pic.flipH && 'scaleX(-1)', pic.flipV && 'scaleY(-1)'].filter(Boolean).join(' ');
  return (
    <div className="wbe-object-picture" style={{ ...style, transform: transform || undefined }} aria-hidden="true">
      <img src={url} alt="" draggable={false} onError={() => setUrl(null)} style={{ left: -c.l * width, top: -c.t * height, width, height }} />
    </div>
  );
}

/** A handle on a selected object's frame: the edges it moves, and where it sits. */
type Edge = 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw';
const HANDLES: Edge[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
/** Smallest an object is dragged down to, in pixels. */
const MIN_SIZE = 16;

/** A drag in progress: of an object's frame (`move`) or one of its handles, so far. */
interface Drag {
  key: string;
  /** The object as it was when the drag began; once the moved one replaces it, the drag no longer shifts it. */
  obj: SheetObject;
  mode: 'move' | Edge;
  x0: number;
  y0: number;
  dx: number;
  dy: number;
}

/** How far each edge of a box moves for a drag of `dx`, `dy`, keeping it at least MIN_SIZE across. */
function edgesFor(mode: Drag['mode'], dx: number, dy: number, width: number, height: number) {
  if (mode === 'move') return { left: dx, top: dy, right: dx, bottom: dy };
  const west = mode.includes('w') ? Math.min(dx, width - MIN_SIZE) : 0;
  const east = mode.includes('e') ? Math.max(dx, MIN_SIZE - width) : 0;
  const north = mode.includes('n') ? Math.min(dy, height - MIN_SIZE) : 0;
  const south = mode.includes('s') ? Math.max(dy, MIN_SIZE - height) : 0;
  return { left: west, top: north, right: east, bottom: south };
}

/**
 * The workbook's charts, pictures and shapes where they sit over the grid:
 * charts drawn from the cells as they are now (when the editor draws
 * charts), pictures, and placeholders for the rest, which are kept but not
 * drawn.
 * Clicking one selects it, as in Excel: its frame then moves it, its handles
 * resize it, and Delete removes it.
 */
export function ObjectsOverlay({
  store,
  view,
  handle,
  pane,
  colVersion,
}: {
  store: WorkbookStore;
  view: SheetView;
  handle: GridHandle;
  /** The band of rows this overlay covers (see Pane); objects anchored in a band above are continued here. */
  pane: number;
  colVersion: number;
}) {
  const t = useT();
  const id = useId();
  const host = useRef<HTMLDivElement>(null);
  const frame = useRef<HTMLDivElement>(null);
  const [boxes, setBoxes] = useState<Box[]>([]);
  const [area, setArea] = useState<CSSProperties>();
  const [drag, setDrag] = useState<Drag>();
  const [menu, setMenu] = useState<{ x: number; y: number; box: Box }>();
  const [tick, setTick] = useState(0);
  useEffect(() => handle.subscribe(() => setTick((n) => n + 1)), [handle]);

  const layoutVersion = view.layoutVersion;
  useLayoutEffect(() => {
    const el = host.current;
    const at = handle.view();
    const band = at.panes[pane];
    const parent = el?.parentElement?.getBoundingClientRect();
    if (!el || !band || !parent) return;
    // Objects stop at the band's scrollbars, as Excel's stop at its window's.
    const next = { left: band.box.x - parent.left, top: band.box.y - parent.top, width: band.box.width - band.scrollbars.right, height: band.box.height - band.scrollbars.bottom };
    setArea((prev) => (prev && (['left', 'top', 'width', 'height'] as const).every((k) => prev[k] === next[k]) ? prev : next));
    const origin = { left: band.box.x, top: band.box.y };
    const [firstRow, lastRow] = band.range;
    // Objects are placed from the first row in view, or from their own if it is further down.
    const scrolled = band.rows[0] ?? firstRow;
    // Objects in scrolling columns slide under the frozen ones (and the row numbers) as Excel's do.
    const freezeX = at.freezeX;
    const out: Box[] = [];
    view.info.objects.forEach((obj, i) => {
      if (!obj.from || obj.kind === 'pivot' || obj.kind === 'control') return;
      const size = objectSize(view, obj);
      if (view.gridCol(obj.from.col) < 0 || view.gridRow(obj.from.row) < 0 || !size || size.width < 8 || size.height < 8) return;
      if (obj.from.row > lastRow) return;
      // Anchored in the frozen rows above: this band shows whatever reaches below them.
      const above = obj.from.row < firstRow ? rowsHeight(view, obj.from.row, firstRow) - obj.from.rowOffPx : 0;
      if (above >= size.height) return;
      const refRow = Math.max(obj.from.row, scrolled);
      // Far above what's shown: nothing of it reaches the band.
      if (refRow - obj.from.row > 5000) return;
      const a = handle.cellBox(refRow, obj.from.col);
      if (!a) return;
      // Anchored from its cell's start edge: the right one, on a sheet read right to left.
      const left = at.rtl ? a.x + a.width - obj.from.colOffPx - size.width : a.x + obj.from.colOffPx;
      const scrolls = obj.from.col >= view.info.frozenCols;
      out.push({
        key: `${i}`,
        id: objectKey(obj),
        obj,
        left: left - origin.left,
        top: a.y - rowsHeight(view, obj.from.row, refRow) + obj.from.rowOffPx - origin.top,
        width: size.width,
        height: size.height,
        clipLeft: scrolls && !at.rtl ? Math.max(0, freezeX - left) : 0,
        clipRight: scrolls && at.rtl ? Math.max(0, left + size.width - freezeX) : 0,
        continued: obj.from.row < firstRow,
        editable: objectEditable(view.info, obj),
      });
    });
    setBoxes(out);
  }, [view, layoutVersion, handle, pane, tick, colVersion]);

  const selected = boxes.find((b) => b.id === store.selectedObject);
  // A newly selected object's frame takes the keyboard, in the pane that draws the object itself.
  useEffect(() => {
    if (selected && !selected.continued && !frame.current?.contains(document.activeElement)) frame.current?.focus({ preventScroll: true });
  }, [selected?.id, selected?.continued]); // eslint-disable-line react-hooks/exhaustive-deps

  /** Starts moving (or, from a handle, resizing) an object, selecting it; a drag shorter than a few pixels only selects. */
  const begin = (e: ReactMouseEvent, b: Box, mode: Drag['mode']) => {
    if (e.button !== 0 || !b.editable) return;
    e.preventDefault();
    e.stopPropagation();
    store.selectObject(b.id);
    const start: Drag = { key: b.id, obj: b.obj, mode, x0: e.clientX, y0: e.clientY, dx: 0, dy: 0 };
    let now = start;
    setDrag(start);
    const move = (ev: MouseEvent) => {
      now = { ...start, dx: ev.clientX - start.x0, dy: ev.clientY - start.y0 };
      setDrag(now);
    };
    const end = (ev: Event) => {
      window.removeEventListener('mousemove', move, true);
      window.removeEventListener('mouseup', end, true);
      window.removeEventListener('keydown', cancel, true);
      if (ev.type === 'mouseup' && Math.abs(now.dx) + Math.abs(now.dy) > 2) {
        // It stays where it was dropped until the moved object takes its place, rather than flashing back first.
        void store.placeObject(b.id, edgesFor(mode, now.dx, now.dy, b.width, b.height)).finally(() => setDrag(undefined));
      } else {
        setDrag(undefined);
      }
      frame.current?.focus({ preventScroll: true });
    };
    // Esc lets go without moving it.
    const cancel = (ev: KeyboardEvent) => {
      if (ev.key !== 'Escape') return;
      ev.preventDefault();
      ev.stopPropagation();
      end(ev);
    };
    window.addEventListener('mousemove', move, true);
    window.addEventListener('mouseup', end, true);
    window.addEventListener('keydown', cancel, true);
  };

  /** Excel's chart dialogs, for a chart. */
  const openEditor = (b: Box) => {
    if (b.editable && b.obj.kind === 'chart') store.set('chartDialog', { key: b.id });
  };

  /** Right-clicking an object selects it and offers what can be done with it. */
  const onMenu = (e: ReactMouseEvent, b: Box) => {
    if (!b.editable) return;
    e.preventDefault();
    e.stopPropagation();
    store.selectObject(b.id);
    setMenu({ x: e.clientX, y: e.clientY, box: b });
  };

  /** The wheel scrolls the sheet under an object, as it does elsewhere on the grid. */
  const wheel = (e: ReactWheelEvent) => handle.scrollBy(e.deltaX, e.deltaY);

  /** Keys on a selected object, as Excel has them: arrows move it, Shift and arrows resize it, Delete removes it, Tab goes on to the next and Esc back to the cells. */
  const onKey = (e: ReactKeyboardEvent, b: Box) => {
    const step = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
    if (step) {
      const [x, y] = step;
      void store.placeObject(b.id, e.shiftKey ? { left: 0, top: 0, right: x, bottom: y } : { left: x, top: y, right: x, bottom: y });
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      void store.deleteObject(b.id).then(() => store.focusGrid());
    } else if (e.key === 'Enter' && b.obj.kind === 'chart') {
      openEditor(b);
    } else if (e.key === 'ContextMenu' || (e.key === 'F10' && e.shiftKey)) {
      const r = (e.target as HTMLElement).getBoundingClientRect();
      setMenu({ x: r.left + 8, y: r.top + 8, box: b });
    } else if (e.key === 'Escape') {
      store.selectObject(undefined);
      store.focusGrid();
    } else if (e.key === 'Tab') {
      const list = store.editableObjects().map(objectKey);
      const at = list.indexOf(b.id);
      store.selectObject(list[(at + (e.shiftKey ? list.length - 1 : 1)) % list.length]);
    } else return;
    e.preventDefault();
    e.stopPropagation();
  };

  const placed = (b: Box): CSSProperties => {
    if (drag?.key !== b.id || drag.obj !== b.obj) return { left: b.left, top: b.top, width: b.width, height: b.height };
    const d = edgesFor(drag.mode, drag.dx, drag.dy, b.width, b.height);
    return { left: b.left + d.left, top: b.top + d.top, width: b.width + d.right - d.left, height: b.height + d.bottom - d.top };
  };

  return (
    <>
      <div className="wbe-objects-overlay" ref={host} style={area}>
        {boxes.map((b) => {
          const style: CSSProperties = { ...placed(b), clipPath: b.clipLeft || b.clipRight ? `inset(0 ${b.clipRight}px 0 ${b.clipLeft}px)` : undefined };
          const placeholder = <Placeholder box={b} style={style} />;
          return b.obj.kind === 'chart' ? (
            <ChartSlot key={b.key} store={store} obj={b.obj} placeholder={placeholder} className="wbe-object-box wbe-object-chart wbe-drawn" style={style} copy={b.continued} />
          ) : b.obj.picture ? (
            <Picture key={b.key} store={store} box={b} style={style} />
          ) : (
            <Placeholder key={b.key} box={b} style={style} />
          );
        })}
        {boxes
          .filter((b) => b.editable && b !== selected)
          .map((b) => (
            <div
              key={`hit-${b.key}`}
              className="wbe-object-hit"
              style={{ ...placed(b), clipPath: b.clipLeft || b.clipRight ? `inset(0 ${b.clipRight}px 0 ${b.clipLeft}px)` : undefined }}
              onMouseDown={(e) => begin(e, b, 'move')}
              onDoubleClick={() => openEditor(b)}
              onContextMenu={(e) => onMenu(e, b)}
              onWheel={wheel}
              aria-hidden="true"
            />
          ))}
        {selected && (
          <div
            ref={selected.continued ? undefined : frame}
            className="wbe-object-frame"
            style={placed(selected)}
            tabIndex={selected.continued ? -1 : 0}
            role="group"
            aria-label={t('objects.selected', { name: selected.obj.name })}
            aria-describedby={`${id}-hint`}
            onMouseDown={(e) => begin(e, selected, 'move')}
            onDoubleClick={() => openEditor(selected)}
            onContextMenu={(e) => onMenu(e, selected)}
            onWheel={wheel}
            onKeyDown={(e) => onKey(e, selected)}
          >
            {HANDLES.map((h) => (
              <span key={h} className={`wbe-object-handle wbe-handle-${h}`} onMouseDown={(e) => begin(e, selected, h)} aria-hidden="true" />
            ))}
            <span id={`${id}-hint`} className="wbe-sr-only">
              {t('objects.selectedHint')}
            </span>
          </div>
        )}
      </div>
      {menu && (
        <ContextMenu
          x={menu.x}
          y={menu.y}
          label={menu.box.obj.name}
          onClose={() => setMenu(undefined)}
          items={[
            ...(menu.box.obj.kind === 'chart' ? [{ label: t('charts.menu.edit'), onSelect: () => openEditor(menu.box) }] : []),
            { label: t('objects.menu.delete'), danger: true, onSelect: () => void store.deleteObject(menu.box.id).then(() => store.focusGrid()) },
          ]}
        />
      )}
    </>
  );
}
