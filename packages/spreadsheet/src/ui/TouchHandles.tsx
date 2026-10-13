import { useEffect, useState, type PointerEvent as ReactPointerEvent, type RefObject } from 'react';
import type { Rect } from '../core/types';
import type { GridHandle } from '../grid/contract';
import type { WorkbookStore } from '../state/store';

/** A block of cells on the grid, in the host's pixels. */
interface Box {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

interface Cell {
  row: number;
  col: number;
}

/** How close to the grid's edge a dragged handle scrolls it, and how far each frame. */
const EDGE = 32;
const STEP = 12;

const span = (a: Cell, b: Cell): Rect => ({ r1: Math.min(a.row, b.row), c1: Math.min(a.col, b.col), r2: Math.max(a.row, b.row), c2: Math.max(a.col, b.col) });

/**
 * On a touch screen, where a finger dragged over the grid scrolls it: round
 * handles on the selection's corners, dragged to select a range, and once
 * "Fill by dragging" is chosen from the cell menu, arrows on its bottom and
 * right edges, dragged to fill down or across, as Excel's phone app has
 * them. Keyboards have Shift and the arrow keys, and Fill down, instead.
 */
export default function TouchHandles({ store, host, handle }: { store: WorkbookStore; host: RefObject<HTMLDivElement | null>; handle: GridHandle }) {
  const [preview, setPreview] = useState<Rect>();
  const filling = store.touchFill;
  // The handles follow the cells as the grid scrolls.
  const [, moved] = useState(0);
  useEffect(() => handle.subscribe(() => moved((n) => n + 1)), [handle]);
  useEffect(() => {
    if (!filling) return;
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && store.set('touchFill', false);
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [filling, store]);

  const sel = store.selection;
  const rect = sel.rects[0];
  const el = host.current;
  /** Where a block of cells is, in the host's pixels. */
  const boxOf = (r: Rect): Box | undefined => {
    const a = handle.cellBox(r.r1, r.c1);
    const b = handle.cellBox(r.r2, r.c2);
    const h = el?.getBoundingClientRect();
    return a && b && h ? { left: a.x - h.left, top: a.y - h.top, right: b.x + b.width - h.left, bottom: b.y + b.height - h.top } : undefined;
  };
  if (!el || !rect || sel.rects.length !== 1 || sel.lines) return null;
  const box = boxOf(rect);
  if (!box) return null;
  const { width, height } = el.getBoundingClientRect();
  const shown = (x: number, y: number) => x >= 0 && y >= 0 && x <= width && y <= height;

  /** Follows a finger from a handle, the grid scrolling while it is held near an edge. */
  const drag = (e: ReactPointerEvent<HTMLElement>, move: (cell: Cell) => void, done: () => void) => {
    e.preventDefault();
    e.stopPropagation();
    const target = e.currentTarget;
    target.setPointerCapture(e.pointerId);
    let x = e.clientX;
    let y = e.clientY;
    const follow = () => {
      const cell = handle.cellNear(x, y);
      if (cell) move(cell);
    };
    let frame = 0;
    const scroll = () => {
      const h = el.getBoundingClientRect();
      const dx = x > h.right - EDGE ? STEP : x < h.left + EDGE ? -STEP : 0;
      const dy = y > h.bottom - EDGE ? STEP : y < h.top + EDGE ? -STEP : 0;
      if (dx || dy) {
        handle.scrollBy(dx, dy);
        follow();
      }
      frame = requestAnimationFrame(scroll);
    };
    frame = requestAnimationFrame(scroll);
    const onMove = (ev: PointerEvent) => {
      x = ev.clientX;
      y = ev.clientY;
      follow();
    };
    const end = () => {
      cancelAnimationFrame(frame);
      target.removeEventListener('pointermove', onMove);
      target.removeEventListener('pointerup', end);
      target.removeEventListener('pointercancel', end);
      done();
    };
    target.addEventListener('pointermove', onMove);
    target.addEventListener('pointerup', end);
    target.addEventListener('pointercancel', end);
  };

  /** A corner handle: the opposite corner stays put, and the active cell with it unless it is left outside. */
  const select = (e: ReactPointerEvent<HTMLElement>, anchor: Cell) =>
    drag(
      e,
      (cell) => {
        const r = span(anchor, cell);
        const keep = sel.row >= r.r1 && sel.row <= r.r2 && sel.col >= r.c1 && sel.col <= r.c2;
        store.setSelection({ row: keep ? sel.row : anchor.row, col: keep ? sel.col : anchor.col, rects: [r] });
      },
      () => undefined,
    );

  /** A fill arrow: the selection's cells continue down (or across) to where it is let go. */
  const fill = (e: ReactPointerEvent<HTMLElement>, across: boolean) => {
    let to = rect;
    drag(
      e,
      (cell) => {
        to = across ? { ...rect, c2: Math.max(rect.c2, cell.col) } : { ...rect, r2: Math.max(rect.r2, cell.row) };
        setPreview(to);
      },
      () => {
        setPreview(undefined);
        store.set('touchFill', false);
        if (to.r2 === rect.r2 && to.c2 === rect.c2) return;
        void store.fill(rect, to).then(() => store.setSelection({ row: sel.row, col: sel.col, rects: [to] }));
      },
    );
  };

  const outline = preview && boxOf(preview);
  const at = (x: number, y: number) => ({ left: x, top: y });
  return (
    <>
      {outline && (
        <div className="wbe-touch-preview" style={{ left: outline.left, top: outline.top, width: outline.right - outline.left, height: outline.bottom - outline.top }} />
      )}
      {filling ? (
        <>
          {shown((box.left + box.right) / 2, box.bottom) && (
            <span className="wbe-touch-handle wbe-touch-fill" style={at((box.left + box.right) / 2, box.bottom)} onPointerDown={(e) => fill(e, false)} aria-hidden="true">
              <svg viewBox="0 0 16 16" width="14" height="14">
                <path d="M8 3v9M4 8l4 4 4-4" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </span>
          )}
          {shown(box.right, (box.top + box.bottom) / 2) && (
            <span className="wbe-touch-handle wbe-touch-fill" style={at(box.right, (box.top + box.bottom) / 2)} onPointerDown={(e) => fill(e, true)} aria-hidden="true">
              <svg viewBox="0 0 16 16" width="14" height="14">
                <path d="M3 8h9M8 4l4 4-4 4" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </span>
          )}
        </>
      ) : (
        <>
          {shown(box.left, box.top) && <span className="wbe-touch-handle" style={at(box.left, box.top)} onPointerDown={(e) => select(e, { row: rect.r2, col: rect.c2 })} aria-hidden="true" />}
          {shown(box.right, box.bottom) && <span className="wbe-touch-handle" style={at(box.right, box.bottom)} onPointerDown={(e) => select(e, { row: rect.r1, col: rect.c1 })} aria-hidden="true" />}
        </>
      )}
    </>
  );
}
