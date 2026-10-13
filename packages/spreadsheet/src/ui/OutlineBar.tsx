import type { DataEditorRef, Rectangle } from '@glideapps/glide-data-grid';
import { useEffect, useLayoutEffect, useRef, useState, type MouseEvent, type RefObject } from 'react';
import { colToLetters } from '@ruled/formula/refs';
import type { OutlineAxis } from '../core/types';
import { useT } from '../i18n/react';
import { outlineDepth, outlineGroups, OUTLINE_LANE, summaryLine, type Axis, type OutlineGroup } from '../state/outline';
import type { SheetView } from '../state/sheetView';
import type { WorkbookStore } from '../state/store';

/** One of the sheet's grids: frozen rows have a grid of their own above the rest. */
export interface OutlinePane {
  ref: RefObject<DataEditorRef | null>;
  /** Its rows, as grid rows of the sheet view. */
  rowStart: number;
  rowCount: number;
}

/** A group's line beside the lines of it in view, along the bar; `cap` is its far end, away from the summary, when in view. */
interface Line {
  level: number;
  a: number;
  b: number;
  cap?: 'start' | 'end';
}

/** A group's button, at the middle of its summary line. */
interface Toggle {
  group: OutlineGroup;
  at: number;
  open: boolean;
}

/** The part of a bar beside one grid's rows (or the columns), which what scrolls out of it is clipped to. */
interface Strip {
  offset: number;
  size: number;
  lines: Line[];
  toggles: Toggle[];
}

interface Geometry {
  rows: Strip[];
  cols?: Strip;
  /** The middle of the column letters' row, where the row level buttons go, and of the row numbers' column, for the column level buttons. */
  rowLevelsAt?: number;
  colLevelsAt?: number;
}

/** The lines and buttons of the groups with lines in view, from where each line is (along the bar). */
function lay(view: SheetView, axis: Axis, o: OutlineAxis | undefined, boxes: Map<number, [number, number]>): Pick<Strip, 'lines' | 'toggles'> {
  const lines: Line[] = [];
  const toggles: Toggle[] = [];
  const keys = [...boxes.keys()];
  if (!keys.length) return { lines, toggles };
  const first = keys[0];
  const last = keys[keys.length - 1];
  const visible = axis === 'rows' ? view.visibleRows : view.visibleCols;
  const after = o?.summaryAfter ?? true;
  for (const g of outlineGroups(o)) {
    const s = summaryLine(o, g);
    if (Math.max(g.end, s) < first || Math.min(g.start, s) > last) continue;
    const span = view.lineSpan(axis, g.start, g.end);
    if (span) {
      let a = Infinity;
      let b = -Infinity;
      for (const [line, [x, y]] of boxes) {
        if (line < g.start || line > g.end) continue;
        a = Math.min(a, x);
        b = Math.max(b, y);
      }
      const capLine = visible[after ? span[0] : span[1]];
      if (a < b) lines.push({ level: g.level, a, b, cap: boxes.has(capLine) ? (after ? 'start' : 'end') : undefined });
    }
    const at = boxes.get(s);
    if (at) toggles.push({ group: g, at: (at[0] + at[1]) / 2, open: !!span });
  }
  return { lines, toggles };
}

/** Where everything in the bars goes, measured from the grids as they now are. */
function measure(el: HTMLElement, view: SheetView, panes: OutlinePane[], regions: (Rectangle | undefined)[], rows: boolean, cols: boolean): Geometry {
  const origin = el.getBoundingClientRect();
  const hosts = [...(el.parentElement?.querySelectorAll<HTMLElement>('.wbe-grid-host') ?? [])];
  const out: Geometry = { rows: [] };
  const head = panes[0]?.ref.current;
  const corner = head?.getBounds(0, -1);
  if (!head || !corner) return out;
  out.rowLevelsAt = corner.y - origin.top + corner.height / 2;
  out.colLevelsAt = corner.x - origin.left + corner.width / 2;
  if (rows) {
    panes.forEach((pane, i) => {
      const grid = pane.ref.current;
      const host = hosts[i]?.getBoundingClientRect();
      if (!grid || !host) return;
      const top = (i === 0 ? corner.y + corner.height : host.top) - origin.top;
      const r = regions[i];
      const from = Math.max(0, (r?.y ?? 0) - 1);
      const to = Math.min(pane.rowCount, r ? r.y + r.height + 2 : 100);
      // Glide's bounds take in the line after a cell, which belongs to the next.
      const boxes = new Map<number, [number, number]>();
      for (let gr = from; gr < to; gr++) {
        const b = grid.getBounds(0, gr);
        if (b) boxes.set(view.visibleRows[pane.rowStart + gr], [b.y - origin.top - top, b.y + b.height - 1 - origin.top - top]);
      }
      out.rows.push({ offset: top, size: host.bottom - origin.top - top, ...lay(view, 'rows', view.info.outline?.rows, boxes) });
    });
  }
  if (cols) {
    // Columns that scroll slide under the frozen ones, and the row numbers.
    let frozen = 0;
    for (let c = 0; c < view.info.frozenCols; c++) if (view.gridCol(c) >= 0) frozen++;
    const left = corner.x + corner.width - 1 - origin.left;
    const edge = head.getBounds(frozen, -1);
    const freezeX = frozen && edge ? edge.x + edge.width - 1 - origin.left : left;
    const r = regions[0];
    const shown = new Set<number>();
    for (let gc = 1; gc <= frozen; gc++) shown.add(gc);
    for (let gc = Math.max(frozen + 1, (r?.x ?? 0) - 1); gc <= Math.min(view.colCount, r ? r.x + r.width + 2 : 40); gc++) shown.add(gc);
    const boxes = new Map<number, [number, number]>();
    for (const gc of shown) {
      const b = head.getBounds(gc, -1);
      if (!b) continue;
      const a = gc > frozen ? Math.max(b.x - origin.left, freezeX) : b.x - origin.left;
      const e = b.x + b.width - 1 - origin.left;
      if (e > a) boxes.set(view.visibleCols[gc - 1], [a - left, e - left]);
    }
    out.cols = { offset: left, size: origin.width - left, ...lay(view, 'cols', view.info.outline?.cols, boxes) };
  }
  return out;
}

/**
 * Excel's outline bars: beside the row numbers for grouped rows, above the
 * column letters for grouped columns. Each group has a line along it and a
 * button on its summary row (or column) that collapses or expands it; the
 * numbered buttons show the outline to that level. They keep in line with
 * the grids as they scroll, frozen panes included.
 */
export function OutlineBars({
  store,
  view,
  panes,
  regions,
  onRegionChange,
  width,
  height,
  colVersion,
}: {
  store: WorkbookStore;
  view: SheetView;
  panes: OutlinePane[];
  /** Each pane's cells in view, as last told. */
  regions: RefObject<(Rectangle | undefined)[]>;
  /** Has `fn` called when they change, until the function it returns is called. */
  onRegionChange: (fn: () => void) => () => void;
  /** The bars' width beside the rows, and height above the columns; 0 for one not shown. */
  width: number;
  height: number;
  colVersion: number;
}) {
  const t = useT();
  const host = useRef<HTMLDivElement>(null);
  const [tick, setTick] = useState(0);
  const [geo, setGeo] = useState<Geometry>({ rows: [] });

  useEffect(() => {
    const changed = () => setTick((n) => n + 1);
    const stop = onRegionChange(changed);
    const el = host.current?.parentElement;
    const watch = el && typeof ResizeObserver !== 'undefined' ? new ResizeObserver(changed) : undefined;
    if (el) watch?.observe(el);
    return () => {
      stop();
      watch?.disconnect();
    };
  }, [onRegionChange]);

  const layoutVersion = view.layoutVersion;
  useLayoutEffect(() => {
    if (host.current) setGeo(measure(host.current, view, panes, regions.current ?? [], width > 0, height > 0));
  }, [view, layoutVersion, panes, regions, width, height, colVersion, tick]);

  // A click leaves the focus where it was, in the grid, as Excel's outline buttons do.
  const keep = (e: MouseEvent) => e.preventDefault();
  const bar = (axis: Axis) => {
    const rows = axis === 'rows';
    const o = view.info.outline?.[axis];
    const disabled = store.outlineLocked(axis);
    const depth = outlineDepth(o);
    const strips = rows ? geo.rows : geo.cols ? [geo.cols] : [];
    const levelsAt = rows ? geo.rowLevelsAt : geo.colLevelsAt;
    const lane = (level: number, extra: number) => (level - 1) * OUTLINE_LANE + extra;
    const along = (a: number) => (rows ? { top: a } : { left: a });
    const across = (level: number, extra: number) => (rows ? { left: lane(level, extra) } : { top: lane(level, extra) });
    const name = (g: OutlineGroup) =>
      rows
        ? g.start === g.end
          ? t('outline.bar.row', { row: g.start + 1 })
          : t('outline.bar.rowRange', { from: g.start + 1, to: g.end + 1 })
        : g.start === g.end
          ? t('outline.bar.column', { column: colToLetters(g.start) })
          : t('outline.bar.columnRange', { from: colToLetters(g.start), to: colToLetters(g.end) });
    return (
      <div className={`wbe-outline wbe-outline-${axis}`} role="group" aria-label={t(rows ? 'outline.bar.rows' : 'outline.bar.columns')} style={rows ? { width } : { height }}>
        {levelsAt !== undefined &&
          Array.from({ length: depth + 1 }, (_, i) => {
            const label = t(rows ? 'outline.bar.rowLevel' : 'outline.bar.columnLevel', { level: i + 1 });
            return (
              <button
                key={`level${i}`}
                type="button"
                className="wbe-outline-button wbe-outline-level"
                style={{ ...across(i + 1, 1), ...along(levelsAt - 7) }}
                aria-label={label}
                title={label}
                disabled={disabled}
                onMouseDown={keep}
                onClick={() => void store.showOutlineLevel(axis, i + 1)}
              >
                {i + 1}
              </button>
            );
          })}
        {strips.map((strip, i) => (
          <div key={i} className="wbe-outline-strip" style={rows ? { top: strip.offset, height: strip.size } : { left: strip.offset, width: strip.size }}>
            {strip.lines.map((l) => (
              <div
                key={`${l.level}:${l.a}`}
                className={`wbe-outline-line ${l.cap ? `wbe-outline-cap-${l.cap}` : ''}`}
                style={rows ? { left: lane(l.level, 7), top: l.a + 2, height: Math.max(0, l.b - l.a - 4) } : { top: lane(l.level, 7), left: l.a + 2, width: Math.max(0, l.b - l.a - 4) }}
              />
            ))}
            {strip.toggles.map(({ group: g, at, open }) => {
              const label = t(open ? 'outline.bar.collapse' : 'outline.bar.expand', { lines: name(g) });
              return (
                <button
                  key={`${g.level}:${g.start}`}
                  type="button"
                  className="wbe-outline-button"
                  style={{ ...across(g.level, 1), ...along(at - 7) }}
                  aria-label={label}
                  aria-expanded={open}
                  title={label}
                  disabled={disabled}
                  onMouseDown={keep}
                  onClick={() => void store.toggleGroup(axis, g)}
                >
                  {open ? '−' : '+'}
                </button>
              );
            })}
          </div>
        ))}
      </div>
    );
  };

  return (
    <div className="wbe-outline-bars" ref={host}>
      {width > 0 && bar('rows')}
      {height > 0 && bar('cols')}
    </div>
  );
}
