import type { DataEditorRef } from '@glideapps/glide-data-grid';
import { useLayoutEffect, useRef, useState, type CSSProperties, type RefObject } from 'react';
import type { SheetObject } from '../core/types';
import type { SheetView } from '../state/sheetView';
import type { WorkbookStore } from '../state/store';
import { ChartSlot } from './ChartSlot';
import { Icon } from './icons';
import type { MessageKey } from '../i18n/message';
import { useT } from '../i18n/react';
import { objectSize, rowsHeight } from './objectGeometry';

interface Box {
  key: string;
  obj: SheetObject;
  left: number;
  top: number;
  width: number;
  height: number;
  /** How much of its left side is under the frozen columns (or the row numbers) and not drawn. */
  clipLeft: number;
  /** The lower part of an object anchored in the frozen rows above; that pane has the object itself. */
  continued: boolean;
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
 * The workbook's charts, pictures and shapes where they sit over the grid:
 * charts drawn from the cells as they are now (when the editor draws
 * charts), and placeholders for the rest, which are kept but not drawn.
 */
export function ObjectsOverlay({
  store,
  view,
  gridRef,
  tick,
  rowStart,
  rowCount,
  top,
}: {
  store: WorkbookStore;
  view: SheetView;
  gridRef: RefObject<DataEditorRef | null>;
  tick: number;
  /** The grid rows this pane shows; objects anchored above them are drawn by the pane above, and continued here. */
  rowStart: number;
  rowCount: number;
  /** Height of the column header above the cells. */
  top: number;
}) {
  const host = useRef<HTMLDivElement>(null);
  const [boxes, setBoxes] = useState<Box[]>([]);

  useLayoutEffect(() => {
    const grid = gridRef.current;
    const el = host.current;
    if (!grid || !el) return;
    const origin = el.getBoundingClientRect();
    const firstRow = view.visibleRows[rowStart];
    // Objects in scrolling columns slide under the frozen ones (and the row numbers) as Excel's do.
    let frozenCols = 0;
    for (let c = 0; c < view.info.frozenCols; c++) if (view.gridCol(c) >= 0) frozenCols++;
    const edge = grid.getBounds(frozenCols, 0);
    const freezeX = edge ? edge.x + edge.width - 1 : origin.left;
    const out: Box[] = [];
    view.info.objects.forEach((obj, i) => {
      if (!obj.from || obj.kind === 'pivot' || obj.kind === 'control') return;
      const gc = view.gridCol(obj.from.col);
      const g = view.gridRow(obj.from.row);
      const size = objectSize(view, obj);
      if (gc < 0 || g < 0 || !size || size.width < 8 || size.height < 8) return;
      const gr = g - rowStart;
      if (gr >= rowCount) return;
      // Anchored in the frozen rows above: this pane shows whatever reaches below them.
      const above = gr < 0 ? rowsHeight(view, obj.from.row, firstRow) - obj.from.rowOffPx : 0;
      if (above >= size.height) return;
      const a = grid.getBounds(gc + 1, Math.max(gr, 0));
      if (!a) return;
      const left = a.x + obj.from.colOffPx;
      out.push({
        key: `${i}`,
        obj,
        left: left - origin.left,
        top: (gr < 0 ? a.y - above : a.y + obj.from.rowOffPx) - origin.top,
        width: size.width,
        height: size.height,
        clipLeft: obj.from.col >= view.info.frozenCols ? Math.max(0, freezeX - left) : 0,
        continued: gr < 0,
      });
    });
    setBoxes(out);
  }, [view, gridRef, tick, rowStart, rowCount]);

  return (
    <div className="wbe-objects-overlay" ref={host} style={{ top }}>
      {boxes.map((b) => {
        const style: CSSProperties = {
          left: b.left,
          top: b.top,
          width: b.width,
          height: b.height,
          clipPath: b.clipLeft ? `inset(0 0 0 ${b.clipLeft}px)` : undefined,
        };
        const placeholder = <Placeholder box={b} style={style} />;
        return b.obj.kind === 'chart' ? (
          <ChartSlot key={b.key} store={store} obj={b.obj} placeholder={placeholder} className="wbe-object-box wbe-object-chart wbe-drawn" style={style} copy={b.continued} />
        ) : (
          <Placeholder key={b.key} box={b} style={style} />
        );
      })}
    </div>
  );
}
