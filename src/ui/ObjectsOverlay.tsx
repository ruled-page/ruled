import type { DataEditorRef } from '@glideapps/glide-data-grid';
import { useLayoutEffect, useRef, useState, type RefObject } from 'react';
import type { SheetObject } from '../core/types';
import type { SheetView } from '../state/sheetView';
import { Icon } from './icons';
import type { MessageKey } from '../i18n/message';
import { useT } from '../i18n/react';

interface Box {
  key: string;
  obj: SheetObject;
  left: number;
  top: number;
  width: number;
  height: number;
}

const LABEL: Record<SheetObject['kind'], MessageKey> = {
  chart: 'objects.chart',
  image: 'objects.picture',
  shape: 'objects.shape',
  pivot: 'objects.pivot',
  control: 'objects.control',
};

/**
 * Placeholders where the workbook's charts, pictures and shapes sit, so it
 * is clear they exist and will be saved, even though they are not drawn.
 */
export function ObjectsOverlay({
  view,
  gridRef,
  tick,
  rowStart,
  rowCount,
  top,
}: {
  view: SheetView;
  gridRef: RefObject<DataEditorRef | null>;
  tick: number;
  /** The grid rows this pane shows; objects anchored outside them are left to the other pane. */
  rowStart: number;
  rowCount: number;
  /** Height of the column header above the cells. */
  top: number;
}) {
  const t = useT();
  const host = useRef<HTMLDivElement>(null);
  const [boxes, setBoxes] = useState<Box[]>([]);

  useLayoutEffect(() => {
    const grid = gridRef.current;
    const el = host.current;
    if (!grid || !el) return;
    const origin = el.getBoundingClientRect();
    const out: Box[] = [];
    view.info.objects.forEach((obj, i) => {
      if (!obj.from || obj.kind === 'pivot' || obj.kind === 'control') return;
      const gc = view.gridCol(obj.from.col);
      const gr = view.gridRow(obj.from.row) - rowStart;
      if (gc < 0 || gr < 0 || gr >= rowCount) return;
      const a = grid.getBounds(gc + 1, gr);
      if (!a) return;
      let width = obj.sizePx?.w ?? 0;
      let height = obj.sizePx?.h ?? 0;
      if (obj.to) {
        const tc = view.gridCol(obj.to.col);
        const tr = view.gridRow(obj.to.row) - rowStart;
        const b = tc >= 0 && tr >= 0 ? grid.getBounds(tc + 1, tr) : undefined;
        if (b) {
          width = b.x + obj.to.colOffPx - (a.x + obj.from.colOffPx);
          height = b.y + obj.to.rowOffPx - (a.y + obj.from.rowOffPx);
        }
      }
      if (width < 8 || height < 8) return;
      out.push({
        key: `${i}`,
        obj,
        left: a.x + obj.from.colOffPx - origin.left,
        top: a.y + obj.from.rowOffPx - origin.top,
        width,
        height,
      });
    });
    setBoxes(out);
  }, [view, gridRef, tick, rowStart, rowCount]);

  return (
    <div className="objects-overlay" ref={host} aria-hidden="true" style={{ top }}>
      {boxes.map((b) => {
        const label = b.obj.title ?? t(LABEL[b.obj.kind]);
        return (
          <div
            key={b.key}
            className={`object-box object-${b.obj.kind} ${b.width < 150 || b.height < 56 ? 'compact' : ''}`}
            style={{ left: b.left, top: b.top, width: b.width, height: b.height }}
            title={t('objects.title', { name: label })}
          >
            <div className="object-label">
              <Icon name={b.obj.kind === 'chart' ? 'chart' : b.obj.kind === 'image' ? 'image' : 'shape'} />
              <span>
                {b.obj.title ?? (b.obj.chartType ? t('objects.typedChart', { type: t.text(b.obj.chartType) }) : label)}
                <em>{t('objects.kept')}</em>
              </span>
            </div>
          </div>
        );
      })}
    </div>
  );
}
