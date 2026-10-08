import {
  ArcElement,
  BarController,
  BarElement,
  CategoryScale,
  Chart,
  DoughnutController,
  Filler,
  Legend,
  LinearScale,
  LineController,
  LineElement,
  LogarithmicScale,
  PieController,
  PointElement,
  ScatterController,
  Title,
} from 'chart.js';
import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import type { SheetObject } from '../core/types';
import { useT } from '../i18n/react';
import { chartSheets, describeChart, plotChart } from '../state/charts';
import type { WorkbookStore } from '../state/store';
import { chartConfig, parsePalette, type ChartLook } from './chartConfig';
import { readGridPalette } from './gridPalette';

Chart.register(
  BarController,
  BarElement,
  LineController,
  LineElement,
  PointElement,
  PieController,
  DoughnutController,
  ArcElement,
  ScatterController,
  CategoryScale,
  LinearScale,
  LogarithmicScale,
  Filler,
  Legend,
  Title,
);

/** The look charts take from the editor's theme tokens. */
function readLook(el: HTMLElement | null): ChartLook {
  const grid = readGridPalette(el);
  const colors = el && typeof getComputedStyle !== 'undefined' ? getComputedStyle(el).getPropertyValue('--wbe-chart-colors') : '';
  return { palette: parsePalette(colors), background: grid.bg, text: grid.text, muted: grid.textMuted, gridline: grid.line, axis: grid.headerLine, font: grid.font };
}

/**
 * A chart drawn with Chart.js from the cells as they are now, redrawn when
 * they change. Screen readers get its name and, in words, its series and
 * values.
 */
export default function ChartView({
  store,
  obj,
  className,
  style,
  copy,
}: {
  store: WorkbookStore;
  obj: SheetObject;
  className?: string;
  style?: CSSProperties;
  copy?: boolean;
}) {
  const t = useT();
  const box = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const chart = useRef<Chart | null>(null);
  const descriptionId = useId();
  const spec = obj.chart!;
  const { sheets, styles, date1904 } = store;

  // Views change in place as cells change; their versions say when to read them again.
  const versions = chartSheets(spec)
    .map((i) => sheets[i]?.version)
    .join();
  const plot = useMemo(
    () => plotChart(spec, { sheets, styles, date1904 }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [spec, sheets, styles, date1904, versions],
  );

  const [look, setLook] = useState<ChartLook | null>(null);
  const themeVersion = store.themeVersion;
  useLayoutEffect(() => {
    const next = readLook(box.current);
    setLook((prev) => (JSON.stringify(prev) === JSON.stringify(next) ? prev : next));
  }, [themeVersion]);

  useEffect(() => {
    if (!look || !canvas.current) return;
    const config = chartConfig(plot, look, t.text, date1904);
    const current = chart.current;
    if (current && (current.config as { type?: string }).type === config.type) {
      current.data = config.data;
      current.options = config.options ?? {};
      current.update();
    } else {
      current?.destroy();
      chart.current = new Chart(canvas.current, config);
    }
  }, [plot, look, t, date1904]);

  useEffect(
    () => () => {
      chart.current?.destroy();
      chart.current = null;
    },
    [],
  );

  const type = obj.chartType ? t.text(obj.chartType) : undefined;
  const name = type
    ? plot.title
      ? t('editor.typedChartTitled', { type, title: plot.title })
      : t('editor.typedChart', { type })
    : (plot.title ?? t('objects.chart'));
  const description = useMemo(() => (copy ? '' : describeChart(plot, t, date1904)), [plot, t, date1904, copy]);

  return (
    <>
      <div
        ref={box}
        className={className}
        style={style}
        role={copy ? undefined : 'img'}
        aria-label={copy ? undefined : name}
        aria-describedby={copy ? undefined : descriptionId}
        aria-hidden={copy || undefined}
      >
        <canvas ref={canvas} />
      </div>
      {!copy && (
        <p id={descriptionId} className="wbe-sr-only">
          {description}
        </p>
      )}
    </>
  );
}
