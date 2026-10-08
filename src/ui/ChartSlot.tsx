import { Component, lazy, Suspense, type CSSProperties, type ReactNode } from 'react';
import type { SheetObject } from '../core/types';
import type { WorkbookStore } from '../state/store';

// The chart library loads only when a workbook has a chart to draw.
const ChartView = lazy(() => import('./ChartView'));

/** Shows the placeholder instead of a chart that fails to load or draw. */
class Fallback extends Component<{ fallback: ReactNode; children: ReactNode }, { failed: boolean }> {
  state = { failed: false };

  static getDerivedStateFromError(): { failed: boolean } {
    return { failed: true };
  }

  render(): ReactNode {
    return this.state.failed ? this.props.fallback : this.props.children;
  }
}

/** A chart, drawn when the editor draws it; otherwise, and while the drawing loads, its placeholder. */
export function ChartSlot({
  store,
  obj,
  placeholder,
  className,
  style,
  copy,
}: {
  store: WorkbookStore;
  obj: SheetObject;
  placeholder: ReactNode;
  className?: string;
  style?: CSSProperties;
  /** A second drawing of the same chart, which screen readers skip. */
  copy?: boolean;
}) {
  if (!obj.chart || obj.chart.unsupported) return placeholder;
  return (
    <Fallback fallback={placeholder}>
      <Suspense fallback={placeholder}>
        <ChartView store={store} obj={obj} className={className} style={style} copy={copy} />
      </Suspense>
    </Fallback>
  );
}
