import { Component, type ReactNode } from 'react';
import { useT } from '../i18n/react';
import type { WorkbookStore } from '../state/store';
import { Icon } from '../ui/icons';

interface Props {
  store: WorkbookStore;
  children: ReactNode;
}

/**
 * Catches an error while the editor draws, so it says so and offers to try
 * again, rather than unmounting, which in a React host would take the
 * host's page down with it. The workbook and its edits live in the store and
 * the worker, so drawing again picks up where it was.
 */
export class EditorBoundary extends Component<Props, { failed: boolean }> {
  state = { failed: false };

  static getDerivedStateFromError(): { failed: boolean } {
    return { failed: true };
  }

  componentDidCatch(error: Error): void {
    this.props.store.reportError(error);
  }

  render(): ReactNode {
    return this.state.failed ? <DrawingFailed retry={() => this.setState({ failed: false })} /> : this.props.children;
  }
}

function DrawingFailed({ retry }: { retry: () => void }) {
  const t = useT();
  return (
    <div className="center-state error" role="alert">
      <Icon name="warn" size={28} />
      <p>{t('editor.drawingFailed')}</p>
      <button className="btn primary" onClick={retry}>
        {t('editor.tryAgain')}
      </button>
    </div>
  );
}
