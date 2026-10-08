import { createRoot } from 'react-dom/client';
import { WorkbookStore, type StoreOptions } from '../state/store';
import { handleFor, type EditorHandle } from './handle';
import { WorkbookEditor } from './WorkbookEditor';

export type { EditorHandle } from './handle';

/**
 * Renders the editor into `element` without needing React in the host page.
 * The element should have a height; the editor fills it.
 */
export function mountWorkbookEditor(element: HTMLElement, options: StoreOptions = {}): EditorHandle {
  const store = new WorkbookStore(options);
  const root = createRoot(element);
  const handle = handleFor(store, () => {
    root.unmount();
    store.dispose();
  });
  root.render(<WorkbookEditor editor={handle} />);
  return handle;
}
