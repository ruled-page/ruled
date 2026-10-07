import { createRoot } from 'react-dom/client';
import { WorkbookStore, type SavedWorkbook, type StoreOptions } from '../state/store';
import type { WorkbookSchema } from '../schema/types';
import type { EditorTheme } from './theme';
import { WorkbookEditor } from './WorkbookEditor';

/** Controls an editor mounted with mountWorkbookEditor. */
export interface EditorHandle {
  /** Opens a workbook from its bytes. */
  open(bytes: ArrayBuffer | Uint8Array, fileName: string): Promise<void>;
  /** The edited workbook as .xlsx bytes, without running save checks or onSave. */
  getWorkbook(): Promise<SavedWorkbook | undefined>;
  /** Runs the same save flow as the Save button (schema checks, then onSave or a download). */
  save(): Promise<void>;
  /** Downloads a copy as .xlsx without saving it to the host. */
  download(): Promise<void>;
  /** There are changes the host hasn't received through a save yet. */
  readonly unsaved: boolean;
  /** Adds or replaces a schema and re-checks the open workbook. */
  registerSchema(schema: WorkbookSchema): void;
  /** Changes the theme tokens while the editor is open. */
  setTheme(theme: EditorTheme | undefined): void;
  /** Cells changed since the workbook was opened. */
  readonly editedCount: number;
  /** Calls `listener` whenever the editor's state changes; returns an unsubscribe function. */
  subscribe(listener: () => void): () => void;
  /** The underlying store, for anything the handle doesn't cover. */
  readonly store: WorkbookStore;
  /** Unmounts the editor and stops its worker. */
  destroy(): void;
}

/**
 * Renders the editor into `element` without needing React in the host page.
 * The element should have a height; the editor fills it.
 */
export function mountWorkbookEditor(element: HTMLElement, options: StoreOptions = {}): EditorHandle {
  const store = new WorkbookStore(options);
  const root = createRoot(element);
  root.render(<WorkbookEditor store={store} />);
  return {
    store,
    open: (bytes, fileName) => store.open(fileName, bytes instanceof Uint8Array ? (bytes.slice().buffer as ArrayBuffer) : bytes),
    getWorkbook: () => store.save(),
    save: () => store.download(),
    download: () => store.exportCopy(),
    get unsaved() {
      return store.unsaved;
    },
    registerSchema: (schema) => store.registerSchema(schema),
    setTheme: (theme) => store.setTheme(theme),
    get editedCount() {
      return store.editedCount;
    },
    subscribe: (listener) => store.subscribe(listener),
    destroy: () => {
      root.unmount();
      store.dispose();
    },
  };
}
