import type { CsvOptions } from '../core/csv';
import type { WorkbookSchema } from '../schema/types';
import { WorkbookStore, type SavedWorkbook, type SaveResult, type StoreOptions } from '../state/store';
import type { EditorTheme } from './theme';

/**
 * What a host controls an editor through, whichever way it shows it: with
 * <WorkbookEditor editor={…} /> in React, mountWorkbookEditor() on any page,
 * or the <workbook-editor> element.
 */
export interface EditorHandle {
  /**
   * Opens a workbook from its bytes, or a CSV file (.csv, .tsv or .txt, by
   * its name) as a workbook of one sheet. Resolves true once it is open, and false
   * when the user chose to keep the workbook already open, with its unsaved
   * changes. Rejects, saying why, when the file can't be opened; the editor
   * shows the reason too. The bytes are copied, so the caller's stay usable.
   */
  open(bytes: ArrayBuffer | Uint8Array, fileName: string): Promise<boolean>;
  /** Runs the same save as the Save button (schema checks, then onSave or a download), resolving with what happened. */
  save(): Promise<SaveResult>;
  /** Downloads a copy as .xlsx without saving it to the host. */
  downloadCopy(): Promise<void>;
  /** The workbook as it is now, as onSave would receive it, without save checks or onSave. */
  getWorkbook(): Promise<SavedWorkbook | undefined>;
  /**
   * Downloads the sheet shown, and only that sheet, as a CSV file (UTF-8 with
   * a byte order mark), as the toolbar's CSV button does. Values that would run
   * as formulas in a spreadsheet app are escaped unless `escapeFormulas` is false.
   */
  downloadCsv(options?: CsvOptions): Promise<void>;
  /** The sheet shown as CSV text, as downloadCsv writes it but without the byte order mark; undefined when it isn't a worksheet. */
  getCsv(options?: CsvOptions): Promise<string | undefined>;
  /** Selects a cell or range, such as "B4" or "Prices!A1:C3", or the range a defined name refers to, and shows it. */
  goTo(ref: string): void;
  /** Adds or replaces a schema and re-checks the open workbook. */
  registerSchema(schema: WorkbookSchema): void;
  removeSchema(id: string): void;
  /** Changes the theme tokens while the editor is open. */
  setTheme(theme: EditorTheme | undefined): void;
  /** Changes the language of the editor's text while it is open. */
  setLanguage(language: string): void;
  /** Changes the name recorded against each action in the change history. */
  setUser(user: string | undefined): void;
  /** Whether a workbook is open, opening or couldn't be opened. */
  readonly status: 'empty' | 'loading' | 'ready' | 'error';
  /** The open workbook's file name. */
  readonly fileName: string;
  /** Cells changed since the workbook was opened. */
  readonly editedCount: number;
  /** There are changes the host hasn't received through a save. */
  readonly unsaved: boolean;
  /** Calls `listener` whenever the editor's state changes; returns an unsubscribe function. */
  subscribe(listener: () => void): () => void;
  /** Stops the editor's worker; mountWorkbookEditor's handle unmounts the editor as well. */
  destroy(): void;
}

const stores = new WeakMap<EditorHandle, WorkbookStore>();

/** The store behind a handle, for the editor's own components. Not part of the public API. */
export function storeOf(handle: EditorHandle): WorkbookStore {
  const store = stores.get(handle);
  if (!store) throw new Error('Not an editor made by createWorkbookEditor().');
  return store;
}

/** A handle over an existing store. */
export function handleFor(store: WorkbookStore, destroy: () => void = () => store.dispose()): EditorHandle {
  const handle: EditorHandle = {
    async open(bytes, fileName) {
      const r = await store.open(fileName, bytes instanceof Uint8Array ? (bytes.slice().buffer as ArrayBuffer) : bytes.slice(0));
      if (r.opened) return true;
      if ('cancelled' in r) return false;
      throw new Error(store.t.text(r.error));
    },
    save: () => store.save(),
    downloadCopy: () => store.exportCopy(),
    getWorkbook: () => store.snapshot(),
    downloadCsv: (options) => store.exportCsv(options),
    getCsv: async (options) => (await store.csv(options))?.text,
    goTo: (ref) => void store.goTo(ref),
    registerSchema: (schema) => store.registerSchema(schema),
    removeSchema: (id) => store.removeSchema(id),
    setTheme: (theme) => store.setTheme(theme),
    setLanguage: (language) => store.setLanguage(language),
    setUser: (user) => store.setUser(user),
    get status() {
      return store.status;
    },
    get fileName() {
      return store.fileName;
    },
    get editedCount() {
      return store.editedCount;
    },
    get unsaved() {
      return store.unsaved;
    },
    subscribe: (listener) => store.subscribe(listener),
    destroy,
  };
  stores.set(handle, store);
  return handle;
}

/**
 * Creates an editor for a React host to render with
 * <WorkbookEditor editor={editor} />. Call destroy() when it is no longer
 * needed, which stops its worker.
 */
export function createWorkbookEditor(options: StoreOptions = {}): EditorHandle {
  return handleFor(new WorkbookStore(options));
}
