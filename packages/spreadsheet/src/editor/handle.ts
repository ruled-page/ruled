import { toA1 } from '@ruled/formula/refs';
import type { CsvOptions } from '../core/csv';
import { isMessage } from '../i18n/message';
import type { FileOrigin } from '../plugin/types';
import type { WorkbookSchema } from '../schema/types';
import { WorkbookStore, type SavedWorkbook, type SaveResult, type StoreOptions } from '../state/store';
import type { EditorTheme } from './theme';
import { pluginParts } from '../ui/lazyPlugins';
import type { MenuOptions } from '../ui/menuCommands';

/** The selection, as `handle.selection` gives it. */
export interface EditorSelection {
  /** The name of the sheet shown. */
  sheet: string;
  /** The active cell, such as "B4". */
  active: string;
  /** The selected ranges, such as "A1:C3" or "B4"; whole rows or columns reach the end of the sheet's used range. */
  ranges: string[];
}

/** A schema issue, as `handle.issues` lists it. */
export interface EditorIssue {
  /** The name of the sheet; undefined for an issue about the whole workbook. */
  sheet?: string;
  /** The cell the issue is on, such as "B4"; undefined for an issue about a sheet or the workbook. Issues are never on a range. */
  cell?: string;
  /**
   * What found it: the key of the editor's own check, such as "schema.check.min", or "custom" for a message
   * the schema's author wrote (a `validate`, `check`, or `onChange` hook, or a message set on a check).
   */
  rule: string;
  severity: 'error' | 'warning';
  /** In the editor's current language. */
  message: string;
}

/** Where a sheet is scrolled to, as `handle.viewport` gives it. */
export interface EditorViewport {
  /** The name of the sheet shown. */
  sheet: string;
  /** The first row shown, wholly or in part, below any frozen rows: its number, as in A1 (row 1 is 1). Hidden rows are never it. */
  row: number;
  /** The first column shown, wholly or in part, right of any frozen columns: its number (column A is 1). */
  col: number;
  /** The cell at `row` and `col`, such as "C10": what to pass to another editor's scrollTo. */
  cell: string;
}

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
  /**
   * Runs the same save as the Save button, resolving with what happened:
   * schema checks then onSave, or without onSave, writing back to the
   * plugin source the workbook came from, or a download.
   */
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
  /**
   * Selects a cell or range, such as "B4" or "Prices!A1:C3", or the range a
   * defined name refers to, and shows it. The grid takes the keyboard focus,
   * as after the name box, unless `focus` is false.
   */
  goTo(ref: string, options?: { focus?: boolean }): void;
  /**
   * Selects exactly what `selection` gives, so one editor can mirror another's: the sheet, the ranges and the active cell,
   * which is shown. The grid takes the keyboard focus unless `focus` is false. Returns false, changing nothing, when the sheet
   * or a range is unknown or badly written, or the active cell is in none of the ranges. A range that cuts through a merged
   * cell grows to take it in, as when selected with the mouse.
   */
  select(selection: EditorSelection, options?: { focus?: boolean }): boolean;
  /**
   * Scrolls a cell, such as "C10" or "Prices!C10", to the top left of the
   * part of the sheet that scrolls (below and right of any frozen panes),
   * switching sheet if needed, as far as the sheet scrolls. The selection
   * and the focus stay where they are. A hidden cell scrolls to the first
   * shown after it. False when `ref` names no cell on a sheet shown.
   */
  scrollTo(ref: string): boolean;
  /**
   * Freezes the top `rows` and left `cols` of a sheet for viewing only: the
   * active sheet, or the one named by `sheet`. It is not a change to the file:
   * no undo step, the workbook isn't marked unsaved, and saving writes the
   * file's own panes. `{ rows: 0, cols: 0 }` shows the file's own panes again.
   * It lasts until the file is opened again, or until the user freezes or
   * unfreezes panes from the menu, which changes the file and replaces it for
   * that sheet; undoing that menu step brings it back, redoing clears it
   * again. `viewport` takes it into account. False when no workbook is
   * open or `sheet` names none.
   */
  freeze(panes: { rows: number; cols: number; sheet?: string }): boolean;
  /**
   * The sheet shown, its active cell and its selected ranges; undefined until
   * a workbook is open. The same object until the selection changes, so a
   * `subscribe` listener can tell it changed by comparing it with the last.
   */
  readonly selection: EditorSelection | undefined;
  /**
   * The schema issues the open workbook has now, errors first, as the issues
   * panel lists them, whether or not the panel is shown; empty without a
   * schema. Checks still running are not in it. The same array until the
   * issues or the language change, so a `subscribe` listener can tell it
   * changed by comparing it with the last.
   */
  readonly issues: readonly EditorIssue[];
  /**
   * The first row and column the sheet shown has in view below and right
   * of any frozen panes; undefined until the grid has drawn it. The same
   * object until it changes. `subscribe` listeners are called when it does,
   * at most once a frame while scrolling.
   */
  readonly viewport: EditorViewport | undefined;
  /**
   * Deletes the draft kept in this browser for the open file, so it isn't
   * offered when the file is opened again. Does nothing with `drafts: false`
   * or for a password-protected workbook, which keep none. It changes
   * nothing else: edits made afterwards are drafted as usual. For no drafts
   * at all, use `drafts: false`; to drop the edits too, open the file again.
   * To delete the draft of a file that isn't open, use `discardDraft(bytes)`,
   * exported by the package.
   */
  discardDraft(): Promise<void>;
  /** Adds or replaces a schema and re-checks the open workbook. */
  registerSchema(schema: WorkbookSchema): void;
  removeSchema(id: string): void;
  /** Changes the theme tokens while the editor is open. */
  setTheme(theme: EditorTheme | undefined): void;
  /** Changes the language of the editor's text while it is open. */
  setLanguage(language: string): void;
  /** Changes the name recorded against each action in the change history. */
  setUser(user: string | undefined): void;
  /** Changes the lists that filling continues, as the `customLists` option gives them. */
  setCustomLists(lists: string[][] | undefined): void;
  /** Changes which menu entries and shortcuts are kept, as the `menus` option gives them. */
  setMenus(menus: MenuOptions | undefined): void;
  /** Makes the workbook read-only, or editable again, as the `readOnly` option does. */
  setReadOnly(readOnly: boolean): void;
  /** Whether the workbook is read-only. */
  readonly readOnly: boolean;
  /** Whether a workbook is open, opening or couldn't be opened. */
  readonly status: 'empty' | 'loading' | 'ready' | 'error';
  /** The open workbook's file name. */
  readonly fileName: string;
  /** Cells changed since the workbook was opened. */
  readonly editedCount: number;
  /** There are changes the host hasn't received through a save. */
  readonly unsaved: boolean;
  /** Where the open workbook came from, when a plugin's file source opened it; undefined for files from the host or this device. */
  readonly origin: FileOrigin | undefined;
  /** Calls `listener` whenever the editor's state changes, the selection and the viewport included; returns an unsubscribe function. */
  subscribe(listener: () => void): () => void;
  /** Stops the editor's worker; mountWorkbookEditor's handle unmounts the editor as well. */
  destroy(): void;
  /**
   * Resolves once every plugin's `setup` has run, including any promise it
   * returned, and the editor can draw what they added; at once for an editor
   * without plugins. It never rejects: a plugin that throws or rejects is
   * reported as a `pluginError` and the rest still count. There is no timeout:
   * a setup that never settles holds it until the editor is destroyed, which
   * resolves it, without starting the plugins if they hadn't started yet.
   */
  readonly ready: Promise<void>;
}

const NO_ISSUES: readonly EditorIssue[] = [];

const stores = new WeakMap<EditorHandle, WorkbookStore>();

/** The store behind a handle, for the editor's own components. Not part of the public API. */
export function storeOf(handle: EditorHandle): WorkbookStore {
  const store = stores.get(handle);
  if (!store) throw new Error('Not an editor made by createWorkbookEditor().');
  return store;
}

/** A handle over an existing store. */
export function handleFor(store: WorkbookStore, destroy: () => void = () => store.dispose()): EditorHandle {
  // Kept until they change, so listeners can compare.
  let selection: EditorSelection | undefined;
  let viewport: EditorViewport | undefined;
  let issues: { ctl: unknown; version: number; language: string; list: readonly EditorIssue[] } = { ctl: undefined, version: 0, language: '', list: NO_ISSUES };
  let ready: Promise<void>;
  const handle: EditorHandle = {
    async open(bytes, fileName) {
      const r = await store.open(fileName, bytes instanceof Uint8Array ? (bytes.slice().buffer as ArrayBuffer) : bytes.slice(0));
      if (r.opened) return true;
      if ('cancelled' in r) return false;
      throw new Error(store.t.text(r.error));
    },
    save: () => store.savePrimary(),
    downloadCopy: () => store.exportCopy(),
    getWorkbook: () => store.snapshot(),
    downloadCsv: (options) => store.exportCsv(options),
    getCsv: async (options) => (await store.csv(options))?.text,
    goTo: (ref, options) => void store.goTo(ref, options),
    select: (to, options) => store.select(to, options),
    scrollTo: (ref) => store.scrollTo(ref),
    freeze: ({ rows, cols, sheet }) => store.status === 'ready' && store.freezeView(rows, cols, sheet),
    discardDraft: () => store.discardDraft(),
    get selection() {
      const now = store.status === 'ready' ? store.commandSelection() : undefined;
      if (!now) return (selection = undefined);
      if (selection && selection.sheet === now.sheet && selection.active === now.active && selection.ranges.join() === now.ranges.join()) return selection;
      return (selection = now);
    },
    get issues() {
      const ctl = store.status === 'ready' ? store.schema : undefined;
      if (!ctl) return (issues = { ctl, version: 0, language: '', list: NO_ISSUES }).list;
      const language = store.t.language;
      if (issues.ctl === ctl && issues.version === store.schemaVersion && issues.language === language) return issues.list;
      const list = ctl.issues().map((i): EditorIssue => {
        return {
          ...(i.sheet >= 0 && { sheet: i.sheetName }),
          ...(i.row !== undefined && i.col !== undefined && { cell: toA1(i.row, i.col) }),
          rule: isMessage(i.message) ? i.message.key : 'custom',
          severity: i.severity,
          message: store.t.text(i.message),
        };
      });
      // Re-checks that find the same issues keep the array.
      const same = issues.list.length === list.length && list.every((n, k) => JSON.stringify(n) === JSON.stringify(issues.list[k]));
      return (issues = { ctl, version: store.schemaVersion, language, list: same ? issues.list : list.length ? list : NO_ISSUES }).list;
    },
    get viewport() {
      const at = store.status === 'ready' ? store.viewport : undefined;
      const sheet = at && at.sheet === store.active ? store.view?.info.name : undefined;
      if (!at || sheet === undefined) return (viewport = undefined);
      if (viewport && viewport.sheet === sheet && viewport.row === at.row + 1 && viewport.col === at.col + 1) return viewport;
      return (viewport = { sheet, row: at.row + 1, col: at.col + 1, cell: toA1(at.row, at.col) });
    },
    registerSchema: (schema) => store.registerSchema(schema),
    removeSchema: (id) => store.removeSchema(id),
    setTheme: (theme) => store.setTheme(theme),
    setLanguage: (language) => store.setLanguage(language),
    setUser: (user) => store.setUser(user),
    setCustomLists: (lists) => store.setCustomLists(lists),
    setMenus: (menus) => store.setMenus(menus),
    setReadOnly: (readOnly) => store.setReadOnly(readOnly),
    get readOnly() {
      return store.readOnly;
    },
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
    get origin() {
      return store.origin;
    },
    subscribe: (listener) => {
      const stop = store.subscribe(listener);
      const stopViewport = store.subscribeViewport(listener);
      return () => {
        stop();
        stopViewport();
      };
    },
    destroy,
    get ready() {
      return ready;
    },
  };
  stores.set(handle, store);
  // Plugins start once their code has loaded, whether or not the editor has rendered; their parts are drawn from a module loaded meanwhile.
  ready = store
    .startPlugins(handle)
    .then(() => Promise.all([store.plugins?.setUp, store.plugins?.active ? pluginParts.load() : undefined]))
    .then(
      () => undefined,
      () => undefined,
    );
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
