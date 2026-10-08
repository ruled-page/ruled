import type { CsvOptions } from '../core/csv';
import type { SavedWorkbook, SaveResult, StoreOptions } from '../state/store';
import type { WorkbookSchema } from '../schema/types';
import type { EditorHandle } from './mount';
import type { EditorTheme } from './theme';

type Mount = (element: HTMLElement, options?: StoreOptions) => EditorHandle;

/** Detail of the `save` event when the element handles saving through events. */
export interface SaveEventDetail extends SavedWorkbook {
  /** Keep the save pending until `promise` settles; a rejection is shown as the error. */
  waitUntil(promise: Promise<unknown>): void;
}

/** Detail of the `change` event. */
export interface ChangeEventDetail {
  /** Cells that differ from the file as it was opened. */
  editedCount: number;
  /** Whether there are edits the host hasn't received through a save. */
  unsaved: boolean;
}

/** Detail of the `opened` event. */
export interface OpenedEventDetail {
  fileName: string;
}

/** Detail of the `editorerror` event. */
export interface EditorErrorEventDetail {
  /** What went wrong, in the editor's language. */
  message: string;
}

/** The events <workbook-editor> fires, for typed listeners. */
export interface WorkbookEditorEventMap extends HTMLElementEventMap {
  /** The editor is ready for a workbook. */
  ready: CustomEvent<void>;
  /** A workbook has opened. */
  opened: CustomEvent<OpenedEventDetail>;
  change: CustomEvent<ChangeEventDetail>;
  save: CustomEvent<SaveEventDetail>;
  /** The workbook from `src` couldn't be loaded or opened, or the editor stopped working. */
  editorerror: CustomEvent<EditorErrorEventDetail>;
}

/** The <workbook-editor> element: what hosts can set, call and listen to. */
export interface WorkbookEditorElement extends HTMLElement {
  /**
   * Editor options. Changing the theme, language, user or schemas of a
   * connected element applies them as it is; changing anything else
   * re-creates the editor, closing the workbook.
   */
  options: StoreOptions;
  /** Opens a workbook or CSV file: true once open, false when the user kept the one already open; rejects saying why it can't. */
  open(bytes: ArrayBuffer | Uint8Array, fileName: string): Promise<boolean>;
  /** Runs the save flow (schema checks, then the `save` event or a download), resolving with what happened. */
  save(): Promise<SaveResult>;
  /** Downloads a copy as .xlsx without saving it to the host. */
  downloadCopy(): Promise<void>;
  /** The workbook as it is now, as a save would hand it over, without the save checks. */
  getWorkbook(): Promise<SavedWorkbook | undefined>;
  /** Downloads the sheet shown as a CSV file, with values that would run as formulas escaped unless `escapeFormulas` is false. */
  downloadCsv(options?: CsvOptions): Promise<void>;
  /** The sheet shown as CSV text; undefined when it isn't a worksheet. */
  getCsv(options?: CsvOptions): Promise<string | undefined>;
  /** Selects a cell or range, such as "B4" or "Prices!A1:C3", and shows it. */
  goTo(ref: string): void;
  registerSchema(schema: WorkbookSchema): void;
  removeSchema(id: string): void;
  setTheme(theme: EditorTheme | undefined): void;
  setLanguage(language: string): void;
  setUser(user: string | undefined): void;
  setCustomLists(lists: string[][] | undefined): void;
  readonly status: EditorHandle['status'];
  readonly fileName: string;
  readonly unsaved: boolean;
  readonly editedCount: number;
  addEventListener<K extends keyof WorkbookEditorEventMap>(type: K, listener: (this: WorkbookEditorElement, ev: WorkbookEditorEventMap[K]) => unknown, options?: boolean | AddEventListenerOptions): void;
  addEventListener(type: string, listener: EventListenerOrEventListenerObject, options?: boolean | AddEventListenerOptions): void;
  removeEventListener<K extends keyof WorkbookEditorEventMap>(type: K, listener: (this: WorkbookEditorElement, ev: WorkbookEditorEventMap[K]) => unknown, options?: boolean | EventListenerOptions): void;
  removeEventListener(type: string, listener: EventListenerOrEventListenerObject, options?: boolean | EventListenerOptions): void;
}

declare global {
  interface HTMLElementTagNameMap {
    'workbook-editor': WorkbookEditorElement;
  }
}

/** Options a connected element applies without re-creating the editor. */
const LIVE: (keyof StoreOptions)[] = ['theme', 'language', 'user', 'schemas', 'customLists'];

/**
 * Defines <workbook-editor>, usable from any framework or plain HTML.
 *
 * Attributes: `src` (URL of a workbook or CSV file to open), `file-name`, and
 * `save-handler="event"` to receive saves as `save` events instead of a
 * download. Properties: `options` (editor options).
 * Methods: open(), save(), downloadCopy(), getWorkbook(), downloadCsv(), getCsv(), goTo(), registerSchema(), removeSchema(), setTheme(), setLanguage(), setUser(), setCustomLists().
 * Events: `ready`, `opened` (detail: { fileName }), `change` (detail: { editedCount, unsaved }),
 * `save` (detail: the saved workbook with its bytes, changes, values and
 * readers; call `detail.waitUntil(promise)` to keep the save pending while
 * you store it), and `editorerror` (detail: { message }) when a workbook from
 * `src` can't be loaded or opened, or the editor stops working.
 */
export function defineWorkbookEditorElement(mount: Mount, tagName = 'workbook-editor'): void {
  if (customElements.get(tagName)) return;

  class Editor extends HTMLElement implements WorkbookEditorElement {
    static observedAttributes = ['src'];
    private opts: StoreOptions = {};
    private handle?: EditorHandle;
    private unsubscribe?: () => void;
    private lastEdited = 0;
    private lastUnsaved = false;
    private lastOpened = '';
    /** A removal that ends the editor unless the element is put back first (moved elsewhere in the page). */
    private teardown?: ReturnType<typeof setTimeout>;
    /** The `src` being fetched, cancelled when another replaces it. */
    private loading?: AbortController;
    /** The editor renders in here, so its own DOM events reach the element after React has seen them. */
    private readonly container = document.createElement('div');

    constructor() {
      super();
      // A framework may set `options` before this element was defined (the editor
      // loaded later); that value sits on the instance and hides the setter.
      if (Object.prototype.hasOwnProperty.call(this, 'options')) {
        const options = (this as { options?: StoreOptions }).options;
        delete (this as { options?: StoreOptions }).options;
        this.opts = options ?? {};
      }
      this.container.style.height = '100%';
      // Inputs inside the editor fire the browser's change events, which
      // bubble. Stop them here, so `change` listeners hear only the editor's.
      this.addEventListener('change', (e) => {
        if (e.target !== this) e.stopImmediatePropagation();
      });
    }

    get options(): StoreOptions {
      return this.opts;
    }

    set options(value: StoreOptions) {
      const before = this.opts;
      const next = value ?? {};
      this.opts = next;
      if (!this.handle) return;
      const changed = (Object.keys({ ...before, ...next }) as (keyof StoreOptions)[]).filter((k) => before[k] !== next[k]);
      if (changed.every((k) => LIVE.includes(k))) {
        if (changed.includes('theme')) this.handle.setTheme(next.theme);
        if (changed.includes('language') && next.language) this.handle.setLanguage(next.language);
        if (changed.includes('user')) this.handle.setUser(next.user);
        if (changed.includes('customLists')) this.handle.setCustomLists(next.customLists);
        if (changed.includes('schemas')) {
          const ids = new Set((next.schemas ?? []).map((s) => s.id));
          for (const s of before.schemas ?? []) if (!ids.has(s.id)) this.handle.removeSchema(s.id);
          for (const s of [...(next.schemas ?? [])].reverse()) this.handle.registerSchema(s);
        }
        return;
      }
      this.end();
      this.connectedCallback();
    }

    connectedCallback(): void {
      if (this.teardown !== undefined) {
        // Moved within the page: the editor carries on as it was.
        clearTimeout(this.teardown);
        this.teardown = undefined;
        if (this.handle) return;
      }
      if (!this.style.display) this.style.display = 'block';
      if (!this.container.isConnected) this.append(this.container);
      const viaEvent = this.getAttribute('save-handler') === 'event';
      this.handle = mount(this.container, {
        ...this.opts,
        onSave: viaEvent ? (file) => this.emitSave(file) : this.opts.onSave,
        onError: (error) => {
          this.fail(error.message);
          this.opts.onError?.(error);
        },
      });
      this.lastOpened = '';
      this.unsubscribe = this.handle.subscribe(() => {
        const handle = this.handle;
        if (!handle) return;
        if (handle.status === 'ready' && handle.fileName !== this.lastOpened) {
          this.lastOpened = handle.fileName;
          this.dispatchEvent(new CustomEvent<OpenedEventDetail>('opened', { detail: { fileName: handle.fileName } }));
        } else if (handle.status !== 'ready') this.lastOpened = '';
        const edited = handle.editedCount;
        const unsaved = handle.unsaved;
        if (edited === this.lastEdited && unsaved === this.lastUnsaved) return;
        this.lastEdited = edited;
        this.lastUnsaved = unsaved;
        this.dispatchEvent(new CustomEvent<ChangeEventDetail>('change', { detail: { editedCount: edited, unsaved } }));
      });
      this.dispatchEvent(new CustomEvent('ready'));
      const src = this.getAttribute('src');
      if (src) void this.load(src);
    }

    disconnectedCallback(): void {
      // Wait a moment: an element moved elsewhere in the page is disconnected and connected again.
      this.teardown = setTimeout(() => {
        this.teardown = undefined;
        this.end();
      }, 0);
    }

    /** Ends the editor and its worker. */
    private end(): void {
      this.loading?.abort();
      this.unsubscribe?.();
      this.handle?.destroy();
      this.handle = undefined;
    }

    attributeChangedCallback(name: string, oldValue: string | null, value: string | null): void {
      if (name === 'src' && value && value !== oldValue && this.handle) void this.load(value);
    }

    private fail(message: string): void {
      this.dispatchEvent(new CustomEvent<EditorErrorEventDetail>('editorerror', { detail: { message } }));
    }

    private async load(src: string): Promise<void> {
      this.loading?.abort();
      const loading = (this.loading = new AbortController());
      try {
        const res = await fetch(src, { signal: loading.signal });
        if (!res.ok) throw new Error(`Could not load ${src}: ${res.status} ${res.statusText}`.trim());
        const bytes = await res.arrayBuffer();
        if (loading.signal.aborted) return;
        const name = this.getAttribute('file-name') ?? decodeURIComponent(src.split('/').pop()?.split('?')[0] || 'workbook.xlsx');
        await this.open(bytes, name);
      } catch (e) {
        if (!loading.signal.aborted) this.fail((e as Error).message);
      }
    }

    private async emitSave(file: SavedWorkbook): Promise<void> {
      const waits: Promise<unknown>[] = [];
      const detail: SaveEventDetail = { ...file, waitUntil: (p) => waits.push(p) };
      this.dispatchEvent(new CustomEvent<SaveEventDetail>('save', { detail }));
      await Promise.all(waits);
    }

    private connected(): EditorHandle {
      if (!this.handle) throw new Error('<workbook-editor> is not connected to the document.');
      return this.handle;
    }

    open(bytes: ArrayBuffer | Uint8Array, fileName: string): Promise<boolean> {
      return this.connected().open(bytes, fileName);
    }

    save(): Promise<SaveResult> {
      return this.connected().save();
    }

    downloadCopy(): Promise<void> {
      return this.connected().downloadCopy();
    }

    getWorkbook(): Promise<SavedWorkbook | undefined> {
      return this.handle?.getWorkbook() ?? Promise.resolve(undefined);
    }

    downloadCsv(options?: CsvOptions): Promise<void> {
      return this.connected().downloadCsv(options);
    }

    getCsv(options?: CsvOptions): Promise<string | undefined> {
      return this.handle?.getCsv(options) ?? Promise.resolve(undefined);
    }

    goTo(ref: string): void {
      this.handle?.goTo(ref);
    }

    registerSchema(schema: WorkbookSchema): void {
      this.opts = { ...this.opts, schemas: [schema, ...(this.opts.schemas ?? []).filter((s) => s.id !== schema.id)] };
      this.handle?.registerSchema(schema);
    }

    removeSchema(id: string): void {
      this.opts = { ...this.opts, schemas: (this.opts.schemas ?? []).filter((s) => s.id !== id) };
      this.handle?.removeSchema(id);
    }

    /** Changes the theme tokens without re-creating the editor. */
    setTheme(theme: EditorTheme | undefined): void {
      this.opts = { ...this.opts, theme };
      this.handle?.setTheme(theme);
    }

    /** Changes the language of the editor's text, such as "de", without re-creating the editor. */
    setLanguage(language: string): void {
      this.opts = { ...this.opts, language };
      this.handle?.setLanguage(language);
    }

    setUser(user: string | undefined): void {
      this.opts = { ...this.opts, user };
      this.handle?.setUser(user);
    }

    /** Changes the lists that filling continues without re-creating the editor. */
    setCustomLists(lists: string[][] | undefined): void {
      this.opts = { ...this.opts, customLists: lists };
      this.handle?.setCustomLists(lists);
    }

    get status(): EditorHandle['status'] {
      return this.handle?.status ?? 'empty';
    }

    get fileName(): string {
      return this.handle?.fileName ?? '';
    }

    get unsaved(): boolean {
      return this.handle?.unsaved ?? false;
    }

    get editedCount(): number {
      return this.handle?.editedCount ?? 0;
    }
  }

  customElements.define(tagName, Editor);
}
