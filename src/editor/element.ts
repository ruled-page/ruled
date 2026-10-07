import type { SavedWorkbook, StoreOptions } from '../state/store';
import type { WorkbookSchema } from '../schema/types';
import type { EditorHandle } from './mount';
import type { EditorTheme } from './theme';

type Mount = (element: HTMLElement, options?: StoreOptions) => EditorHandle;

/** Detail of the `save` event when the element handles saving through events. */
export interface SaveEventDetail extends SavedWorkbook {
  /** Keep the save pending until `promise` settles; a rejection is shown as the error. */
  waitUntil(promise: Promise<unknown>): void;
}

/**
 * Defines <workbook-editor>, usable from any framework or plain HTML.
 *
 * Attributes: `src` (URL of a workbook to open), `file-name`, and
 * `save-handler="event"` to receive saves as `save` events instead of a
 * download. Properties: `options` (editor options, set before connecting).
 * Methods: open(), save(), download(), getWorkbook(), registerSchema(), setTheme(), setLanguage().
 * Events: `ready`, `change` (detail: { editedCount, unsaved }), `save` (detail:
 * the saved workbook with its bytes, changes, values and readers; call
 * `detail.waitUntil(promise)` to keep the save pending while you store it).
 */
export function defineWorkbookEditorElement(mount: Mount, tagName = 'workbook-editor'): void {
  if (customElements.get(tagName)) return;

  class WorkbookEditorElement extends HTMLElement {
    static observedAttributes = ['src'];
    private opts: StoreOptions = {};
    private handle?: EditorHandle;
    private unsubscribe?: () => void;
    private lastEdited = 0;
    private lastUnsaved = false;
    /** The editor renders in here, so its own DOM events reach the element after React has seen them. */
    private readonly container = document.createElement('div');

    constructor() {
      super();
      this.container.style.height = '100%';
      // Inputs inside the editor fire the browser's change events, which
      // bubble. Stop them here, so `change` listeners hear only the editor's.
      this.addEventListener('change', (e) => {
        if (e.target !== this) e.stopImmediatePropagation();
      });
    }

    /** Editor options. Setting them on a connected element re-creates the editor. */
    get options(): StoreOptions {
      return this.opts;
    }

    set options(value: StoreOptions) {
      this.opts = value ?? {};
      if (this.handle) {
        this.disconnectedCallback();
        this.connectedCallback();
      }
    }

    connectedCallback(): void {
      if (!this.style.display) this.style.display = 'block';
      if (!this.container.isConnected) this.append(this.container);
      const viaEvent = this.getAttribute('save-handler') === 'event';
      this.handle = mount(this.container, {
        ...this.opts,
        onSave: viaEvent ? (file) => this.emitSave(file) : this.opts.onSave,
      });
      this.unsubscribe = this.handle.subscribe(() => {
        const edited = this.handle?.editedCount ?? 0;
        const unsaved = this.handle?.unsaved ?? false;
        if (edited === this.lastEdited && unsaved === this.lastUnsaved) return;
        this.lastEdited = edited;
        this.lastUnsaved = unsaved;
        this.dispatchEvent(new CustomEvent('change', { detail: { editedCount: edited, unsaved } }));
      });
      this.dispatchEvent(new CustomEvent('ready'));
      const src = this.getAttribute('src');
      if (src) void this.load(src);
    }

    disconnectedCallback(): void {
      this.unsubscribe?.();
      this.handle?.destroy();
      this.handle = undefined;
    }

    attributeChangedCallback(name: string, oldValue: string | null, value: string | null): void {
      if (name === 'src' && value && value !== oldValue && this.handle) void this.load(value);
    }

    private async load(src: string): Promise<void> {
      const res = await fetch(src);
      if (!res.ok) throw new Error(`Could not load ${src}: ${res.status}`);
      const name = this.getAttribute('file-name') ?? decodeURIComponent(src.split('/').pop()?.split('?')[0] || 'workbook.xlsx');
      await this.open(await res.arrayBuffer(), name);
    }

    private async emitSave(file: SavedWorkbook): Promise<void> {
      const waits: Promise<unknown>[] = [];
      const detail: SaveEventDetail = { ...file, waitUntil: (p) => waits.push(p) };
      this.dispatchEvent(new CustomEvent<SaveEventDetail>('save', { detail }));
      await Promise.all(waits);
    }

    open(bytes: ArrayBuffer | Uint8Array, fileName: string): Promise<void> {
      if (!this.handle) throw new Error('<workbook-editor> is not connected to the document.');
      return this.handle.open(bytes, fileName);
    }

    save(): Promise<void> {
      return this.handle?.save() ?? Promise.resolve();
    }

    /** Downloads a copy as .xlsx without saving it to the host. */
    download(): Promise<void> {
      return this.handle?.download() ?? Promise.resolve();
    }

    get unsaved(): boolean {
      return this.handle?.unsaved ?? false;
    }

    getWorkbook() {
      return this.handle?.getWorkbook() ?? Promise.resolve(undefined);
    }

    registerSchema(schema: WorkbookSchema): void {
      this.handle?.registerSchema(schema);
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

    get editedCount(): number {
      return this.handle?.editedCount ?? 0;
    }
  }

  customElements.define(tagName, WorkbookEditorElement);
}
