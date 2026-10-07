import type { SavedWorkbook, StoreOptions } from '../state/store';
import type { WorkbookSchema } from '../schema/types';
import type { EditorHandle } from './mount';

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
 * Methods: open(), save(), getWorkbook(), registerSchema().
 * Events: `ready`, `change` (detail: { editedCount }), `save`.
 */
export function defineWorkbookEditorElement(mount: Mount, tagName = 'workbook-editor'): void {
  if (customElements.get(tagName)) return;

  class WorkbookEditorElement extends HTMLElement {
    static observedAttributes = ['src'];
    private opts: StoreOptions = {};
    private handle?: EditorHandle;
    private unsubscribe?: () => void;
    private lastEdited = 0;

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
      const viaEvent = this.getAttribute('save-handler') === 'event';
      this.handle = mount(this, {
        ...this.opts,
        onSave: viaEvent ? (file) => this.emitSave(file) : this.opts.onSave,
      });
      this.unsubscribe = this.handle.subscribe(() => {
        const edited = this.handle?.editedCount ?? 0;
        if (edited === this.lastEdited) return;
        this.lastEdited = edited;
        this.dispatchEvent(new CustomEvent('change', { detail: { editedCount: edited } }));
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

    getWorkbook() {
      return this.handle?.getWorkbook() ?? Promise.resolve(undefined);
    }

    registerSchema(schema: WorkbookSchema): void {
      this.handle?.registerSchema(schema);
    }

    get editedCount(): number {
      return this.handle?.editedCount ?? 0;
    }
  }

  customElements.define(tagName, WorkbookEditorElement);
}
