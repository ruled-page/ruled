/**
 * Self-contained build for hosts without React (Angular, plain HTML): React,
 * the worker and the styles are all bundled, and importing this module
 * defines the <workbook-editor> element.
 */
import type { StoreOptions } from '../state/store';
import InlineWorker from '../worker/workbook.worker.ts?worker&inline';
import editorCss from './editor.css?inline';
import { defineWorkbookEditorElement } from './element';
import { mountWorkbookEditor as mount, type EditorHandle } from './mount';

function injectStyles(): void {
  if (document.getElementById('workbook-editor-styles')) return;
  const style = document.createElement('style');
  style.id = 'workbook-editor-styles';
  style.textContent = editorCss;
  document.head.appendChild(style);
}

export function mountWorkbookEditor(element: HTMLElement, options: StoreOptions = {}): EditorHandle {
  injectStyles();
  return mount(element, { createWorker: () => new InlineWorker(), ...options });
}

defineWorkbookEditorElement(mountWorkbookEditor);

export {
  defineWorkbookEditorElement,
  type ChangeEventDetail,
  type EditorErrorEventDetail,
  type OpenedEventDetail,
  type SaveEventDetail,
  type WorkbookEditorElement,
  type WorkbookEditorEventMap,
} from './element';
export type { EditorHandle } from './handle';
export { SaveConflictError } from '../state/store';
export { de, en, msg, type Localized, type Message, type Text, type Translations } from '../i18n';
export type { EditorFeatures, HistoryEntry, SavedWorkbook, SaveResult, StoreOptions as EditorOptions } from '../state/store';
export type { CellChange } from '../core/types';
export type { CsvOptions } from '../core/csv';
export type { LinkClick } from '../state/links';
export { defineSchema } from '../schema/types';
export type * from '../schema/types';
export { schemaFromJSON, schemaToJSON } from '../schema/json';
