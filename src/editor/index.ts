/**
 * Public API of the workbook editor.
 *
 * React hosts render <WorkbookEditor store={new WorkbookStore(options)} />.
 * Other hosts call mountWorkbookEditor(element, options), or use the
 * <workbook-editor> custom element from the standalone build.
 */
export { WorkbookEditor, type WorkbookEditorProps } from './WorkbookEditor';
export { mountWorkbookEditor, type EditorHandle } from './mount';
export { WorkbookStore, type EditorFeatures, type SavedWorkbook, type StoreOptions as EditorOptions } from '../state/store';
export { defineSchema } from '../schema/types';
export type * from '../schema/types';
export { schemaFromJSON, schemaToJSON } from '../schema/json';
