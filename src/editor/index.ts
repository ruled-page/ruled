/**
 * Public API of the workbook editor.
 *
 * React hosts render <WorkbookEditor store={new WorkbookStore(options)} />.
 * Other hosts call mountWorkbookEditor(element, options), or use the
 * <workbook-editor> custom element from the standalone build.
 */
export { WorkbookEditor, type WorkbookEditorProps } from './WorkbookEditor';
export { mountWorkbookEditor, type EditorHandle } from './mount';
export {
  SaveConflictError,
  WorkbookStore,
  type EditorFeatures,
  type HistoryEntry,
  type SavedWorkbook,
  type StoreOptions as EditorOptions,
} from '../state/store';
export type { CellChange } from '../core/types';
export { themeProperty, type EditorTheme, type ThemeToken } from './theme';
export type { IconName, IconOverrides } from '../ui/icons';
export { defineSchema } from '../schema/types';
export type * from '../schema/types';
export { schemaFromJSON, schemaToJSON } from '../schema/json';
