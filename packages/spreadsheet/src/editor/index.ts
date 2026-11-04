/**
 * Public API of the workbook editor.
 *
 * React hosts render <WorkbookEditor editor={useWorkbookEditor(options)} />, or
 * manage one from createWorkbookEditor(options) themselves. <Workbook>, the
 * two in one, takes props a React Server Component can give it.
 * Other hosts call mountWorkbookEditor(element, options), or use the
 * <workbook-editor> custom element from the standalone build. Each gives an
 * EditorHandle to control the editor with.
 */
export { WorkbookEditor, type WorkbookEditorProps } from './WorkbookEditor';
export { Workbook, type WorkbookProps } from './Workbook';
export type { SavedWorkbookData } from './plainData';
export { createWorkbookEditor, type EditorHandle, type EditorSelection, type EditorViewport, type EditorIssue } from './handle';
export { useEditorState, useWorkbookEditor, type EditorState, type WorkbookEditorHookOptions, type WorkbookFile } from './hooks';
export { mountWorkbookEditor } from './mount';
export { discardDraft } from './drafts';
export {
  SaveConflictError,
  type EditorFeatures,
  type SaveAnswer,
  type SaveResult,
  type HistoryEntry,
  type SavedWorkbook,
  type StoreOptions as EditorOptions,
} from '../state/store';
export type { CellChange } from '../core/types';
export type { CsvOptions } from '../core/csv';
export type { LinkClick } from '../state/links';
export { themeProperty, type EditorTheme, type ThemeToken } from './theme';
export type { IconName, IconOverrides } from '../ui/icons';
/** The ids of the cell and heading menus' entries, for the `menus` option, and which of them edit the workbook. */
export { MENU_COMMANDS, type MenuCommandId, type MenuCommandInfo, type MenuEntryId, type MenuOptions } from '../ui/menuCommands';
/**
 * Languages: `translations` in the options adds languages or rewords
 * messages, keyed like the built-in catalogues. Start a new language from
 * `en`, which has every message; `de` is the German one.
 */
export { de, en, type Translations } from '../i18n';
/**
 * Text the editor shows: written text, a message to translate (`msg(key,
 * params)`), or host-written text in several languages (`Localized`, such as
 * { en: 'Price', de: 'Preis' }).
 */
export { msg, type Localized, type Message, type Text } from '../i18n/message';
export { defineSchema } from '../schema/types';
export type * from '../schema/types';
export { schemaFromJSON, schemaToJSON } from '../schema/json';
