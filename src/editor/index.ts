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
