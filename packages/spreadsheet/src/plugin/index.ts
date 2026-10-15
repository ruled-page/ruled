/**
 * What plugins are written against: `@ruled/sheet/plugin`. A plugin imports
 * only these types, so it never brings a second copy of the editor with it;
 * `definePlugin` is optional, for type inference.
 */
import type { RuledPlugin } from './types';

export type * from './types';
export type { EditorHandle } from '../editor/handle';
export type { EditorEvent, SavedWorkbook, SaveConflictError } from '../state/store';
export type { Localized, Text } from '../i18n/message';
export type { Translations } from '../i18n';

/** Returns the plugin as given; it only helps TypeScript infer its type. */
export function definePlugin<T extends RuledPlugin>(plugin: T): T {
  return plugin;
}
