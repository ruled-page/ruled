/**
 * The self-contained build, React, the worker and the styles included, which
 * defines the `<workbook-editor>` element when imported:
 * `import '@sitepen/spreadsheet/standalone'`.
 *
 * @module @sitepen/spreadsheet/standalone
 */
import { de as deMessages, en as enMessages } from '../../packages/spreadsheet/src/editor/standalone';

export * from '../../packages/spreadsheet/src/editor/standalone';
export type { MessageKey } from '../../packages/spreadsheet/src/i18n/message';

/** The English messages, every one, keyed as `translations` takes them: start a new language from a copy of it. */
export const en: Record<string, string> = enMessages;
/** The German messages. */
export const de: Record<string, string> = deMessages;

/**
 * A message for the editor to translate when it shows it: a key of the
 * catalogues, such as `'toolbar.save'` (typed as `MessageKey` in the
 * package), and the values for its placeholders.
 */
export interface Message {
  key: string;
  params?: Record<string, unknown>;
}

/** Makes a message for the editor to translate when it shows it: `msg('schema.check.required')`. */
export declare function msg(key: string, params?: Record<string, unknown>): Message;
