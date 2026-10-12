/**
 * Messages as data: what code that runs in the worker (or without an editor)
 * hands over for the editor to translate when it shows it. Nothing here
 * loads i18next or the catalogues, so the worker stays small.
 */
import type { en } from './messages/en';

type PluralSuffix = 'zero' | 'one' | 'two' | 'few' | 'many' | 'other';
/** A message's key as code uses it: plural forms share one ("rows" for rows_one and rows_other). */
export type MessageKey = keyof typeof en extends infer K ? (K extends `${infer Base}_${PluralSuffix}` ? Base : K) : never;

export type MessageParams = Record<string, unknown>;

/** A message to translate when it is shown. */
export interface Message {
  key: MessageKey;
  params?: MessageParams;
}

/** Text a host wrote in several languages: { en: 'Price', de: 'Preis', 'de-CH': 'Preis (CHF)' }. */
export type LocalizedMap = Readonly<Record<string, string>>;

/**
 * Text to show: as written (from the workbook or the host), a message to
 * translate, or host-written text in several languages.
 */
export type Text = string | Message | LocalizedMap;

export function msg(key: MessageKey, params?: MessageParams): Message {
  return params ? { key, params } : { key };
}

export function isMessage(x: unknown): x is Message {
  return typeof x === 'object' && x !== null && typeof (x as Message).key === 'string';
}

/**
 * Host-written text in one language or several: "Price" or
 * { en: 'Price', de: 'Preis', 'de-CH': 'Preis (CHF)' }. The most specific
 * match for the editor's language is shown, then English, then the first.
 */
export type Localized = string | LocalizedMap;

/**
 * An error people see, such as a change that would split a table: it carries
 * the message to translate, across the worker boundary too.
 */
export class TextError extends Error {
  constructor(readonly text: Text) {
    super(typeof text === 'string' ? text : isMessage(text) ? text.key : pickLanguage(text, 'en'));
    this.name = 'TextError';
  }
}

/** { en: 'Price', de: 'Preis' }: a plain object of strings (arrays, for list formatting, are left alone). */
export function isLocalizedMap(v: unknown): v is LocalizedMap {
  return typeof v === 'object' && v !== null && !Array.isArray(v) && !isMessage(v) && Object.values(v).every((x) => typeof x === 'string');
}

/** What to show for an error: the message it carries, or its own words. */
export function errorText(err: unknown): Text {
  const text = (err as { text?: unknown } | null)?.text;
  if (typeof text === 'string' || isMessage(text) || isLocalizedMap(text)) return text;
  return err instanceof Error ? err.message : String(err);
}

/** Host-written text in the language that fits best, without loading any catalogue. */
export function pickLanguage(text: Localized | undefined | null, language: string): string {
  if (text == null) return '';
  if (typeof text === 'string') return text;
  const lower = Object.fromEntries(Object.entries(text).map(([k, v]) => [k.toLowerCase(), v]));
  const lng = language.toLowerCase();
  return lower[lng] ?? lower[lng.split('-')[0]] ?? lower.en ?? Object.values(text)[0] ?? '';
}

