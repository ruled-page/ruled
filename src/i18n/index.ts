/**
 * Translation, with i18next. Every word the editor shows comes from the
 * `wbe` namespace: English is built in and complete, German is built in, and
 * a host can add any language (or its own wording for some messages) with
 * `translations`. Messages are i18next JSON: {{name}} placeholders,
 * {{count, number}} formatting and _one/_other plural keys.
 *
 * Each editor has its own i18next instance, so a host that uses i18next
 * itself keeps its own setup.
 *
 * Code that runs in the worker returns messages as data
 * (`msg('core.file.tooBig', { size })`) rather than sentences, and the editor
 * translates them when it shows them, so changing language changes everything
 * already on screen. Text from the workbook or the host's schema is shown as
 * written; schemas can give it in several languages (see `Localized`).
 */
import i18next, { type i18n as I18n } from 'i18next';
import { isLocalizedMap, isMessage, pickLanguage, type Localized, type MessageKey, type MessageParams, type Text } from './message';
import { de } from './messages/de';
import { en } from './messages/en';

export * from './message';

export { en } from './messages/en';
export { de } from './messages/de';

/** The i18next namespace the editor's messages live in. */
export const NAMESPACE = 'wbe';

/** Every message, in English: the catalogue translations follow. */
export type Messages = Record<keyof typeof en, string>;

/**
 * Translations by language: whole catalogues, or a few messages to word
 * differently. Keys are the English catalogue's, with the plural forms the
 * language needs (a Polish catalogue has rows_few and rows_many as well).
 */
export type Translations = Readonly<Record<string, Readonly<Record<string, string>>>>;

/** The languages with a built-in catalogue. */
export const BUILT_IN: Translations = { en, de };

/** An i18next instance of the editor's own, in `language`, with the host's translations added. */
export function createI18n(language: string, translations?: Translations): I18n {
  const resources: Record<string, Record<string, Record<string, string>>> = {};
  for (const source of [BUILT_IN, translations ?? {}]) {
    for (const [lng, messages] of Object.entries(source)) {
      resources[lng] = { [NAMESPACE]: { ...resources[lng]?.[NAMESPACE], ...(messages as Record<string, string>) } };
    }
  }
  const instance = i18next.createInstance();
  void instance.init({
    resources,
    lng: language,
    fallbackLng: 'en',
    ns: [NAMESPACE],
    defaultNS: NAMESPACE,
    // Keys are flat ("toolbar.open"); React escapes what it renders.
    keySeparator: false,
    nsSeparator: false,
    interpolation: { escapeValue: false },
    initAsync: false,
  });
  return instance;
}

export interface Translator {
  (key: MessageKey, params?: MessageParams): string;
  /** Translates a message, or returns written text as it is. */
  text(text: Text | undefined | null): string;
  /** Picks the right language from host-written text. */
  localized(text: Localized | undefined | null): string;
  number(n: number, options?: Intl.NumberFormatOptions): string;
  date(d: Date | number, options?: Intl.DateTimeFormatOptions): string;
  /** "a, b and c" (or "a, b or c") in the language's style. */
  list(items: string[], type?: 'conjunction' | 'disjunction'): string;
  /** The language shown, such as "en-US" or "de". */
  readonly language: string;
  readonly dir: 'ltr' | 'rtl';
}

/** Translation functions for code outside React, following the instance's current language. */
export function translator(i18n: I18n): Translator {
  const language = () => i18n.language || 'en';
  const resolve = (params: MessageParams | undefined): MessageParams | undefined => {
    if (!params) return params;
    let out: MessageParams | undefined;
    // Messages and host-written text in several languages inside a message are resolved first.
    for (const [k, v] of Object.entries(params)) if (isMessage(v) || isLocalizedMap(v)) (out ??= { ...params })[k] = t.text(v);
    return out ?? params;
  };
  const t = ((key: MessageKey, params?: MessageParams) => i18n.t(key, { ns: NAMESPACE, ...resolve(params) }) as string) as Translator;
  Object.defineProperties(t, {
    language: { get: language },
    dir: { get: () => i18n.dir(language()) },
  });
  Object.assign(t, {
    text: (text: Text | undefined | null) =>
      text == null ? '' : typeof text === 'string' ? text : isMessage(text) ? t(text.key, text.params) : pickLanguage(text, language()),
    localized: (text: Localized | undefined | null) => pickLanguage(text, language()),
    number: (n: number, options?: Intl.NumberFormatOptions) => n.toLocaleString(language(), options),
    date: (d: Date | number, options?: Intl.DateTimeFormatOptions) => new Date(d).toLocaleString(language(), options ?? { dateStyle: 'medium' }),
    list: (items: string[], type: 'conjunction' | 'disjunction' = 'conjunction') => new Intl.ListFormat(language(), { type }).format(items),
  });
  return t;
}

/** English, for code with no editor to say otherwise (the Node API, tests). */
export const english: Translator = translator(createI18n('en'));
