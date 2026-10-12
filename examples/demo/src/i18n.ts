/**
 * The demo page's own text, translated with the global i18next instance, the
 * way a host that already uses i18next would. The editor keeps a separate
 * instance for its messages, so the two never share keys or settings; the
 * demo switches the language of both together.
 */
import i18next from 'i18next';
import { initReactI18next } from 'react-i18next';
import { de } from './locales/de';
import { en } from './locales/en';

export { i18next };

/** The namespace of the demo's own text. */
export const DEMO_NS = 'demo';

/** The languages the demo offers, each named in its own language. */
export const LANGUAGES = [
  { code: 'en', name: 'English' },
  { code: 'de', name: 'Deutsch' },
] as const;

export const LANGUAGE_STORAGE_KEY = 'demo-language';

/** The language chosen last time, else German for a German browser, else English. */
function firstLanguage(): string {
  try {
    const stored = localStorage.getItem(LANGUAGE_STORAGE_KEY);
    if (LANGUAGES.some((l) => l.code === stored)) return stored!;
  } catch {
    // No stored choice.
  }
  return navigator.language.toLowerCase().startsWith('de') ? 'de' : 'en';
}

void i18next.use(initReactI18next).init({
  resources: { en: { [DEMO_NS]: en }, de: { [DEMO_NS]: de } },
  lng: firstLanguage(),
  fallbackLng: 'en',
  ns: [DEMO_NS],
  defaultNS: DEMO_NS,
  // Keys are flat ("welcome.heading"); React escapes what it renders.
  keySeparator: false,
  nsSeparator: false,
  interpolation: { escapeValue: false },
  initAsync: false,
});

/** The page's language and title follow the demo's language. */
function updatePage(): void {
  document.documentElement.lang = i18next.language;
  document.title = i18next.t('page.title');
}
updatePage();
i18next.on('languageChanged', updatePage);
