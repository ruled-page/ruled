/**
 * The portal's own text, with the global i18next instance, as a host that
 * already uses i18next would have it. The editor keeps its own instance; the
 * portal tells it when the language changes.
 */
import i18next from 'i18next';
import { initReactI18next } from 'react-i18next';
import { de } from './locales/de';
import { en } from './locales/en';

export { i18next };

export const NS = 'portal';

/** The languages the portal offers, each named in its own language. */
export const LANGUAGES = [
  { code: 'en', name: 'English' },
  { code: 'de', name: 'Deutsch' },
] as const;

const STORAGE_KEY = 'portal-language';

function firstLanguage(): string {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (LANGUAGES.some((l) => l.code === stored)) return stored!;
  } catch {
    // No stored choice.
  }
  return navigator.language.toLowerCase().startsWith('de') ? 'de' : 'en';
}

void i18next.use(initReactI18next).init({
  resources: { en: { [NS]: en }, de: { [NS]: de } },
  lng: firstLanguage(),
  fallbackLng: 'en',
  ns: [NS],
  defaultNS: NS,
  keySeparator: false,
  nsSeparator: false,
  interpolation: { escapeValue: false },
  initAsync: false,
});

const sync = (lng: string) => {
  document.documentElement.lang = lng;
  document.title = i18next.t('app.title', { ns: NS });
};
sync(i18next.language);
i18next.on('languageChanged', sync);

export function chooseLanguage(lng: string): void {
  void i18next.changeLanguage(lng);
  try {
    localStorage.setItem(STORAGE_KEY, lng);
  } catch {
    // Remembering the choice is a nicety.
  }
}
