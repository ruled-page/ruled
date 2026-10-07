import { StrictMode, useState, type ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import { I18nextProvider, useTranslation } from 'react-i18next';
import { WorkbookEditor, WorkbookStore } from '../editor';
import { DEMO_NS, i18next, LANGUAGE_STORAGE_KEY, LANGUAGES } from './i18n';
import { DEMO_SCHEMAS } from './schemas';
import { DEMO_THEMES } from './themes';
import { Welcome } from './Welcome';
import './demo.css';

// A consumer passes its own schemas and save handling here; the demo
// registers the example schemas and lets saves download the file. The
// editor's text starts in the demo's language; numbers typed into cells are
// still read in the browser's locale.
const store = new WorkbookStore({ schemas: DEMO_SCHEMAS, language: i18next.language });
if (import.meta.env.DEV) (window as unknown as { __store: WorkbookStore }).__store = store;

/**
 * The editor renders toolbarEnd and emptyState inside its own i18next
 * provider; this hands the demo's parts back to the host's instance.
 */
function HostText({ children }: { children: ReactNode }) {
  return (
    <I18nextProvider i18n={i18next} defaultNS={DEMO_NS}>
      {children}
    </I18nextProvider>
  );
}

/** The theme chosen last time. Earlier versions of the demo stored the theme's English name. */
function storedTheme(): string {
  try {
    const stored = localStorage.getItem('demo-theme')?.toLowerCase();
    if (DEMO_THEMES.some((t) => t.id === stored)) return stored!;
  } catch {
    // No stored choice.
  }
  return 'default';
}

const firstTheme = storedTheme();
store.setTheme(DEMO_THEMES.find((t) => t.id === firstTheme)?.theme);

/** Lets the demo switch between example themes, to show the editor can be restyled throughout. */
function ThemePicker() {
  const { t } = useTranslation(DEMO_NS);
  const [id, setId] = useState(firstTheme);
  const choose = (next: string) => {
    setId(next);
    store.setTheme(DEMO_THEMES.find((theme) => theme.id === next)?.theme);
    try {
      localStorage.setItem('demo-theme', next);
    } catch {
      // Remembering the choice is a nicety.
    }
  };
  return (
    <select className="demo-theme" value={id} onChange={(e) => choose(e.target.value)} aria-label={t('picker.theme')} title={t('picker.theme')}>
      {DEMO_THEMES.map((theme) => (
        <option key={theme.id} value={theme.id}>
          {t(theme.label)}
        </option>
      ))}
    </select>
  );
}

/** Switches the language of the demo's own text and of the editor together. */
function LanguagePicker() {
  const { t, i18n } = useTranslation(DEMO_NS);
  const choose = (lng: string) => {
    void i18next.changeLanguage(lng);
    store.setLanguage(lng);
    try {
      localStorage.setItem(LANGUAGE_STORAGE_KEY, lng);
    } catch {
      // Remembering the choice is a nicety.
    }
  };
  return (
    <select className="demo-language" value={i18n.language} onChange={(e) => choose(e.target.value)} aria-label={t('picker.language')} title={t('picker.language')}>
      {LANGUAGES.map((l) => (
        <option key={l.code} value={l.code} lang={l.code}>
          {l.name}
        </option>
      ))}
    </select>
  );
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <WorkbookEditor
      store={store}
      toolbarEnd={
        <HostText>
          <ThemePicker />
          <LanguagePicker />
        </HostText>
      }
      emptyState={(openPicker) => (
        <HostText>
          <Welcome store={store} onOpen={openPicker} />
        </HostText>
      )}
    />
  </StrictMode>,
);
