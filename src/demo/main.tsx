import { StrictMode, useEffect, useState, type ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import { I18nextProvider, useTranslation } from 'react-i18next';
import { ENGINES, type EngineId } from '../core/engineIds';
import { WorkbookEditor } from '../editor';
import { handleFor } from '../editor/handle';
import { WorkbookStore } from '../state/store';
import { DEMO_NS, i18next, LANGUAGE_STORAGE_KEY, LANGUAGES } from './i18n';
import { DEMO_SCHEMAS } from './schemas';
import { DEMO_THEMES } from './themes';
import { openSample, SAMPLES, SCHEMA_SAMPLES, Welcome } from './Welcome';
import { Select } from '../ui/Select';
import './demo.css';

// A consumer passes its own schemas and save handling here; the demo
// registers the example schemas and lets saves download the file. The
// editor's text starts in the demo's language; numbers typed into cells are
// still read in the browser's locale.
// The demo reaches into the editor's store for its welcome screen; hosts use the handle.
/** The engine to calculate with: ?engine= in the link, else the one chosen last time. */
function firstEngine(): EngineId {
  const known = (id: string | null | undefined): id is EngineId => ENGINES.some((e) => e.id === id);
  const asked = new URLSearchParams(location.search).get('engine');
  if (known(asked)) return asked;
  try {
    const stored = localStorage.getItem('demo-engine');
    if (known(stored)) return stored;
  } catch {
    // No stored choice.
  }
  return 'native';
}

/** The samples, by file name, which the address names so a reload opens the same one. */
const SAMPLE_FILES = new Set([...SAMPLES, ...SCHEMA_SAMPLES].map((s) => s.file));

/** Keeps ?sample= in the address naming the sample open, or none for another file or none at all. */
function rememberSample(fileName: string | undefined): void {
  const url = new URL(location.href);
  if (fileName && SAMPLE_FILES.has(fileName)) url.searchParams.set('sample', fileName);
  else url.searchParams.delete('sample');
  if (url.href !== location.href) history.replaceState(history.state, '', url);
}

// The status bar says which engine is calculating and how long things take, for comparing engines.
const store = new WorkbookStore({
  schemas: DEMO_SCHEMAS,
  language: i18next.language,
  engine: firstEngine(),
  diagnostics: true,
  onEvent: (e) => {
    if (e.type === 'opened') rememberSample(e.fileName);
  },
});
const asked = new URLSearchParams(location.search).get('sample');
if (asked && SAMPLE_FILES.has(asked)) void openSample(store, asked);
const editor = handleFor(store);
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
    <Select className="demo-theme" value={id} onChange={(e) => choose(e.target.value)} aria-label={t('picker.theme')} title={t('picker.theme')}>
      {DEMO_THEMES.map((theme) => (
        <option key={theme.id} value={theme.id}>
          {t(theme.label)}
        </option>
      ))}
    </Select>
  );
}

/** Back to the list of samples, closing the workbook (the page asks first if it has unsaved edits). */
function SamplesButton() {
  const { t } = useTranslation(DEMO_NS);
  const [open, setOpen] = useState(!!store.fileName);
  useEffect(() => {
    const stop = store.subscribe(() => setOpen(!!store.fileName));
    return () => void stop();
  }, []);
  if (!open) return null;
  return (
    <button
      className="wbe-btn"
      onClick={() => {
        const url = new URL(location.href);
        url.searchParams.delete('sample');
        location.assign(url);
      }}
    >
      {t('toolbar.samples')}
    </button>
  );
}

/** Lets the demo calculate with a different engine, opening the workbook again with it. */
function EnginePicker() {
  const { t } = useTranslation(DEMO_NS);
  const [id, setId] = useState<EngineId>(store.engine);
  const choose = (next: EngineId) => {
    setId(next);
    void store.setEngine(next);
    try {
      localStorage.setItem('demo-engine', next);
    } catch {
      // Remembering the choice is a nicety.
    }
  };
  return (
    <Select className="demo-engine" value={id} onChange={(e) => choose(e.target.value as EngineId)} aria-label={t('picker.engine')} title={t('picker.engine')}>
      {ENGINES.map((engine) => (
        <option key={engine.id} value={engine.id}>
          {engine.name}
        </option>
      ))}
    </Select>
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
    <Select className="demo-language" value={i18n.language} onChange={(e) => choose(e.target.value)} aria-label={t('picker.language')} title={t('picker.language')}>
      {LANGUAGES.map((l) => (
        <option key={l.code} value={l.code} lang={l.code}>
          {l.name}
        </option>
      ))}
    </Select>
  );
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <WorkbookEditor
      editor={editor}
      toolbarEnd={
        <HostText>
          <SamplesButton />
          <EnginePicker />
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
