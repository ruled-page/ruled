import { StrictMode, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { WorkbookEditor, WorkbookStore } from '../editor';
import { DEMO_SCHEMAS } from './schemas';
import { DEMO_THEMES } from './themes';
import { Welcome } from './Welcome';
import './demo.css';

// A consumer passes its own schemas and save handling here; the demo
// registers the example schemas and lets saves download the file.
const store = new WorkbookStore({ schemas: DEMO_SCHEMAS });
if (import.meta.env.DEV) (window as unknown as { __store: WorkbookStore }).__store = store;

/** Lets the demo switch between example themes, to show the editor can be restyled throughout. */
function ThemePicker() {
  const [name, setName] = useState(() => {
    try {
      return localStorage.getItem('demo-theme') ?? 'Default';
    } catch {
      return 'Default';
    }
  });
  const choose = (next: string) => {
    setName(next);
    store.setTheme(DEMO_THEMES.find((t) => t.name === next)?.theme);
    try {
      localStorage.setItem('demo-theme', next);
    } catch {
      // Remembering the choice is a nicety.
    }
  };
  return (
    <select className="demo-theme" value={name} onChange={(e) => choose(e.target.value)} aria-label="Theme" title="Theme">
      {DEMO_THEMES.map((t) => (
        <option key={t.name}>{t.name}</option>
      ))}
    </select>
  );
}

try {
  store.setTheme(DEMO_THEMES.find((t) => t.name === localStorage.getItem('demo-theme'))?.theme);
} catch {
  // No stored choice.
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <WorkbookEditor store={store} toolbarEnd={<ThemePicker />} emptyState={(openPicker) => <Welcome store={store} onOpen={openPicker} />} />
  </StrictMode>,
);
