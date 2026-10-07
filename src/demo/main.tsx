import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { WorkbookEditor, WorkbookStore } from '../editor';
import { DEMO_SCHEMAS } from './schemas';
import { Welcome } from './Welcome';
import './demo.css';

// A consumer passes its own schemas and save handling here; the demo
// registers the example schemas and lets saves download the file.
const store = new WorkbookStore({ schemas: DEMO_SCHEMAS });
if (import.meta.env.DEV) (window as unknown as { __store: WorkbookStore }).__store = store;

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <WorkbookEditor store={store} emptyState={(openPicker) => <Welcome store={store} onOpen={openPicker} />} />
  </StrictMode>,
);
