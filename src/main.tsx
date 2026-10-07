import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import { DEMO_SCHEMAS } from './demo/schemas';
import { WorkbookStore } from './state/store';
import './styles.css';

// A consumer passes its own schemas here; the demo registers the examples.
const store = new WorkbookStore({ schemas: DEMO_SCHEMAS });
if (import.meta.env.DEV) (window as unknown as { __store: WorkbookStore }).__store = store;

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App store={store} />
  </StrictMode>,
);
