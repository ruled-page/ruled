import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import { WorkbookStore } from './state/store';
import './styles.css';

const store = new WorkbookStore();
if (import.meta.env.DEV) (window as unknown as { __store: WorkbookStore }).__store = store;

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App store={store} />
  </StrictMode>,
);
