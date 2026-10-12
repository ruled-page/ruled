/**
 * A procurement portal with the workbook editor embedded in it: suppliers
 * upload a price list, check and correct it against the portal's rules, and
 * submit it; the portal reads the values into its own records. Everything
 * runs in the browser.
 */
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import './i18n';
import { App } from './App';
import './portal.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
