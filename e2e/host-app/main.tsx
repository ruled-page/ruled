import { createRoot } from 'react-dom/client';
import { WorkbookEditor, WorkbookStore } from 'fake-spreadsheet';
import 'fake-spreadsheet/style.css';

// The library as a React app gets it from npm, with nothing configured.
const store = new WorkbookStore({ drafts: false, warnOnUnload: false });
createRoot(document.getElementById('root')!).render(<WorkbookEditor store={store} />);
