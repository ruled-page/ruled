import { version } from 'react';
import { createRoot } from 'react-dom/client';
import { createWorkbookEditor, WorkbookEditor } from 'workbook-editor';
import 'workbook-editor/style.css';

// The library as a React app gets it from npm, with nothing configured.
const editor = createWorkbookEditor({ drafts: false, warnOnUnload: false });
document.documentElement.dataset.react = version;
createRoot(document.getElementById('root')!).render(<WorkbookEditor editor={editor} />);
