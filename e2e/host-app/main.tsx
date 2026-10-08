import { createRoot } from 'react-dom/client';
import { createWorkbookEditor, WorkbookEditor } from 'fake-spreadsheet';
import 'fake-spreadsheet/style.css';

// The library as a React app gets it from npm, with nothing configured.
const editor = createWorkbookEditor({ drafts: false, warnOnUnload: false });
createRoot(document.getElementById('root')!).render(<WorkbookEditor editor={editor} />);
