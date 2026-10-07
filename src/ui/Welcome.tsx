import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';

export const SAMPLES = [
  {
    file: 'kitchen-sink.xlsx',
    title: 'Sales report',
    detail: 'Charts, a chart sheet, a table, validation, notes, pictures, and a few formulas this editor can’t calculate.',
  },
  {
    file: 'pivot.xlsx',
    title: 'Pivot table',
    detail: 'A pivot table plus filled-down formulas, which Excel stores in a compact shared form.',
  },
  {
    file: 'big.xlsx',
    title: '100,000 rows',
    detail: '1.1 million cells and 300,000 formulas. Create it with npm run samples.',
  },
];

export async function openSample(store: WorkbookStore, file: string): Promise<void> {
  const res = await fetch(`${import.meta.env.BASE_URL}samples/${file}`);
  const bytes = res.ok ? await res.arrayBuffer() : new ArrayBuffer(0);
  // Dev servers answer missing files with the app's HTML, so check for a zip signature.
  const head = new Uint8Array(bytes, 0, Math.min(2, bytes.byteLength));
  if (head[0] !== 0x50 || head[1] !== 0x4b) {
    store.toast('error', `Sample ${file} was not found. Run “npm run samples” to create it.`);
    return;
  }
  await store.open(file, bytes);
}

export function Welcome({ store, onOpen }: { store: WorkbookStore; onOpen: () => void }) {
  return (
    <div className="welcome">
      <div className="welcome-card">
        <h1>Edit an Excel workbook in the browser</h1>
        <p className="lede">
          Open a .xlsx file, change cells, paste from Excel, drag to fill, and save it back as .xlsx. Formulas recalculate as you
          edit. Charts, pivot tables, macros and formatting are kept in the saved file, even the ones this editor can’t display.
        </p>
        <button className="dropzone" onClick={onOpen}>
          <Icon name="open" size={22} />
          <span>
            <strong>Choose a workbook</strong> or drop it anywhere on this page
          </span>
          <span className="muted">.xlsx or .xlsm, processed entirely in your browser</span>
        </button>
        <h2>Or try a sample</h2>
        <ul className="samples-list">
          {SAMPLES.map((s) => (
            <li key={s.file}>
              <button onClick={() => void openSample(store, s.file)}>
                <Icon name="sheet" />
                <span>
                  <strong>{s.title}</strong>
                  <span>{s.detail}</span>
                </span>
              </button>
            </li>
          ))}
        </ul>
      </div>
    </div>
  );
}
