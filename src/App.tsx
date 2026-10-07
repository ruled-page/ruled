import { useEffect, useRef, useState, useSyncExternalStore } from 'react';
import type { WorkbookStore } from './state/store';
import { Banner, SchemaBanner, SheetTabs, StatusBar, Toolbar } from './ui/Chrome';
import { FormulaBar } from './ui/FormulaBar';
import { Icon } from './ui/icons';
import { ConfirmDialog, Toasts } from './ui/Overlays';
import { SchemaDesigner } from './ui/SchemaDesigner';
import { SheetGrid } from './ui/SheetGrid';
import { SidePanel } from './ui/SidePanel';
import { Welcome } from './ui/Welcome';

function isTextField(el: EventTarget | null): boolean {
  return el instanceof HTMLElement && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable);
}

export function App({ store }: { store: WorkbookStore }) {
  useSyncExternalStore(store.subscribe, store.getVersion);
  const fileInput = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const openPicker = () => fileInput.current?.click();

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey;
      if (!mod || store.status !== 'ready') return;
      const key = e.key.toLowerCase();
      if (key === 's') {
        e.preventDefault();
        void store.download();
      } else if (!isTextField(e.target) && !store.confirmRequest && !store.designerOpen) {
        if (key === 'z' && !e.shiftKey) {
          e.preventDefault();
          void store.undo();
        } else if ((key === 'z' && e.shiftKey) || key === 'y') {
          e.preventDefault();
          void store.redo();
        }
      }
    };
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (store.editedCount > 0) e.preventDefault();
    };
    window.addEventListener('keydown', onKey);
    window.addEventListener('beforeunload', onBeforeUnload);
    return () => {
      window.removeEventListener('keydown', onKey);
      window.removeEventListener('beforeunload', onBeforeUnload);
    };
  }, [store]);

  const openFile = async (file: File | undefined) => {
    if (file) await store.open(file.name, await file.arrayBuffer());
  };

  const view = store.view;
  const selection = store.selection;
  const ready = store.status === 'ready' && view;

  return (
    <div
      className={`app ${store.panel ? 'with-report' : ''}`}
      onDragOver={(e) => {
        if (e.dataTransfer.types.includes('Files')) {
          e.preventDefault();
          setDragging(true);
        }
      }}
      onDragLeave={(e) => {
        if (e.currentTarget === e.target) setDragging(false);
      }}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        void openFile(e.dataTransfer.files[0]);
      }}
    >
      <Toolbar store={store} onOpen={openPicker} />
      <input
        ref={fileInput}
        type="file"
        accept=".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        hidden
        onChange={(e) => {
          void openFile(e.target.files?.[0]);
          e.target.value = '';
        }}
      />
      {ready && <SchemaBanner store={store} />}
      {ready && <Banner store={store} />}
      {ready && view.info.kind === 'worksheet' && <FormulaBar store={store} view={view} selection={selection} />}
      <div className="workspace">
        <main className="sheet-area">
          {store.status === 'empty' && <Welcome store={store} onOpen={openPicker} />}
          {store.status === 'loading' && (
            <div className="center-state">
              <span className="spinner" />
              <p>{store.stage}</p>
            </div>
          )}
          {store.status === 'error' && (
            <div className="center-state error">
              <Icon name="warn" size={28} />
              <p>{store.error}</p>
              <button className="btn" onClick={openPicker}>
                Open another file
              </button>
            </div>
          )}
          {ready && view.info.kind === 'worksheet' && (
            <SheetGrid
              store={store}
              view={view}
              dataVersion={view.version * 31 + store.styles.length}
              schemaVersion={store.schemaVersion}
              selection={selection}
            />
          )}
          {ready && view.info.kind !== 'worksheet' && (
            <div className="center-state">
              <Icon name="chart" size={32} />
              <h2>{view.info.name}</h2>
              <p>
                {view.info.objects.find((o) => o.kind === 'chart')?.chartType ?? 'This'} chart
                {view.info.objects.find((o) => o.kind === 'chart')?.title ? ` “${view.info.objects.find((o) => o.kind === 'chart')?.title}”` : ''} fills
                this whole sheet. It is kept in the file and Excel redraws it from your edited data; this editor doesn’t display charts.
              </p>
            </div>
          )}
        </main>
        {ready && <SidePanel store={store} />}
      </div>
      {ready && <SheetTabs store={store} />}
      <StatusBar store={store} view={ready ? view : undefined} selection={ready ? selection : undefined} />
      {dragging && (
        <div className="drop-overlay">
          <Icon name="open" size={28} />
          <p>Drop to open</p>
        </div>
      )}
      {store.designerOpen && <SchemaDesigner store={store} />}
      <ConfirmDialog store={store} />
      <Toasts store={store} />
    </div>
  );
}
