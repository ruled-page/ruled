import { useEffect, useRef, useState, useSyncExternalStore, type ReactNode } from 'react';
import type { WorkbookStore } from '../state/store';
import { Banner, SchemaBanner, SheetTabs, StatusBar, Toolbar } from '../ui/Chrome';
import { FilterPicker } from '../ui/FilterPicker';
import { FindPanel } from '../ui/FindPanel';
import { FormulaBar } from '../ui/FormulaBar';
import { GridMenu } from '../ui/GridMenu';
import { Icon } from '../ui/icons';
import { ConfirmDialog, Toasts } from '../ui/Overlays';
import { SchemaDesigner } from '../ui/SchemaDesigner';
import { SheetPanes } from '../ui/SheetPanes';
import { SidePanel } from '../ui/SidePanel';
import './editor.css';

function isTextField(el: EventTarget | null): boolean {
  return el instanceof HTMLElement && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable);
}

/**
 * Glide renders cell editors into an element with id "portal". Create it
 * when the host page doesn't have one, inside the editor's style scope.
 */
function ensurePortal(): void {
  let portal = document.getElementById('portal');
  if (!portal) {
    portal = document.createElement('div');
    portal.id = 'portal';
    document.body.appendChild(portal);
  }
  portal.classList.add('wbe', 'wbe-portal');
}

export interface WorkbookEditorProps {
  store: WorkbookStore;
  /** Shown before a workbook is opened; receives a function that opens the file picker. */
  emptyState?: (openPicker: () => void) => ReactNode;
  className?: string;
}

function DefaultEmptyState({ openPicker, canOpen }: { openPicker: () => void; canOpen: boolean }) {
  return (
    <div className="center-state">
      {canOpen ? (
        <button className="dropzone" onClick={openPicker}>
          <Icon name="open" size={22} />
          <span>
            <strong>Choose a workbook</strong> or drop it here
          </span>
          <span className="muted">.xlsx or .xlsm, processed entirely in your browser</span>
        </button>
      ) : (
        <p className="muted">No workbook is open.</p>
      )}
    </div>
  );
}

/** The workbook editor: toolbar, formula bar, grid, sheet tabs and panels. */
export function WorkbookEditor({ store, emptyState, className }: WorkbookEditorProps) {
  useSyncExternalStore(store.subscribe, store.getVersion);
  const root = useRef<HTMLDivElement>(null);
  const fileInput = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const canOpen = store.features.open;
  const openPicker = () => {
    if (canOpen) fileInput.current?.click();
  };

  useEffect(ensurePortal, []);

  useEffect(() => {
    // Shortcuts only apply while focus is in this editor (or its popups).
    const inEditor = (target: EventTarget | null) =>
      target instanceof Node && (root.current?.contains(target) || document.getElementById('portal')?.contains(target));
    const onKey = (e: KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey;
      if (!mod || store.status !== 'ready' || !inEditor(e.target)) return;
      const key = e.key.toLowerCase();
      if (key === 's' && store.features.save) {
        e.preventDefault();
        void store.download();
      } else if (key === 'f' && !store.designerOpen) {
        e.preventDefault();
        store.set('findOpen', true);
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
      if (store.warnOnUnload && store.editedCount > 0) e.preventDefault();
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
    <div className={`wbe wbe-root ${className ?? ''}`} ref={root}>
      <div
        className={`app ${store.panel ? 'with-report' : ''}`}
        onDragOver={(e) => {
          if (canOpen && e.dataTransfer.types.includes('Files')) {
            e.preventDefault();
            setDragging(true);
          }
        }}
        onDragLeave={(e) => {
          if (e.currentTarget === e.target) setDragging(false);
        }}
        onDrop={(e) => {
          if (!canOpen) return;
          e.preventDefault();
          setDragging(false);
          void openFile(e.dataTransfer.files[0]);
        }}
      >
        <Toolbar store={store} onOpen={openPicker} />
        {canOpen && (
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
        )}
        {ready && <SchemaBanner store={store} />}
        {ready && <Banner store={store} />}
        {ready && view.info.kind === 'worksheet' && <FormulaBar store={store} view={view} selection={selection} />}
        <div className="workspace">
          <main className="sheet-area">
            {store.status === 'empty' && (emptyState ? emptyState(openPicker) : <DefaultEmptyState openPicker={openPicker} canOpen={canOpen} />)}
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
                {canOpen && (
                  <button className="btn" onClick={openPicker}>
                    Open another file
                  </button>
                )}
              </div>
            )}
            {ready && view.info.kind === 'worksheet' && (
              <SheetPanes
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
                  {view.info.objects.find((o) => o.kind === 'chart')?.title ? ` “${view.info.objects.find((o) => o.kind === 'chart')?.title}”` : ''}{' '}
                  fills this whole sheet. It is kept in the file and Excel redraws it from your edited data; this editor doesn’t display charts.
                </p>
              </div>
            )}
            {ready && store.findOpen && <FindPanel store={store} />}
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
        {ready && <GridMenu store={store} />}
        {ready && <FilterPicker store={store} />}
        <ConfirmDialog store={store} />
        <Toasts store={store} />
      </div>
    </div>
  );
}
