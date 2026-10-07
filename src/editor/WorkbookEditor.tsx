import { useEffect, useId, useRef, useState, useSyncExternalStore, type CSSProperties, type ReactNode } from 'react';
import type { WorkbookStore } from '../state/store';
import { Banner, SchemaBanner, SheetTabs, StatusBar, Toolbar } from '../ui/Chrome';
import { FilterPicker } from '../ui/FilterPicker';
import { FindPanel } from '../ui/FindPanel';
import { FormulaBar } from '../ui/FormulaBar';
import { GridMenu } from '../ui/GridMenu';
import { Icon, IconContext } from '../ui/icons';
import { ColumnWidthDialog, ShortcutsDialog } from '../ui/KeyboardDialogs';
import { MOD } from '../ui/keys';
import { firstFocusable } from '../ui/useDialog';
import { copyTheme, themeStyle } from './theme';
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
/** The editor's parts, in the order F6 visits them. */
const REGIONS = ['.toolbar', '.formula-bar', '.sheet-area', '.report', '.sheet-tabs'];
/** While the schema designer waits for cells to be picked, F6 goes between the sheet and its bar. */
const PICKING = ['.pick-bar', '.sheet-area'];

/** Moves the focus to the next (or previous) part of the editor, as F6 does in Excel. */
function nextRegion(root: HTMLElement, store: WorkbookStore, back: boolean): void {
  const present = (root.querySelector('.pick-bar') ? PICKING : REGIONS)
    .map((s) => root.querySelector<HTMLElement>(s))
    .filter((el): el is HTMLElement => !!el);
  if (!present.length) return;
  const at = present.findIndex((el) => el.contains(document.activeElement));
  const n = present.length;
  const next = present[at < 0 ? (back ? n - 1 : 0) : (at + (back ? n - 1 : 1)) % n];
  if (next.matches('.sheet-area') && store.view?.info.kind === 'worksheet') return store.focusGrid();
  const target =
    next.querySelector<HTMLElement>('[role="tab"][aria-selected="true"], [aria-current="true"]') ?? firstFocusable(next) ?? next;
  target.focus();
}

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
  /** The host's own controls, at the end of the toolbar before the Save button. */
  toolbarEnd?: ReactNode;
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
export function WorkbookEditor({ store, emptyState, toolbarEnd, className }: WorkbookEditorProps) {
  useSyncExternalStore(store.subscribe, store.getVersion);
  const root = useRef<HTMLDivElement>(null);
  const hintId = useId();
  const fileInput = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const canOpen = store.features.open;
  const openPicker = () => {
    if (canOpen) fileInput.current?.click();
  };

  useEffect(ensurePortal, []);

  // Glide opens cell editors in one portal for the whole page; it takes the
  // theme of the editor being used, however the host set that theme.
  const themeVersion = store.themeVersion;
  useEffect(() => {
    const el = root.current;
    const portal = document.getElementById('portal');
    if (!el || !portal) return;
    const sync = () => copyTheme(el, portal);
    sync();
    el.addEventListener('focusin', sync);
    el.addEventListener('pointerdown', sync, true);
    return () => {
      el.removeEventListener('focusin', sync);
      el.removeEventListener('pointerdown', sync, true);
    };
  }, [themeVersion]);

  useEffect(() => {
    // Shortcuts only apply while focus is in this editor (or its popups).
    const inEditor = (target: EventTarget | null) =>
      target instanceof Node && (root.current?.contains(target) || document.getElementById('portal')?.contains(target));
    const onKey = (e: KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey;
      if (!inEditor(e.target) || !root.current) return;
      // F6 is Excel's key for this; browsers that keep F6 for their address bar leave Ctrl+F6.
      if (e.key === 'F6' && !e.altKey && !e.metaKey && (!store.designerOpen || store.designerPick) && !store.confirmRequest) {
        e.preventDefault();
        nextRegion(root.current, store, e.shiftKey);
        return;
      }
      if (!mod) return;
      const key = e.key.toLowerCase();
      if (key === '/') {
        e.preventDefault();
        store.set('shortcutsOpen', true);
        return;
      }
      if (store.status !== 'ready') return;
      if (key === 'g' && !store.designerOpen) {
        e.preventDefault();
        root.current.querySelector<HTMLInputElement>('.name-box')?.focus();
      } else if (key === 's' && store.features.save) {
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
        } else if ((key === 'd' || key === 'r') && !e.shiftKey && !e.altKey) {
          e.preventDefault();
          void store.fillDown(key === 'r');
        }
      }
    };
    const onBeforeUnload = (e: BeforeUnloadEvent) => {
      if (store.warnOnUnload && store.unsaved) e.preventDefault();
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
    <div className={`wbe wbe-root ${className ?? ''}`} ref={root} style={themeStyle(store.theme) as CSSProperties}>
      <IconContext.Provider value={store.icons}>
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
          <Toolbar store={store} onOpen={openPicker} extra={toolbarEnd} />
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
            <main className="sheet-area" aria-label={ready ? `Sheet ${view.info.name}` : undefined} aria-describedby={ready ? `${hintId}-hint` : undefined}>
              {ready && (
                <p id={`${hintId}-hint`} className="sr-only">
                  Arrow keys and Tab move between cells; type to replace a cell, or press Enter to edit it. Shift+F10 opens the cell menu, F6 moves to the
                  next part of the editor, and {MOD}/ lists every shortcut.
                </p>
              )}
              {store.status === 'empty' && (emptyState ? emptyState(openPicker) : <DefaultEmptyState openPicker={openPicker} canOpen={canOpen} />)}
              {store.status === 'loading' && (
                <div className="center-state" role="status">
                  <span className="spinner" aria-hidden="true" />
                  <p>{store.stage}</p>
                </div>
              )}
              {store.status === 'error' && (
                <div className="center-state error" role="alert">
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
          {ready && <ColumnWidthDialog store={store} />}
          <ShortcutsDialog store={store} />
          <ConfirmDialog store={store} />
          <Toasts store={store} />
        </div>
      </IconContext.Provider>
    </div>
  );
}
