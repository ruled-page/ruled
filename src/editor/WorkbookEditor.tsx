import { useEffect, useId, useRef, useState, useSyncExternalStore, type CSSProperties, type ReactNode } from 'react';
import type { SheetView } from '../state/sheetView';
import type { WorkbookStore } from '../state/store';
import { Banner, SchemaBanner, SheetTabs, StatusBar, Toolbar } from '../ui/Chrome';
import { ChartSlot } from '../ui/ChartSlot';
import { CommentPopup } from '../ui/CommentPopup';
import { FilterPicker } from '../ui/FilterPicker';
import { FindPanel } from '../ui/FindPanel';
import { FormatBar } from '../ui/FormatBar';
import { FormulaBar } from '../ui/FormulaBar';
import { PORTAL_ID } from '../ui/glideDom';
import { GridMenu } from '../ui/GridMenu';
import { Icon, IconContext } from '../ui/icons';
import { ShortcutsDialog, SizeDialog } from '../ui/KeyboardDialogs';
import { NameDialog, NameManager } from '../ui/NameDialogs';
import { APPLE, shortcutKey, shortcutsKey } from '../ui/keys';
import { firstFocusable } from '../ui/useDialog';
import { EditorBoundary } from './ErrorBoundary';
import { storeOf, type EditorHandle } from './handle';
import { copyTheme, themeStyle } from './theme';
import { I18nextProvider, Trans } from 'react-i18next';
import { NAMESPACE } from '../i18n';
import { useT } from '../i18n/react';
import { ConfirmDialog, Toasts } from '../ui/Overlays';
import { SchemaDesigner } from '../ui/SchemaDesigner';
import { SheetPanes } from '../ui/SheetPanes';
import { SidePanel } from '../ui/SidePanel';
import './editor.css';

function isTextField(el: EventTarget | null): boolean {
  return el instanceof HTMLElement && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable);
}

/** The editor's parts, in the order F6 visits them. */
const REGIONS = ['.wbe-toolbar', '.wbe-format-bar', '.wbe-formula-bar', '.wbe-sheet-area', '.wbe-report', '.wbe-sheet-tabs'];
/** While the schema designer waits for cells to be picked, F6 goes between the sheet and its bar. */
const PICKING = ['.wbe-pick-bar', '.wbe-sheet-area'];

/** Moves the focus to the next (or previous) part of the editor, as F6 does in Excel. */
function nextRegion(root: HTMLElement, store: WorkbookStore, back: boolean): void {
  const present = (root.querySelector('.wbe-pick-bar') ? PICKING : REGIONS)
    .map((s) => root.querySelector<HTMLElement>(s))
    .filter((el): el is HTMLElement => !!el);
  if (!present.length) return;
  const at = present.findIndex((el) => el.contains(document.activeElement));
  const n = present.length;
  const next = present[at < 0 ? (back ? n - 1 : 0) : (at + (back ? n - 1 : 1)) % n];
  if (next.matches('.wbe-sheet-area') && store.view?.info.kind === 'worksheet') return store.focusGrid();
  const target =
    next.querySelector<HTMLElement>('[role="tab"][aria-selected="true"], [aria-current="true"]') ?? firstFocusable(next) ?? next;
  target.focus();
}

/** The element Glide opens cell editors in, one for the page, inside the editor's style scope. */
function ensurePortal(): void {
  if (document.getElementById(PORTAL_ID)) return;
  const portal = document.createElement('div');
  portal.id = PORTAL_ID;
  portal.className = 'wbe wbe-portal';
  document.body.appendChild(portal);
}

export interface WorkbookEditorProps {
  /** The editor to show, from createWorkbookEditor(). */
  editor: EditorHandle;
  /** Shown before a workbook is opened; receives a function that opens the file picker. */
  emptyState?: (openPicker: () => void) => ReactNode;
  /** The host's own controls, at the end of the toolbar before the Save button. */
  toolbarEnd?: ReactNode;
  className?: string;
}

function DefaultEmptyState({ openPicker, canOpen, csv }: { openPicker: () => void; canOpen: boolean; csv: boolean }) {
  const t = useT();
  return (
    <div className="wbe-center-state">
      {canOpen ? (
        <button className="wbe-dropzone" onClick={openPicker}>
          <Icon name="open" size={22} />
          <span>
            <Trans i18nKey="editor.chooseWorkbook" components={{ b: <strong /> }} />
          </span>
          <span className="wbe-muted">{t(csv ? 'editor.chooseWorkbookNoteCsv' : 'editor.chooseWorkbookNote')}</span>
        </button>
      ) : (
        <p className="wbe-muted">{t('editor.noWorkbook')}</p>
      )}
    </div>
  );
}

/** A chart sheet's chart, filling the sheet; or, for one the editor doesn't draw, what it is. */
function ChartSheet({ store, view }: { store: WorkbookStore; view: SheetView }) {
  const t = useT();
  const chart = view.info.objects.find((o) => o.kind === 'chart');
  const type = chart?.chartType ? t.text(chart.chartType) : undefined;
  const title = chart?.title;
  const what = type
    ? title
      ? t('editor.typedChartTitled', { type, title })
      : t('editor.typedChart', { type })
    : title
      ? t('editor.thisChartTitled', { title })
      : t('editor.thisChart');
  const placeholder = (
    <div className="wbe-center-state">
      <Icon name="chart" size={32} />
      <h2>{view.info.name}</h2>
      <p>{t('editor.chartSheet', { chart: what })}</p>
    </div>
  );
  return chart ? <ChartSlot store={store} obj={chart} placeholder={placeholder} className="wbe-chart-sheet" /> : placeholder;
}

/** The workbook editor: toolbar, formula bar, grid, sheet tabs and panels. */
export function WorkbookEditor({ editor, emptyState, toolbarEnd, className }: WorkbookEditorProps) {
  const store = storeOf(editor);
  useSyncExternalStore(store.subscribe, store.getVersion);
  const root = useRef<HTMLDivElement>(null);

  useEffect(ensurePortal, []);

  // Phone-sized layout follows the editor's own width, which in a host's layout can be less than the window's.
  useEffect(() => {
    const el = root.current;
    if (!el || typeof ResizeObserver === 'undefined') return;
    const watch = new ResizeObserver(([entry]) => store.setNarrow(entry.contentRect.width <= 720));
    watch.observe(el);
    return () => watch.disconnect();
  }, [store]);

  // Glide opens cell editors in one portal for the whole page; it takes the
  // theme and language of the editor being used, however the host set them.
  const themeVersion = store.themeVersion;
  const language = store.t.language;
  useEffect(() => {
    const el = root.current;
    const portal = document.getElementById(PORTAL_ID);
    if (!el || !portal) return;
    const sync = () => {
      copyTheme(el, portal);
      portal.lang = el.lang;
      portal.dir = el.dir;
    };
    sync();
    el.addEventListener('focusin', sync);
    el.addEventListener('pointerdown', sync, true);
    return () => {
      el.removeEventListener('focusin', sync);
      el.removeEventListener('pointerdown', sync, true);
    };
  }, [themeVersion, language]);

  useEffect(() => {
    // Shortcuts only apply while focus is in this editor (or its popups).
    const inEditor = (target: EventTarget | null) =>
      target instanceof Node && (root.current?.contains(target) || document.getElementById(PORTAL_ID)?.contains(target));
    const onKey = (e: KeyboardEvent) => {
      const command = e.metaKey || e.ctrlKey;
      if (!inEditor(e.target) || !root.current) return;
      // F6 is Excel's key for this; browsers that keep F6 for their address bar leave Ctrl+F6.
      if (e.key === 'F6' && !e.altKey && !e.metaKey && (!store.designerOpen || store.designerPick) && !store.confirmRequest) {
        e.preventDefault();
        nextRegion(root.current, store, e.shiftKey);
        return;
      }
      if (!command) return;
      const key = shortcutKey(e);
      if (key === '/') {
        e.preventDefault();
        store.set('shortcutsOpen', true);
        return;
      }
      if (store.status !== 'ready') return;
      // Browser shortcuts are taken only for the same job (save, find, undo…), never
      // for another: ⌘R still reloads, ⌘D still bookmarks, ⌘G still finds the next match.
      if (key === 's' && store.features.save) {
        e.preventDefault();
        void store.save();
      } else if (key === 'f' && !store.designerOpen) {
        e.preventDefault();
        store.set('findOpen', true);
      } else if (!isTextField(e.target) && !store.confirmRequest && !store.designerOpen) {
        if (key === 'z' && !e.shiftKey) {
          e.preventDefault();
          void store.undo();
        } else if ((key === 'z' && e.shiftKey) || (key === 'y' && !APPLE)) {
          e.preventDefault();
          void store.redo();
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

  return (
    <div
      className={`wbe wbe-root ${className ?? ''}`}
      ref={root}
      style={themeStyle(store.theme) as CSSProperties}
      lang={store.t.language}
      dir={store.t.dir}
    >
      <I18nextProvider i18n={store.i18n} defaultNS={NAMESPACE}>
        <IconContext.Provider value={store.icons}>
          <EditorBoundary store={store}>
            <EditorApp store={store} emptyState={emptyState} toolbarEnd={toolbarEnd} />
          </EditorBoundary>
        </IconContext.Provider>
      </I18nextProvider>
    </div>
  );
}

/** Everything inside the editor's root, where its language and icons are provided. */
function EditorApp({ store, emptyState, toolbarEnd }: Omit<WorkbookEditorProps, 'editor'> & { store: WorkbookStore }) {
  const t = useT();
  const hintId = useId();
  const fileInput = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const canOpen = store.features.open;
  const openPicker = () => {
    if (canOpen) fileInput.current?.click();
  };

  const openFile = async (file: File | undefined) => {
    if (file) await store.open(file.name, await file.arrayBuffer(), { csv: store.features.csv });
  };

  const view = store.view;
  const selection = store.selection;
  const ready = store.status === 'ready' && view;

  return (
    <div
      className={`wbe-app ${store.panel ? 'wbe-with-report' : ''}`}
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
          accept={`.xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet${store.features.csv ? ',.csv,.tsv,.txt,text/csv,text/tab-separated-values' : ''}`}
          hidden
          onChange={(e) => {
            void openFile(e.target.files?.[0]);
            e.target.value = '';
          }}
        />
      )}
      {ready && <SchemaBanner store={store} />}
      {ready && <Banner store={store} />}
      {ready && store.features.format && <FormatBar store={store} />}
      {ready && view.info.kind === 'worksheet' && <FormulaBar store={store} view={view} selection={selection} />}
      <div className="wbe-workspace">
        {/* A labelled region, not <main>: the host page has its own. */}
        <section
          className="wbe-sheet-area"
          aria-label={ready ? t('editor.sheetLabel', { name: view.info.name }) : undefined}
          aria-describedby={ready ? `${hintId}-hint` : undefined}
        >
          {ready && (
            <p id={`${hintId}-hint`} className="wbe-sr-only">
              {t('editor.sheetHint', { key: shortcutsKey(t) })}
            </p>
          )}
          {store.status === 'empty' && (emptyState ? emptyState(openPicker) : <DefaultEmptyState openPicker={openPicker} canOpen={canOpen} csv={store.features.csv} />)}
          {store.status === 'loading' && (
            <div className="wbe-center-state" role="status">
              <span className="wbe-spinner" aria-hidden="true" />
              <p>{t.text(store.stage)}</p>
            </div>
          )}
          {store.status === 'error' && (
            <div className="wbe-center-state wbe-error" role="alert">
              <Icon name="warn" size={28} />
              <p>{t.text(store.error)}</p>
              {store.canReopen && (
                <button className="wbe-btn wbe-primary" onClick={() => void store.reopen()}>
                  {t('editor.reopen', { fileName: store.fileName })}
                </button>
              )}
              {canOpen && (
                <button className="wbe-btn" onClick={openPicker}>
                  {t('editor.openAnother')}
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
          {ready && view.info.kind !== 'worksheet' && <ChartSheet store={store} view={view} />}
          {ready && store.findOpen && <FindPanel store={store} />}
        </section>
        {ready && <SidePanel store={store} />}
      </div>
      {ready && <SheetTabs store={store} />}
      <StatusBar store={store} view={ready ? view : undefined} selection={ready ? selection : undefined} />
      {dragging && (
        <div className="wbe-drop-overlay">
          <Icon name="open" size={28} />
          <p>{t('editor.dropToOpen')}</p>
        </div>
      )}
      {store.designerOpen && <SchemaDesigner store={store} />}
      {ready && <GridMenu store={store} />}
      {ready && <FilterPicker store={store} />}
      {ready && <SizeDialog store={store} axis="cols" />}
      {ready && <SizeDialog store={store} axis="rows" />}
      {ready && <CommentPopup store={store} />}
      {ready && <NameManager store={store} />}
      {ready && <NameDialog store={store} />}
      <ShortcutsDialog store={store} />
      <ConfirmDialog store={store} />
      <Toasts store={store} />
    </div>
  );
}
