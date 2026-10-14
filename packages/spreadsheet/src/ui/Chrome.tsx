import type { DragEvent, KeyboardEvent, ReactNode, RefObject } from 'react';
import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { quoteSheetName, toA1 } from '@ruled/formula/refs';
import { VT, type Finding } from '../core/types';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { BrandMark, Icon } from './icons';
import { RichText } from './RichText';
import { RenameSheetDialog, SheetMenu, sheetNamed, sheetSteps, UnhideSheetDialog, type AfterSheetChange } from './SheetDialogs';
import { mod, shortcutsKey } from './keys';
import { useT } from '../i18n/react';

function SchemaChip({ store }: { store: WorkbookStore }) {
  const t = useT();
  const ctl = store.schema;
  if (!ctl) {
    if (!store.schemaMismatch) return null;
    return (
      <button className="wbe-chip wbe-chip-warn" onClick={() => store.set('designerOpen', true)} title={t('schemaChip.noneTitle')}>
        <Icon name="warn" size={14} /> <span className="wbe-label">{t('schemaChip.none')}</span>
      </button>
    );
  }
  const { errors, warnings, pending } = ctl.summary();
  const tone = errors ? 'wbe-chip-error' : warnings ? 'wbe-chip-warn' : 'wbe-chip-ok';
  const status = errors
    ? t('schemaChip.errors', { count: errors })
    : warnings
      ? t('schemaChip.warnings', { count: warnings })
      : pending
        ? t('schemaChip.checking')
        : t('schemaChip.allPass');
  const name = t.localized(ctl.schema.name);
  return (
    <button
      className={`wbe-chip ${tone}`}
      onClick={() => store.set('panel', store.panel === 'issues' ? null : 'issues')}
      title={t('schemaChip.title', { name })}
      aria-label={t('schemaChip.label', { name, status })}
      aria-expanded={store.panel === 'issues'}
    >
      <Icon name={errors || warnings ? 'warn' : 'check'} size={14} />
      <span className="wbe-label">{name}</span>
      <span className="wbe-chip-status">
        {pending > 0 && <span className="wbe-spinner wbe-small" />} {status}
      </span>
    </button>
  );
}

/**
 * The toolbar's labels fold into icons in three steps as it runs out of room
 * (see editor.css); past the last, its actions take a row of their own.
 */
const FOLDS = 4;

/**
 * Folds the toolbar to the first step at which it fits: what its labels need
 * depends on the language, the theme's font and what the host adds, so it
 * measures itself rather than folding at fixed widths.
 */
function fold(el: HTMLElement): void {
  const room = el.clientWidth;
  let step = 0;
  for (; step < FOLDS; step++) {
    el.dataset.fold = String(step);
    el.style.width = 'max-content';
    const need = el.getBoundingClientRect().width;
    el.style.width = '';
    if (need <= room + 0.5) break;
  }
  el.dataset.fold = String(step);
}

/** The toolbar's element, folded again whenever its width, text or theme changes. */
function useFolding(): RefObject<HTMLDivElement | null> {
  const bar = useRef<HTMLDivElement>(null);
  const seen = useRef('');
  useLayoutEffect(() => {
    const el = bar.current;
    if (!el) return;
    // Cheap to read, unlike sizes: what is in the toolbar, and the theme on the editor's root.
    const now = `${el.textContent}|${el.getElementsByTagName('*').length}|${el.closest('.wbe')?.getAttribute('style') ?? ''}`;
    if (now === seen.current) return;
    seen.current = now;
    fold(el);
  });
  useEffect(() => {
    const el = bar.current;
    if (!el || typeof ResizeObserver === 'undefined') return;
    let frame = 0;
    let width = el.clientWidth;
    const refold = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => fold(el));
    };
    const observer = new ResizeObserver(() => {
      if (el.clientWidth === width) return;
      width = el.clientWidth;
      refold();
    });
    observer.observe(el);
    // Labels measured before the theme's font loaded were measured in another: fold again once it has.
    document.fonts?.addEventListener('loadingdone', refold);
    return () => {
      observer.disconnect();
      document.fonts?.removeEventListener('loadingdone', refold);
      cancelAnimationFrame(frame);
    };
  }, []);
  return bar;
}

export function Toolbar({ store, onOpen, extra }: { store: WorkbookStore; onOpen: () => void; extra?: ReactNode }) {
  const t = useT();
  const bar = useFolding();
  const ready = store.status === 'ready';
  const notable = store.findings.filter((f) => f.severity === 'warn').length;
  const target = ready ? store.recordTarget : undefined;
  // Tables that can grow keep their row buttons in place, so selecting a cell never moves the toolbar.
  const growable = ready ? (store.schema?.binding.sheets.flatMap((s) => s.tables).filter((table) => !table.fixed) ?? []) : [];
  const across = (target?.table ?? growable[0])?.byRow === false;
  const outside = across ? t('toolbar.selectColumnFirst') : t('toolbar.selectRowFirst');
  const exact = store.numbers.display === 'full';
  // On a phone the less used controls are behind "More", in a panel under the toolbar.
  const [more, setMore] = useState(false);
  const moreId = useId();
  const moreButton = useRef<HTMLButtonElement>(null);
  const morePanel = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!more) return;
    morePanel.current?.querySelector<HTMLElement>('button:not(:disabled), input:not(:disabled), select, a[href]')?.focus();
    const away = (e: PointerEvent) => {
      const target = e.target as Node;
      if (!morePanel.current?.contains(target) && !moreButton.current?.contains(target)) setMore(false);
    };
    // Safari doesn't focus a checkbox that is tapped, so the focus can fall from the panel to the page; Escape still closes it.
    const escape = (e: globalThis.KeyboardEvent) => {
      if (e.key !== 'Escape' || document.activeElement !== document.body) return;
      setMore(false);
      moreButton.current?.focus();
    };
    document.addEventListener('pointerdown', away, true);
    document.addEventListener('keydown', escape);
    return () => {
      document.removeEventListener('pointerdown', away, true);
      document.removeEventListener('keydown', escape);
    };
  }, [more]);
  return (
    // Not a <header>: a page that embeds the editor has its own banner landmark.
    <div className="wbe-toolbar" ref={bar}>
      <div className="wbe-brand">
        <span className="wbe-brand-mark" aria-hidden="true">
          <BrandMark />
        </span>
        <span className="wbe-brand-name">{t('toolbar.brand')}</span>
        {store.fileName && <span className="wbe-file-name">{store.fileName}</span>}
        {store.editedCount > 0 && (
          <button
            className="wbe-badge wbe-badge-edit"
            onClick={() => store.set('panel', store.panel === 'changes' ? null : 'changes')}
            title={t('toolbar.changesTitle')}
            aria-expanded={store.panel === 'changes'}
          >
            {t('toolbar.changes', { count: store.editedCount })}
          </button>
        )}
        {ready && <SchemaChip store={store} />}
        {store.busy && <span className="wbe-spinner" role="status" aria-label={t('toolbar.working')} />}
      </div>
      <div className="wbe-toolbar-actions">
        {store.features.open && (
          <button className="wbe-btn" onClick={onOpen} title={t('toolbar.openTitle')} aria-label={t('toolbar.open')}>
            <Icon name="open" /> <span className="wbe-label">{t('toolbar.open')}</span>
          </button>
        )}
        <div className="wbe-btn-group">
          {/* Marked unavailable rather than disabled, so undoing the last step doesn't drop the focus. */}
          <button className="wbe-btn wbe-icon" aria-disabled={!store.undoStack.length} onClick={() => void store.undo()} title={t('toolbar.undoTitle', { key: mod(t, 'Z') })} aria-label={t('toolbar.undo')}>
            <Icon name="undo" />
          </button>
          <button className="wbe-btn wbe-icon" aria-disabled={!store.redoStack.length} onClick={() => void store.redo()} title={t('toolbar.redoTitle', { key: mod(t, 'Y') })} aria-label={t('toolbar.redo')}>
            <Icon name="redo" />
          </button>
        </div>
        {growable.length > 0 && (
          <div className="wbe-btn-group">
            <button
              className="wbe-btn"
              disabled={!target}
              onClick={() => void store.insertRecord()}
              title={target ? t(across ? 'toolbar.insertColumnTitle' : 'toolbar.insertRowTitle') : outside}
              aria-label={t(across ? 'toolbar.insertColumn' : 'toolbar.insertRow')}
            >
              <Icon name="plus" /> <span className="wbe-label">{t(across ? 'toolbar.insertColumn' : 'toolbar.insertRow')}</span>
            </button>
            <button
              className="wbe-btn"
              disabled={!target || target.append}
              onClick={() => void store.deleteRecords()}
              title={target ? t(across ? 'toolbar.deleteColumnTitle' : 'toolbar.deleteRowTitle') : outside}
              aria-label={t(across ? 'toolbar.deleteColumn' : 'toolbar.deleteRow')}
            >
              <Icon name="minus" /> <span className="wbe-label">{t(across ? 'toolbar.deleteColumn' : 'toolbar.deleteRow')}</span>
            </button>
          </div>
        )}
        <div
          className="wbe-toolbar-more"
          id={moreId}
          ref={morePanel}
          data-open={more || undefined}
          onKeyDown={(e) => {
            if (e.key !== 'Escape' || !more) return;
            e.stopPropagation();
            setMore(false);
            moreButton.current?.focus();
          }}
          // A button done with closes the panel; toggles and pickers keep it open.
          onClick={(e) => (e.target as Element).closest('button') && setMore(false)}
        >
        <label className={`wbe-toggle ${!ready ? 'wbe-disabled' : ''}`} title={t('toolbar.exactTitle')}>
          <input type="checkbox" checked={exact} disabled={!ready} onChange={(e) => store.set('exactNumbers', e.target.checked)} />
          <span className="wbe-toggle-long">{t('toolbar.exact')}</span>
          <span className="wbe-toggle-short" aria-hidden="true">
            {t('toolbar.exactShort')}
          </span>
        </label>
        <label className={`wbe-toggle ${!ready ? 'wbe-disabled' : ''}`} title={t('toolbar.highlightTitle')}>
          <input type="checkbox" checked={store.highlightEdits} disabled={!ready} onChange={(e) => store.set('highlightEdits', e.target.checked)} />
          <span className="wbe-toggle-long">{t('toolbar.highlight')}</span>
          <span className="wbe-toggle-short" aria-hidden="true">
            {t('toolbar.highlightShort')}
          </span>
        </label>
        {store.schema && (
          <label className={`wbe-toggle ${!ready ? 'wbe-disabled' : ''}`} title={t('toolbar.labelsTitle')}>
            <input type="checkbox" checked={store.showLabels} disabled={!ready} onChange={(e) => store.set('showLabels', e.target.checked)} />
            <span className="wbe-toggle-long">{t('toolbar.labels')}</span>
            <span className="wbe-toggle-short" aria-hidden="true">
              {t('toolbar.labelsShort')}
            </span>
          </label>
        )}
        {ready && store.features.format && (
          <label className="wbe-toggle wbe-phone-only">
            <input type="checkbox" checked={store.formatBarShown} onChange={(e) => store.set('formatBarShown', e.target.checked)} />
            <span className="wbe-toggle-long">{t('toolbar.formatBar')}</span>
          </label>
        )}
        {store.features.schemas && (
          <button className={`wbe-btn ${store.designerOpen ? 'wbe-active' : ''}`} onClick={() => store.set('designerOpen', true)} title={t('toolbar.schemasTitle')} aria-label={t('toolbar.schemas')}>
            <Icon name="schema" /> <span className="wbe-label">{t('toolbar.schemas')}</span>
          </button>
        )}
        {store.features.report && (
          <button
            className={`wbe-btn ${store.panel === 'report' ? 'wbe-active' : ''}`}
            disabled={!ready}
            aria-expanded={store.panel === 'report'}
            onClick={() => store.set('panel', store.panel === 'report' ? null : 'report')}
            title={t('toolbar.report')}
            aria-label={notable ? t('toolbar.reportNotable', { count: notable }) : t('toolbar.report')}
          >
            <Icon name="info" /> <span className="wbe-label">{t('toolbar.report')}</span>
            {notable > 0 && <span className="wbe-count">{t.number(notable)}</span>}
          </button>
        )}
        <button className="wbe-btn wbe-icon" onClick={() => store.set('shortcutsOpen', true)} title={t('toolbar.shortcutsTitle', { key: shortcutsKey(t) })} aria-label={t('toolbar.shortcuts')}>
          <Icon name="keyboard" /> <span className="wbe-more-label">{t('toolbar.shortcuts')}</span>
        </button>
        {extra}
        {store.features.csv && store.features.download && (
          <button
            className="wbe-btn"
            disabled={!ready || store.view?.info.kind !== 'worksheet'}
            aria-disabled={store.busy}
            onClick={() => !store.busy && void store.exportCsv()}
            title={t('toolbar.csvTitle')}
            aria-label={t('toolbar.csvLabel')}
          >
            <Icon name="download" /> {t('toolbar.csv')}
          </button>
        )}
        {store.features.save && store.hasSaveHandler && store.features.download && (
          <button
            className="wbe-btn wbe-icon"
            disabled={!ready}
            aria-disabled={store.busy}
            onClick={() => !store.busy && void store.exportCopy()}
            title={t('toolbar.downloadTitle')}
            aria-label={t('toolbar.download')}
          >
            <Icon name="download" /> <span className="wbe-more-label">{t('toolbar.download')}</span>
          </button>
        )}
        </div>
        <button
          ref={moreButton}
          className="wbe-btn wbe-icon wbe-toolbar-more-btn"
          aria-expanded={more}
          aria-controls={moreId}
          aria-label={t('toolbar.more')}
          title={t('toolbar.more')}
          onClick={() => setMore(!more)}
        >
          <Icon name="more" />
        </button>
        {store.features.save && (
          // Busy for a moment while saving, without taking the focus away from the button.
          <button className="wbe-btn wbe-primary" disabled={!ready} aria-disabled={store.busy} onClick={() => !store.busy && void store.save()} title={t('toolbar.saveTitle', { key: mod(t, 'S') })}>
            <Icon name={store.hasSaveHandler ? 'check' : 'download'} />{' '}
            {store.hasSaveHandler ? (
              store.saveLabel
            ) : (
              <>
                <span className="wbe-save-long">{store.saveLabel}</span>
                <span className="wbe-save-short" aria-hidden="true">
                  {t('toolbar.save')}
                </span>
              </>
            )}
          </button>
        )}
      </div>
    </div>
  );
}

export function SchemaBanner({ store }: { store: WorkbookStore }) {
  const t = useT();
  const m = store.schemaMismatch;
  if (store.schemaBannerDismissed) return null;
  const dismiss = (
    <button
      className="wbe-btn wbe-icon wbe-ghost"
      onClick={() => {
        store.set('schemaBannerDismissed', true);
        store.focusGrid();
      }}
      aria-label={t('banner.dismiss')}
    >
      <Icon name="close" />
    </button>
  );
  if (!m && store.schemaNote) {
    return (
      <div className="wbe-banner" role="status">
        <Icon name="schema" />
        <div className="wbe-banner-text">
          <p>{t.text(store.schemaNote)}</p>
        </div>
        {dismiss}
      </div>
    );
  }
  if (!m) return null;
  return (
    <div className="wbe-banner wbe-banner-warn" role="status">
      <Icon name="warn" />
      <div className="wbe-banner-text">
        <p>
          {m.name ? (
            <RichText i18nKey="banner.mismatchNamed" values={{ name: t.localized(m.name) }} components={{ b: <strong /> }} />
          ) : (
            t('banner.mismatch', { count: m.tried })
          )}
        </p>
        {m.problems.length > 0 && (
          <ul>
            {m.problems.slice(0, 3).map((p, i) => (
              <li key={i}>{t.text(p)}</li>
            ))}
            {m.problems.length > 3 && <li>{t('banner.moreProblems', { count: m.problems.length - 3 })}</li>}
          </ul>
        )}
      </div>
      {store.features.schemas && (
        <button className="wbe-link" onClick={() => store.set('designerOpen', true)}>
          {t('banner.openDesigner')}
        </button>
      )}
      {dismiss}
    </div>
  );
}

export function Banner({ store }: { store: WorkbookStore }) {
  const t = useT();
  const summary = useMemo(() => {
    const count = (kinds: Finding['kind'][]) => store.findings.filter((f) => kinds.includes(f.kind)).length;
    const parts: string[] = [];
    const charts = count(['chart', 'chartsheet']);
    const pivots = count(['pivot']);
    if (charts) parts.push(t('banner.charts', { count: charts }));
    if (pivots) parts.push(t('banner.pivots', { count: pivots }));
    if (count(['macro'])) parts.push(t('banner.macros'));
    if (count(['activeContent'])) parts.push(t('banner.activeContent'));
    if (count(['externalLink'])) parts.push(t('banner.externalLinks'));
    const notCalculated = (store.calc?.unsupported ?? 0) + (store.calc?.mismatched ?? 0);
    if (notCalculated) parts.push(t('banner.notCalculated', { count: notCalculated }));
    return parts;
  }, [store.findings, store.calc, t]);

  if (store.bannerDismissed || !summary.length || !store.features.report) return null;
  return (
    <div className="wbe-banner" role="status">
      <Icon name="info" />
      <p>
        <RichText i18nKey="banner.workbookHas" values={{ list: t.list(summary) }} components={{ b: <strong /> }} />
      </p>
      <button className="wbe-link" onClick={() => store.set('panel', 'report')}>
        {t('banner.reviewDetails')}
      </button>
      <button
        className="wbe-btn wbe-icon wbe-ghost"
        onClick={() => {
          store.set('bannerDismissed', true);
          store.focusGrid();
        }}
        aria-label={t('banner.dismiss')}
      >
        <Icon name="close" />
      </button>
    </div>
  );
}

/**
 * Says when a workbook opened with live recalculation off, being too big to
 * recalculate as people type, and offers to turn it on, which the status
 * bar also offers.
 */
export function RecalcBanner({ store }: { store: WorkbookStore }) {
  const t = useT();
  const calc = store.calc;
  if (!calc?.off || store.recalcBannerDismissed) return null;
  return (
    <div className="wbe-banner wbe-banner-warn" role="status">
      <Icon name="warn" />
      <p>{t('banner.recalcOff', { count: calc.formulaCells })}</p>
      {store.recalcStarting ? (
        <span className="wbe-muted">
          <span className="wbe-spinner wbe-small" /> {t('status.building')}
        </span>
      ) : (
        <button className="wbe-link" onClick={() => void store.enableRecalc()} title={t('status.turnOnTitle')}>
          {t('banner.turnOn')}
        </button>
      )}
      <button
        className="wbe-btn wbe-icon wbe-ghost"
        onClick={() => {
          store.set('recalcBannerDismissed', true);
          store.focusGrid();
        }}
        aria-label={t('banner.dismiss')}
      >
        <Icon name="close" />
      </button>
    </div>
  );
}

export function SheetTabs({ store }: { store: WorkbookStore }) {
  const t = useT();
  const hidden = store.sheets.filter((s) => s.info.state !== 'visible').length;
  const bound = new Set(store.schema?.binding.sheets.map((s) => s.sheet) ?? []);
  const schemaName = t.localized(store.schema?.schema.name);
  const tabs = useRef(new Map<number, HTMLButtonElement>());
  const { visible, left, right } = sheetSteps(store);
  // The sheet shown is the one tab stop, so many sheets aren't a long run of Tab presses.
  const stop = visible.includes(store.active) ? store.active : visible[0];

  // After a change to the sheets, the focus goes to the tab it concerns, which may now be elsewhere.
  const [focusTo, setFocusTo] = useState<number | null>(null);
  useEffect(() => {
    if (focusTo === null) return;
    tabs.current.get(focusTo)?.focus();
    setFocusTo(null);
  }, [focusTo]);
  const after: AfterSheetChange = (change, focus) => void change.then((made) => setFocusTo(focus(made)));

  const openMenu = (sheet: number, x: number, y: number) => store.set('sheetMenu', { sheet, x, y });
  // The arrow keys, Home and End move between sheets; Enter or Space shows the one with the focus.
  // Alt+Shift+arrows move the sheet, F2 renames it, and Shift+F10 or the menu key opens its menu.
  const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    const at = visible.findIndex((i) => tabs.current.get(i) === document.activeElement);
    if (at < 0) return;
    const sheet = visible[at];
    const name = store.sheets[sheet].info.name;
    const arrow = e.key === 'ArrowRight' ? right : e.key === 'ArrowLeft' ? left : 0;
    if (arrow && e.altKey && e.shiftKey && !e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      after(store.moveSheetBy(sheet, arrow), sheetNamed(store, name));
      return;
    }
    if (e.key === 'F2') {
      e.preventDefault();
      store.set('renameSheetFor', sheet);
      return;
    }
    if ((e.key === 'F10' && e.shiftKey) || e.key === 'ContextMenu') {
      e.preventDefault();
      const r = tabs.current.get(sheet)!.getBoundingClientRect();
      openMenu(sheet, r.left, r.bottom);
      return;
    }
    if (e.altKey || e.ctrlKey || e.metaKey) return;
    const to = e.key === 'Home' ? 0 : e.key === 'End' ? visible.length - 1 : arrow ? (at + arrow + visible.length) % visible.length : -1;
    if (to < 0) return;
    e.preventDefault();
    tabs.current.get(visible[to])?.focus();
  };

  // Dragging a tab moves its sheet: the drop goes before the sheet at `before`, in workbook order.
  const [drag, setDrag] = useState<{ from: number; before?: number } | null>(null);
  const dragOver = (e: DragEvent<HTMLButtonElement>, i: number) => {
    if (!drag) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const r = e.currentTarget.getBoundingClientRect();
    const later = (e.clientX > r.left + r.width / 2) === (right > 0);
    const before = later ? i + 1 : i;
    if (before !== drag.before) setDrag({ from: drag.from, before });
  };
  const drop = (e: DragEvent<HTMLButtonElement>) => {
    e.preventDefault();
    // Not a file to open, which dropping on the editor otherwise is.
    e.stopPropagation();
    if (drag?.before !== undefined) after(store.dropSheet(drag.from, drag.before), sheetNamed(store, store.sheets[drag.from].info.name));
    setDrag(null);
  };
  // Where the drop would go: at the start of the tab it goes before or, past the last tab or before a hidden sheet, the end of the tab before.
  const marker = (i: number) => {
    if (drag?.before === undefined || drag.before === drag.from || drag.before === drag.from + 1) return '';
    if (drag.before === i) return 'wbe-drop-start';
    const next = visible.find((v) => v > i);
    return drag.before > i && (next === undefined || drag.before < next) ? 'wbe-drop-end' : '';
  };

  return (
    <>
      <nav className="wbe-sheet-tabs" aria-label={t('tabs.label')}>
        {/* Named by the navigation it is in. */}
        <div className="wbe-sheet-tablist" role="tablist" onKeyDown={onKeyDown}>
          {visible.map((i) => {
            const s = store.sheets[i];
            return (
              <button
                key={i}
                ref={(el) => {
                  if (el) tabs.current.set(i, el);
                  else tabs.current.delete(i);
                }}
                role="tab"
                aria-selected={i === store.active}
                tabIndex={i === stop ? 0 : -1}
                className={`wbe-tab ${i === store.active ? 'wbe-active' : ''} ${marker(i)}`}
                style={s.info.tabColor ? { boxShadow: `inset 0 -3px 0 ${s.info.tabColor}` } : undefined}
                onClick={() => store.setActive(i)}
                onDoubleClick={() => store.set('renameSheetFor', i)}
                onContextMenu={(e) => {
                  e.preventDefault();
                  openMenu(i, e.clientX, e.clientY);
                }}
                draggable
                onDragStart={(e) => {
                  e.dataTransfer.effectAllowed = 'move';
                  e.dataTransfer.setData('text/plain', s.info.name);
                  setDrag({ from: i });
                }}
                onDragOver={(e) => dragOver(e, i)}
                onDrop={drop}
                onDragEnd={() => setDrag(null)}
                title={bound.has(i) ? t('tabs.checkedBy', { name: schemaName }) : undefined}
              >
                {s.info.kind === 'chartsheet' && <Icon name="chart" size={13} />}
                {s.info.protected && <Icon name="lock" size={12} />}
                {bound.has(i) && <Icon name="schema" size={12} />}
                {s.info.name}
                {/* What the icons show, for screen readers. */}
                {s.info.kind === 'chartsheet' && <span className="wbe-sr-only">{`, ${t('tabs.chartSheet')}`}</span>}
                {s.info.protected && <span className="wbe-sr-only">{`, ${t('tabs.protected')}`}</span>}
                {bound.has(i) && <span className="wbe-sr-only">{`, ${t('tabs.checkedBy', { name: schemaName })}`}</span>}
              </button>
            );
          })}
        </div>
        <button className="wbe-btn wbe-icon wbe-ghost wbe-tabs-btn" onClick={() => after(store.addSheet(), () => store.active)} title={t('sheets.addTitle')} aria-label={t('sheets.add')}>
          <Icon name="plus" size={14} />
        </button>
        {/* The shown sheet's tab menu, for touch screens (and anyone who looks for it here). */}
        <button
          className="wbe-btn wbe-icon wbe-ghost wbe-tabs-btn"
          aria-label={t('sheets.menuButton')}
          title={t('sheets.menuButton')}
          aria-haspopup="menu"
          onClick={(e) => {
            const r = e.currentTarget.getBoundingClientRect();
            openMenu(store.active, r.left, r.top);
          }}
        >
          <Icon name="more" size={14} />
        </button>
        {hidden > 0 &&
          (store.hiddenSheets.length ? (
            <button className="wbe-link wbe-tabs-note" aria-haspopup="dialog" onClick={() => store.set('unhideOpen', true)} title={t('sheets.unhideButtonTitle')}>
              {t('tabs.hidden', { count: hidden })}
            </button>
          ) : (
            <span className="wbe-tabs-note" title={t('tabs.hiddenTitle')}>
              {t('tabs.hidden', { count: hidden })}
            </span>
          ))}
      </nav>
      <SheetMenu store={store} after={after} />
      <RenameSheetDialog store={store} after={after} />
      <UnhideSheetDialog store={store} after={after} />
    </>
  );
}

/** Which engine is calculating, what it cost to download, and how long opening, calculating and the last edit took. */
/** Names the first cell where a circular reference closes, as Excel's status bar does; choosing it goes there. */
function CircularReadout({ store }: { store: WorkbookStore }) {
  const t = useT();
  const first = store.circular[0];
  const name = store.sheets[first.sheet]?.info.name;
  const ref = toA1(first.row, first.col);
  const where = first.sheet === store.active || !name ? ref : `${quoteSheetName(name)}!${ref}`;
  return (
    <button className="wbe-link wbe-status-circular" onClick={() => store.goTo(name ? `${quoteSheetName(name)}!${ref}` : ref)}>
      {t('status.circular', { cell: where })}
    </button>
  );
}

function EngineReadout({ store }: { store: WorkbookStore }) {
  const t = useT();
  const info = store.engineInfo;
  if (!info) return null;
  const parts = [info.name];
  parts.push(info.transferBytes || info.decodedBytes ? t('status.engineSize', { kB: Math.round((info.transferBytes || info.decodedBytes) / 1024) }) : t('status.engineCached'));
  if (store.openMs !== undefined) parts.push(t('status.openMs', { ms: store.openMs }));
  if (store.calcStats) parts.push(t('status.calcMs', { ms: store.calcStats.ms }));
  if (store.editStats) parts.push(t('status.editMs', { ms: store.editStats.ms }));
  const heap = store.editStats?.heapBytes ?? store.calcStats?.heapBytes;
  if (heap !== undefined) parts.push(t('status.heap', { mb: Math.round(heap / 1048576) }));
  return <span className="wbe-status-engine">{parts.join(' · ')}</span>;
}

export function StatusBar({ store, view, selection }: { store: WorkbookStore; view?: SheetView; selection?: Selection }) {
  const t = useT();
  const stats = useMemo(() => {
    if (!view || !selection) return null;
    const total = selection.rects.reduce((n, r) => n + (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1), 0);
    if (total < 2 || total > 3_000_000) return null;
    let count = 0;
    let nums = 0;
    let sum = 0;
    for (const r of selection.rects) {
      view.forEach(r, (_row, _col, rec) => {
        if (rec.t === VT.None) return;
        count++;
        if (rec.t === VT.Number) {
          nums++;
          sum += rec.n ?? 0;
        }
      });
    }
    return { count, nums, sum };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [view, view?.version, selection]);

  const calc = store.calc;
  const fmt = (n: number) => t.number(n, { maximumFractionDigits: store.numbers.display === 'full' ? 15 : 4 });
  return (
    // Not a <footer>, which would be a landmark of the page the editor is in.
    <div className="wbe-status-bar">
      <span className="wbe-status-calc" role="status">
        {store.status === 'ready' && !calc && (
          <>
            <span className="wbe-spinner wbe-small" /> {t('status.building')}
          </>
        )}
        {calc?.off && (
          <>
            <span className="wbe-dot" /> {t('status.recalcOff', { count: calc.formulaCells })}
            {store.recalcStarting ? (
              <span className="wbe-muted">
                <span className="wbe-spinner wbe-small" /> {t('status.building')}
              </span>
            ) : (
              <button className="wbe-link" onClick={() => void store.enableRecalc()} title={t('status.turnOnTitle')}>
                {t('status.turnOn')}
              </button>
            )}
          </>
        )}
        {calc && !calc.off && (
          <>
            <span className="wbe-dot wbe-ok" /> {t('status.live', { count: calc.liveCells })}
            {calc.unsupported + calc.mismatched > 0 &&
              (store.features.report ? (
                <button className="wbe-link" onClick={() => store.set('panel', 'report')}>
                  {t('status.notCalculated', { count: calc.unsupported + calc.mismatched })}
                </button>
              ) : (
                <span className="wbe-muted">{t('status.notCalculated', { count: calc.unsupported + calc.mismatched })}</span>
              ))}
          </>
        )}
      </span>
      {store.circular.length > 0 && <CircularReadout store={store} />}
      {view?.filtering && (
        <span className="wbe-status-filter">
          {t('status.filtered', { count: view.filteredCount })}
          <button className="wbe-link" onClick={() => void store.clearFilter()}>
            {t('status.showAll')}
          </button>
        </span>
      )}
      {store.diagnostics && <EngineReadout store={store} />}
      {stats && (
        <span className="wbe-status-stats">
          {stats.nums > 0 && <span>{t('status.average', { value: fmt(stats.sum / stats.nums) })}</span>}
          <span>{t('status.count', { value: t.number(stats.count) })}</span>
          {stats.nums > 0 && <span>{t('status.sum', { value: fmt(stats.sum) })}</span>}
        </span>
      )}
    </div>
  );
}
