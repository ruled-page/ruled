import type { ReactNode } from 'react';
import { useMemo } from 'react';
import { VT, type Finding } from '../core/types';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { Trans } from 'react-i18next';
import { Icon } from './icons';
import { mod, shortcutsKey } from './keys';
import { useT } from '../i18n/react';

function SchemaChip({ store }: { store: WorkbookStore }) {
  const t = useT();
  const ctl = store.schema;
  if (!ctl) {
    if (!store.schemaMismatch) return null;
    return (
      <button className="chip chip-warn" onClick={() => store.set('designerOpen', true)} title={t('schemaChip.noneTitle')}>
        <Icon name="warn" size={14} /> <span className="label">{t('schemaChip.none')}</span>
      </button>
    );
  }
  const { errors, warnings, pending } = ctl.summary();
  const tone = errors ? 'chip-error' : warnings ? 'chip-warn' : 'chip-ok';
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
      className={`chip ${tone}`}
      onClick={() => store.set('panel', store.panel === 'issues' ? null : 'issues')}
      title={t('schemaChip.title', { name })}
      aria-label={t('schemaChip.label', { name, status })}
      aria-expanded={store.panel === 'issues'}
    >
      <Icon name={errors || warnings ? 'warn' : 'check'} size={14} />
      <span className="label">{name}</span>
      <span className="chip-status">
        {pending > 0 && <span className="spinner small" />} {status}
      </span>
    </button>
  );
}

export function Toolbar({ store, onOpen, extra }: { store: WorkbookStore; onOpen: () => void; extra?: ReactNode }) {
  const t = useT();
  const ready = store.status === 'ready';
  const notable = store.findings.filter((f) => f.severity === 'warn').length;
  const target = ready ? store.recordTarget : undefined;
  // Tables that can grow keep their row buttons in place, so selecting a cell never moves the toolbar.
  const growable = ready ? (store.schema?.binding.sheets.flatMap((s) => s.tables).filter((table) => !table.fixed) ?? []) : [];
  const across = (target?.table ?? growable[0])?.byRow === false;
  const outside = across ? t('toolbar.selectColumnFirst') : t('toolbar.selectRowFirst');
  const exact = store.numbers.display === 'full';
  return (
    <header className="toolbar">
      <div className="brand">
        <span className="brand-mark" aria-hidden="true">
          <Icon name="sheet" size={16} />
        </span>
        <span className="brand-name">{t('toolbar.brand')}</span>
        {store.fileName && <span className="file-name">{store.fileName}</span>}
        {store.editedCount > 0 && (
          <button className="badge badge-edit" onClick={() => store.set('panel', store.panel === 'changes' ? null : 'changes')} title={t('toolbar.changesTitle')}>
            {t('toolbar.changes', { count: store.editedCount })}
          </button>
        )}
        {ready && <SchemaChip store={store} />}
        {store.busy && <span className="spinner" role="status" aria-label={t('toolbar.working')} />}
      </div>
      <div className="toolbar-actions">
        {store.features.open && (
          <button className="btn" onClick={onOpen} title={t('toolbar.openTitle')} aria-label={t('toolbar.open')}>
            <Icon name="open" /> <span className="label">{t('toolbar.open')}</span>
          </button>
        )}
        <div className="btn-group">
          <button className="btn icon" disabled={!store.undoStack.length} onClick={() => void store.undo()} title={t('toolbar.undoTitle', { key: mod(t, 'Z') })} aria-label={t('toolbar.undo')}>
            <Icon name="undo" />
          </button>
          <button className="btn icon" disabled={!store.redoStack.length} onClick={() => void store.redo()} title={t('toolbar.redoTitle', { key: mod(t, 'Y') })} aria-label={t('toolbar.redo')}>
            <Icon name="redo" />
          </button>
        </div>
        {growable.length > 0 && (
          <div className="btn-group">
            <button
              className="btn"
              disabled={!target}
              onClick={() => void store.insertRecord()}
              title={target ? t(across ? 'toolbar.insertColumnTitle' : 'toolbar.insertRowTitle') : outside}
              aria-label={t(across ? 'toolbar.insertColumn' : 'toolbar.insertRow')}
            >
              <Icon name="plus" /> <span className="label">{t(across ? 'toolbar.insertColumn' : 'toolbar.insertRow')}</span>
            </button>
            <button
              className="btn"
              disabled={!target || target.append}
              onClick={() => void store.deleteRecords()}
              title={target ? t(across ? 'toolbar.deleteColumnTitle' : 'toolbar.deleteRowTitle') : outside}
              aria-label={t(across ? 'toolbar.deleteColumn' : 'toolbar.deleteRow')}
            >
              <Icon name="minus" /> <span className="label">{t(across ? 'toolbar.deleteColumn' : 'toolbar.deleteRow')}</span>
            </button>
          </div>
        )}
        <label className={`toggle ${!ready ? 'disabled' : ''}`} title={t('toolbar.exactTitle')}>
          <input type="checkbox" checked={exact} disabled={!ready} onChange={(e) => store.set('exactNumbers', e.target.checked)} />
          <span>{t('toolbar.exact')}</span>
        </label>
        <label className={`toggle ${!ready ? 'disabled' : ''}`} title={t('toolbar.highlightTitle')}>
          <input type="checkbox" checked={store.highlightEdits} disabled={!ready} onChange={(e) => store.set('highlightEdits', e.target.checked)} />
          <span>{t('toolbar.highlight')}</span>
        </label>
        {store.features.schemas && (
          <button className={`btn ${store.designerOpen ? 'active' : ''}`} onClick={() => store.set('designerOpen', true)} title={t('toolbar.schemasTitle')} aria-label={t('toolbar.schemas')}>
            <Icon name="schema" /> <span className="label">{t('toolbar.schemas')}</span>
          </button>
        )}
        {store.features.report && (
          <button
            className={`btn ${store.panel === 'report' ? 'active' : ''}`}
            disabled={!ready}
            onClick={() => store.set('panel', store.panel === 'report' ? null : 'report')}
            title={t('toolbar.report')}
            aria-label={notable ? t('toolbar.reportNotable', { count: notable }) : t('toolbar.report')}
          >
            <Icon name="info" /> <span className="label">{t('toolbar.report')}</span>
            {notable > 0 && <span className="count">{t.number(notable)}</span>}
          </button>
        )}
        <button className="btn icon" onClick={() => store.set('shortcutsOpen', true)} title={t('toolbar.shortcutsTitle', { key: shortcutsKey(t) })} aria-label={t('toolbar.shortcuts')}>
          <Icon name="keyboard" />
        </button>
        {extra}
        {store.features.save && store.hasSaveHandler && store.features.download && (
          <button className="btn icon" disabled={!ready || store.busy} onClick={() => void store.exportCopy()} title={t('toolbar.downloadTitle')} aria-label={t('toolbar.download')}>
            <Icon name="download" />
          </button>
        )}
        {store.features.save && (
          <button className="btn primary" disabled={!ready || store.busy} onClick={() => void store.download()} title={t('toolbar.saveTitle', { key: mod(t, 'S') })}>
            <Icon name={store.hasSaveHandler ? 'check' : 'download'} /> {store.saveLabel}
          </button>
        )}
      </div>
    </header>
  );
}

export function SchemaBanner({ store }: { store: WorkbookStore }) {
  const t = useT();
  const m = store.schemaMismatch;
  if (store.schemaBannerDismissed) return null;
  const dismiss = (
    <button className="btn icon ghost" onClick={() => store.set('schemaBannerDismissed', true)} aria-label={t('banner.dismiss')}>
      <Icon name="close" />
    </button>
  );
  if (!m && store.schemaNote) {
    return (
      <div className="banner" role="status">
        <Icon name="schema" />
        <div className="banner-text">
          <p>{t.text(store.schemaNote)}</p>
        </div>
        {dismiss}
      </div>
    );
  }
  if (!m) return null;
  return (
    <div className="banner banner-warn" role="status">
      <Icon name="warn" />
      <div className="banner-text">
        <p>
          {m.name ? (
            <Trans i18nKey="banner.mismatchNamed" values={{ name: t.localized(m.name) }} components={{ b: <strong /> }} />
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
        <button className="link" onClick={() => store.set('designerOpen', true)}>
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
    <div className="banner" role="status">
      <Icon name="info" />
      <p>
        <Trans i18nKey="banner.workbookHas" values={{ list: t.list(summary) }} components={{ b: <strong /> }} />
      </p>
      <button className="link" onClick={() => store.set('panel', 'report')}>
        {t('banner.reviewDetails')}
      </button>
      <button className="btn icon ghost" onClick={() => store.set('bannerDismissed', true)} aria-label={t('banner.dismiss')}>
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
  return (
    <nav className="sheet-tabs" aria-label={t('tabs.label')}>
      {store.sheets.map((s, i) =>
        s.info.state !== 'visible' ? null : (
          <button
            key={i}
            className={`tab ${i === store.active ? 'active' : ''}`}
            aria-current={i === store.active ? 'true' : undefined}
            style={s.info.tabColor ? { boxShadow: `inset 0 -3px 0 ${s.info.tabColor}` } : undefined}
            onClick={() => store.setActive(i)}
            title={bound.has(i) ? t('tabs.checkedBy', { name: schemaName }) : undefined}
          >
            {s.info.kind === 'chartsheet' && <Icon name="chart" size={13} />}
            {s.info.protected && <Icon name="lock" size={12} />}
            {bound.has(i) && <Icon name="schema" size={12} />}
            {s.info.name}
          </button>
        ),
      )}
      {hidden > 0 && (
        <span className="tabs-note" title={t('tabs.hiddenTitle')}>
          {t('tabs.hidden', { count: hidden })}
        </span>
      )}
    </nav>
  );
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
    <footer className="status-bar">
      <span className="status-calc" role="status">
        {store.status === 'ready' && !calc && (
          <>
            <span className="spinner small" /> {t('status.building')}
          </>
        )}
        {calc?.off && (
          <>
            <span className="dot" /> {t('status.recalcOff', { count: calc.formulaCells })}
            {store.recalcStarting ? (
              <span className="muted">
                <span className="spinner small" /> {t('status.building')}
              </span>
            ) : (
              <button className="link" onClick={() => void store.enableRecalc()} title={t('status.turnOnTitle')}>
                {t('status.turnOn')}
              </button>
            )}
          </>
        )}
        {calc && !calc.off && (
          <>
            <span className="dot ok" /> {t('status.live', { count: calc.liveCells })}
            {calc.unsupported + calc.mismatched > 0 &&
              (store.features.report ? (
                <button className="link" onClick={() => store.set('panel', 'report')}>
                  {t('status.notCalculated', { count: calc.unsupported + calc.mismatched })}
                </button>
              ) : (
                <span className="muted">{t('status.notCalculated', { count: calc.unsupported + calc.mismatched })}</span>
              ))}
          </>
        )}
      </span>
      {view?.filter && (
        <span className="status-filter">
          {t('status.filtered', { count: view.filteredCount })}
          <button className="link" onClick={() => store.clearFilter()}>
            {t('status.showAll')}
          </button>
        </span>
      )}
      {stats && (
        <span className="status-stats">
          {stats.nums > 0 && <span>{t('status.average', { value: fmt(stats.sum / stats.nums) })}</span>}
          <span>{t('status.count', { value: t.number(stats.count) })}</span>
          {stats.nums > 0 && <span>{t('status.sum', { value: fmt(stats.sum) })}</span>}
        </span>
      )}
    </footer>
  );
}
