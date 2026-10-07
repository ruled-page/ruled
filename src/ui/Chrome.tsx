import { useMemo } from 'react';
import { VT, type Finding } from '../core/types';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { Icon } from './icons';

export function Toolbar({ store, onOpen }: { store: WorkbookStore; onOpen: () => void }) {
  const ready = store.status === 'ready';
  const notable = store.findings.filter((f) => f.severity === 'warn').length;
  return (
    <header className="toolbar">
      <div className="brand">
        <span className="brand-mark" aria-hidden="true">
          <Icon name="sheet" size={16} />
        </span>
        <span className="brand-name">Workbook editor</span>
        {store.fileName && <span className="file-name">{store.fileName}</span>}
        {store.editedCount > 0 && (
          <span className="badge badge-edit">
            {store.editedCount.toLocaleString()} {store.editedCount === 1 ? 'change' : 'changes'}
          </span>
        )}
        {store.busy && <span className="spinner" aria-label="Working" />}
      </div>
      <div className="toolbar-actions">
        <button className="btn" onClick={onOpen}>
          <Icon name="open" /> <span className="label">Open</span>
        </button>
        <div className="btn-group">
          <button className="btn icon" disabled={!store.undoStack.length} onClick={() => void store.undo()} title="Undo (⌘Z)" aria-label="Undo">
            <Icon name="undo" />
          </button>
          <button className="btn icon" disabled={!store.redoStack.length} onClick={() => void store.redo()} title="Redo (⇧⌘Z)" aria-label="Redo">
            <Icon name="redo" />
          </button>
        </div>
        <label className={`toggle ${!ready ? 'disabled' : ''}`}>
          <input type="checkbox" checked={store.highlightEdits} disabled={!ready} onChange={(e) => store.set('highlightEdits', e.target.checked)} />
          <span>Highlight changes</span>
        </label>
        <button className={`btn ${store.reportOpen ? 'active' : ''}`} disabled={!ready} onClick={() => store.set('reportOpen', !store.reportOpen)}>
          <Icon name="info" /> <span className="label">What's in this file</span>
          {notable > 0 && <span className="count">{notable}</span>}
        </button>
        <button className="btn primary" disabled={!ready || store.busy} onClick={() => void store.download()} title="Save as Excel (⌘S)">
          <Icon name="download" /> Save as Excel
        </button>
      </div>
    </header>
  );
}

export function Banner({ store }: { store: WorkbookStore }) {
  const summary = useMemo(() => {
    const count = (kinds: Finding['kind'][]) => store.findings.filter((f) => kinds.includes(f.kind)).length;
    const parts: string[] = [];
    const charts = count(['chart', 'chartsheet']);
    const pivots = count(['pivot']);
    if (charts) parts.push(`${charts} chart${charts === 1 ? '' : 's'}`);
    if (pivots) parts.push(`${pivots} pivot table${pivots === 1 ? '' : 's'}`);
    if (count(['macro'])) parts.push('macros');
    if (count(['externalLink'])) parts.push('links to other workbooks');
    const notCalculated = (store.calc?.unsupported ?? 0) + (store.calc?.mismatched ?? 0);
    if (notCalculated) parts.push(`${notCalculated.toLocaleString()} formula${notCalculated === 1 ? '' : 's'} this editor can't recalculate`);
    return parts;
  }, [store.findings, store.calc]);

  if (store.bannerDismissed || !summary.length) return null;
  const list = summary.length > 1 ? `${summary.slice(0, -1).join(', ')} and ${summary[summary.length - 1]}` : summary[0];
  return (
    <div className="banner" role="status">
      <Icon name="info" />
      <p>
        This workbook has {list}. <strong>Everything is kept when you save</strong>, even what this editor can't show.
      </p>
      <button className="link" onClick={() => store.set('reportOpen', true)}>
        Review details
      </button>
      <button className="btn icon ghost" onClick={() => store.set('bannerDismissed', true)} aria-label="Dismiss">
        <Icon name="close" />
      </button>
    </div>
  );
}

export function SheetTabs({ store }: { store: WorkbookStore }) {
  const hidden = store.sheets.filter((s) => s.info.state !== 'visible').length;
  return (
    <nav className="sheet-tabs" aria-label="Sheets">
      {store.sheets.map((s, i) =>
        s.info.state !== 'visible' ? null : (
          <button
            key={i}
            className={`tab ${i === store.active ? 'active' : ''}`}
            style={s.info.tabColor ? { boxShadow: `inset 0 -3px 0 ${s.info.tabColor}` } : undefined}
            onClick={() => store.setActive(i)}
          >
            {s.info.kind === 'chartsheet' && <Icon name="chart" size={13} />}
            {s.info.protected && <Icon name="lock" size={12} />}
            {s.info.name}
          </button>
        ),
      )}
      {hidden > 0 && (
        <span className="tabs-note" title="Hidden sheets are kept in the file but not shown, as in Excel.">
          {hidden} hidden
        </span>
      )}
    </nav>
  );
}

export function StatusBar({ store, view, selection }: { store: WorkbookStore; view?: SheetView; selection?: Selection }) {
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
  const fmt = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: 4 });
  return (
    <footer className="status-bar">
      <span className="status-calc">
        {store.status === 'ready' && !calc && (
          <>
            <span className="spinner small" /> Building calculation engine…
          </>
        )}
        {calc && (
          <>
            <span className="dot ok" /> {calc.liveCells.toLocaleString()} formulas recalculate live
            {calc.unsupported + calc.mismatched > 0 && (
              <button className="link" onClick={() => store.set('reportOpen', true)}>
                {(calc.unsupported + calc.mismatched).toLocaleString()} not calculated here
              </button>
            )}
          </>
        )}
      </span>
      {stats && (
        <span className="status-stats">
          {stats.nums > 0 && <span>Average: {fmt(stats.sum / stats.nums)}</span>}
          <span>Count: {stats.count.toLocaleString()}</span>
          {stats.nums > 0 && <span>Sum: {fmt(stats.sum)}</span>}
        </span>
      )}
    </footer>
  );
}
