import { useMemo } from 'react';
import { VT, type Finding } from '../core/types';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { Icon } from './icons';

function SchemaChip({ store }: { store: WorkbookStore }) {
  const ctl = store.schema;
  if (!ctl) {
    if (!store.schemaMismatch) return null;
    return (
      <button className="chip chip-warn" onClick={() => store.set('designerOpen', true)} title="No schema matched; open the schema designer">
        <Icon name="warn" size={14} /> <span className="label">No schema</span>
      </button>
    );
  }
  const { errors, warnings, pending } = ctl.summary();
  const tone = errors ? 'chip-error' : warnings ? 'chip-warn' : 'chip-ok';
  const status = errors
    ? `${errors} error${errors === 1 ? '' : 's'}`
    : warnings
      ? `${warnings} warning${warnings === 1 ? '' : 's'}`
      : pending
        ? 'Checking…'
        : 'All checks pass';
  return (
    <button className={`chip ${tone}`} onClick={() => store.set('panel', store.panel === 'issues' ? null : 'issues')} title={`Schema: ${ctl.schema.name}`}>
      <Icon name={errors || warnings ? 'warn' : 'check'} size={14} />
      <span className="label">{ctl.schema.name}</span>
      <span className="chip-status">
        {pending > 0 && <span className="spinner small" />} {status}
      </span>
    </button>
  );
}

export function Toolbar({ store, onOpen }: { store: WorkbookStore; onOpen: () => void }) {
  const ready = store.status === 'ready';
  const notable = store.findings.filter((f) => f.severity === 'warn').length;
  const target = ready ? store.recordTarget : undefined;
  const noun = target?.table.byRow === false ? 'column' : 'row';
  const exact = store.numbers.display === 'full';
  return (
    <header className="toolbar">
      <div className="brand">
        <span className="brand-mark" aria-hidden="true">
          <Icon name="sheet" size={16} />
        </span>
        <span className="brand-name">Workbook editor</span>
        {store.fileName && <span className="file-name">{store.fileName}</span>}
        {store.editedCount > 0 && (
          <button className="badge badge-edit" onClick={() => store.set('panel', store.panel === 'changes' ? null : 'changes')} title="Show what changed">
            {store.editedCount.toLocaleString()} {store.editedCount === 1 ? 'change' : 'changes'}
          </button>
        )}
        {ready && <SchemaChip store={store} />}
        {store.busy && <span className="spinner" aria-label="Working" />}
      </div>
      <div className="toolbar-actions">
        {store.features.open && (
          <button className="btn" onClick={onOpen} title="Open a workbook">
            <Icon name="open" /> <span className="label">Open</span>
          </button>
        )}
        <div className="btn-group">
          <button className="btn icon" disabled={!store.undoStack.length} onClick={() => void store.undo()} title="Undo (⌘Z)" aria-label="Undo">
            <Icon name="undo" />
          </button>
          <button className="btn icon" disabled={!store.redoStack.length} onClick={() => void store.redo()} title="Redo (⇧⌘Z)" aria-label="Redo">
            <Icon name="redo" />
          </button>
        </div>
        {target && (
          <div className="btn-group">
            <button className="btn" onClick={() => void store.insertRecord()} title={`Insert a ${noun} into this table`}>
              <Icon name="plus" /> <span className="label">Insert {noun}</span>
            </button>
            <button className="btn" disabled={target.append} onClick={() => void store.deleteRecords()} title={`Delete the selected ${noun}s from this table`}>
              <Icon name="minus" /> <span className="label">Delete {noun}</span>
            </button>
          </div>
        )}
        <label className={`toggle ${!ready ? 'disabled' : ''}`} title="Show every stored digit and copy exact values">
          <input type="checkbox" checked={exact} disabled={!ready} onChange={(e) => store.set('exactNumbers', e.target.checked)} />
          <span>Exact numbers</span>
        </label>
        <label className={`toggle ${!ready ? 'disabled' : ''}`} title="Tint cells changed in this session">
          <input type="checkbox" checked={store.highlightEdits} disabled={!ready} onChange={(e) => store.set('highlightEdits', e.target.checked)} />
          <span>Highlight changes</span>
        </label>
        {store.features.schemas && (
          <button className={`btn ${store.designerOpen ? 'active' : ''}`} onClick={() => store.set('designerOpen', true)} title="Create, import and export schemas">
            <Icon name="schema" /> <span className="label">Schemas</span>
          </button>
        )}
        <button className={`btn ${store.panel === 'report' ? 'active' : ''}`} disabled={!ready} onClick={() => store.set('panel', store.panel === 'report' ? null : 'report')} title="What's in this file">
          <Icon name="info" /> <span className="label">What's in this file</span>
          {notable > 0 && <span className="count">{notable}</span>}
        </button>
        {store.features.save && (
          <button className="btn primary" disabled={!ready || store.busy} onClick={() => void store.download()} title="Save (⌘S)">
            <Icon name="download" /> Save{store.hasSaveHandler ? '' : ' as Excel'}
          </button>
        )}
      </div>
    </header>
  );
}

export function SchemaBanner({ store }: { store: WorkbookStore }) {
  const m = store.schemaMismatch;
  if (!m && store.schemaNote) {
    return (
      <div className="banner" role="status">
        <Icon name="schema" />
        <div className="banner-text">
          <p>{store.schemaNote}</p>
        </div>
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
            <>
              This file doesn't match the <strong>{m.name}</strong> schema, so it's open as a plain spreadsheet with no schema checks.
            </>
          ) : (
            <>
              This file doesn't match {m.tried === 1 ? 'the registered schema' : `any of the ${m.tried} registered schemas`}, so it's open as a
              plain spreadsheet with no schema checks.
            </>
          )}
        </p>
        {m.problems.length > 0 && (
          <ul>
            {m.problems.slice(0, 3).map((p) => (
              <li key={p}>{p}</li>
            ))}
            {m.problems.length > 3 && <li>…and {m.problems.length - 3} more.</li>}
          </ul>
        )}
      </div>
      <button className="link" onClick={() => store.set('designerOpen', true)}>
        Open schema designer
      </button>
    </div>
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
      <button className="link" onClick={() => store.set('panel', 'report')}>
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
  const bound = new Set(store.schema?.binding.sheets.map((s) => s.sheet) ?? []);
  return (
    <nav className="sheet-tabs" aria-label="Sheets">
      {store.sheets.map((s, i) =>
        s.info.state !== 'visible' ? null : (
          <button
            key={i}
            className={`tab ${i === store.active ? 'active' : ''}`}
            style={s.info.tabColor ? { boxShadow: `inset 0 -3px 0 ${s.info.tabColor}` } : undefined}
            onClick={() => store.setActive(i)}
            title={bound.has(i) ? `Checked by the “${store.schema?.schema.name}” schema` : undefined}
          >
            {s.info.kind === 'chartsheet' && <Icon name="chart" size={13} />}
            {s.info.protected && <Icon name="lock" size={12} />}
            {bound.has(i) && <Icon name="schema" size={12} />}
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
  const fmt = (n: number) => n.toLocaleString(undefined, { maximumFractionDigits: store.numbers.display === 'full' ? 15 : 4 });
  return (
    <footer className="status-bar">
      <span className="status-calc">
        {store.status === 'ready' && !calc && (
          <>
            <span className="spinner small" /> Building calculation engine…
          </>
        )}
        {calc?.off && (
          <>
            <span className="dot" /> Live recalculation off ({calc.formulaCells.toLocaleString()} formulas)
            {store.recalcStarting ? (
              <span className="muted">
                <span className="spinner small" /> Building calculation engine…
              </span>
            ) : (
              <button className="link" onClick={() => void store.enableRecalc()} title="Building the engine for a workbook this size takes a while and uses a lot of memory">
                Turn on
              </button>
            )}
          </>
        )}
        {calc && !calc.off && (
          <>
            <span className="dot ok" /> {calc.liveCells.toLocaleString()} formulas recalculate live
            {calc.unsupported + calc.mismatched > 0 && (
              <button className="link" onClick={() => store.set('panel', 'report')}>
                {(calc.unsupported + calc.mismatched).toLocaleString()} not calculated here
              </button>
            )}
          </>
        )}
      </span>
      {view?.filter && (
        <span className="status-filter">
          Filtered: {view.filteredCount.toLocaleString()} row{view.filteredCount === 1 ? '' : 's'} hidden
          <button className="link" onClick={() => store.clearFilter()}>
            Show all
          </button>
        </span>
      )}
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
