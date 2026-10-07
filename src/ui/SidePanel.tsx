import { useEffect } from 'react';
import { parseA1, toA1 } from '../core/refs';
import type { CellChange } from '../core/types';
import type { Issue } from '../schema/controller';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { ReportBody } from './ReportPanel';
import { Tabs } from './Tabs';

function IssuesBody({ store }: { store: WorkbookStore }) {
  const ctl = store.schema;
  if (!ctl) {
    return (
      <div className="report-body">
        <p className="muted">No schema applies to this workbook, so there is nothing to check. Open the schema designer to create one from it.</p>
        <button className="btn" onClick={() => store.set('designerOpen', true)}>
          <Icon name="schema" /> Open schema designer
        </button>
      </div>
    );
  }
  const issues = ctl.issues();
  const { errors, warnings, pending } = ctl.summary();
  const go = (i: Issue) => {
    if (i.sheet < 0 || i.row === undefined || i.col === undefined) return;
    store.setActive(i.sheet);
    store.setSelection({ row: i.row, col: i.col, rects: [{ r1: i.row, c1: i.col, r2: i.row, c2: i.col }] });
  };
  return (
    <div className="report-body">
      <p className="report-lede schema-lede">
        Checked against the <strong>{ctl.schema.name}</strong> schema.
        {ctl.schema.description ? <span className="muted"> {ctl.schema.description}</span> : null}
      </p>
      <div className="calc-summary">
        <div>
          <strong className={errors ? 'tone-error-text' : ''}>{errors}</strong>
          <span>errors</span>
        </div>
        <div>
          <strong>{warnings}</strong>
          <span>warnings</span>
        </div>
        <div>
          <strong>{pending ? <span className="spinner small" /> : 0}</strong>
          <span>checks running</span>
        </div>
      </div>
      {issues.length === 0 && !pending && (
        <p className="all-clear">
          <Icon name="check" /> Everything matches the schema.
        </p>
      )}
      <ul className="findings">
        {issues.map((i, n) => (
          <li key={n} className={`finding sev-${i.severity === 'error' ? 'error' : 'warn'}`}>
            <button className="finding-main" onClick={() => go(i)} disabled={i.row === undefined}>
              <span className="finding-title">{i.message}</span>
              {i.sheetName && (
                <span className="finding-where">
                  {i.sheetName}
                  {i.row !== undefined && i.col !== undefined ? ` · ${toA1(i.row, i.col)}` : ''}
                </span>
              )}
            </button>
          </li>
        ))}
      </ul>
      <section>
        <h3>Markers</h3>
        <ul className="legend">
          <li>
            <span className="swatch tri-right red" /> Breaks a schema rule.
          </li>
          <li>
            <span className="swatch tri-right orange" /> Warning from a schema rule.
          </li>
          <li>
            <span className="swatch dot-blue" /> A check is still running.
          </li>
          <li>
            <span className="swatch label-cell" /> Labels fixed by the schema.
          </li>
        </ul>
      </section>
    </div>
  );
}

function ChangesBody({ store }: { store: WorkbookStore }) {
  useEffect(() => {
    void store.refreshChanges();
  }, [store]);
  const changes = store.changeList;
  const go = (sheetName: string, ref: string) => {
    const sheet = store.sheets.findIndex((s) => s.info.name === sheetName);
    const pos = parseA1(ref);
    if (sheet < 0 || !pos) return;
    store.reveal(sheet, pos.row, pos.col);
  };
  const time = (t: number) => new Date(t).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  return (
    <div className="report-body">
      <p className="report-lede schema-lede">
        {changes.length ? summarize(changes) : 'Nothing has changed since the file was opened.'}
      </p>
      {changes.length > 0 && (
        <div className="changes-table">
          <table>
            <thead>
              <tr>
                <th>Cell</th>
                <th>Was</th>
                <th>Now</th>
              </tr>
            </thead>
            <tbody>
              {changes.slice(0, 500).map((c) => (
                <tr key={`${c.sheet}!${c.ref}`} onClick={() => go(c.sheet, c.ref)}>
                  <td className="mono">
                    {/* The row is clickable too; the button is the same for keyboards and screen readers. */}
                    <button className="cell-link" title="Go to this cell" onClick={(e) => {
                      e.stopPropagation();
                      go(c.sheet, c.ref);
                    }}>
                      {store.sheets.length > 1 ? `${c.sheet}!` : ''}
                      {c.ref}
                    </button>
                  </td>
                  {c.note ? (
                    <td colSpan={2} className="change-note">
                      {c.note}
                    </td>
                  ) : (
                    <>
                      <td>{c.before === '' ? <span className="muted">(empty)</span> : c.before}</td>
                      <td>{c.formatOnly ? <span className="muted">format only</span> : c.after === '' ? <span className="muted">(empty)</span> : c.after}</td>
                    </>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
          {changes.length > 500 && <p className="muted small">Showing the first 500.</p>}
        </div>
      )}
      <section>
        <h3>History</h3>
        {store.history.length === 0 && <p className="muted">No edits yet.</p>}
        <ol className="history">
          {[...store.history].reverse().slice(0, 200).map((h, i) => (
            <li key={i}>
              <span className="history-time">{time(h.at)}</span>
              <span>
                <strong>{h.action}</strong> · {h.cells.toLocaleString()} cell{h.cells === 1 ? '' : 's'}
                {h.user ? ` · ${h.user}` : ''}
                {h.refs.length > 0 && <span className="muted"> ({h.refs.join(', ')}{h.cells > h.refs.length ? ', …' : ''})</span>}
              </span>
            </li>
          ))}
        </ol>
      </section>
    </div>
  );
}

/** "3 cells differ from the file as it was opened, and 1 change to rows or columns." */
function summarize(changes: CellChange[]): string {
  const structural = changes.filter((c) => c.note).length;
  const cells = changes.length - structural;
  const parts = [];
  if (cells) parts.push(`${cells.toLocaleString()} cell${cells === 1 ? '' : 's'} differ from the file as it was opened`);
  if (structural) parts.push(`${structural.toLocaleString()} change${structural === 1 ? '' : 's'} to rows, columns or where cells are`);
  return `${parts.join(', and ')}.`;
}

export function SidePanel({ store }: { store: WorkbookStore }) {
  const panel = store.panel;
  if (!panel) return null;
  return (
    <aside className="report" aria-label="Workbook details">
      <Tabs
        label="Workbook details"
        tabs={[
          { id: 'issues', label: 'Checks' },
          { id: 'changes', label: 'Changes' },
          { id: 'report', label: "What's in this file" },
        ]}
        value={panel}
        onChange={(id) => store.set('panel', id)}
        head={(tabs) => (
          <div className="report-head">
            {tabs}
            <button className="btn icon ghost" onClick={() => store.set('panel', null)} aria-label="Close">
              <Icon name="close" />
            </button>
          </div>
        )}
      >
        {panel === 'issues' ? <IssuesBody store={store} /> : panel === 'changes' ? <ChangesBody store={store} /> : <ReportBody store={store} />}
      </Tabs>
    </aside>
  );
}
