import { toA1 } from '../core/refs';
import type { Issue } from '../schema/controller';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { ReportBody } from './ReportPanel';

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

export function SidePanel({ store }: { store: WorkbookStore }) {
  const panel = store.panel;
  if (!panel) return null;
  return (
    <aside className="report" aria-label={panel === 'issues' ? 'Schema issues' : "What's in this file"}>
      <div className="report-head">
        <div className="panel-tabs" role="tablist">
          <button role="tab" aria-selected={panel === 'issues'} className={panel === 'issues' ? 'active' : ''} onClick={() => store.set('panel', 'issues')}>
            Schema checks
          </button>
          <button role="tab" aria-selected={panel === 'report'} className={panel === 'report' ? 'active' : ''} onClick={() => store.set('panel', 'report')}>
            What's in this file
          </button>
        </div>
        <button className="btn icon ghost" onClick={() => store.set('panel', null)} aria-label="Close">
          <Icon name="close" />
        </button>
      </div>
      {panel === 'issues' ? <IssuesBody store={store} /> : <ReportBody store={store} />}
    </aside>
  );
}
