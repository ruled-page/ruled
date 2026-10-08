import { useEffect } from 'react';
import { nameProblem } from '../core/names';
import { parseA1, quoteSheetName, toA1 } from '../core/refs';
import type { CellChange } from '../core/types';
import type { Issue } from '../schema/controller';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { RichText } from './RichText';
import { ReportBody } from './ReportPanel';
import { Tabs } from './Tabs';
import type { Translator } from '../i18n';
import { useT } from '../i18n/react';

function IssuesBody({ store }: { store: WorkbookStore }) {
  const t = useT();
  const ctl = store.schema;
  if (!ctl) {
    return (
      <div className="wbe-report-body">
        <p className="wbe-muted">{t('checks.noSchema')}</p>
        <button className="wbe-btn" onClick={() => store.set('designerOpen', true)}>
          <Icon name="schema" /> {t('checks.openDesigner')}
        </button>
      </div>
    );
  }
  const issues = ctl.issues();
  const { errors, warnings, pending } = ctl.summary();
  const go = (i: Issue) => {
    if (i.sheet < 0 || i.row === undefined || i.col === undefined) return;
    store.reveal(i.sheet, i.row, i.col);
    // On a phone the panel covers the sheet: get out of the way of the cell.
    if (store.narrow) store.set('panel', null);
  };
  return (
    <div className="wbe-report-body">
      <p className="wbe-report-lede wbe-schema-lede">
        <RichText i18nKey="checks.checkedAgainst" values={{ name: t.localized(ctl.schema.name) }} components={{ b: <strong /> }} />
        {ctl.schema.description ? <span className="wbe-muted"> {t.localized(ctl.schema.description)}</span> : null}
      </p>
      <div className="wbe-calc-summary">
        <div>
          <strong className={errors ? 'wbe-tone-error-text' : ''}>{t.number(errors)}</strong>
          <span>{t('checks.errors', { count: errors })}</span>
        </div>
        <div>
          <strong>{t.number(warnings)}</strong>
          <span>{t('checks.warnings', { count: warnings })}</span>
        </div>
        <div>
          <strong>{pending ? <span className="wbe-spinner wbe-small" /> : t.number(0)}</strong>
          <span>{t('checks.running')}</span>
        </div>
      </div>
      {issues.length === 0 && !pending && (
        <p className="wbe-all-clear">
          <Icon name="check" /> {t('checks.allClear')}
        </p>
      )}
      <ul className="wbe-findings">
        {issues.map((i, n) => (
          <li key={n} className={`wbe-finding wbe-sev-${i.severity === 'error' ? 'error' : 'warn'}`}>
            <button className="wbe-finding-main" onClick={() => go(i)} disabled={i.row === undefined}>
              {/* Said as well as coloured. */}
              <span className="wbe-finding-title">{t(i.severity === 'error' ? 'issue.error' : 'issue.warning', { message: t.text(i.message) })}</span>
              {i.sheetName && (
                <span className="wbe-finding-where">
                  {i.sheetName}
                  {i.row !== undefined && i.col !== undefined ? ` · ${toA1(i.row, i.col)}` : ''}
                </span>
              )}
            </button>
          </li>
        ))}
      </ul>
      <section>
        <h3>{t('checks.markers')}</h3>
        <ul className="wbe-legend">
          <li>
            <span className="wbe-swatch wbe-tri-right wbe-red" /> {t('checks.markerError')}
          </li>
          <li>
            <span className="wbe-swatch wbe-tri-right wbe-orange" /> {t('checks.markerWarning')}
          </li>
          <li>
            <span className="wbe-swatch wbe-dot-blue" /> {t('checks.markerPending')}
          </li>
          <li>
            <span className="wbe-swatch wbe-label-cell" /> {t('checks.markerLabel')}
          </li>
        </ul>
      </section>
    </div>
  );
}

function ChangesBody({ store }: { store: WorkbookStore }) {
  const t = useT();
  useEffect(() => {
    void store.refreshChanges();
  }, [store]);
  const changes = store.changeList;
  const go = (sheetName: string, ref: string) => {
    const sheet = store.sheets.findIndex((s) => s.info.name === sheetName);
    const pos = parseA1(ref);
    // A change to a whole sheet shows the sheet, if it can be shown.
    if (sheet >= 0 && ref === '' && store.sheets[sheet].info.state === 'visible') {
      store.setActive(sheet);
      return;
    }
    // A change to a defined name: its sheet, if any, is the one it is local to.
    if (!pos && ref !== '' && !nameProblem(ref)) {
      store.goTo(sheetName ? `${quoteSheetName(sheetName)}!${ref}` : ref);
      return;
    }
    if (sheet < 0 || !pos) return;
    store.reveal(sheet, pos.row, pos.col);
    if (store.narrow) store.set('panel', null);
  };
  const time = (at: number) => t.date(at, { hour: '2-digit', minute: '2-digit', second: '2-digit' });
  const empty = <span className="wbe-muted">{t('changes.empty')}</span>;
  return (
    <div className="wbe-report-body">
      <p className="wbe-report-lede wbe-schema-lede">{changes.length ? summarize(t, changes) : t('changes.none')}</p>
      {changes.length > 0 && (
        <div className="wbe-changes-table">
          <table>
            <thead>
              <tr>
                <th>{t('changes.cell')}</th>
                <th>{t('changes.was')}</th>
                <th>{t('changes.now')}</th>
              </tr>
            </thead>
            <tbody>
              {changes.slice(0, SHOWN).map((c, i) => (
                // Keyed by position: a cell's note can change as well as its value, and a sheet can change more than once.
                <tr key={`${i}:${c.sheet}!${c.ref}${c.comment ? '#note' : ''}`} onClick={() => go(c.sheet, c.ref)}>
                  <td className="wbe-mono">
                    {/* The row is clickable too; the button is the same for keyboards and screen readers. */}
                    <button className="wbe-cell-link" title={t('changes.goTo')} onClick={(e) => {
                      e.stopPropagation();
                      go(c.sheet, c.ref);
                    }}>
                      {c.ref === '' ? c.sheet : store.sheets.length > 1 && c.sheet ? `${c.sheet}!${c.ref}` : c.ref}
                    </button>
                  </td>
                  {c.note ? (
                    <td colSpan={2} className="wbe-change-note">
                      {t.text(c.note)}
                    </td>
                  ) : c.comment ? (
                    <>
                      <td className="wbe-change-lines">{c.before === '' ? empty : c.before}</td>
                      <td className="wbe-change-lines">
                        <span className="wbe-change-note">{t.text(c.comment)}</span>
                        {c.after !== '' && `\n${c.after}`}
                      </td>
                    </>
                  ) : (
                    <>
                      <td>{c.before === '' ? empty : c.before}</td>
                      <td>{c.formatOnly ? <span className="wbe-muted">{t('changes.formatOnly')}</span> : c.after === '' ? empty : c.after}</td>
                    </>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
          {changes.length > SHOWN && <p className="wbe-muted wbe-small">{t('changes.first', { n: SHOWN })}</p>}
        </div>
      )}
      <section>
        <h3>{t('changes.history')}</h3>
        {store.history.length === 0 && <p className="wbe-muted">{t('changes.noEdits')}</p>}
        <ol className="wbe-history">
          {[...store.history].reverse().slice(0, 200).map((h, i) => (
            <li key={i}>
              <span className="wbe-history-time">{time(h.at)}</span>
              <span>
                <strong>{t.text(h.action)}</strong> · {t('changes.historyCells', { count: h.cells })}
                {h.user ? ` · ${h.user}` : ''}
                {h.refs.length > 0 && <span className="wbe-muted"> ({h.refs.join(', ')}{h.cells > h.refs.length ? ', …' : ''})</span>}
              </span>
            </li>
          ))}
        </ol>
      </section>
    </div>
  );
}

/** How many changed cells the Changes panel lists. */
const SHOWN = 500;

/** "3 cells differ from the file as it was opened, and 1 change to rows, columns or where cells are." */
function summarize(t: Translator, changes: CellChange[]): string {
  const structural = changes.filter((c) => c.note).length;
  // A cell whose value and note both changed is one cell.
  const cells = new Set(changes.filter((c) => !c.note).map((c) => `${c.sheet}!${c.ref}`)).size;
  if (cells && structural) {
    return t('changes.summaryBoth', { cells: t('changes.cells', { count: cells }), structural: t('changes.structural', { count: structural }) });
  }
  return cells ? t('changes.summaryCells', { count: cells }) : t('changes.summaryStructural', { count: structural });
}

export function SidePanel({ store }: { store: WorkbookStore }) {
  const t = useT();
  const panel = store.panel;
  if (!panel || (panel === 'report' && !store.features.report)) return null;
  return (
    <aside className="wbe-report" aria-label={t('panel.label')}>
      <Tabs
        tabs={[
          { id: 'issues', label: t('panel.checks') },
          { id: 'changes', label: t('panel.changes') },
          ...(store.features.report ? [{ id: 'report' as const, label: t('panel.report') }] : []),
        ]}
        value={panel}
        onChange={(id) => store.set('panel', id)}
        head={(tabs) => (
          <div className="wbe-report-head">
            {tabs}
            <button
              className="wbe-btn wbe-icon wbe-ghost"
              onClick={() => {
                store.set('panel', null);
                store.focusGrid();
              }}
              aria-label={t('panel.close')}
            >
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
