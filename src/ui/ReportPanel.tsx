import { parseA1, parseRange } from '../core/refs';
import type { Finding, FindingKind } from '../core/types';
import type { WorkbookStore } from '../state/store';
import type { MessageKey } from '../i18n/message';
import { useT } from '../i18n/react';
import { isNarrow } from '../state/viewport';

type GroupId = 'kept' | 'recalc' | 'limits';

const GROUPS: { id: GroupId; title: MessageKey; blurb: MessageKey; kinds: FindingKind[] }[] = [
  {
    id: 'kept',
    title: 'report.keptTitle',
    blurb: 'report.keptBlurb',
    kinds: [
      'chart', 'chartsheet', 'pivot', 'image', 'shape', 'control', 'macro', 'externalLink', 'connection', 'slicer',
      'conditionalFormat', 'dataValidation', 'comment', 'sparkline', 'hyperlink', 'cellImage', 'hiddenSheet',
    ],
  },
  {
    id: 'recalc',
    title: 'report.recalcTitle',
    blurb: 'report.recalcBlurb',
    kinds: ['formulaEngine'],
  },
  {
    id: 'limits',
    title: 'report.limitsTitle',
    blurb: 'report.limitsBlurb',
    kinds: ['table', 'protection', 'arrayFormula', 'mergedCells', 'frozenRows', 'autoFilter', 'richText'],
  },
];

/** How many formulas the worker lists as samples at most. */
const MAX_SAMPLES = 50;

/** The "What's in this file" list: what the editor keeps but can't show, and its limits. */
export function ReportBody({ store }: { store: WorkbookStore }) {
  const t = useT();
  const go = (f: Finding) => {
    if (!f.sheet) return;
    const i = store.sheets.findIndex((s) => s.info.name === f.sheet);
    if (i < 0 || store.sheets[i].info.state !== 'visible') return;
    store.setActive(i);
    const rect = f.where ? parseRange(f.where.split(',')[0].trim()) : null;
    if (rect && store.sheets[i].info.kind === 'worksheet') {
      store.reveal(i, rect.r1, rect.c1);
      store.setSelection({ row: rect.r1, col: rect.c1, rects: [rect] });
    }
    if (isNarrow()) store.set('panel', null);
  };
  const calc = store.calc;
  const duration = (ms: number) =>
    ms < 1000 ? t('report.milliseconds', { n: t.number(ms) }) : t('report.seconds', { n: t.number(ms / 1000, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) });
  const openMs = (store.timings.unzipMs ?? 0) + (store.timings.parseMs ?? 0);

  return (
      <div className="report-body">
        <p className="report-lede">{t('report.lede')}</p>
        {GROUPS.map((g) => {
          const items = store.findings
            .filter((f) => g.kinds.includes(f.kind))
            .sort((a, b) => (a.severity === b.severity ? g.kinds.indexOf(a.kind) - g.kinds.indexOf(b.kind) : a.severity === 'warn' ? -1 : 1));
          const recalc = g.id === 'recalc';
          if (!items.length && !recalc) return null;
          return (
            <section key={g.id}>
              <h3>{t(g.title)}</h3>
              <p className="blurb">{t(g.blurb)}</p>
              {recalc && calc && (
                <div className="calc-summary">
                  <div>
                    <strong>{t.number(calc.liveCells)}</strong>
                    <span>{t('report.live', { count: calc.liveCells })}</span>
                  </div>
                  <div>
                    <strong>{t.number(calc.unsupported + calc.mismatched)}</strong>
                    <span>{t('report.notCalculated', { count: calc.unsupported + calc.mismatched })}</span>
                  </div>
                  <div>
                    <strong>{duration(calc.buildMs)}</strong>
                    <span>{t('report.checkTime')}</span>
                  </div>
                </div>
              )}
              {recalc && !calc && <p className="muted">{t('report.building')}</p>}
              <ul className="findings">
                {items.map((f, i) => (
                  <li key={i} className={`finding sev-${f.severity}`}>
                    <button className="finding-main" onClick={() => go(f)} disabled={!f.sheet}>
                      <span className="finding-title">{t.text(f.title)}</span>
                      {f.sheet && (
                        <span className="finding-where">
                          {f.sheet}
                          {f.where && parseA1(f.where.split(':')[0]) ? ` · ${f.where}` : ''}
                        </span>
                      )}
                    </button>
                    <p>{t.text(f.detail)}</p>
                  </li>
                ))}
              </ul>
              {recalc && calc && calc.samples.length > 0 && (
                <details className="samples">
                  <summary>
                    {calc.samples.length === MAX_SAMPLES
                      ? t('report.samplesFirst', { n: calc.samples.length })
                      : t('report.samples', { n: calc.samples.length })}
                  </summary>
                  <table>
                    <tbody>
                      {calc.samples.map((s, i) => (
                        <tr key={i}>
                          <td className="mono">
                            {s.sheet}!{s.ref}
                          </td>
                          <td className="mono formula">{s.formula}</td>
                          <td>
                            {t.text(s.reason)}
                            {s.excel !== undefined && (
                              <span className="muted">
                                {' '}
                                {t('report.sampleValues', { excel: t.text(s.excel), engine: t.text(s.engine) })}
                              </span>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </details>
              )}
            </section>
          );
        })}
        <section>
          <h3>{t('report.markers')}</h3>
          <ul className="legend">
            <li>
              <span className="swatch tri grey" /> {t('report.markerUnsupported')}
            </li>
            <li>
              <span className="swatch tri amber" /> {t('report.markerStale')}
            </li>
            <li>
              <span className="swatch edited" /> {t('report.markerEdited', { toggle: t('toolbar.highlight') })}
            </li>
          </ul>
        </section>
        <p className="timings muted">
          {calc ? t('report.openedAndChecked', { open: openMs, check: calc.buildMs }) : t('report.opened', { open: openMs })}
        </p>
      </div>
  );
}
