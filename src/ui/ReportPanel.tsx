import { parseA1, parseRange } from '../core/refs';
import type { Finding, FindingKind } from '../core/types';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';

const GROUPS: { title: string; blurb: string; kinds: FindingKind[] }[] = [
  {
    title: 'Kept in the file, not shown here',
    blurb: 'Saved exactly as they were. Excel redraws charts and refreshes pivot tables from your edited data when the file is opened.',
    kinds: [
      'chart', 'chartsheet', 'pivot', 'image', 'shape', 'control', 'macro', 'externalLink', 'connection', 'slicer',
      'conditionalFormat', 'dataValidation', 'comment', 'sparkline', 'hyperlink', 'cellImage', 'hiddenSheet',
    ],
  },
  {
    title: 'Recalculation',
    blurb: 'Formulas recalculate as you edit. The few this editor cannot calculate keep Excel’s value and are recalculated by Excel on open.',
    kinds: ['formulaEngine'],
  },
  {
    title: 'Editing limits',
    blurb: 'Cells this editor will not change, so that saving cannot damage the workbook.',
    kinds: ['table', 'protection', 'arrayFormula', 'mergedCells', 'frozenRows', 'autoFilter', 'richText'],
  },
];

export function ReportPanel({ store }: { store: WorkbookStore }) {
  const go = (f: Finding) => {
    if (!f.sheet) return;
    const i = store.sheets.findIndex((s) => s.info.name === f.sheet);
    if (i < 0 || store.sheets[i].info.state !== 'visible') return;
    store.setActive(i);
    const rect = f.where ? parseRange(f.where.split(',')[0].trim()) : null;
    if (rect && store.sheets[i].info.kind === 'worksheet') {
      store.setSelection({ row: rect.r1, col: rect.c1, rects: [rect] });
    }
  };
  const calc = store.calc;

  return (
    <aside className="report" aria-label="What's in this file">
      <div className="report-head">
        <div>
          <h2>What's in this file</h2>
          <p>Nothing listed here is removed when you save.</p>
        </div>
        <button className="btn icon ghost" onClick={() => store.set('reportOpen', false)} aria-label="Close">
          <Icon name="close" />
        </button>
      </div>
      <div className="report-body">
        {GROUPS.map((g) => {
          const items = store.findings
            .filter((f) => g.kinds.includes(f.kind))
            .sort((a, b) => (a.severity === b.severity ? g.kinds.indexOf(a.kind) - g.kinds.indexOf(b.kind) : a.severity === 'warn' ? -1 : 1));
          if (!items.length && g.title !== 'Recalculation') return null;
          return (
            <section key={g.title}>
              <h3>{g.title}</h3>
              <p className="blurb">{g.blurb}</p>
              {g.title === 'Recalculation' && calc && (
                <div className="calc-summary">
                  <div>
                    <strong>{calc.liveCells.toLocaleString()}</strong>
                    <span>recalculate live</span>
                  </div>
                  <div>
                    <strong>{(calc.unsupported + calc.mismatched).toLocaleString()}</strong>
                    <span>not calculated here</span>
                  </div>
                  <div>
                    <strong>{calc.buildMs < 1000 ? `${calc.buildMs} ms` : `${(calc.buildMs / 1000).toFixed(1)} s`}</strong>
                    <span>to check against Excel</span>
                  </div>
                </div>
              )}
              {g.title === 'Recalculation' && !calc && <p className="muted">Building the calculation engine…</p>}
              <ul className="findings">
                {items.map((f, i) => (
                  <li key={i} className={`finding sev-${f.severity}`}>
                    <button className="finding-main" onClick={() => go(f)} disabled={!f.sheet}>
                      <span className="finding-title">{f.title}</span>
                      {f.sheet && (
                        <span className="finding-where">
                          {f.sheet}
                          {f.where && parseA1(f.where.split(':')[0]) ? ` · ${f.where}` : ''}
                        </span>
                      )}
                    </button>
                    <p>{f.detail}</p>
                  </li>
                ))}
              </ul>
              {g.title === 'Recalculation' && calc && calc.samples.length > 0 && (
                <details className="samples">
                  <summary>Formulas not calculated here ({calc.samples.length === 50 ? 'first 50' : calc.samples.length})</summary>
                  <table>
                    <tbody>
                      {calc.samples.map((s, i) => (
                        <tr key={i}>
                          <td className="mono">
                            {s.sheet}!{s.ref}
                          </td>
                          <td className="mono formula">{s.formula}</td>
                          <td>
                            {s.reason}
                            {s.excel !== undefined && (
                              <span className="muted">
                                {' '}
                                (Excel {s.excel}, here {s.engine})
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
          <h3>Cell markers</h3>
          <ul className="legend">
            <li>
              <span className="swatch tri grey" /> Formula this editor can't calculate; shows Excel's saved value.
            </li>
            <li>
              <span className="swatch tri amber" /> Its inputs changed; Excel recalculates it when the file is opened.
            </li>
            <li>
              <span className="swatch edited" /> Changed in this session (toggle with “Highlight changes”).
            </li>
          </ul>
        </section>
        <p className="timings muted">
          Opened in {((store.timings.unzipMs ?? 0) + (store.timings.parseMs ?? 0)).toLocaleString()} ms
          {calc ? `; formulas checked in ${calc.buildMs.toLocaleString()} ms` : ''}.
        </p>
      </div>
    </aside>
  );
}
