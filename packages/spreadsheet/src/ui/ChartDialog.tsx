import { quoteSheetName, rangeToA1 } from '@ruled/formula/refs';
import { useId, useRef, useState } from 'react';
import { chartKindOf, seriesRefsOf, type ChartEdit, type LegendPosition } from '../core/chartEdit';
import type { ChartSpec } from '../core/charts';
import { NEW_CHART_KINDS, type NewChartKind } from '../core/chartWriter';
import { objectKey } from '../core/objectKey';
import type { Rect } from '../core/types';
import { useT } from '../i18n/react';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { Select } from './Select';
import { useDialog } from './useDialog';

const LEGENDS: (LegendPosition | 'none')[] = ['none', 'top', 'bottom', 'left', 'right'];

interface Draft {
  kind: NewChartKind | '';
  title: string;
  legend: LegendPosition | 'none';
  range: string;
  byColumns: boolean;
  series: { name: string; values: string }[];
  categories: string;
}

const shown = (ref: string | undefined) => (ref ? `=${ref.replace(/^=/, '')}` : '');

/** The range a chart's data comes from, when it is one block on one sheet, and whether its series run down the columns. */
function rangeOf(spec: ChartSpec | undefined, sheetName: (i: number) => string): { range: string; byColumns: boolean } {
  const series = spec?.groups.flatMap((g) => g.series) ?? [];
  const refs = series.flatMap((s) => [s.name?.ref, s.values.ref, s.categories?.ref]).filter((r): r is NonNullable<typeof r> => !!r);
  const tall = series.every((s) => !s.values.ref || s.values.ref.rect.r2 > s.values.ref.rect.r1 || s.values.ref.rect.c2 === s.values.ref.rect.c1);
  if (!refs.length || refs.some((r) => r.sheet !== refs[0].sheet)) return { range: '', byColumns: tall };
  const rect: Rect = {
    r1: Math.min(...refs.map((r) => r.rect.r1)),
    c1: Math.min(...refs.map((r) => r.rect.c1)),
    r2: Math.max(...refs.map((r) => r.rect.r2)),
    c2: Math.max(...refs.map((r) => r.rect.c2)),
  };
  const a1 = rangeToA1(rect).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
  return { range: `=${quoteSheetName(sheetName(refs[0].sheet))}!${a1}`, byColumns: tall };
}

function draftOf(spec: ChartSpec | undefined, sheetName: (i: number) => string): Draft {
  const refs = seriesRefsOf(spec);
  const legend = spec?.legend?.position;
  return {
    kind: chartKindOf(spec) ?? '',
    title: spec?.title?.cache.filter((v) => v !== null).join('') ?? '',
    legend: legend === 'top' || legend === 'bottom' || legend === 'left' || legend === 'right' ? legend : 'none',
    ...rangeOf(spec, sheetName),
    series: refs.map((s) => ({ name: s.name && /!|^=/.test(s.name) ? shown(s.name) : (s.name ?? ''), values: shown(s.values) })),
    categories: shown(refs[0]?.categories),
  };
}

/**
 * Excel's chart dialogs in one: Change Chart Type, the chart's title and
 * legend, and Select Data, with the data range (its series down the columns
 * or along the rows) or each series' name and values and the category
 * labels. OK changes only what was changed.
 */
export function ChartDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const id = useId();
  const at = store.chartDialog;
  const obj = at && store.view?.info.objects.find((o) => o.kind === 'chart' && objectKey(o) === at.key);
  const dialog = useRef<HTMLFormElement>(null);
  const initial = useRef<HTMLSelectElement>(null);
  const [draft, setDraft] = useState<Draft | null>(null);
  const [start, setStart] = useState<Draft | null>(null);
  const opened = useRef<typeof at>(undefined);
  if (opened.current !== at) {
    opened.current = at;
    if (at && obj) {
      const d = draftOf(obj.chart, (i) => store.sheets[i]?.info.name ?? '');
      setDraft(d);
      setStart(d);
    }
  }
  const close = () => store.set('chartDialog', undefined);
  useDialog(dialog, !!at, close, { modal: true, initial });
  if (!at || !obj || !draft || !start) return null;

  const set = (patch: Partial<Draft>) => setDraft({ ...draft, ...patch });
  const setSeries = (i: number, patch: Partial<Draft['series'][number]>) => set({ series: draft.series.map((s, j) => (j === i ? { ...s, ...patch } : s)) });
  const fromRange = draft.range !== start.range || draft.byColumns !== start.byColumns;
  const scatter = (draft.kind || start.kind) === 'scatter';
  const submit = () => {
    const edit: ChartEdit = {};
    if (draft.kind && draft.kind !== start.kind) edit.kind = draft.kind;
    if (draft.title !== start.title) edit.title = draft.title.trim();
    if (draft.legend !== start.legend) edit.legend = draft.legend === 'none' ? null : draft.legend;
    if (fromRange) edit.range = { ref: draft.range.replace(/^=/, ''), byColumns: draft.byColumns };
    else if (JSON.stringify([draft.series, draft.categories]) !== JSON.stringify([start.series, start.categories])) {
      edit.series = draft.series.filter((s) => s.values.trim()).map((s) => ({ name: s.name, values: s.values, categories: draft.categories }));
    }
    close();
    void store.editChart(at.key, edit);
  };

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-chart-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        tabIndex={-1}
        noValidate
        onMouseDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          submit();
        }}
      >
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t('charts.edit.heading')}</h2>
          <div className="wbe-chart-row">
            <div className="wbe-field">
              <label htmlFor={`${id}-kind`}>{t('charts.edit.type')}</label>
              <Select id={`${id}-kind`} ref={initial} value={draft.kind} onChange={(e) => set({ kind: e.target.value as NewChartKind })} aria-describedby={`${id}-kind-hint`}>
                {!start.kind && <option value="">{t('charts.edit.keepType')}</option>}
                {NEW_CHART_KINDS.map((k) => (
                  <option key={k} value={k}>
                    {t(`core.chartType.${k}`)}
                  </option>
                ))}
              </Select>
            </div>
            <div className="wbe-field">
              <label htmlFor={`${id}-legend`}>{t('charts.edit.legend')}</label>
              <Select id={`${id}-legend`} value={draft.legend} onChange={(e) => set({ legend: e.target.value as Draft['legend'] })}>
                {LEGENDS.map((l) => (
                  <option key={l} value={l}>
                    {t(`charts.edit.legend.${l}`)}
                  </option>
                ))}
              </Select>
            </div>
          </div>
          <p id={`${id}-kind-hint`} className="wbe-muted wbe-small">
            {t('charts.edit.typeHint')}
          </p>
          <div className="wbe-field">
            <label htmlFor={`${id}-name`}>{t('charts.edit.title')}</label>
            <input id={`${id}-name`} value={draft.title} autoComplete="off" onChange={(e) => set({ title: e.target.value })} aria-describedby={`${id}-name-hint`} />
            <p id={`${id}-name-hint`} className="wbe-muted wbe-small">
              {t('charts.edit.titleHint')}
            </p>
          </div>

          <fieldset className="wbe-chart-data">
            <legend>{t('charts.edit.data')}</legend>
            <div className="wbe-chart-row">
              <div className="wbe-field">
                <label htmlFor={`${id}-range`}>{t('charts.edit.range')}</label>
                <input id={`${id}-range`} value={draft.range} autoComplete="off" spellCheck={false} onChange={(e) => set({ range: e.target.value })} aria-describedby={`${id}-range-hint`} />
              </div>
              <div className="wbe-field">
                <label htmlFor={`${id}-by`}>{t('charts.edit.seriesIn')}</label>
                <Select id={`${id}-by`} value={draft.byColumns ? 'columns' : 'rows'} onChange={(e) => set({ byColumns: e.target.value === 'columns' })}>
                  <option value="columns">{t('charts.edit.columns')}</option>
                  <option value="rows">{t('charts.edit.rows')}</option>
                </Select>
              </div>
            </div>
            <p id={`${id}-range-hint`} className="wbe-muted wbe-small">
              {t('charts.edit.rangeHint')}
            </p>
            <div className="wbe-chart-series" role="group" aria-label={t('charts.edit.series')}>
              <span className="wbe-small wbe-muted" aria-hidden="true">
                {t('charts.edit.seriesName')}
              </span>
              <span className="wbe-small wbe-muted" aria-hidden="true">
                {t('charts.edit.seriesValues')}
              </span>
              <span />
              {draft.series.map((s, i) => {
                const label = t('charts.edit.seriesLabel', { n: i + 1 });
                return (
                  <div key={i} className="wbe-chart-series-row">
                    <input
                      value={s.name}
                      disabled={fromRange}
                      autoComplete="off"
                      spellCheck={false}
                      onChange={(e) => setSeries(i, { name: e.target.value })}
                      aria-label={`${label}: ${t('charts.edit.seriesName')}`}
                    />
                    <input
                      value={s.values}
                      disabled={fromRange}
                      autoComplete="off"
                      spellCheck={false}
                      onChange={(e) => setSeries(i, { values: e.target.value })}
                      aria-label={`${label}: ${t('charts.edit.seriesValues')}`}
                    />
                    <button
                      type="button"
                      className="wbe-btn wbe-icon"
                      disabled={fromRange || draft.series.length < 2}
                      onClick={() => set({ series: draft.series.filter((_, j) => j !== i) })}
                      title={t('charts.edit.removeSeries', { n: i + 1 })}
                      aria-label={t('charts.edit.removeSeries', { n: i + 1 })}
                    >
                      <Icon name="trash" />
                    </button>
                  </div>
                );
              })}
            </div>
            <button type="button" className="wbe-btn" disabled={fromRange} onClick={() => set({ series: [...draft.series, { name: '', values: '' }] })}>
              <Icon name="plus" />
              {t('charts.edit.addSeries')}
            </button>
            <div className="wbe-field">
              <label htmlFor={`${id}-categories`}>{t(scatter ? 'charts.edit.xValues' : 'charts.edit.categories')}</label>
              <input id={`${id}-categories`} value={draft.categories} disabled={fromRange} autoComplete="off" spellCheck={false} onChange={(e) => set({ categories: e.target.value })} />
            </div>
          </fieldset>

          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary">
              {t('dialog.ok')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}
