import { useId, useRef, useState } from 'react';
import { isDateFormat } from 'numfmt';
import { parseInput } from '../core/input';
import type { SeriesOptions } from '../core/types';
import type { MessageKey } from '../i18n/message';
import { useT } from '../i18n/react';
import { cleanLists, splitEntries } from '../state/fillLists';
import type { WorkbookStore } from '../state/store';
import { useDialog } from './useDialog';

interface Draft extends Omit<SeriesOptions, 'step' | 'stop'> {
  /** As typed, in the user's locale. */
  step: string;
  stop: string;
}

/** A number typed in the user's locale, as a cell reads it: "1,5" in Germany. Dates count only where `dates` allows. */
function readNumber(store: WorkbookStore, text: string, dates: boolean): number | undefined {
  const parsed = parseInput(text.trim(), 'General', store.locale, store.date1904);
  if (parsed.error || parsed.content.kind !== 'number') return undefined;
  if (!dates && parsed.numFmt && isDateFormat(parsed.numFmt)) return undefined;
  return parsed.content.value;
}

const BY = ['rows', 'columns'] as const;
const TYPES = ['linear', 'growth', 'date', 'autofill'] as const;
const UNITS = ['day', 'weekday', 'month', 'year'] as const;

/**
 * Excel's Fill Series dialog (Home › Fill › Series) for the selection: in
 * rows or columns, Linear, Growth, Date (by day, weekday, month or year) or
 * AutoFill, with a step and a stop value, or a trend fitted to the selected
 * numbers. Numbers and dates are read in the user's locale.
 */
export function SeriesDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const open = store.seriesDialog;
  const dialog = useRef<HTMLFormElement>(null);
  const initial = useRef<HTMLInputElement>(null);
  const stepInput = useRef<HTMLInputElement>(null);
  const stopInput = useRef<HTMLInputElement>(null);
  const id = useId();
  const [draft, setDraft] = useState<Draft | null>(null);
  const [tried, setTried] = useState(false);
  const opened = useRef(false);
  if (opened.current !== open) {
    opened.current = open;
    if (open) {
      const start = store.seriesStart();
      setDraft({ ...start, step: t.number(start.step), stop: '' });
      setTried(false);
    }
  }
  const close = () => store.set('seriesDialog', false);
  useDialog(dialog, open, close, { modal: true, initial });
  if (!open || !draft) return null;

  const set = (patch: Partial<Draft>) => setDraft({ ...draft, ...patch });
  const numbers = draft.type !== 'autofill';
  const trendable = draft.type === 'linear' || draft.type === 'growth';
  const trend = trendable && draft.trend;
  const usesStep = numbers && !trend;
  const step = readNumber(store, draft.step, false);
  const stop = numbers && draft.stop.trim() ? readNumber(store, draft.stop, true) : undefined;
  const sel = store.selection;
  const rect = sel.rects[0];
  const oneCell = !rect || (rect.r1 === rect.r2 && rect.c1 === rect.c2);

  const stepError: MessageKey | undefined = !usesStep
    ? undefined
    : !draft.step.trim()
      ? 'fill.series.stepMissing'
      : step === undefined
        ? 'fill.series.stepInvalid'
        : draft.type === 'date' && draft.dateUnit !== 'day' && !Number.isInteger(step)
          ? 'fill.series.stepWhole'
          : undefined;
  const stopError: MessageKey | undefined = numbers && draft.stop.trim() && stop === undefined ? 'fill.series.stopInvalid' : undefined;
  const rangeError: MessageKey | undefined = !oneCell ? undefined : !numbers ? 'fill.series.oneCellAutofill' : draft.stop.trim() ? undefined : 'fill.series.oneCell';
  const shown = (error: MessageKey | undefined) => (tried ? error : undefined);
  const example = { example: t.number(1.5) };

  const submit = () => {
    if (stepError || stopError || rangeError) {
      setTried(true);
      if (stepError) stepInput.current?.focus();
      else if (stopError || (rangeError && numbers)) stopInput.current?.focus();
      return;
    }
    close();
    void store.fillSeries({
      by: draft.by,
      type: draft.type,
      dateUnit: draft.dateUnit,
      step: usesStep ? step! : 1,
      ...(stop !== undefined && { stop }),
      trend,
    });
  };

  const radio = <K extends 'by' | 'type' | 'dateUnit'>(key: K, value: Draft[K], label: string, first?: boolean) => (
    <label key={String(value)}>
      <input type="radio" name={`${id}-${key}`} ref={first ? initial : undefined} checked={draft[key] === value} onChange={() => set({ [key]: value } as Partial<Draft>)} />
      {label}
    </label>
  );

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal wbe-series-dialog"
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
          <h2 id={`${id}-title`}>{t('fill.series.title')}</h2>
          <div className="wbe-series-groups">
            <fieldset className="wbe-series-group">
              <legend>{t('fill.series.by')}</legend>
              {BY.map((by) => radio('by', by, t(`fill.series.${by}`), draft.by === by))}
            </fieldset>
            <fieldset className="wbe-series-group">
              <legend>{t('fill.series.type')}</legend>
              {TYPES.map((type) => radio('type', type, t(`fill.series.${type}`)))}
            </fieldset>
            <fieldset className="wbe-series-group" disabled={draft.type !== 'date'}>
              <legend>{t('fill.series.dateUnit')}</legend>
              {UNITS.map((unit) => radio('dateUnit', unit, t(`fill.series.${unit}`)))}
            </fieldset>
          </div>
          <label className="wbe-series-trend">
            <input type="checkbox" checked={trend} disabled={!trendable} onChange={(e) => set({ trend: e.target.checked })} aria-describedby={`${id}-trend`} />
            {t('fill.series.trend')}
          </label>
          <p id={`${id}-trend`} className="wbe-muted wbe-small">
            {t(numbers ? 'fill.series.trendHint' : 'fill.series.autofillHint')}
          </p>
          <div className="wbe-series-values">
            <div className="wbe-field">
              <label htmlFor={`${id}-step`}>{t('fill.series.step')}</label>
              <input
                id={`${id}-step`}
                ref={stepInput}
                value={draft.step}
                inputMode="decimal"
                autoComplete="off"
                spellCheck={false}
                disabled={!usesStep}
                onChange={(e) => set({ step: e.target.value })}
                aria-invalid={!!shown(stepError)}
                aria-describedby={shown(stepError) ? `${id}-step-error` : undefined}
              />
              {shown(stepError) && (
                <p id={`${id}-step-error`} className="wbe-field-error">
                  {t(stepError!, example)}
                </p>
              )}
            </div>
            <div className="wbe-field">
              <label htmlFor={`${id}-stop`}>{t('fill.series.stop')}</label>
              <input
                id={`${id}-stop`}
                ref={stopInput}
                value={draft.stop}
                autoComplete="off"
                spellCheck={false}
                disabled={!numbers}
                onChange={(e) => set({ stop: e.target.value })}
                aria-invalid={!!shown(stopError ?? (numbers ? rangeError : undefined))}
                aria-describedby={[`${id}-stop-hint`, shown(stopError ?? (numbers ? rangeError : undefined)) && `${id}-stop-error`].filter(Boolean).join(' ')}
              />
              {shown(stopError ?? (numbers ? rangeError : undefined)) && (
                <p id={`${id}-stop-error`} className="wbe-field-error">
                  {t((stopError ?? rangeError)!, { example: t.number(100) })}
                </p>
              )}
            </div>
          </div>
          <p id={`${id}-stop-hint`} className="wbe-muted wbe-small">
            {t('fill.series.stopHint')}
          </p>
          {shown(rangeError) && !numbers && (
            <p className="wbe-series-error" role="alert">
              {t(rangeError!)}
            </p>
          )}
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn wbe-series-lists" aria-haspopup="dialog" onClick={() => store.set('customListsOpen', true)}>
              {t('fill.series.customLists')}
            </button>
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary">
              {t('fill.series.fill')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}

/** Where a list in the dialog comes from, and which it is. */
type Picked = 'new' | `${'built' | 'host' | 'own'}:${number}`;

/**
 * Excel's Custom Lists dialog: the lists that dragging the fill handle and
 * AutoFill continue. The built-in ones and the host's can be read; the
 * user's own can be added, changed and deleted, and are kept in this browser.
 */
export function CustomListsDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const open = store.customListsOpen;
  const dialog = useRef<HTMLDivElement>(null);
  const listBox = useRef<HTMLSelectElement>(null);
  const entriesInput = useRef<HTMLTextAreaElement>(null);
  const id = useId();
  const [picked, setPicked] = useState<Picked>('new');
  const [entries, setEntries] = useState('');
  const [tried, setTried] = useState(false);
  const [status, setStatus] = useState<MessageKey | undefined>();
  const opened = useRef(false);
  if (opened.current !== open) {
    opened.current = open;
    if (open) {
      setPicked('new');
      setEntries('');
      setTried(false);
      setStatus(undefined);
    }
  }
  const close = () => store.set('customListsOpen', false);
  useDialog(dialog, open, close, { modal: true });
  if (!open) return null;

  const groups = { built: store.builtInFillLists, host: store.hostFillLists, own: store.userLists };
  const listOf = (key: Picked) => {
    if (key === 'new') return undefined;
    const [group, i] = key.split(':') as [keyof typeof groups, string];
    return groups[group][Number(i)];
  };
  const choose = (key: Picked) => {
    setPicked(key);
    setEntries(listOf(key)?.join('\n') ?? '');
    setTried(false);
    setStatus(undefined);
  };
  const editable = picked === 'new' || picked.startsWith('own:');
  const list = cleanLists([splitEntries(entries)])[0];
  const error = tried && editable && !list;
  const keyOf = (l: string[]) => l.join('\n').toLowerCase();

  const save = () => {
    if (!list) {
      setTried(true);
      entriesInput.current?.focus();
      return;
    }
    const own = store.userLists;
    const at = picked === 'new' ? own.length : Number(picked.slice(4));
    store.setUserLists(picked === 'new' ? [...own, list] : own.map((l, i) => (i === at ? list : l)));
    // A list the same as another one is kept once.
    const now = store.userLists.findIndex((l) => keyOf(l) === keyOf(list));
    setPicked(now < 0 ? 'new' : `own:${now}`);
    setEntries(list.join('\n'));
    setTried(false);
    setStatus(picked === 'new' ? 'fill.lists.added' : 'fill.lists.saved');
  };
  const remove = async () => {
    if (!picked.startsWith('own:')) return;
    if (await store.deleteUserList(Number(picked.slice(4)))) {
      choose('new');
      setStatus('fill.lists.deleted');
      // The Delete button it was pressed with is now disabled.
      listBox.current?.focus();
    }
  };
  const importCells = () => {
    setEntries(store.selectedTexts().join('\n'));
    setStatus(undefined);
    entriesInput.current?.focus();
  };

  const option = (group: keyof typeof groups) =>
    groups[group].map((l, i) => (
      <option key={`${group}:${i}`} value={`${group}:${i}`}>
        {l.join(', ')}
      </option>
    ));

  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <div
        ref={dialog}
        className="wbe-modal wbe-modal-wide wbe-lists-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        tabIndex={-1}
        onMouseDown={(e) => e.stopPropagation()}
      >
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t('fill.lists.title')}</h2>
          <div className="wbe-lists-body">
            <div className="wbe-field">
              <label htmlFor={`${id}-lists`}>{t('fill.lists.lists')}</label>
              <select id={`${id}-lists`} ref={listBox} size={10} value={picked} onChange={(e) => choose(e.target.value as Picked)}>
                <option value="new">{t('fill.lists.new')}</option>
                {groups.own.length > 0 && <optgroup label={t('fill.lists.yours')}>{option('own')}</optgroup>}
                {groups.host.length > 0 && <optgroup label={t('fill.lists.host')}>{option('host')}</optgroup>}
                <optgroup label={t('fill.lists.builtIn')}>{option('built')}</optgroup>
              </select>
            </div>
            <div className="wbe-field">
              <label htmlFor={`${id}-entries`}>{t('fill.lists.entries')}</label>
              <textarea
                id={`${id}-entries`}
                ref={entriesInput}
                rows={9}
                value={entries}
                readOnly={!editable}
                spellCheck={false}
                onChange={(e) => {
                  setEntries(e.target.value);
                  setStatus(undefined);
                }}
                aria-invalid={error}
                aria-describedby={[`${id}-hint`, error && `${id}-error`].filter(Boolean).join(' ')}
              />
              {error && (
                <p id={`${id}-error`} className="wbe-field-error">
                  {t('fill.lists.tooFew')}
                </p>
              )}
            </div>
          </div>
          <p id={`${id}-hint`} className="wbe-muted wbe-small">
            {t(editable ? 'fill.lists.entriesHint' : 'fill.lists.readOnly')}
          </p>
          {store.listsNotKept && <p className="wbe-muted wbe-small">{t('fill.lists.notKept')}</p>}
          <div className="wbe-lists-actions">
            <button type="button" className="wbe-btn" onClick={importCells} disabled={!editable}>
              {t('fill.lists.import')}
            </button>
            <button type="button" className="wbe-btn" onClick={save} disabled={!editable}>
              {t(picked === 'new' ? 'fill.lists.add' : 'fill.lists.save')}
            </button>
            <button type="button" className="wbe-btn" onClick={() => void remove()} disabled={!picked.startsWith('own:')}>
              {t('fill.lists.delete')}
            </button>
            <span className="wbe-lists-status" role="status">
              {status ? t(status) : ''}
            </span>
          </div>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn wbe-primary" onClick={close}>
              {t('fill.lists.close')}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
