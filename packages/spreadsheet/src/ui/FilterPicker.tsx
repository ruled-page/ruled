import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { colToLetters } from '@ruled/formula/refs';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { useDialog } from './useDialog';
import { msg } from '../i18n/message';
import { useT } from '../i18n/react';

type ConditionItem = 'equal' | 'notEqual' | 'beginsWith' | 'endsWith' | 'contains' | 'notContains' | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual' | 'between' | 'top10' | 'aboveAverage' | 'belowAverage' | 'before' | 'after' | 'custom';

/** What Excel's Text, Number and Date Filters offer. */
const CONDITION_ITEMS: Record<'text' | 'number' | 'date', ConditionItem[]> = {
  text: ['equal', 'notEqual', 'beginsWith', 'endsWith', 'contains', 'notContains', 'custom'],
  number: ['equal', 'notEqual', 'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'between', 'top10', 'aboveAverage', 'belowAverage', 'custom'],
  date: ['equal', 'before', 'after', 'between', 'custom'],
};

/**
 * Excel's filter menu for a column: sort the filter's range by it, filter
 * it by a condition (Text, Number or Date Filters), or pick which of its
 * values to show.
 */
export function FilterPicker({ store }: { store: WorkbookStore }) {
  const t = useT();
  const id = useId();
  const at = store.filterPicker;
  const choices = useMemo(() => (at ? store.filterChoices(at.col, at.rect) : undefined), [store, at]);
  const [checked, setChecked] = useState<Set<string>>(new Set());
  const [search, setSearch] = useState('');
  const [conditions, setConditions] = useState(false);
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ left: 0, top: 0 });

  useEffect(() => {
    if (choices) setChecked(new Set(choices.allowed ?? choices.values.map((v) => v.text)));
    setSearch('');
    setConditions(false);
  }, [choices]);

  useLayoutEffect(() => {
    const el = ref.current;
    if (!el || !at) return;
    const r = el.getBoundingClientRect();
    setPos({ left: Math.max(4, Math.min(at.x, window.innerWidth - r.width - 4)), top: Math.max(4, Math.min(at.y, window.innerHeight - r.height - 4)) });
  }, [at, choices]);

  useEffect(() => {
    if (!at) return;
    const away = (e: Event) => {
      if (!ref.current?.contains(e.target as Node)) store.set('filterPicker', undefined);
    };
    window.addEventListener('mousedown', away, true);
    return () => window.removeEventListener('mousedown', away, true);
  }, [store, at]);

  useDialog(ref, !!(at && choices), () => store.set('filterPicker', undefined), { modal: true });

  useEffect(() => {
    if (at && !choices) {
      store.toast('info', msg('filter.noData'));
      store.set('filterPicker', undefined);
    }
  }, [store, at, choices]);

  if (!at || !choices) return null;
  const view = store.view;
  const heading = view ? store.shownText(view, choices.rect.r1, at.col) : '';
  const shown = choices.values.filter((v) => v.text.toLowerCase().includes(search.toLowerCase()));
  const all = shown.every((v) => checked.has(v.text));
  const toggle = (text: string) => {
    const next = new Set(checked);
    if (next.has(text)) next.delete(text);
    else next.add(text);
    setChecked(next);
  };
  // Sorting the filter's range by this column, as the top of Excel's filter menu does.
  const kind = store.sortLines(choices.rect, true, false).find((l) => l.index === at.col)?.kind ?? 'text';
  const sort = (descending: boolean) => {
    void store.sortBy({ rect: choices.rect, keys: [{ col: at.col, descending }], header: true, caseSensitive: false, across: false });
    store.set('filterPicker', undefined);
  };
  // Excel's Text, Number or Date Filters: a rule picked opens the Custom AutoFilter dialog, starting on it.
  const items = CONDITION_ITEMS[choices.kind];
  const pick = (item: ConditionItem) => {
    store.set('filterPicker', undefined);
    if (item === 'aboveAverage' || item === 'belowAverage') {
      void store.setFilterCondition(choices.rect, at.col, { kind: 'average', above: item === 'aboveAverage' });
      return;
    }
    const preset =
      item === 'between'
        ? { first: 'greaterThanOrEqual' as const, second: 'lessThanOrEqual' as const, and: true }
        : { first: item === 'before' ? ('lessThan' as const) : item === 'after' ? ('greaterThan' as const) : item === 'custom' || item === 'top10' ? ('equal' as const) : item, and: true };
    store.set('conditionDialog', { col: at.col, rect: choices.rect, mode: item === 'top10' ? 'top10' : 'custom', preset });
  };
  const apply = () => {
    const everything = choices.values.every((v) => checked.has(v.text));
    void store.setFilter(choices.rect, at.col, everything ? undefined : checked);
    store.set('filterPicker', undefined);
  };

  return (
    <div className="wbe-filter-picker" ref={ref} style={pos} role="dialog" aria-modal="true" aria-labelledby={`${id}-title`}>
      <div id={`${id}-title`} className="wbe-filter-title">
        {heading ? t('filter.titleHeading', { heading }) : t('filter.titleColumn', { column: colToLetters(at.col) })}
      </div>
      <div className="wbe-filter-sort">
        <button className="wbe-btn wbe-ghost" onClick={() => sort(false)}>
          {t(kind === 'number' ? 'sort.ascNumber' : kind === 'date' ? 'sort.ascDate' : 'sort.ascText')}
        </button>
        <button className="wbe-btn wbe-ghost" onClick={() => sort(true)}>
          {t(kind === 'number' ? 'sort.descNumber' : kind === 'date' ? 'sort.descDate' : 'sort.descText')}
        </button>
      </div>
      <div className="wbe-filter-conditions">
        <button className="wbe-btn wbe-ghost" aria-expanded={conditions} onClick={() => setConditions(!conditions)}>
          <span>{t(choices.kind === 'number' ? 'filter.numberFilters' : choices.kind === 'date' ? 'filter.dateFilters' : 'filter.textFilters')}</span>
          {choices.condition && <Icon name="check" size={14} />}
          <Icon name={conditions ? 'chevronUp' : 'chevronDown'} size={14} />
        </button>
        {conditions && (
          <ul>
            {items.map((item) => (
              <li key={item}>
                <button className="wbe-btn wbe-ghost" onClick={() => pick(item)}>
                  {t(`filter.item.${item}`)}
                </button>
              </li>
            ))}
          </ul>
        )}
      </div>
      <input className="wbe-filter-search" placeholder={t('filter.search')} aria-label={t('filter.search')} value={search} onChange={(e) => setSearch(e.target.value)} autoFocus />
      <label className="wbe-filter-all">
        <input
          type="checkbox"
          checked={all}
          onChange={() => {
            const next = new Set(checked);
            for (const v of shown) {
              if (all) next.delete(v.text);
              else next.add(v.text);
            }
            setChecked(next);
          }}
        />
        {t(search ? 'filter.selectAllShown' : 'filter.selectAll')}
      </label>
      <ul className="wbe-filter-values">
        {shown.map((v) => (
          <li key={v.text}>
            <label>
              <input type="checkbox" checked={checked.has(v.text)} onChange={() => toggle(v.text)} />
              <span className={v.text ? undefined : 'wbe-muted'}>{v.text || t('filter.blanks')}</span>
              <span className="wbe-filter-count">{t.number(v.count)}</span>
            </label>
          </li>
        ))}
      </ul>
      <p className="wbe-filter-note">{t('filter.note')}</p>
      <div className="wbe-filter-actions">
        {(choices.allowed || choices.condition) && (
          <button
            className="wbe-btn wbe-link"
            onClick={() => {
              void store.setFilter(choices.rect, at.col, undefined);
              store.set('filterPicker', undefined);
            }}
          >
            {t('filter.clear')}
          </button>
        )}
        <span className="wbe-spacer" />
        <button className="wbe-btn" onClick={() => store.set('filterPicker', undefined)}>
          {t('dialog.cancel')}
        </button>
        <button className="wbe-btn wbe-primary" onClick={apply} disabled={!checked.size}>
          {t('filter.show')}
        </button>
      </div>
    </div>
  );
}
