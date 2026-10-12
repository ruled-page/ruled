import { useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { colToLetters } from '@sitepen/formula-engine/refs';
import type { WorkbookStore } from '../state/store';
import { useDialog } from './useDialog';
import { msg } from '../i18n/message';
import { useT } from '../i18n/react';

/**
 * Picks which values of a column to show, like Excel's filter list. Rows are
 * hidden on screen only; the file's own filter is left as it was.
 */
export function FilterPicker({ store }: { store: WorkbookStore }) {
  const t = useT();
  const id = useId();
  const at = store.filterPicker;
  const choices = useMemo(() => (at ? store.filterChoices(at.col) : undefined), [store, at]);
  const [checked, setChecked] = useState<Set<string>>(new Set());
  const [search, setSearch] = useState('');
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ left: 0, top: 0 });

  useEffect(() => {
    if (choices) setChecked(new Set(choices.allowed ?? choices.values.map((v) => v.text)));
    setSearch('');
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
        {choices.allowed && (
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
