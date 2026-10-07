import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { colToLetters } from '../core/refs';
import type { WorkbookStore } from '../state/store';
import { useDialog } from './useDialog';

/**
 * Picks which values of a column to show, like Excel's filter list. Rows are
 * hidden on screen only; the file's own filter is left as it was.
 */
export function FilterPicker({ store }: { store: WorkbookStore }) {
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
      store.toast('info', 'Select a cell in a block of data with headings to filter it.');
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
    store.setFilter(choices.rect, at.col, everything ? undefined : checked);
    store.set('filterPicker', undefined);
  };

  return (
    <div className="filter-picker" ref={ref} style={pos} role="dialog" aria-label="Filter by values">
      <div className="filter-title">Show rows where {heading ? `“${heading}”` : `column ${colToLetters(at.col)}`} is</div>
      <input className="filter-search" placeholder="Search values" value={search} onChange={(e) => setSearch(e.target.value)} autoFocus />
      <label className="filter-all">
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
        Select all{search ? ' shown' : ''}
      </label>
      <ul className="filter-values">
        {shown.map((v) => (
          <li key={v.text}>
            <label>
              <input type="checkbox" checked={checked.has(v.text)} onChange={() => toggle(v.text)} />
              <span className={v.text ? undefined : 'muted'}>{v.text || '(Blanks)'}</span>
              <span className="filter-count">{v.count.toLocaleString()}</span>
            </label>
          </li>
        ))}
      </ul>
      <p className="filter-note">Filtering changes what you see here. The file is saved with its own filter, as it was.</p>
      <div className="filter-actions">
        {choices.allowed && (
          <button
            className="btn link"
            onClick={() => {
              store.setFilter(choices.rect, at.col, undefined);
              store.set('filterPicker', undefined);
            }}
          >
            Clear
          </button>
        )}
        <span className="spacer" />
        <button className="btn" onClick={() => store.set('filterPicker', undefined)}>
          Cancel
        </button>
        <button className="btn primary" onClick={apply} disabled={!checked.size}>
          Show
        </button>
      </div>
    </div>
  );
}
