import { useEffect, useId, useMemo, useRef, useState } from 'react';
import type { CellEditorProps, EditMove } from '../grid/contract';
import { Icon } from './icons';
import { useT } from '../i18n/react';

/** Options containing the filter, those starting with it first. */
function filterOptions(options: readonly string[], filter: string): string[] {
  const f = filter.trim().toLowerCase();
  if (!f) return [...options];
  const starts = options.filter((o) => o.toLowerCase().startsWith(f));
  const contains = options.filter((o) => !o.toLowerCase().startsWith(f) && o.toLowerCase().includes(f));
  return [...starts, ...contains];
}

/**
 * Dropdown editor for enum fields. Typing filters the list, arrow keys move,
 * Enter or Tab picks, Escape or clicking elsewhere leaves the cell unchanged.
 */
export function EnumEditor({ text, typed, finish, options: allowed, box: target }: CellEditorProps) {
  const t = useT();
  const options = useMemo(() => allowed ?? [], [allowed]);
  const current = text.replace(/^'/, '');
  const [filter, setFilter] = useState(typed ? text : '');
  const filtered = useMemo(() => filterOptions(options, filter), [options, filter]);
  const [active, setActive] = useState(() => Math.max(0, filtered.indexOf(current)));
  const list = useRef<HTMLUListElement>(null);
  const id = useId();

  useEffect(() => {
    const exact = filtered.findIndex((o) => o === current);
    setActive(filter ? 0 : Math.max(0, exact));
  }, [filtered, filter, current]);

  useEffect(() => {
    list.current?.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' });
  }, [active]);

  const choose = (option: string, movement: EditMove) => finish(option, movement);

  return (
    <div className="wbe-enum-editor" style={{ minWidth: Math.max(target.width, 160) }}>
      <input
        className="wbe-enum-filter"
        // oxlint-disable-next-line jsx-a11y/no-autofocus -- the list opens because the user asked for it, and typing filters it
        autoFocus
        value={filter}
        placeholder={current || t('dropdown.placeholder')}
        role="combobox"
        aria-label={t('dropdown.filter')}
        aria-expanded="true"
        aria-autocomplete="list"
        aria-controls={`${id}-list`}
        aria-activedescendant={filtered[active] !== undefined ? `${id}-${active}` : undefined}
        aria-describedby={current ? `${id}-clear-hint` : undefined}
        style={{ height: target.height }}
        onChange={(e) => setFilter(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
            e.preventDefault();
            e.stopPropagation();
            const step = e.key === 'ArrowDown' ? 1 : -1;
            setActive((i) => (filtered.length ? (i + step + filtered.length) % filtered.length : 0));
          } else if (e.key === 'Enter' || e.key === 'Tab') {
            // Handled here rather than by the grid, which would save the unchanged value.
            e.preventDefault();
            e.stopPropagation();
            const pick = filtered[active];
            const movement: EditMove = e.key === 'Enter' ? [0, 1] : [e.shiftKey ? -1 : 1, 0];
            if (pick !== undefined) choose(pick, movement);
            else if (!filter.trim()) finish(undefined, movement);
          } else if (e.key === 'Delete' && !filter && current) {
            // The keyboard's way to the Clear button.
            e.preventDefault();
            e.stopPropagation();
            choose('', [0, 0]);
          }
        }}
      />
      <ul
        className="wbe-enum-options"
        id={`${id}-list`}
        role="listbox"
        aria-label={t('dropdown.values')}
        ref={list}
        // The list can't be scrolled by the user directly (see editor.css); the wheel scrolls it here.
        onWheel={(e) => {
          e.currentTarget.scrollTop += e.deltaY * (e.deltaMode === 1 ? 23 : 1);
        }}
      >
        {filtered.map((o, i) => (
          <li
            key={o}
            id={`${id}-${i}`}
            role="option"
            aria-selected={i === active}
            className={i === active ? 'wbe-active' : ''}
            onMouseEnter={() => setActive(i)}
            onMouseDown={(e) => {
              e.preventDefault();
              choose(o, [0, 0]);
            }}
          >
            <span>{o}</span>
            {o === current && (
              <>
                <Icon name="check" size={14} />
                <span className="wbe-sr-only">{t('dropdown.current')}</span>
              </>
            )}
          </li>
        ))}
      </ul>
      {/* Outside the list, which holds only values, and said when it appears. */}
      <div role="status">{filtered.length === 0 && <p className="wbe-enum-empty">{t('dropdown.noMatch', { filter })}</p>}</div>
      {current && (
        <button
          className="wbe-enum-clear"
          // Keeps the focus in the filter, which would otherwise close the editor.
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => choose('', [0, 0])}
        >
          {t('dropdown.clear')}
        </button>
      )}
      {current && (
        <p id={`${id}-clear-hint`} className="wbe-sr-only">
          {t('dropdown.clearHint')}
        </p>
      )}
    </div>
  );
}
