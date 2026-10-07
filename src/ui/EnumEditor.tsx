import type { GridCell, ProvideEditorComponent, TextCell } from '@glideapps/glide-data-grid';
import { useEffect, useId, useMemo, useRef, useState } from 'react';
import { Icon } from './icons';

/** A text cell whose field only allows a fixed list of values. */
export type EnumCell = TextCell & { enumOptions?: readonly string[] };

export function enumOptionsOf(cell: GridCell): readonly string[] | undefined {
  return (cell as EnumCell).enumOptions;
}

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
export const EnumEditor: ProvideEditorComponent<GridCell> = ({ value, onFinishedEditing, initialValue, target }) => {
  const cell = value as EnumCell;
  const options = useMemo(() => cell.enumOptions ?? [], [cell.enumOptions]);
  const current = cell.data.replace(/^'/, '');
  const [filter, setFilter] = useState(initialValue ?? '');
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

  const choose = (option: string, movement: readonly [-1 | 0 | 1, -1 | 0 | 1]) => {
    onFinishedEditing({ ...cell, data: option }, movement);
  };

  return (
    <div className="enum-editor" style={{ minWidth: Math.max(target.width, 160) }}>
      <input
        className="enum-filter"
        autoFocus
        value={filter}
        placeholder={current || 'Choose a value'}
        role="combobox"
        aria-label="Filter values"
        aria-expanded="true"
        aria-autocomplete="list"
        aria-controls={`${id}-list`}
        aria-activedescendant={filtered[active] !== undefined ? `${id}-${active}` : undefined}
        style={{ height: target.height }}
        onChange={(e) => setFilter(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
            e.preventDefault();
            e.stopPropagation();
            const step = e.key === 'ArrowDown' ? 1 : -1;
            setActive((i) => (filtered.length ? (i + step + filtered.length) % filtered.length : 0));
          } else if (e.key === 'Enter' || e.key === 'Tab') {
            // Handled here rather than by Glide, which would save the unchanged value.
            e.preventDefault();
            e.stopPropagation();
            const pick = filtered[active];
            const movement: readonly [-1 | 0 | 1, -1 | 0 | 1] = e.key === 'Enter' ? [0, 1] : [e.shiftKey ? -1 : 1, 0];
            if (pick !== undefined) choose(pick, movement);
            else if (!filter.trim()) onFinishedEditing(undefined, movement);
          }
        }}
      />
      <ul className="enum-options" id={`${id}-list`} role="listbox" aria-label="Values" ref={list}>
        {filtered.map((o, i) => (
          <li
            key={o}
            id={`${id}-${i}`}
            role="option"
            aria-selected={i === active}
            className={i === active ? 'active' : ''}
            onMouseEnter={() => setActive(i)}
            onMouseDown={(e) => {
              e.preventDefault();
              choose(o, [0, 0]);
            }}
          >
            <span>{o}</span>
            {o === current && <Icon name="check" size={14} />}
          </li>
        ))}
        {filtered.length === 0 && <li className="empty">No value matches “{filter}”</li>}
      </ul>
      {current && (
        <button
          className="enum-clear"
          // Keeps the focus in the filter, which would otherwise close the editor.
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => choose('', [0, 0])}
        >
          Clear
        </button>
      )}
    </div>
  );
};
