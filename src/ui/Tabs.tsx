import { useId, useRef, type ReactNode } from 'react';

export interface Tab<T extends string> {
  id: T;
  label: ReactNode;
}

/**
 * Tabs as screen readers expect them: the selected tab is the one tab stop,
 * the arrow keys (and Home and End) move between tabs, and `children` is the
 * panel the selected tab shows.
 */
export function Tabs<T extends string>({
  tabs,
  value,
  onChange,
  label,
  head,
  children,
}: {
  tabs: Tab<T>[];
  value: T;
  onChange: (id: T) => void;
  label: string;
  /** Shown beside the tabs, such as a close button. */
  head?: (tabs: ReactNode) => ReactNode;
  children: ReactNode;
}) {
  const id = useId();
  const refs = useRef<(HTMLButtonElement | null)[]>([]);
  const at = tabs.findIndex((t) => t.id === value);
  const list = (
    <div
      className="wbe-panel-tabs"
      role="tablist"
      aria-label={label}
      onKeyDown={(e) => {
        const step = (e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0) * (getComputedStyle(e.currentTarget).direction === 'rtl' ? -1 : 1);
        const to = e.key === 'Home' ? 0 : e.key === 'End' ? tabs.length - 1 : step ? (at + step + tabs.length) % tabs.length : -1;
        if (to < 0) return;
        e.preventDefault();
        onChange(tabs[to].id);
        refs.current[to]?.focus();
      }}
    >
      {tabs.map((t, i) => (
        <button
          key={t.id}
          ref={(el) => {
            refs.current[i] = el;
          }}
          role="tab"
          id={`${id}-${t.id}`}
          aria-controls={`${id}-panel`}
          aria-selected={i === at}
          tabIndex={i === at ? 0 : -1}
          className={i === at ? 'wbe-active' : ''}
          onClick={() => onChange(t.id)}
        >
          {t.label}
        </button>
      ))}
    </div>
  );
  return (
    <>
      {head ? head(list) : list}
      <div className="wbe-tab-panel" role="tabpanel" id={`${id}-panel`} aria-labelledby={`${id}-${value}`}>
        {children}
      </div>
    </>
  );
}
