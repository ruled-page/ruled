import { useEffect, useLayoutEffect, useRef, useState, type KeyboardEvent as ReactKeyboardEvent, type ReactNode } from 'react';

export type MenuItem =
  /** An item; `checked` makes it one of a set of choices, with whether it's the one chosen, and `sample` shows what it gives. */
  | { label: string; onSelect: () => void; disabled?: boolean; hint?: string; danger?: boolean; checked?: boolean; sample?: ReactNode; title?: string }
  /** A submenu, opened by hovering it, or with the arrow key towards it, Enter or Space. */
  | { label: string; items: MenuItem[]; disabled?: boolean }
  | { separator: true };

const enabled = (menu: HTMLElement | null) => [...(menu?.querySelectorAll<HTMLButtonElement>(':scope > button:not(:disabled)') ?? [])];
const rtl = (el: Element) => getComputedStyle(el).direction === 'rtl';

/**
 * A right-click menu at a point, closed by choosing an item, Escape, Tab, or
 * clicking or scrolling elsewhere. Its items are reached with the arrow keys,
 * Home and End, as ARIA's menu pattern has it, not with Tab.
 */
export function ContextMenu({ x, y, items, onClose, label }: { x: number; y: number; items: MenuItem[]; onClose: () => void; label: string }) {
  const root = useRef<HTMLDivElement>(null);
  const menu = useRef<HTMLDivElement>(null);
  const sub = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ left: x, top: y });
  /** The open submenu: its item's index, where it shows, and whether the keyboard opened it (so it takes the focus). */
  const [open, setOpen] = useState<{ index: number; left: number; top: number; focus: boolean } | null>(null);
  // Give focus back to what had it (the grid), so its shortcuts keep working.
  const opener = useRef(document.activeElement as HTMLElement | null);
  // A layout effect, so a picker or panel the chosen item opens can take focus after.
  useLayoutEffect(() => () => opener.current?.focus({ preventScroll: true }), []);

  // Keep the menu on screen.
  useLayoutEffect(() => {
    const el = menu.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    setPos({ left: Math.max(4, Math.min(x, window.innerWidth - r.width - 4)), top: Math.max(4, Math.min(y, window.innerHeight - r.height - 4)) });
    enabled(el)[0]?.focus();
  }, [x, y]);

  // A submenu opens beside its item, on whichever side has room, and takes the focus when the keyboard opened it.
  useLayoutEffect(() => {
    const el = sub.current;
    if (!el || !open) return;
    const r = el.getBoundingClientRect();
    const item = enabled(menu.current).find((b) => Number(b.dataset.index) === open.index)?.getBoundingClientRect();
    if (item) {
      const right = item.right + r.width + 4 <= window.innerWidth;
      const left = rtl(el) || !right ? Math.max(4, item.left - r.width) : item.right;
      const top = Math.max(4, Math.min(item.top - 4, window.innerHeight - r.height - 4));
      if (left !== open.left || top !== open.top) setOpen({ ...open, left, top });
    }
    if (open.focus) enabled(el)[0]?.focus();
  }, [open]);

  useEffect(() => {
    const away = (e: Event) => {
      if (!root.current?.contains(e.target as Node)) onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !sub.current?.contains(document.activeElement)) {
        e.preventDefault();
        onClose();
      } else if (e.key === 'Tab') {
        e.preventDefault();
        onClose();
      }
    };
    window.addEventListener('mousedown', away, true);
    window.addEventListener('wheel', onClose, { passive: true });
    window.addEventListener('keydown', key, true);
    window.addEventListener('blur', onClose);
    return () => {
      window.removeEventListener('mousedown', away, true);
      window.removeEventListener('wheel', onClose);
      window.removeEventListener('keydown', key, true);
      window.removeEventListener('blur', onClose);
    };
  }, [onClose]);

  /** Arrow keys, Home and End within one menu. */
  const move = (e: ReactKeyboardEvent<HTMLDivElement>): boolean => {
    const buttons = enabled(e.currentTarget);
    const i = buttons.indexOf(document.activeElement as HTMLButtonElement);
    const to =
      e.key === 'ArrowDown' ? (i + 1) % buttons.length : e.key === 'ArrowUp' ? (i + buttons.length - 1) % buttons.length : e.key === 'Home' ? 0 : e.key === 'End' ? buttons.length - 1 : -1;
    if (to < 0) return false;
    e.preventDefault();
    e.stopPropagation();
    buttons[to]?.focus();
    return true;
  };

  const openSub = (index: number, focus: boolean) => setOpen((o) => (o?.index === index && !focus ? o : { index, left: pos.left, top: pos.top, focus }));

  const onMenuKey = (e: ReactKeyboardEvent<HTMLDivElement>) => {
    if (move(e)) return;
    const index = Number((document.activeElement as HTMLElement | null)?.dataset.index);
    const item = items[index];
    const toward = rtl(e.currentTarget) ? 'ArrowLeft' : 'ArrowRight';
    if (item && 'items' in item && (e.key === toward || e.key === 'Enter' || e.key === ' ')) {
      e.preventDefault();
      openSub(index, true);
    }
  };

  const onSubKey = (e: ReactKeyboardEvent<HTMLDivElement>) => {
    if (move(e)) return;
    const back = rtl(e.currentTarget) ? 'ArrowRight' : 'ArrowLeft';
    if (e.key === back || e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      const index = open?.index;
      setOpen(null);
      enabled(menu.current).find((b) => Number(b.dataset.index) === index)?.focus();
    }
  };

  const choose = (item: Extract<MenuItem, { onSelect: () => void }>) => {
    onClose();
    item.onSelect();
  };

  // The first item is the menu's one tab stop, so a menu tall enough to scroll can be reached from the keyboard.
  const render = (list: MenuItem[], top: boolean, first = list.findIndex((item) => !('separator' in item) && !item.disabled)) =>
    list.map((item, i) =>
      'separator' in item ? (
        // oxlint-disable-next-line jsx-a11y/control-has-associated-label -- a separator, not a control
        <div key={i} className="wbe-menu-separator" role="separator" />
      ) : 'items' in item ? (
        <button
          key={i}
          role="menuitem"
          tabIndex={i === first ? 0 : -1}
          data-index={i}
          aria-haspopup="menu"
          aria-expanded={open?.index === i}
          disabled={item.disabled}
          onClick={() => openSub(i, true)}
          onMouseEnter={() => openSub(i, false)}
        >
          <span>{item.label}</span>
          <span className="wbe-menu-arrow" aria-hidden="true">
            ›
          </span>
        </button>
      ) : (
        <button
          key={i}
          role={item.checked === undefined ? 'menuitem' : 'menuitemradio'}
          aria-checked={item.checked}
          tabIndex={i === first ? 0 : -1}
          data-index={i}
          className={item.danger ? 'wbe-danger' : undefined}
          disabled={item.disabled}
          title={item.title}
          onClick={() => choose(item)}
          onMouseEnter={top ? () => setOpen(null) : undefined}
        >
          {item.checked !== undefined && (
            <span className="wbe-menu-check" aria-hidden="true">
              {item.checked ? '✓' : ''}
            </span>
          )}
          {item.sample}
          <span>{item.label}</span>
          {item.hint && <kbd>{item.hint}</kbd>}
        </button>
      ),
    );

  const opened = open ? items[open.index] : undefined;
  return (
    // No box of its own: a menu opened from a toolbar button would otherwise be a flex item there, its gap moving the buttons.
    <div ref={root} style={{ display: 'contents' }} onContextMenu={(e) => e.preventDefault()}>
      <div className="wbe-context-menu" role="menu" aria-label={label} ref={menu} style={pos} onKeyDown={onMenuKey}>
        {render(items, true)}
      </div>
      {opened && 'items' in opened && (
        <div className="wbe-context-menu" role="menu" aria-label={opened.label} ref={sub} style={{ left: open!.left, top: open!.top }} onKeyDown={onSubKey}>
          {render(opened.items, false)}
        </div>
      )}
    </div>
  );
}
