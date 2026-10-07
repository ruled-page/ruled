import { useEffect, useLayoutEffect, useRef, useState } from 'react';

export type MenuItem =
  | { label: string; onSelect: () => void; disabled?: boolean; hint?: string; danger?: boolean }
  | { separator: true };

/** A right-click menu at a point, closed by choosing an item, Escape, or clicking or scrolling elsewhere. */
export function ContextMenu({ x, y, items, onClose }: { x: number; y: number; items: MenuItem[]; onClose: () => void }) {
  const ref = useRef<HTMLDivElement>(null);
  const [pos, setPos] = useState({ left: x, top: y });
  // Give focus back to what had it (the grid), so its shortcuts keep working.
  const opener = useRef(document.activeElement as HTMLElement | null);
  // A layout effect, so a picker or panel the chosen item opens can take focus after.
  useLayoutEffect(() => () => opener.current?.focus({ preventScroll: true }), []);

  // Keep the menu on screen.
  useLayoutEffect(() => {
    const el = ref.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    setPos({ left: Math.max(4, Math.min(x, window.innerWidth - r.width - 4)), top: Math.max(4, Math.min(y, window.innerHeight - r.height - 4)) });
    el.querySelector<HTMLButtonElement>('button:not(:disabled)')?.focus();
  }, [x, y]);

  useEffect(() => {
    const away = (e: Event) => {
      if (!ref.current?.contains(e.target as Node)) onClose();
    };
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        onClose();
      }
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        const buttons = [...(ref.current?.querySelectorAll<HTMLButtonElement>('button:not(:disabled)') ?? [])];
        const i = buttons.indexOf(document.activeElement as HTMLButtonElement);
        buttons[(i + (e.key === 'ArrowDown' ? 1 : buttons.length - 1)) % buttons.length]?.focus();
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

  return (
    <div className="context-menu" role="menu" ref={ref} style={pos} onContextMenu={(e) => e.preventDefault()}>
      {items.map((item, i) =>
        'separator' in item ? (
          <div key={i} className="menu-separator" role="separator" />
        ) : (
          <button
            key={i}
            role="menuitem"
            className={item.danger ? 'danger' : undefined}
            disabled={item.disabled}
            onClick={() => {
              onClose();
              item.onSelect();
            }}
          >
            <span>{item.label}</span>
            {item.hint && <kbd>{item.hint}</kbd>}
          </button>
        ),
      )}
    </div>
  );
}
