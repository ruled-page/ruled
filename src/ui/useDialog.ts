import { useEffect, useRef, type RefObject } from 'react';

const FOCUSABLE = [
  'a[href]',
  'button:not([disabled])',
  'input:not([disabled]):not([type="hidden"])',
  'select:not([disabled])',
  'textarea:not([disabled])',
  '[tabindex]:not([tabindex="-1"])',
].join(',');

/** The first control in `el` that takes the keyboard focus. */
export function firstFocusable(el: ParentNode): HTMLElement | null {
  return [...el.querySelectorAll<HTMLElement>(FOCUSABLE)].find((x) => x.getClientRects().length > 0) ?? null;
}

/** Open dialogs, the newest last: only the newest answers Escape and Tab. */
const open: HTMLElement[] = [];

/**
 * Keyboard behaviour for a dialog or popup, while `isOpen`. When it opens, the focus moves
 * into it (to `initial`, else its first control); Escape calls `onClose`; in a
 * modal one, Tab and Shift+Tab go round its controls without leaving it; and
 * when it closes, the focus goes back to what had it before.
 */
export function useDialog(
  ref: RefObject<HTMLElement | null>,
  isOpen: boolean,
  onClose: () => void,
  options: { modal?: boolean; initial?: RefObject<HTMLElement | null> } = {},
): void {
  const close = useRef(onClose);
  close.current = onClose;
  const { modal = false, initial } = options;
  // What had the focus as it opened, noted before any control inside takes it with autoFocus.
  const opener = useRef<HTMLElement | null>(null);
  const wasOpen = useRef(false);
  if (isOpen !== wasOpen.current) {
    wasOpen.current = isOpen;
    if (isOpen) opener.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
  }

  useEffect(() => {
    const el = ref.current;
    if (!isOpen || !el) return;
    // Opened from a menu that has closed since: back to where the menu gave the focus.
    const before = opener.current?.isConnected || el.contains(document.activeElement) ? opener.current : (document.activeElement as HTMLElement | null);
    open.push(el);
    if (!el.contains(document.activeElement)) (initial?.current ?? el.querySelector<HTMLElement>(FOCUSABLE) ?? el).focus({ preventScroll: true });

    const onKey = (e: KeyboardEvent) => {
      if (open[open.length - 1] !== el) return;
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        close.current();
        return;
      }
      if (e.key !== 'Tab' || !modal) return;
      const items = [...el.querySelectorAll<HTMLElement>(FOCUSABLE)].filter((x) => x.getClientRects().length > 0);
      if (!items.length) return;
      const first = items[0];
      const last = items[items.length - 1];
      const inside = el.contains(document.activeElement);
      if (e.shiftKey && (!inside || document.activeElement === first)) {
        e.preventDefault();
        last.focus();
      } else if (!e.shiftKey && (!inside || document.activeElement === last)) {
        e.preventDefault();
        first.focus();
      }
    };
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('keydown', onKey, true);
      open.splice(open.indexOf(el), 1);
      // Back to the opener, unless the user has already moved the focus elsewhere.
      const now = document.activeElement;
      if (before?.isConnected && (!now || now === document.body || el.contains(now))) before.focus({ preventScroll: true });
    };
  }, [ref, isOpen, modal, initial]);
}
