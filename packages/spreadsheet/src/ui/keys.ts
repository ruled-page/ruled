import type { Translator } from '../i18n';

/** Apple devices write shortcuts with ⌘; everything else with Ctrl. */
export const APPLE = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);

/** A shortcut with the platform's modifier, as the language writes it: ⌘Z on Apple devices, Ctrl+Z (Strg+Z in German) elsewhere. */
export function mod(t: Translator, key: string, apple = APPLE): string {
  return apple ? `⌘${key}` : t('keys.mod', { key });
}

/**
 * The key for the list of shortcuts. On Apple devices ⌃/ as well as ⌘/,
 * because Safari keeps ⌘/ for itself (Show Status Bar) and never passes it on.
 */
export function shortcutsKey(t: Translator, apple = APPLE): string {
  return apple ? '⌘/, ⌃/' : mod(t, '/', apple);
}

/** Redo as the platform writes it: ⇧⌘Z on Apple devices (⌘Y is the browser's history there), Ctrl+Y elsewhere. */
export function redoKey(t: Translator, apple = APPLE): string {
  return apple ? '⇧⌘Z' : mod(t, 'Y', apple);
}

/**
 * Whether a key was pressed with the platform's command key: ⌘ (or ⌃) on
 * Apple devices, Ctrl elsewhere, where the Meta key (Windows, Super) is the
 * system's.
 */
export function commandKey(e: Pick<KeyboardEvent, 'ctrlKey' | 'metaKey'>, apple = APPLE): boolean {
  return e.ctrlKey || (apple && e.metaKey);
}

/**
 * The key of a shortcut such as Ctrl+Z: the letter the key types, so the
 * shortcut follows the keyboard's layout, or where it types no Latin letter
 * (a Cyrillic layout, or a browser that doesn't say) the key in that place on
 * a US keyboard. "/" is matched by place too, as on many layouts it is a
 * shifted key.
 */
export function shortcutKey(e: Pick<KeyboardEvent, 'key' | 'code'>): string {
  if (e.code === 'Slash' || e.key === '/') return '/';
  if (/^[a-z]$/i.test(e.key)) return e.key.toLowerCase();
  if (/^Key[A-Z]$/.test(e.code)) return e.code.slice(3).toLowerCase();
  return e.key.toLowerCase();
}
