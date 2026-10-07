import type { Translator } from '../i18n';

/** Apple devices write shortcuts with ⌘; everything else with Ctrl. */
const APPLE = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);

/** A shortcut with the platform's modifier, as the language writes it: ⌘Z on Apple devices, Ctrl+Z (Strg+Z in German) elsewhere. */
export function mod(t: Translator, key: string): string {
  return APPLE ? `⌘${key}` : t('keys.mod', { key });
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
