import type { Translator } from '../i18n';

/** Apple devices write shortcuts with ⌘; everything else with Ctrl. */
const APPLE = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);

/** A shortcut with the platform's modifier, as the language writes it: ⌘Z on Apple devices, Ctrl+Z (Strg+Z in German) elsewhere. */
export function mod(t: Translator, key: string): string {
  return APPLE ? `⌘${key}` : t('keys.mod', { key });
}
