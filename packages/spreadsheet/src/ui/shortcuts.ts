/**
 * The editor's keyboard shortcuts as data, for the list Ctrl+/ shows and for
 * the documentation: by group, each with its keys and what it does, as the
 * language writes them, for Apple devices or others (by default, the one it runs on).
 */
import type { Translator } from '../i18n';
import { APPLE, mod, redoKey, shortcutsKey } from './keys';

/** The shortcuts by group, with keys and what they do as the language writes them. */
export function shortcuts(t: Translator, apple = APPLE): { group: string; keys: [string, string][] }[] {
  const m = (...keys: string[]) => keys.map((k) => mod(t, k, apple)).join(', ');
  return [
    {
      group: t('shortcuts.moving'),
      keys: [
        [t('keys.arrows'), t('shortcuts.moveCells')],
        [t('keys.tab'), t('shortcuts.moveRightLeft')],
        [t('keys.shiftArrows'), t('shortcuts.extend')],
        [t('keys.f6'), t('shortcuts.regions')],
      ],
    },
    {
      group: t('shortcuts.editing'),
      keys: [
        [t('keys.typing'), t('shortcuts.typing')],
        [t('keys.enter'), t('shortcuts.enter')],
        [t('keys.escape'), t('shortcuts.escape')],
        [t('keys.delete'), t('shortcuts.clear')],
        [`${m('Z')}, ${redoKey(t, apple)}`, t('shortcuts.undoRedo')],
        [m('X', 'C', 'V'), t('shortcuts.clipboard')],
        [t('keys.accept'), t('shortcuts.complete')],
        [t('keys.f4'), t('shortcuts.absolute')],
        [m('+', '-'), t('shortcuts.insertDelete')],
        [apple ? '⇧⌘F' : t('keys.filterToggle'), t('shortcuts.filterToggle')],
        [apple ? '⌥↓' : t('keys.altDown'), t('shortcuts.filterMenu')],
        [apple ? '⌥F1' : t('keys.altF1'), t('shortcuts.insertChart')],
        [apple ? '⌃⌥5' : t('keys.ctrlAlt5'), t('shortcuts.selectObjects')],
        [apple ? '⌥F5' : t('keys.altF5'), t('shortcuts.refreshPivots')],
        [apple ? '⌃⌥F5' : t('keys.ctrlAltF5'), t('shortcuts.refreshAllPivots')],
      ],
    },
    {
      group: t('shortcuts.commands'),
      keys: [
        [t('keys.menu'), t('shortcuts.menu')],
        [t('keys.shiftF2'), t('shortcuts.note')],
        [mod(t, t('keys.enter'), apple), t('links.shortcut')],
        [apple ? '⌥⇧→' : t('outline.keys.group'), t('outline.shortcuts.group')],
        [apple ? '⌥⇧←' : t('outline.keys.ungroup'), t('outline.shortcuts.ungroup')],
        [m('F'), t('shortcuts.find')],
        [m('G'), t('shortcuts.findNext')],
        [m('S'), t('shortcuts.save')],
        [shortcutsKey(t, apple), t('shortcuts.list')],
      ],
    },
    {
      group: t('sheets.shortcuts.group'),
      keys: [
        [t('keys.menu'), t('sheets.shortcuts.menu')],
        [apple ? '⌥⇧←, ⌥⇧→' : t('sheets.keys.move'), t('sheets.shortcuts.move')],
        [t('sheets.keys.rename'), t('sheets.shortcuts.rename')],
      ],
    },
  ];
}
