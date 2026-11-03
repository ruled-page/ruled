/**
 * The editor's keyboard shortcuts as data, for the list Ctrl+/ shows and for
 * the documentation: by group, each with its keys and what it does, as the
 * language writes them, for Apple devices or others (by default, the one it runs on).
 */
import type { Translator } from '../i18n';
import { APPLE, mod, redoKey, shortcutsKey } from './keys';
import type { MenuCommandId } from './menuCommands';

/** A shortcut's keys, what it does, and the menu entries it stands for, if any. */
type Shortcut = [keys: string, does: string, ...commands: MenuCommandId[]];

/**
 * The shortcuts by group, with keys and what they do as the language writes
 * them; those standing for menu entries only while `on` has one of them on.
 */
export function shortcuts(t: Translator, apple = APPLE, on: (command: MenuCommandId) => boolean = () => true): { group: string; keys: [string, string][] }[] {
  const m = (...keys: string[]) => keys.map((k) => mod(t, k, apple)).join(', ');
  const groups: { group: string; keys: Shortcut[] }[] = [
    {
      group: t('shortcuts.moving'),
      keys: [
        [t('keys.arrows'), t('shortcuts.moveCells')],
        [t('keys.tab'), t('shortcuts.moveRightLeft')],
        [t('keys.shiftArrows'), t('shortcuts.extend')],
        [apple ? '⌘↑, ⌘↓, ⌘←, ⌘→' : t('keys.ctrlArrows'), t('shortcuts.dataEdge')],
        [apple ? 'fn←, ⌘fn←, ⌘fn→' : t('keys.home'), t('shortcuts.home')],
        [apple ? 'fn↑, fn↓, ⌥fn↑, ⌥fn↓' : t('keys.pages'), t('shortcuts.pages')],
        [apple ? '⇧⌫, ⌘⌫' : t('keys.backspaces'), t('shortcuts.backspaces')],
        [t('keys.f6'), t('shortcuts.regions')],
      ],
    },
    {
      group: t('shortcuts.editing'),
      keys: [
        [t('keys.typing'), t('shortcuts.typing')],
        [t('keys.enter'), t('shortcuts.enter')],
        [t('keys.f2'), t('shortcuts.f2')],
        [t('keys.escape'), t('shortcuts.escape')],
        [t('keys.delete'), t('shortcuts.clear')],
        [`${m('Z')}, ${redoKey(t, apple)}`, t('shortcuts.undoRedo')],
        [m('X', 'C', 'V'), t('shortcuts.clipboard'), 'menu.cut', 'menu.copy', 'menu.paste'],
        [t('keys.accept'), t('shortcuts.complete')],
        [t('keys.f4'), t('shortcuts.absolute')],
        [m('+', '-'), t('shortcuts.insertDelete'), 'menu.insertCells', 'menu.deleteCells', 'menu.insertRowsAbove', 'menu.deleteRows', 'menu.insertColumnsLeft', 'menu.deleteColumns'],
        [apple ? '⇧⌘F' : t('keys.filterToggle'), t('shortcuts.filterToggle'), 'menu.filterOn'],
        [apple ? '⌥↓' : t('keys.altDown'), t('shortcuts.filterMenu')],
        [apple ? '⌥F1' : t('keys.altF1'), t('shortcuts.insertChart'), 'core.chartType.column'],
        [apple ? '⌃⌥5' : t('keys.ctrlAlt5'), t('shortcuts.selectObjects')],
        [apple ? '⌥F5' : t('keys.altF5'), t('shortcuts.refreshPivots'), 'pivot.menu.refresh'],
        [apple ? '⌃⌥F5' : t('keys.ctrlAltF5'), t('shortcuts.refreshAllPivots'), 'pivot.menu.refreshAll'],
      ],
    },
    {
      group: t('shortcuts.commands'),
      keys: [
        [t('keys.menu'), t('shortcuts.menu')],
        [t('keys.shiftF2'), t('shortcuts.note'), 'menu.newNote', 'menu.editNote', 'menu.showComments'],
        [mod(t, t('keys.enter'), apple), t('links.shortcut'), 'links.menu.open'],
        [apple ? '⌥⇧→' : t('outline.keys.group'), t('outline.shortcuts.group'), 'outline.menu.groupRows', 'outline.menu.groupColumns'],
        [apple ? '⌥⇧←' : t('outline.keys.ungroup'), t('outline.shortcuts.ungroup'), 'outline.menu.ungroupRows', 'outline.menu.ungroupColumns'],
        [m('F'), t('shortcuts.find'), 'menu.find'],
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
  return groups.map(({ group, keys }) => ({ group, keys: keys.filter(([, , ...ids]) => !ids.length || ids.some(on)).map(([k, does]): [string, string] => [k, does]) }));
}
