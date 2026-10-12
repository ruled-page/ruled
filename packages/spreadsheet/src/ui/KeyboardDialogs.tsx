import { useId, useRef, useState } from 'react';
import { colToLetters } from '@sitepen/formula-engine/refs';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { APPLE, mod, redoKey, shortcutsKey } from './keys';
import { useDialog } from './useDialog';
import type { Translator } from '../i18n';
import { useT } from '../i18n/react';

/** The shortcuts by group, with keys and what they do as the language writes them. */
function shortcuts(t: Translator): { group: string; keys: [string, string][] }[] {
  const m = (...keys: string[]) => keys.map((k) => mod(t, k)).join(', ');
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
        [`${m('Z')}, ${redoKey(t)}`, t('shortcuts.undoRedo')],
        [m('X', 'C', 'V'), t('shortcuts.clipboard')],
        [t('keys.accept'), t('shortcuts.complete')],
        [t('keys.f4'), t('shortcuts.absolute')],
      ],
    },
    {
      group: t('shortcuts.commands'),
      keys: [
        [t('keys.menu'), t('shortcuts.menu')],
        [t('keys.shiftF2'), t('shortcuts.note')],
        [mod(t, t('keys.enter')), t('links.shortcut')],
        [APPLE ? '⌥⇧→' : t('outline.keys.group'), t('outline.shortcuts.group')],
        [APPLE ? '⌥⇧←' : t('outline.keys.ungroup'), t('outline.shortcuts.ungroup')],
        [m('F'), t('shortcuts.find')],
        [m('G'), t('shortcuts.findNext')],
        [m('S'), t('shortcuts.save')],
        [shortcutsKey(t), t('shortcuts.list')],
      ],
    },
    {
      group: t('sheets.shortcuts.group'),
      keys: [
        [t('keys.menu'), t('sheets.shortcuts.menu')],
        [APPLE ? '⌥⇧←, ⌥⇧→' : t('sheets.keys.move'), t('sheets.shortcuts.move')],
        [t('sheets.keys.rename'), t('sheets.shortcuts.rename')],
      ],
    },
  ];
}

/** Every keyboard shortcut, since the grid's keys differ from a web page's. */
export function ShortcutsDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const dialog = useRef<HTMLDivElement>(null);
  const id = useId();
  const close = () => store.set('shortcutsOpen', false);
  useDialog(dialog, store.shortcutsOpen, close, { modal: true });
  if (!store.shortcutsOpen) return null;
  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <div ref={dialog} className="wbe-modal wbe-modal-wide wbe-shortcuts" role="dialog" aria-modal="true" aria-labelledby={`${id}-title`} tabIndex={-1} onMouseDown={(e) => e.stopPropagation()}>
        <div className="wbe-modal-content">
          <div className="wbe-shortcuts-head">
            <h2 id={`${id}-title`}>{t('shortcuts.title')}</h2>
            <button className="wbe-btn wbe-icon wbe-ghost" onClick={close} aria-label={t('shortcuts.close')}>
              <Icon name="close" />
            </button>
          </div>
          {shortcuts(t).map((g) => (
            <table key={g.group}>
              <caption>{g.group}</caption>
              <tbody>
                {g.keys.map(([keys, what]) => (
                  <tr key={keys}>
                    <th scope="row">
                      <kbd>{keys}</kbd>
                    </th>
                    <td>{what}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          ))}
        </div>
      </div>
    </div>
  );
}

/** Sets a column's width, or a row's height, from the keyboard, as dragging its edge does. */
export function SizeDialog({ store, axis }: { store: WorkbookStore; axis: 'cols' | 'rows' }) {
  const t = useT();
  const cols = axis === 'cols';
  const line = cols ? store.columnWidthFor : store.rowHeightFor;
  const dialog = useRef<HTMLFormElement>(null);
  const input = useRef<HTMLInputElement>(null);
  const id = useId();
  const [text, setText] = useState('');
  const open = line !== undefined;
  const opened = useRef<number | undefined>(undefined);
  if (opened.current !== line) {
    opened.current = line;
    if (open) setText(t.number(cols ? store.columnWidth(line) : store.rowHeightPt(line), { useGrouping: false, maximumFractionDigits: 2 }));
  }
  const close = () => store.set(cols ? 'columnWidthFor' : 'rowHeightFor', undefined);
  useDialog(dialog, open, close, { modal: true, initial: input });
  if (!open) return null;
  const size = Number(text.replace(',', '.'));
  const valid = text.trim() !== '' && size >= 0 && size <= (cols ? 255 : 409);
  const apply = () => {
    if (!valid) return;
    void (cols ? store.setColumnWidth(line, size) : store.setRowHeight(line, size));
    close();
  };
  const k = cols ? 'columnWidth' : 'rowHeight';
  return (
    <div className="wbe-modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="wbe-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        tabIndex={-1}
        onMouseDown={(e) => e.stopPropagation()}
        onSubmit={(e) => {
          e.preventDefault();
          apply();
        }}
      >
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{cols ? t('columnWidth.title', { column: colToLetters(line) }) : t('rowHeight.title', { row: line + 1 })}</h2>
          <label className="wbe-field">
            <span>{t(`${k}.label`)}</span>
            <input
              ref={input}
              inputMode="decimal"
              value={text}
              onChange={(e) => setText(e.target.value)}
              onFocus={(e) => e.target.select()}
              aria-invalid={!valid}
              aria-describedby={`${id}-note`}
            />
          </label>
          <p id={`${id}-note`} className="wbe-muted wbe-small">
            {t(`${k}.note`)}
          </p>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary" disabled={!valid}>
              {t(`${k}.set`)}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}
