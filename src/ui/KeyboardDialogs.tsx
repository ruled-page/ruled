import { useId, useRef, useState } from 'react';
import { colToLetters } from '../core/refs';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { mod } from './keys';
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
        [m('G'), t('shortcuts.goTo')],
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
        [m('Z', 'Y'), t('shortcuts.undoRedo')],
        [m('X', 'C', 'V'), t('shortcuts.clipboard')],
        [m('D', 'R'), t('shortcuts.fill')],
      ],
    },
    {
      group: t('shortcuts.commands'),
      keys: [
        [t('keys.menu'), t('shortcuts.menu')],
        [m('F'), t('shortcuts.find')],
        [m('S'), t('shortcuts.save')],
        [m('/'), t('shortcuts.list')],
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
    <div className="modal-backdrop" onMouseDown={close}>
      <div ref={dialog} className="modal modal-wide shortcuts" role="dialog" aria-modal="true" aria-labelledby={`${id}-title`} tabIndex={-1} onMouseDown={(e) => e.stopPropagation()}>
        <div className="modal-content">
          <div className="shortcuts-head">
            <h2 id={`${id}-title`}>{t('shortcuts.title')}</h2>
            <button className="btn icon ghost" onClick={close} aria-label={t('shortcuts.close')}>
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

/** Sets a column's width from the keyboard, as dragging its edge does. */
export function ColumnWidthDialog({ store }: { store: WorkbookStore }) {
  const t = useT();
  const col = store.columnWidthFor;
  const dialog = useRef<HTMLFormElement>(null);
  const input = useRef<HTMLInputElement>(null);
  const id = useId();
  const [text, setText] = useState('');
  const open = col !== undefined;
  const opened = useRef<number | undefined>(undefined);
  if (opened.current !== col) {
    opened.current = col;
    if (open) setText(t.number(store.columnWidth(col), { useGrouping: false, maximumFractionDigits: 2 }));
  }
  const close = () => store.set('columnWidthFor', undefined);
  useDialog(dialog, open, close, { modal: true, initial: input });
  if (!open) return null;
  const chars = Number(text.replace(',', '.'));
  const valid = text.trim() !== '' && chars >= 0 && chars <= 255;
  const apply = () => {
    if (!valid) return;
    store.setColumnWidth(col, chars);
    close();
  };
  return (
    <div className="modal-backdrop" onMouseDown={close}>
      <form
        ref={dialog}
        className="modal"
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
        <div className="modal-content">
          <h2 id={`${id}-title`}>{t('columnWidth.title', { column: colToLetters(col) })}</h2>
          <label className="field">
            <span>{t('columnWidth.label')}</span>
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
          <p id={`${id}-note`} className="muted small">
            {t('columnWidth.note')}
          </p>
          <div className="modal-actions">
            <button type="button" className="btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="btn primary" disabled={!valid}>
              {t('columnWidth.set')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}
