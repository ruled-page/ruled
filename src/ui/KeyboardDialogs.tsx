import { useId, useRef, useState } from 'react';
import { colToLetters } from '../core/refs';
import type { WorkbookStore } from '../state/store';
import { Icon } from './icons';
import { MOD } from './keys';
import { useDialog } from './useDialog';

const SHORTCUTS: { group: string; keys: [string, string][] }[] = [
  {
    group: 'Moving around',
    keys: [
      ['Arrow keys', 'Move between cells'],
      ['Tab, Shift+Tab', 'Move right, left'],
      ['Shift+Arrow keys', 'Select more cells'],
      [`${MOD}G`, 'Go to a cell: type a reference such as B4 or Prices!B4'],
      ['F6, Shift+F6', 'Move to the next or previous part of the editor: toolbar, formula bar, sheet, side panel, sheet tabs'],
    ],
  },
  {
    group: 'Editing',
    keys: [
      ['Typing', "Replaces the cell's contents"],
      ['Enter', 'Edit the cell, or finish editing and move down'],
      ['Escape', 'Cancel editing'],
      ['Delete, Backspace', 'Clear the selected cells'],
      [`${MOD}Z, ${MOD}Y`, 'Undo, redo'],
      [`${MOD}X, ${MOD}C, ${MOD}V`, 'Cut, copy, paste'],
      [`${MOD}D, ${MOD}R`, 'Fill down, fill right'],
    ],
  },
  {
    group: 'Commands',
    keys: [
      ['Shift+F10, Menu key', 'Cell menu: insert and delete rows and columns, column width, sort, filter'],
      [`${MOD}F`, 'Find and replace'],
      [`${MOD}S`, 'Save'],
      [`${MOD}/`, 'This list'],
    ],
  },
];

/** Every keyboard shortcut, since the grid's keys differ from a web page's. */
export function ShortcutsDialog({ store }: { store: WorkbookStore }) {
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
            <h2 id={`${id}-title`}>Keyboard shortcuts</h2>
            <button className="btn icon ghost" onClick={close} aria-label="Close">
              <Icon name="close" />
            </button>
          </div>
          {SHORTCUTS.map((g) => (
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
  const col = store.columnWidthFor;
  const dialog = useRef<HTMLFormElement>(null);
  const input = useRef<HTMLInputElement>(null);
  const id = useId();
  const [text, setText] = useState('');
  const open = col !== undefined;
  const opened = useRef<number | undefined>(undefined);
  if (opened.current !== col) {
    opened.current = col;
    if (open) setText(String(store.columnWidth(col)));
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
          <h2 id={`${id}-title`}>Width of column {colToLetters(col)}</h2>
          <label className="field">
            <span>Width in characters, as Excel measures it (0 to 255)</span>
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
            Changes how wide it is shown here. The saved file keeps the widths it had.
          </p>
          <div className="modal-actions">
            <button type="button" className="btn" onClick={close}>
              Cancel
            </button>
            <button type="submit" className="btn primary" disabled={!valid}>
              Set width
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}
