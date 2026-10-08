import { useId, useRef, useState } from 'react';
import { MAX_SHEET_NAME } from '../core/sheetRules';
import type { WorkbookStore } from '../state/store';
import { ContextMenu, type MenuItem } from './ContextMenu';
import { useDialog } from './useDialog';
import { useT } from '../i18n/react';

/**
 * Runs a change to the sheets, then moves the focus to the tab it concerns,
 * from whether it was made.
 */
export type AfterSheetChange = (change: Promise<boolean>, focus: (made: boolean) => number) => void;

/** The sheet's position in the workbook now, by its name; -1 once it has gone. */
export const sheetNamed = (store: WorkbookStore, name: string) => () => store.sheets.findIndex((s) => s.info.name === name);

/**
 * The sheets' visible positions, and the step through the workbook that
 * goes left (-1, or 1 where the editor is mirrored) or right.
 */
export function sheetSteps(store: WorkbookStore): { visible: number[]; left: -1 | 1; right: -1 | 1 } {
  const rtl = store.t.dir === 'rtl';
  return { visible: store.sheets.flatMap((s, i) => (s.info.state === 'visible' ? [i] : [])), left: rtl ? 1 : -1, right: rtl ? -1 : 1 };
}

/** A sheet tab's menu: insert, rename, delete, move, hide and unhide. */
export function SheetMenu({ store, after }: { store: WorkbookStore; after: AfterSheetChange }) {
  const t = useT();
  const menu = store.sheetMenu;
  const view = menu ? store.sheets[menu.sheet] : undefined;
  if (!menu || !view) return null;
  const { sheet } = menu;
  const { visible, left, right } = sheetSteps(store);
  const at = visible.indexOf(sheet);
  const move = (step: -1 | 1) => after(store.moveSheetBy(sheet, step), sheetNamed(store, view.info.name));
  const items: MenuItem[] = [
    // Excel inserts the new sheet before the one whose tab was chosen.
    { label: t('sheets.menu.insert'), onSelect: () => after(store.addSheet(sheet), () => store.active) },
    { label: t('sheets.menu.rename'), onSelect: () => store.set('renameSheetFor', sheet) },
    { label: t('sheets.menu.delete'), danger: true, onSelect: () => after(store.deleteSheet(sheet), (made) => (made ? store.active : sheet)) },
    { separator: true },
    { label: t('sheets.menu.moveLeft'), disabled: visible[at + left] === undefined, onSelect: () => move(left) },
    { label: t('sheets.menu.moveRight'), disabled: visible[at + right] === undefined, onSelect: () => move(right) },
    { separator: true },
    { label: t('sheets.menu.hide'), onSelect: () => after(store.hideSheet(sheet), (made) => (made ? store.active : sheet)) },
    { label: t('sheets.menu.unhide'), disabled: !store.hiddenSheets.length, onSelect: () => store.set('unhideOpen', true) },
  ];
  return <ContextMenu x={menu.x} y={menu.y} items={items} label={t('sheets.menu.label', { sheet: view.info.name })} onClose={() => store.set('sheetMenu', undefined)} />;
}

/** Renames a sheet, saying as the name is typed why Excel wouldn't take it. */
export function RenameSheetDialog({ store, after }: { store: WorkbookStore; after: AfterSheetChange }) {
  const t = useT();
  const sheet = store.renameSheetFor;
  const view = sheet !== undefined ? store.sheets[sheet] : undefined;
  const dialog = useRef<HTMLFormElement>(null);
  const input = useRef<HTMLInputElement>(null);
  const id = useId();
  const [text, setText] = useState('');
  const opened = useRef<number | undefined>(undefined);
  if (opened.current !== sheet) {
    opened.current = sheet;
    if (view) setText(view.info.name);
  }
  const close = () => store.set('renameSheetFor', undefined);
  useDialog(dialog, !!view, close, { modal: true, initial: input });
  if (sheet === undefined || !view) return null;
  const problem = store.sheetProblem({ type: 'renameSheet', sheet, name: text });
  const apply = () => {
    // The dialog stays open, saying why, until the name is one Excel takes.
    if (problem) {
      input.current?.focus();
      return;
    }
    close();
    after(store.renameSheet(sheet, text), () => sheet);
  };
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
          <h2 id={`${id}-title`}>{t('sheets.rename.title', { sheet: view.info.name })}</h2>
          <label className="wbe-field">
            <span>{t('sheets.rename.label')}</span>
            <input
              ref={input}
              value={text}
              spellCheck={false}
              onChange={(e) => setText(e.target.value)}
              onFocus={(e) => e.target.select()}
              aria-invalid={!!problem}
              aria-describedby={problem ? `${id}-error ${id}-hint` : `${id}-hint`}
            />
          </label>
          <p id={`${id}-error`} className="wbe-field-error" aria-live="polite">
            {problem ? t.text(problem) : ''}
          </p>
          <p id={`${id}-hint`} className="wbe-muted wbe-small">
            {t('sheets.rename.hint', { max: MAX_SHEET_NAME })}
          </p>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary">
              {t('sheets.rename.save')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}

/** Shows a hidden sheet again; sheets hidden so that only Excel's Visual Basic editor can show them aren't offered. */
export function UnhideSheetDialog({ store, after }: { store: WorkbookStore; after: AfterSheetChange }) {
  const t = useT();
  const hidden = store.hiddenSheets;
  const open = store.unhideOpen && hidden.length > 0;
  const dialog = useRef<HTMLFormElement>(null);
  const first = useRef<HTMLInputElement>(null);
  const id = useId();
  const [choice, setChoice] = useState<number | undefined>(undefined);
  const wasOpen = useRef(false);
  if (wasOpen.current !== open) {
    wasOpen.current = open;
    if (open) setChoice(hidden[0]);
  }
  const close = () => store.set('unhideOpen', false);
  useDialog(dialog, open, close, { modal: true, initial: first });
  if (!open) return null;
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
          const sheet = choice ?? hidden[0];
          close();
          after(store.unhideSheet(sheet), () => store.active);
        }}
      >
        <div className="wbe-modal-content">
          <h2 id={`${id}-title`}>{t('sheets.unhide.title')}</h2>
          <fieldset className="wbe-choices">
            <legend>{t('sheets.unhide.label')}</legend>
            {hidden.map((i, k) => (
              <label key={i} className="wbe-choice">
                <input ref={k === 0 ? first : undefined} type="radio" name={`${id}-sheet`} checked={choice === i} onChange={() => setChoice(i)} />
                <span>{store.sheets[i].info.name}</span>
              </label>
            ))}
          </fieldset>
          <div className="wbe-modal-actions">
            <button type="button" className="wbe-btn" onClick={close}>
              {t('dialog.cancel')}
            </button>
            <button type="submit" className="wbe-btn wbe-primary">
              {t('sheets.unhide.confirm')}
            </button>
          </div>
        </div>
      </form>
    </div>
  );
}
