import { useEffect, useId, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { toA1 } from '@ruled/formula/refs';
import { PORTAL_ID, type CellEditorProps } from '../grid/contract';
import type { WorkbookStore } from '../state/store';
import { useT } from '../i18n/react';
import { AssistPopup, assistKey, FormulaText, HintAnnouncer, pickFunction, useAutoComplete, useFormulaAssist } from './FormulaAssist';

/**
 * The editor for text cells: a text area that grows with its text, where
 * Shift+Enter starts a new line, with the formula bar's help while a formula
 * is typed: functions offered as their names are typed, argument hints,
 * references coloured here and outlined on the grid, and F4. Text is
 * finished from the column as it is typed, as Excel's AutoComplete does.
 */
export function CellEntry({ store, text, onText, finish, typed, selectAll, readOnly, wrap, box: target, popupClass }: CellEditorProps & { store: WorkbookStore }) {
  const t = useT();
  const editable = !readOnly;
  const area = useRef<HTMLTextAreaElement>(null);
  const box = useRef<HTMLDivElement>(null);
  const anchor = useRef<HTMLSpanElement>(null);
  const id = useId();
  const [sel, setSel] = useState<[number, number]>(() => [selectAll ? 0 : text.length, text.length]);
  const [popup, setPopup] = useState<{ left: number; top: number }>();
  const pendingSel = useRef<[number, number]>(undefined);
  const assist = useFormulaAssist(store, text, sel[0], sel[1], editable);
  const colored = editable && text.startsWith('=');
  const { row, col } = store.selection;
  const auto = useAutoComplete(store, row, col, editable && store.features.autoComplete);
  const apply = (next: string, start: number, end: number) => {
    onText(next);
    setSel([start, end]);
    pendingSel.current = [start, end];
  };

  // Opens focused, with the text selected or the caret at its end.
  // Opened by typing, it holds the key typed, which AutoComplete can already finish.
  useEffect(() => {
    const el = area.current;
    if (!el || el.disabled) return;
    el.focus();
    const completed = typed ? auto.changed(text, text.length, text.length, 'insertText') : undefined;
    if (completed) apply(completed.text, ...completed.sel);
    else el.setSelectionRange(selectAll ? 0 : el.value.length, el.value.length);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // The grid outlines the references of a formula typed here.
  useEffect(() => store.showFormulaRefs(text), [store, text]);
  useEffect(() => () => store.showFormulaRefs(undefined), [store]);

  // Once drawn: the caret where the help put it, and the popup under the name it is about.
  useLayoutEffect(() => {
    if (pendingSel.current) {
      area.current?.setSelectionRange(...pendingSel.current);
      pendingSel.current = undefined;
    }
    const at = anchor.current?.getBoundingClientRect();
    const under = box.current?.getBoundingClientRect();
    const next = at && under ? { left: Math.round(Math.max(0, Math.min(at.left, window.innerWidth - 320))), top: Math.round(under.bottom + 6) } : undefined;
    setPopup((prev) => (prev?.left === next?.left && prev?.top === next?.top ? prev : next));
  }, [text, sel, assist.anchor]);

  const portal = document.getElementById(PORTAL_ID);
  return (
    <div
      className={`wbe-cell-entry ${colored ? 'wbe-colored' : ''}`}
      ref={box}
      // A wrapped cell is edited at its column's width, growing down as Excel's does, less the editor's padding.
      style={wrap ? { width: Math.max(16, target.width - 17) } : undefined}
    >
      <textarea
        ref={area}
        className="wbe-cell-entry-input"
        value={text}
        disabled={!editable}
        dir="auto"
        spellCheck={!colored}
        aria-label={t('formulaBar.cell', { ref: toA1(row, col) })}
        aria-autocomplete="list"
        aria-controls={assist.completion ? `${id}-list` : undefined}
        aria-activedescendant={assist.completion ? `${id}-list-${assist.active}` : undefined}
        aria-describedby={assist.hint && !assist.completion ? `${id}-hint` : undefined}
        onChange={(e) => {
          const start = e.target.selectionStart;
          const completed = auto.changed(e.target.value, start, e.target.selectionEnd, (e.nativeEvent as InputEvent).inputType);
          if (completed) {
            apply(completed.text, ...completed.sel);
          } else {
            onText(e.target.value);
            setSel([start, e.target.selectionEnd]);
          }
          assist.typed(e.target.value, start);
        }}
        onSelect={(e) => setSel([e.currentTarget.selectionStart, e.currentTarget.selectionEnd])}
        onKeyDown={(e) => {
          const el = e.currentTarget;
          // Stopped here, the list's keys don't also reach the grid, which would finish editing.
          if (assistKey(e, assist, text, el.selectionStart, el.selectionEnd, apply)) {
            e.preventDefault();
            e.stopPropagation();
            return;
          }
          // An entry AutoComplete finished goes in written as it is in the column, where the grid would take the text as typed.
          const entry = auto.final(text);
          if (entry !== text && ((e.key === 'Enter' && !e.shiftKey) || e.key === 'Tab')) {
            e.preventDefault();
            e.stopPropagation();
            finish(entry, e.key === 'Tab' ? [e.shiftKey ? -1 : 1, 0] : [0, 1]);
          }
        }}
      />
      <div className="wbe-cell-entry-text" aria-hidden="true">
        {colored ? <FormulaText text={text} refs={assist.refs} anchor={assist.anchor} anchorRef={anchor} /> : text}
        {'\n'}
      </div>
      {portal &&
        popup &&
        createPortal(
          <AssistPopup
            assist={assist}
            listId={`${id}-list`}
            hintId={`${id}-hint`}
            className={`wbe-fx-floating ${popupClass}`}
            style={popup}
            onPick={(i) => pickFunction(assist, text, i, apply)}
          />,
          portal,
        )}
      <HintAnnouncer hint={assist.hint} />
    </div>
  );
}
