import type { GridCell, ProvideEditorComponent, TextCell } from '@glideapps/glide-data-grid';
import { useEffect, useId, useLayoutEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { toA1 } from '@sitepen/formula-engine/refs';
import type { WorkbookStore } from '../state/store';
import { useT } from '../i18n/react';
import { AssistPopup, assistKey, FormulaText, HintAnnouncer, pickFunction, useFormulaAssist } from './FormulaAssist';
import { GLIDE_CLICK_OUTSIDE_IGNORE, PORTAL_ID } from './glideDom';

/**
 * The editor for text cells, in place of Glide's own and working as it does
 * (a text area that grows with its text; Shift+Enter starts a new line),
 * with the formula bar's help while a formula is typed: functions offered as
 * their names are typed, argument hints, references coloured here and
 * outlined on the grid, and F4.
 */
export function textCellEditor(store: WorkbookStore): ProvideEditorComponent<GridCell> {
  return (props) => <CellEntry store={store} {...props} />;
}

type EditorProps = Parameters<ProvideEditorComponent<GridCell>>[0];

function CellEntry({ store, value, onChange, isHighlighted, validatedSelection }: EditorProps & { store: WorkbookStore }) {
  const t = useT();
  const cell = value as TextCell;
  const text = cell.data;
  const editable = cell.readonly !== true;
  const area = useRef<HTMLTextAreaElement>(null);
  const box = useRef<HTMLDivElement>(null);
  const anchor = useRef<HTMLSpanElement>(null);
  const id = useId();
  const [sel, setSel] = useState<[number, number]>(() => [isHighlighted ? 0 : text.length, text.length]);
  const [popup, setPopup] = useState<{ left: number; top: number }>();
  const pendingSel = useRef<[number, number]>(undefined);
  const assist = useFormulaAssist(text, sel[0], sel[1], editable);
  const colored = editable && text.startsWith('=');
  const { row, col } = store.selection;
  const apply = (next: string, start: number, end: number) => {
    onChange({ ...cell, data: next });
    setSel([start, end]);
    pendingSel.current = [start, end];
  };

  // Opens focused, with the text selected or the caret at its end, as Glide's own editor does.
  useEffect(() => {
    const el = area.current;
    if (!el || el.disabled) return;
    el.focus();
    el.setSelectionRange(isHighlighted ? 0 : el.value.length, el.value.length);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  useLayoutEffect(() => {
    if (validatedSelection === undefined) return;
    const [start, end] = typeof validatedSelection === 'number' ? [validatedSelection, null] : validatedSelection;
    area.current?.setSelectionRange(start, end);
  }, [validatedSelection]);

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
    <div className={`wbe-cell-entry ${colored ? 'wbe-colored' : ''}`} ref={box}>
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
          onChange({ ...cell, data: e.target.value });
          setSel([start, e.target.selectionEnd]);
          assist.typed(e.target.value, start);
        }}
        onSelect={(e) => setSel([e.currentTarget.selectionStart, e.currentTarget.selectionEnd])}
        onKeyDown={(e) => {
          const el = e.currentTarget;
          // Stopped here, the list's keys don't also reach the grid, which would finish editing.
          if (assistKey(e, assist, text, el.selectionStart, el.selectionEnd, apply)) {
            e.preventDefault();
            e.stopPropagation();
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
            className={`wbe-fx-floating ${GLIDE_CLICK_OUTSIDE_IGNORE}`}
            style={popup}
            onPick={(i) => pickFunction(assist, text, i, apply)}
          />,
          portal,
        )}
      <HintAnnouncer hint={assist.hint} />
    </div>
  );
}
