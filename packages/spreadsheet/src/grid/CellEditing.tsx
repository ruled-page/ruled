import { useEffect, useRef, type CSSProperties, type KeyboardEvent as ReactKeyboardEvent } from 'react';
import { createPortal } from 'react-dom';
import type { WorkbookStore } from '../state/store';
import { CellEntry } from '../ui/CellEditor';
import { EnumEditor } from '../ui/EnumEditor';
import type { GridPalette } from '../ui/gridPalette';
import { PORTAL_ID, type Box, type EditMove } from './contract';

/** Popups the cell editors open take this class, so clicks in them don't end editing. */
export const EDITOR_POPUP = 'wbe-editor-popup';

/**
 * A cell being edited. Opened by typing, Excel's Enter mode: the arrow
 * keys finish the entry and move. Opened with Enter, F2 or a double-click,
 * its Edit mode: they move the caret. F2 switches between the two.
 */
export interface Edit {
  row: number;
  col: number;
  text: string;
  typed: boolean;
  selectAll: boolean;
  mode: 'enter' | 'edit';
  /** The values a list cell allows, which it picks from a dropdown. */
  options?: readonly string[];
  wrap: boolean;
}

const ARROWS: Record<string, EditMove> = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] };

interface Props {
  store: WorkbookStore;
  edit: Edit;
  /** The cell's box on the page; undefined while it is scrolled out of view. */
  box: Box | undefined;
  palette: GridPalette;
  /** The sheet reads right to left, where the left arrow goes on to the next column. */
  rtl: boolean;
  onText: (text: string) => void;
  /** Ends editing: entering `text`, or leaving the cell as it was for undefined, then moving. */
  finish: (text: string | undefined, move: EditMove) => void;
  onMode: () => void;
}

/**
 * The cell editor over the cell, in the editor's portal so it can grow past
 * the grid: the text editor with the formula bar's help, or a list cell's
 * dropdown. Enter finishes and moves down (Shift+Enter or Alt+Enter start a
 * new line), Tab across, Escape gives up, and a click elsewhere keeps what
 * was typed.
 */
export function CellEditorBox({ store, edit, box, palette, rtl, onText, finish, onMode }: Props) {
  const wrapper = useRef<HTMLDivElement>(null);
  const latest = useRef({ text: edit.text, finish });
  latest.current = { text: edit.text, finish };
  useEffect(() => {
    const down = (e: MouseEvent) => {
      const target = e.target as Element;
      if (wrapper.current?.contains(target) || target.closest?.(`.${EDITOR_POPUP}`)) return;
      latest.current.finish(edit.options ? undefined : latest.current.text, [0, 0]);
    };
    window.addEventListener('mousedown', down, true);
    return () => window.removeEventListener('mousedown', down, true);
  }, [edit.options]);

  const onKeyDown = (e: ReactKeyboardEvent<HTMLDivElement>) => {
    // The formula help and the dropdown take their own keys first.
    if (e.defaultPrevented) return;
    const mods = e.ctrlKey || e.metaKey || e.altKey;
    if (e.key === 'Escape') finish(undefined, [0, 0]);
    else if (e.key === 'Enter' && e.altKey && !edit.options) {
      const area = e.target as HTMLTextAreaElement;
      area.setRangeText('\n', area.selectionStart, area.selectionEnd, 'end');
      onText(area.value);
    } else if (e.key === 'Enter' && !e.shiftKey && !mods) finish(edit.text, [0, 1]);
    else if (e.key === 'Tab' && !mods) finish(edit.text, [e.shiftKey ? -1 : 1, 0]);
    else if (e.key === 'F2' && !mods && !e.shiftKey) onMode();
    else if (ARROWS[e.key] && edit.mode === 'enter' && !mods && !e.shiftKey && !edit.text.startsWith('=')) {
      const [dx, dy] = ARROWS[e.key];
      finish(edit.text, [rtl ? (-dx as EditMove[0]) : dx, dy]);
    }
    else return;
    e.preventDefault();
    e.stopPropagation();
  };

  const portal = typeof document === 'undefined' ? null : document.getElementById(PORTAL_ID);
  if (!portal || !box) return null;
  const style = {
    left: box.x,
    top: box.y,
    minWidth: box.width,
    minHeight: box.height,
    ...(edit.wrap && { width: box.width }),
    background: palette.bg,
    '--wbe-entry-font': palette.font,
    '--wbe-entry-text': palette.text,
  } as CSSProperties;
  const props = {
    text: edit.text,
    onText,
    finish,
    typed: edit.typed,
    selectAll: edit.selectAll,
    readOnly: false,
    wrap: edit.wrap,
    box,
    options: edit.options,
    popupClass: EDITOR_POPUP,
  };
  return createPortal(
    <div className={`wbe-cell-editor ${edit.options ? 'wbe-list' : ''}`} ref={wrapper} style={style} onKeyDown={onKeyDown}>
      {edit.options ? <EnumEditor {...props} /> : <CellEntry store={store} {...props} />}
    </div>,
    portal,
  );
}
