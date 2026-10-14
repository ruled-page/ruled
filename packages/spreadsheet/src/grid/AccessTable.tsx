import { inRect } from '@ruled/formula/refs';
import { memo, useLayoutEffect, useRef, type HTMLAttributes, type RefObject } from 'react';
import type { Box } from './contract';
import { effectiveStyle } from '../state/cellStyle';
import { checkboxState } from '../state/checkbox';
import type { SheetView } from '../state/sheetView';
import type { Selection, WorkbookStore } from '../state/store';
import { useT } from '../i18n/react';
import { cellOptions } from './lists';
import { spokenText } from './spoken';

interface Props {
  store: WorkbookStore;
  view: SheetView;
  selection: Selection;
  /** A cell's text as the grid draws it. */
  textOf: RefObject<(row: number, col: number) => string>;
  /** Changes whenever what the cells show may have changed. */
  version: string;
  /** The sheet rows and columns to give, in order: those in view, and the active cell's. */
  rows: number[];
  cols: number[];
  /** Set to the active cell's element, which takes the focus. */
  activeRef: RefObject<HTMLTableCellElement | null>;
  /** How many of the sheet's first rows are frozen: they come first, as a group of their own. */
  frozenRows: number;
  /** The active cell takes typing: a key, a dead key or an input method's composition starts editing it. */
  typing: boolean;
  /** Where the active cell is, in the table's own pixels, for an input method's window to open beside it. */
  activeBox: () => Box | undefined;
  /** Called with the text an input method composed in the active cell. */
  onComposed: (text: string) => void;
}

/**
 * A cell of the table, its text set here rather than by React, so that an
 * input method composing in the active cell can't leave React's text out of
 * step with the page.
 */
function Cell({ text, cellRef, ...props }: HTMLAttributes<HTMLTableCellElement> & { text: string; cellRef?: RefObject<HTMLTableCellElement | null> }) {
  const own = useRef<HTMLTableCellElement>(null);
  useLayoutEffect(() => {
    const el = own.current;
    if (el && el.textContent !== text && !el.classList.contains('wbe-composing')) el.textContent = text;
    if (cellRef) cellRef.current = el;
  });
  return <td ref={own} {...props} />;
}

/**
 * The grid for screen readers: the cells in view, and the active cell
 * wherever it is, numbered as the sheet numbers them and each named by its
 * reference, what it shows and what else is true of it. The active cell
 * has the focus while the grid has it, and moves as the selection does.
 */
export const AccessTable = memo(
  AccessTableNow,
  // Built again only when what it gives changes, not each time the grid renders as it scrolls.
  (a, b) => a.view === b.view && a.selection === b.selection && a.rows === b.rows && a.cols === b.cols && a.version === b.version && a.typing === b.typing && a.frozenRows === b.frozenRows,
);

function AccessTableNow({ store, view, selection, textOf, rows, cols, activeRef, frozenRows, typing, activeBox, onComposed }: Props) {
  const t = useT();
  const selected = (row: number, col: number) => selection.lines === 'all' || selection.rects.some((r) => (selection.lines === 'rows' ? row >= r.r1 && row <= r.r2 : selection.lines === 'cols' ? col >= r.c1 && col <= r.c2 : inRect(r, row, col)));
  /** What a cell shows: its text as drawn (a number shown as #### in full), or whether its checkbox is ticked. */
  const shown = (row: number, col: number) => {
    const rec = view.get(row, col);
    const style = effectiveStyle(view, store.styles, row, col, rec ? rec.style : view.styleAt(row, col, null));
    const box = rec ? checkboxState(rec, style) : style.checkbox ? false : undefined;
    if (box !== undefined) return t(box ? 'checkbox.ticked' : 'checkbox.unticked');
    return textOf.current(row, col);
  };
  // The active cell sets this once rendered; a table without one (its line hidden) leaves it unset.
  activeRef.current = null;
  /** An input method starts composing: the cell shows what it composes, over the grid's own cell. */
  const composing = (e: { currentTarget: HTMLTableCellElement }) => {
    const el = e.currentTarget;
    const box = activeBox();
    el.textContent = '';
    el.classList.add('wbe-composing');
    if (box) Object.assign(el.style, { left: `${box.x}px`, top: `${box.y}px`, minWidth: `${box.width}px`, height: `${box.height}px` });
  };
  /** What it composed goes to the cell editor, and the cell goes back to saying what it holds. */
  const composed = (e: { currentTarget: HTMLTableCellElement; data: string }) => {
    const el = e.currentTarget;
    el.classList.remove('wbe-composing');
    el.removeAttribute('style');
    el.textContent = '';
    if (e.data) onComposed(e.data);
  };
  const rowsOf = (list: number[]) =>
    list.map((row) => (
      <tr key={row} aria-rowindex={row + 1}>
        {cols.map((col) => {
          const active = row === selection.row && col === selection.col;
          const text = shown(row, col);
          return (
            <Cell
              key={col}
              text={text}
              aria-colindex={col + 1}
              aria-selected={selected(row, col)}
              // oxlint-disable-next-line jsx-a11y/role-supports-aria-props -- ARIA 1.2 gives grid cells aria-readonly; a td in a grid is one
              aria-readonly={!!view.readOnlyReason(row, col)}
              aria-label={spokenText(store, view, row, col, view.mergeAt(row, col), text, !!cellOptions(store, view, row, col))}
              tabIndex={active ? 0 : -1}
              cellRef={active ? activeRef : undefined}
              {...(active && typing && { contentEditable: 'plaintext-only', spellCheck: false, autoCapitalize: 'off', onCompositionStart: composing, onCompositionEnd: composed })}
            />
          );
        })}
      </tr>
    ));
  return (
    <table
      className="wbe-grid-access"
      role="grid"
      dir={view.info.rightToLeft ? 'rtl' : undefined}
      aria-label={t('grid.label', { name: view.info.name })}
      aria-rowcount={(view.visibleRows[view.rowCount - 1] ?? -1) + 1}
      aria-colcount={(view.visibleCols[view.colCount - 1] ?? -1) + 1}
      aria-multiselectable="true"
    >
      {frozenRows > 0 && <thead aria-label={t('grid.labelFrozen', { name: view.info.name })}>{rowsOf(rows.filter((r) => r < frozenRows))}</thead>}
      <tbody>{rowsOf(rows.filter((r) => r >= frozenRows))}</tbody>
    </table>
  );
}
