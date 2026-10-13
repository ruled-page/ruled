import { inRect } from '@ruled/formula/refs';
import type { RefObject } from 'react';
import { effectiveStyle } from '../../state/cellStyle';
import { checkboxState } from '../../state/checkbox';
import type { SheetView } from '../../state/sheetView';
import type { Selection, WorkbookStore } from '../../state/store';
import { useT } from '../../i18n/react';
import { cellOptions } from '../lists';
import { spokenText } from '../spoken';
import type { Painter, PaintInput } from './draw';

interface Props {
  store: WorkbookStore;
  view: SheetView;
  selection: Selection;
  /** The painter and what it last drew, for the cells' text as drawn. */
  drawn: { painter: Painter; input: PaintInput } | undefined;
  /** The sheet rows and columns to give, in order: those in view, and the active cell's. */
  rows: number[];
  cols: number[];
  /** Set to the active cell's element, which takes the focus. */
  activeRef: RefObject<HTMLTableCellElement | null>;
  /** How many of the sheet's first rows are frozen: they come first, as a group of their own. */
  frozenRows: number;
}

/**
 * The grid for screen readers: the cells in view, and the active cell
 * wherever it is, numbered as the sheet numbers them and each named by its
 * reference, what it shows and what else is true of it. The active cell
 * has the focus while the grid has it, and moves as the selection does.
 */
export function AccessTable({ store, view, selection, drawn, rows, cols, activeRef, frozenRows }: Props) {
  const t = useT();
  const selected = (row: number, col: number) => selection.lines === 'all' || selection.rects.some((r) => (selection.lines === 'rows' ? row >= r.r1 && row <= r.r2 : selection.lines === 'cols' ? col >= r.c1 && col <= r.c2 : inRect(r, row, col)));
  /** What a cell shows: its text as drawn (a number shown as #### in full), or whether its checkbox is ticked. */
  const shown = (row: number, col: number) => {
    const rec = view.get(row, col);
    const style = effectiveStyle(view, store.styles, row, col, rec ? rec.style : view.styleAt(row, col, null));
    const box = rec ? checkboxState(rec, style) : style.checkbox ? false : undefined;
    if (box !== undefined) return t(box ? 'checkbox.ticked' : 'checkbox.unticked');
    return drawn ? drawn.painter.cellText(drawn.input, row, col) : '';
  };
  const rowsOf = (list: number[]) =>
    list.map((row) => (
      <tr key={row} aria-rowindex={row + 1}>
        {cols.map((col) => {
          const active = row === selection.row && col === selection.col;
          const text = shown(row, col);
          return (
            <td
              key={col}
              aria-colindex={col + 1}
              aria-selected={selected(row, col)}
              // oxlint-disable-next-line jsx-a11y/role-supports-aria-props -- ARIA 1.2 gives grid cells aria-readonly; a td in a grid is one
              aria-readonly={!!view.readOnlyReason(row, col)}
              aria-label={spokenText(store, view, row, col, view.mergeAt(row, col), text, !!cellOptions(store, view, row, col))}
              tabIndex={active ? 0 : -1}
              ref={active ? activeRef : undefined}
            >
              {text}
            </td>
          );
        })}
      </tr>
    ));
  return (
    <table
      className="wbe-grid-access"
      role="grid"
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
