import type { SheetView } from '../state/sheetView';
import type { WorkbookStore } from '../state/store';

/**
 * The values a cell's dropdown offers: its schema field's, or else those of
 * Excel's own list validation; undefined for a cell with neither, or one
 * that can't be edited.
 */
export function cellOptions(store: WorkbookStore, view: SheetView, row: number, col: number): readonly string[] | undefined {
  if (view.readOnlyReason(row, col)) return undefined;
  const field = store.schema?.fieldAt(view.info.index, row, col)?.field;
  if (field?.type === 'enum' && field.values?.length) return field.values;
  if (!view.info.listValidations.length) return undefined;
  const dv = store.listAt(view.info.index, row, col);
  const options = dv?.inCellDropdown ? store.listOptions(dv, row, col) : undefined;
  return options?.length ? options : undefined;
}
