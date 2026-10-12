import { VT, type DisplayStyle } from '../core/types';
import type { CellRecord } from './format';

/**
 * Whether a cell shows Excel 365's checkbox, and ticked or not: TRUE ticked,
 * FALSE or nothing unticked. Undefined when it shows none, for a cell
 * without the checkbox format, or with a number, text or error in it.
 */
export function checkboxState(rec: CellRecord | null | undefined, style: DisplayStyle | undefined): boolean | undefined {
  if (!style?.checkbox) return undefined;
  if (!rec || rec.t === VT.None) return false;
  return rec.t === VT.Bool ? rec.n === 1 : undefined;
}
