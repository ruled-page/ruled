import { msg, type Message } from '../i18n/message';
import type { StructuralOp } from './refTransform';
import type { SheetInfo } from './types';

/**
 * A change to the workbook's sheets rather than to their cells. Each is made
 * by rebuilding the workbook, as inserting rows is, so it is undone the same way.
 */
export type SheetOp =
  /** A new, empty worksheet called `name`, at position `at`. */
  | { type: 'addSheet'; at: number; name: string }
  | { type: 'renameSheet'; sheet: number; name: string }
  /** References to the sheet become #REF!, and what only it used goes with it. */
  | { type: 'deleteSheet'; sheet: number }
  /** The sheet moves to position `to` in the new order. */
  | { type: 'moveSheet'; sheet: number; to: number }
  | { type: 'hideSheet'; sheet: number; hidden: boolean };

/** Any change that rebuilds the workbook: one that moves cells, or one to its sheets. */
export type WorkbookOp = StructuralOp | SheetOp;

const SHEET_OPS = new Set<WorkbookOp['type']>(['addSheet', 'renameSheet', 'deleteSheet', 'moveSheet', 'hideSheet']);

export function isSheetOp(op: WorkbookOp): op is SheetOp {
  return SHEET_OPS.has(op.type);
}

/** Excel's longest sheet name. */
export const MAX_SHEET_NAME = 31;

/** What makes a name unusable for a sheet, by Excel's rules, beside the names of the workbook's other sheets. */
export function sheetNameProblem(name: string, others: string[]): Message | undefined {
  if (!name.trim()) return msg('sheets.name.blank');
  if (name.length > MAX_SHEET_NAME) return msg('sheets.name.tooLong', { max: MAX_SHEET_NAME });
  // oxlint-disable-next-line no-control-regex -- line breaks and tabs can't be typed into a tab either
  if (/[:\\/?*[\]\u0000-\u001f]/.test(name)) return msg('sheets.name.characters');
  if (name.startsWith("'") || name.endsWith("'")) return msg('sheets.name.apostrophe');
  if (name.toUpperCase() === 'HISTORY') return msg('sheets.name.reserved');
  const upper = name.toUpperCase();
  if (others.some((o) => o.toUpperCase() === upper)) return msg('sheets.name.taken', { name });
  return undefined;
}

/** The name Excel gives a new sheet: `base` ("Sheet") numbered one past the highest so numbered, and free. */
export function nextSheetName(base: string, names: string[]): string {
  const prefix = base.toUpperCase();
  let n = 0;
  for (const name of names) {
    const upper = name.toUpperCase();
    const rest = upper.slice(prefix.length);
    if (upper.startsWith(prefix) && /^\d{1,9}$/.test(rest)) n = Math.max(n, Number(rest));
  }
  const taken = new Set(names.map((x) => x.toUpperCase()));
  n++;
  while (taken.has(`${prefix}${n}`)) n++;
  return `${base}${n}`;
}

/** Where each sheet is after the change, by its position before; null for the sheet deleted. */
export function sheetIndexMap(op: SheetOp, count: number): (number | null)[] {
  // order[new position] = old position, -1 for the sheet added.
  const order = Array.from({ length: count }, (_, i) => i);
  if (op.type === 'addSheet') order.splice(op.at, 0, -1);
  if (op.type === 'deleteSheet') order.splice(op.sheet, 1);
  if (op.type === 'moveSheet') order.splice(op.to, 0, ...order.splice(op.sheet, 1));
  const map: (number | null)[] = new Array(count).fill(null);
  order.forEach((old, i) => {
    if (old >= 0) map[old] = i;
  });
  return map;
}

/** The map the other way round: where each sheet was before, by its position after; null for the sheet added. */
export function invertIndexMap(map: (number | null)[], count: number): (number | null)[] {
  const out: (number | null)[] = new Array(count).fill(null);
  map.forEach((to, from) => {
    if (to !== null) out[to] = from;
  });
  return out;
}

type SheetFacts = Pick<SheetInfo, 'name' | 'state'>;

/** Why Excel would refuse the change, if it would. */
export function sheetOpProblem(op: SheetOp, sheets: SheetFacts[], structureLocked: boolean): Message | undefined {
  if (structureLocked) return msg('sheets.refuse.protected');
  if (op.type === 'addSheet') {
    if (op.at < 0 || op.at > sheets.length) return msg('sheets.refuse.noSheet');
    return sheetNameProblem(op.name, sheets.map((s) => s.name));
  }
  const sheet = sheets[op.sheet];
  if (!sheet) return msg('sheets.refuse.noSheet');
  const othersVisible = sheets.some((s, i) => i !== op.sheet && s.state === 'visible');
  switch (op.type) {
    case 'renameSheet':
      return sheetNameProblem(op.name, sheets.filter((_s, i) => i !== op.sheet).map((s) => s.name));
    case 'deleteSheet':
      return othersVisible ? undefined : msg('sheets.refuse.deleteLast');
    case 'moveSheet':
      return op.to >= 0 && op.to < sheets.length ? undefined : msg('sheets.refuse.noSheet');
    case 'hideSheet':
      if (!op.hidden) return sheet.state === 'veryHidden' ? msg('sheets.refuse.veryHidden', { sheet: sheet.name }) : undefined;
      return othersVisible || sheet.state !== 'visible' ? undefined : msg('sheets.refuse.hideLast');
  }
}

/** The visible sheet nearest `at`: it, or the first after it, or the last before it; -1 if none is. */
export function nearestVisible(sheets: Pick<SheetInfo, 'state'>[], at: number): number {
  for (let i = Math.max(0, at); i < sheets.length; i++) if (sheets[i].state === 'visible') return i;
  for (let i = Math.min(at, sheets.length) - 1; i >= 0; i--) if (sheets[i].state === 'visible') return i;
  return -1;
}
