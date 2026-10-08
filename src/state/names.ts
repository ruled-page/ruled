import { tokenize } from '@borgar/fx';
import { displayFormula } from '../core/formula';
import {
  findName,
  isListedName,
  isName,
  nameOfSelection,
  nameProblem,
  nameTarget,
  refersToFormula,
  refersToProblem,
  selectionFormula,
  type NameTarget,
} from '../core/names';
import { VT, type DefinedName, type NameRef } from '../core/types';
import { msg, type Text } from '../i18n/message';
import type { WorkbookStore } from './store';

/**
 * Defined names as the editor's screens use them: what the name box shows
 * and offers, going to names, defining, editing and deleting them, and the
 * Name Manager's list and previews.
 */

type Targeted = { name: DefinedName; target: NameTarget };

/** Per list of names (a new one comes with every change), per active sheet: the listed names that refer to cells. */
const cache = new WeakMap<DefinedName[], Map<number, Targeted[]>>();

const sheetNames = (store: WorkbookStore) => store.sheets.map((s) => s.info.name);
const byName = (a: DefinedName, b: DefinedName) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }) || (a.localSheet ?? -1) - (b.localSheet ?? -1);

/** The names people see and edit, as Excel's Name Manager lists them: alphabetically, the workbook's before a sheet's. */
export function listedNames(store: WorkbookStore): DefinedName[] {
  return store.names.filter(isListedName).sort(byName);
}

/** The listed names that refer to cells, with those cells; a reference without a sheet is read on the active sheet. */
function targets(store: WorkbookStore): Targeted[] {
  let bySheet = cache.get(store.names);
  if (!bySheet) cache.set(store.names, (bySheet = new Map()));
  let list = bySheet.get(store.active);
  if (!list) {
    const sheets = sheetNames(store);
    list = store.names.filter(isListedName).flatMap((name) => {
      const target = nameTarget(name, sheets, store.active);
      return target ? [{ name, target }] : [];
    });
    bySheet.set(store.active, list);
  }
  return list;
}

/** The cells a name refers to, when it is a plain range. */
export function targetOf(store: WorkbookStore, n: DefinedName): NameTarget | undefined {
  return nameTarget(n, sheetNames(store), store.active);
}

/** The listed name that text typed on the active sheet means, if any: the sheet's own before the workbook's. */
export function nameFor(store: WorkbookStore, text: string): DefinedName | undefined {
  return findName(store.names.filter(isListedName), text, store.active, sheetNames(store));
}

/** The cells a name typed in the name box refers to, when it is a name for a range. */
export function namedRange(store: WorkbookStore, text: string): NameTarget | undefined {
  const n = nameFor(store, text);
  return n && targetOf(store, n);
}

/** The name the name box shows when the selection is exactly a named range, as Excel's does. */
export function selectionName(store: WorkbookStore): string | undefined {
  return nameOfSelection(targets(store), store.active, store.selection.rects);
}

/**
 * The names the name box offers on the active sheet, as Excel's does: those
 * that refer to cells, the workbook's and the sheet's own (which hides the
 * workbook's of the same name), alphabetically.
 */
export function nameBoxNames(store: WorkbookStore): { name: string; refersTo: string }[] {
  const out = new Map<string, DefinedName>();
  for (const { name } of targets(store)) {
    if (name.localSheet !== undefined && name.localSheet !== store.active) continue;
    const key = name.name.toUpperCase();
    if (!out.has(key) || name.localSheet !== undefined) out.set(key, name);
  }
  return [...out.values()].sort(byName).map((n) => ({ name: n.name, refersTo: displayFormula(n.formula).replace(/^=/, '') }));
}

/** Why `name` can't name something in a scope (replacing the name `from`): Excel's rules, or a name already there. */
export function nameError(store: WorkbookStore, name: string, localSheet: number | undefined, from?: NameRef): Text | undefined {
  // A name kept as the file has it is let be, even where these rules are stricter than Excel's.
  const problem = from && from.name === name ? undefined : nameProblem(name);
  if (problem) return problem;
  const clash = store.names.find((n) => isName(n, { name, localSheet }) && !(from && isName(n, from)));
  if (!clash) return undefined;
  return localSheet === undefined
    ? msg('names.error.takenWorkbook', { name: clash.name })
    : msg('names.error.takenSheet', { name: clash.name, sheet: store.sheets[localSheet]?.info.name ?? '' });
}

/** Why a refers-to formula, as typed, can't be used. */
export function refersToError(store: WorkbookStore, text: string): Text | undefined {
  return refersToProblem(text, sheetNames(store));
}

/** A refers-to formula as typed, in the form the file keeps, with references that name no sheet on the active sheet. */
export function refersToFile(store: WorkbookStore, text: string): string {
  return refersToFormula(text, store.view?.info.name ?? '');
}

/** The selection as a refers-to formula, as Excel fills it in: "=Data!$A$1:$B$3". */
export function selectionRefersTo(store: WorkbookStore): string {
  return selectionFormula(store.view?.info.name ?? '', store.selection.rects);
}

/** Defines a name, or changes the one `from` identifies, as one step to undo. Resolves with why it was refused, if it was. */
export async function defineName(store: WorkbookStore, to: DefinedName, from?: NameRef): Promise<Text | undefined> {
  const problem = nameError(store, to.name, to.localSheet, from);
  if (problem) return problem;
  const r = await store.changeName({ ...(from && { from }), to }, msg(from ? 'names.history.edited' : 'names.history.defined', { name: to.name }));
  return r?.rejected;
}

/** Deletes a name, as one step to undo, once the user has confirmed it. Resolves true if it was deleted. */
export async function deleteName(store: WorkbookStore, n: DefinedName): Promise<boolean> {
  const ok = await store.confirm({
    title: msg('names.delete.title', { name: n.name }),
    body: msg('names.delete.body'),
    confirmLabel: msg('names.delete.confirm'),
  });
  if (!ok) return false;
  const r = await store.changeName({ from: { name: n.name, localSheet: n.localSheet }, to: null }, msg('names.history.deleted', { name: n.name }));
  return !!r && !r.rejected;
}

/**
 * Enter in the name box, as in Excel: goes to a cell, a range or a name, or
 * names the selection, for the whole workbook, with a new name. Says why
 * when it can do neither. Resolves true once it has done one.
 */
export async function nameBoxEnter(store: WorkbookStore, text: string): Promise<boolean> {
  const typed = text.trim();
  if (store.goTo(typed)) return true;
  const place = store.placeOf(typed);
  const sheet = place && store.sheets[place.sheet].info;
  if (sheet && sheet.state !== 'visible') {
    store.toast('warn', msg(sheet.state === 'veryHidden' ? 'names.goTo.veryHiddenSheet' : 'names.goTo.hiddenSheet', { sheet: sheet.name }));
    return false;
  }
  const named = nameFor(store, typed);
  const view = store.view;
  let problem: Text | undefined;
  if (named) problem = msg('names.goTo.notRange', { name: named.name, formula: displayFormula(named.formula) });
  else if (nameProblem(typed)) problem = msg('names.goTo.unknown', { text: typed });
  else if (view?.info.kind !== 'worksheet') problem = msg('names.goTo.notWorksheet');
  else problem = nameError(store, typed, undefined);
  if (problem) {
    store.toast('warn', problem);
    return false;
  }
  if (await defineName(store, { name: typed, formula: refersToFile(store, selectionRefersTo(store)) })) return false;
  store.focusGrid();
  return true;
}

/** How many values a range's preview shows before "…". */
const PREVIEW = 8;

/**
 * What a name refers to, as Excel's Name Manager previews it: a cell's
 * value, the first values of a range in braces ({"a","b";"c","d"}), or a
 * constant. Empty for anything that needs calculating.
 */
export function nameValue(store: WorkbookStore, n: DefinedName): string {
  if (n.formula.includes('#REF!')) return '#REF!';
  const target = targetOf(store, n);
  const view = target && store.sheets[target.sheet];
  if (target && view && target.rects.length === 1) {
    const r = target.rects[0];
    const text = (row: number, col: number) => {
      const shown = store.shownText(view, row, col);
      return view.get(row, col)?.t === VT.String ? `"${shown}"` : shown;
    };
    if (r.r1 === r.r2 && r.c1 === r.c2) return text(r.r1, r.c1);
    const extent = view.extent();
    const rows: string[] = [];
    let count = 0;
    for (let row = r.r1; row <= Math.min(r.r2, Math.max(extent.maxRow, r.r1)) && count < PREVIEW; row++) {
      const cells: string[] = [];
      for (let col = r.c1; col <= Math.min(r.c2, Math.max(extent.maxCol, r.c1)) && count < PREVIEW; col++, count++) cells.push(text(row, col));
      rows.push(cells.join(','));
    }
    const more = (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) > count;
    return `{${rows.join(';')}${more ? '…' : ''}}`;
  }
  let tokens;
  try {
    tokens = tokenize(n.formula, { allowTernary: false }).filter((t) => t.type !== 'whitespace');
  } catch {
    return '';
  }
  const [only] = tokens;
  if (tokens.length === 1 && (only.type === 'number' || only.type === 'string' || only.type === 'bool' || only.type === 'error')) return only.value;
  return '';
}

/** Whether a name's value is an error, as Excel's Name Manager filters them. */
export function nameHasError(store: WorkbookStore, n: DefinedName): boolean {
  return nameValue(store, n).startsWith('#');
}

/** Opens the dialog that defines a name for the selection, as Excel's Define Name does. */
export function openDefineName(store: WorkbookStore): void {
  store.set('nameDialog', { name: '', refersTo: selectionRefersTo(store), comment: '' });
}
