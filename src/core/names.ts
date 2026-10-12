/**
 * Defined names as Excel has them: which names are valid, which name a
 * reference means on a sheet, and the cells a name refers to. Used by the
 * page as well as the worker, so nothing here may load the calculation
 * engine.
 */
import { parseA1Ref, stringifyTokens, tokenize, type Token } from '@borgar/fx';
import { msg, type Message } from '../i18n/message';
import { toFileFormula } from '@sitepen/formula-engine/formula';
import { validateFormula } from './input';
import { colToLetters, lettersToCol, MAX_COLS, MAX_ROWS, quoteSheetName } from '@sitepen/formula-engine/refs';
import type { DefinedName, NameRef, Rect } from './types';

/** Excel's limit on a name's length. */
export const MAX_NAME_LENGTH = 255;

/** Names, like sheet names, ignore case: "TaxRate" and "TAXRATE" are one name. */
export function sameName(a: string, b: string): boolean {
  return a.toUpperCase() === b.toUpperCase();
}

/** Excel's own names: print areas, filters and the like (_xlnm.), and newer functions (_xlfn.). */
export function isBuiltinName(name: string): boolean {
  return /^_xl(?:nm|fn)\./i.test(name);
}

/** Names people see and edit, as Excel's Name Manager lists them: not hidden, and not Excel's own. */
export function isListedName(n: DefinedName): boolean {
  return !n.hidden && !isBuiltinName(n.name);
}

/** Whether `n` is the name `ref` identifies: the same name, in any case, with the same scope. */
export function isName(n: DefinedName, ref: NameRef): boolean {
  return n.localSheet === ref.localSheet && sameName(n.name, ref.name);
}

const START = /^[\p{L}_\\]/u;
const REST = /^[\p{L}\p{M}\p{Nd}_.]*$/u;

/** Whether a name reads as a cell, A1 or R1C1 ("R", "C", "R2C3", "RC"), which can't be a name. */
function readsAsReference(name: string): boolean {
  const a1 = /^([A-Za-z]{1,3})(\d{1,7})$/.exec(name);
  if (a1 && lettersToCol(a1[1]) < MAX_COLS && Number(a1[2]) >= 1 && Number(a1[2]) <= MAX_ROWS) return true;
  return /^[Rr]\d*(?:[Cc]\d*)?$|^[Cc]\d*$/.test(name);
}

/**
 * What is wrong with a name by Excel's rules, or undefined when it is valid:
 * it starts with a letter, an underscore or a backslash; it holds only
 * letters, digits, periods and underscores after that; it isn't a cell
 * reference or TRUE or FALSE; and it is at most 255 characters long.
 */
export function nameProblem(name: string): Message | undefined {
  if (!name) return msg('names.error.empty');
  if (name.length > MAX_NAME_LENGTH) return msg('names.error.tooLong', { max: MAX_NAME_LENGTH });
  if (!START.test(name)) return msg('names.error.start');
  const bad = Array.from(name.slice(1)).find((c) => !REST.test(c));
  if (bad !== undefined) return bad === ' ' ? msg('names.error.space') : msg('names.error.character', { character: bad });
  if (readsAsReference(name)) return msg('names.error.reference', { name });
  if (isBuiltinName(name) || sameName(name, 'TRUE') || sameName(name, 'FALSE')) return msg('names.error.reserved', { name });
  return undefined;
}

/** A name's scope, as the index of the sheet it is local to (-1 for the workbook), for sorting. */
const scopeOrder = (n: DefinedName) => n.localSheet ?? -1;

/**
 * The name `text` means on a sheet, as a formula there would read it: one
 * local to the sheet before one for the whole workbook, ignoring case.
 * "Sheet!Name" means one local to that sheet, or the workbook's.
 */
export function findName(names: DefinedName[], text: string, sheet: number, sheetNames: string[]): DefinedName | undefined {
  let ref;
  try {
    ref = parseA1Ref(text.trim(), { allowTernary: false });
  } catch {
    return undefined;
  }
  if (!ref || !('name' in ref) || !ref.name || (ref.context?.length ?? 0) > 1) return undefined;
  const context = ref.context?.[0];
  const local = context === undefined ? sheet : sheetNames.findIndex((s) => sameName(s, context));
  if (local < 0) return undefined;
  const name = ref.name;
  return names.find((n) => n.localSheet === local && sameName(n.name, name)) ?? names.find((n) => n.localSheet === undefined && sameName(n.name, name));
}

/** The cells a name refers to: ranges on one sheet. */
export interface NameTarget {
  sheet: number;
  rects: Rect[];
}

function tokens(formula: string): Token[] | undefined {
  try {
    return tokenize(formula, { mergeRefs: true, allowTernary: false });
  } catch {
    return undefined;
  }
}

/**
 * The cells a name's formula refers to when it is a plain reference: a range,
 * or ranges joined with commas, on one sheet of this workbook. A reference
 * without a sheet is on the name's own sheet, or else on `sheet`. Undefined
 * for anything else: a value, a calculation (OFFSET, INDEX…), or a reference
 * to another workbook.
 */
export function nameTarget(n: Pick<DefinedName, 'formula' | 'localSheet'>, sheetNames: string[], sheet: number): NameTarget | undefined {
  const solid = tokens(n.formula)?.filter((t) => t.type !== 'whitespace');
  if (!solid?.length || solid.length % 2 === 0) return undefined;
  let at: number | undefined;
  const rects: Rect[] = [];
  for (let i = 0; i < solid.length; i++) {
    const t = solid[i];
    if (i % 2) {
      if (t.type !== 'operator' || t.value !== ',') return undefined;
      continue;
    }
    if ((t.type !== 'range' && t.type !== 'range_beam') || t.value.includes('[')) return undefined;
    const ref = parseA1Ref(t.value, { allowTernary: false });
    if (!ref || !('range' in ref) || !ref.range || (ref.context?.length ?? 0) > 1) return undefined;
    const context = ref.context?.[0];
    const s = context === undefined ? (n.localSheet ?? sheet) : sheetNames.findIndex((x) => sameName(x, context));
    if (s < 0 || (at !== undefined && s !== at)) return undefined;
    at = s;
    const r = ref.range;
    const top = r.top ?? 0;
    const left = r.left ?? 0;
    const bottom = r.bottom ?? (r.top === null ? MAX_ROWS - 1 : top);
    const right = r.right ?? (r.left === null ? MAX_COLS - 1 : left);
    rects.push({ r1: Math.min(top, bottom), c1: Math.min(left, right), r2: Math.max(top, bottom), c2: Math.max(left, right) });
  }
  return at === undefined ? undefined : { sheet: at, rects };
}

const sameRects = (a: Rect[], b: Rect[]) =>
  a.length === b.length && a.every((r) => b.some((x) => x.r1 === r.r1 && x.c1 === r.c1 && x.r2 === r.r2 && x.c2 === r.c2));

/**
 * The name the name box shows for a selection that is exactly what a name
 * refers to, as Excel's does: a name local to the sheet before one for the
 * workbook, then the first alphabetically.
 */
export function nameOfSelection(targets: { name: DefinedName; target: NameTarget }[], sheet: number, rects: Rect[]): string | undefined {
  const fits = targets.filter(({ name, target }) => (name.localSheet === undefined || name.localSheet === sheet) && target.sheet === sheet && sameRects(target.rects, rects));
  fits.sort((a, b) => scopeOrder(b.name) - scopeOrder(a.name) || a.name.name.localeCompare(b.name.name));
  return fits[0]?.name.name;
}

/** A range as an absolute reference: "$A$1:$B$3", "$A$1", whole columns "$B:$C", whole rows "$2:$4". */
export function absoluteRef(r: Rect): string {
  if (r.r1 === 0 && r.r2 === MAX_ROWS - 1) return `$${colToLetters(r.c1)}:$${colToLetters(r.c2)}`;
  if (r.c1 === 0 && r.c2 === MAX_COLS - 1) return `$${r.r1 + 1}:$${r.r2 + 1}`;
  const a = `$${colToLetters(r.c1)}$${r.r1 + 1}`;
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:$${colToLetters(r.c2)}$${r.r2 + 1}`;
}

/** What a name for the selection refers to, as Excel fills it in: "=Data!$A$1:$B$3", areas joined with commas. */
export function selectionFormula(sheetName: string, rects: Rect[]): string {
  return '=' + rects.map((r) => `${quoteSheetName(sheetName)}!${absoluteRef(r)}`).join(',');
}

/**
 * Why a refers-to formula, as typed ("=Data!$A$1"), can't be used: it is
 * empty, Excel couldn't read it, or it names a sheet the workbook hasn't got.
 */
export function refersToProblem(text: string, sheetNames: string[]): Message | undefined {
  const body = text.trim().replace(/^=/, '').trim();
  if (!body) return msg('names.error.refersToEmpty');
  const list = validateFormula(body) === undefined ? tokens(body) : undefined;
  if (!list) return msg('names.error.refersToSyntax');
  for (const t of list) {
    if ((t.type !== 'range' && t.type !== 'range_beam' && t.type !== 'range_named') || t.value.includes('[')) continue;
    const ref = parseA1Ref(t.value, { allowTernary: false });
    const sheet = ref?.context?.find((s) => !sheetNames.some((x) => sameName(x, s)));
    if (sheet !== undefined) return msg('names.error.noSheet', { sheet });
  }
  return undefined;
}

/**
 * A refers-to formula as typed, in the form the file keeps it: without "=",
 * newer functions prefixed, and references that name no sheet put on
 * `sheetName`, as Excel does when a name is defined.
 */
export function refersToFormula(text: string, sheetName: string): string {
  const body = text.trim().replace(/^=/, '').trim();
  const list = tokens(body);
  if (!list) return body;
  const sheet = quoteSheetName(sheetName);
  return toFileFormula(stringifyTokens(list.map((t) => ((t.type === 'range' || t.type === 'range_beam') && !t.value.includes('!') ? { ...t, value: `${sheet}!${t.value}` } : t))));
}
