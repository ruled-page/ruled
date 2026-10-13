import { parseA1Ref, tokenize, type Token } from '@borgar/fx';
import { stripFunctionPrefixes } from '@ruled/formula/formula';
import type { FunctionHelp } from './functionCatalog';
import { MAX_COLS, MAX_ROWS } from '@ruled/formula/refs';
import type { Rect } from './types';

/**
 * Help while a formula is typed, as Excel gives it: the functions a name
 * being typed could be, the arguments of the function the caret is in, the
 * references in the formula in their colours, and F4's cycling of a
 * reference between relative and absolute. Formulas here are as typed,
 * starting with "=".
 */

/**
 * The list of functions, with their arguments and descriptions, loaded the
 * first time a formula is typed, so pages that never see one don't carry it.
 * Until it has loaded, no functions are offered or hinted at.
 */
let catalog: typeof import('./functionCatalog') | undefined;
let loading: Promise<void> | undefined;

/** Loads the list of functions; resolves at once when it is there already. */
export function loadFunctionHelp(): Promise<void> {
  return (loading ??= import('./functionCatalog').then((m) => {
    catalog = m;
  }));
}

/** Whether the list of functions has loaded. */
export const functionHelpLoaded = () => catalog !== undefined;

/** How many colours references are given in turn; the theme has a --wbe-formula-ref token for each. */
export const REF_COLORS = 7;

/** A function name, or the start of one, at the end of the text. */
const NAME_AT_END = /[A-Za-z_][A-Za-z0-9_.]*$/;
const NAME_CHAR = /[A-Za-z0-9_.]/;
/** What a function name can follow: an operator, a separator, an opening bracket or a space. */
const BEFORE_NAME = /[=+\-*/^&<>(,;{\s]/;
/** The longest function name, with its file prefix, is well under this. */
const LONGEST_NAME = 64;

interface Call {
  /** The function, or undefined for plain brackets. */
  name?: string;
  /** Where its name starts. */
  start: number;
  /** Which argument the scan is in, from 0. */
  arg: number;
  /** Array constants open inside it, whose commas separate values, not arguments. */
  braces: number;
}

/** Whether position `at` is inside text, a quoted sheet name or a table reference, and the calls around it, innermost last. */
function contextAt(text: string, at: number): { literal: boolean; calls: Call[] } {
  const calls: Call[] = [];
  let quote = '';
  let brackets = 0;
  for (let i = 1; i < at; i++) {
    const ch = text[i];
    if (quote) {
      // A doubled quote closes and reopens, which comes to the same.
      if (ch === quote) quote = '';
    } else if (brackets) {
      if (ch === '[') brackets++;
      else if (ch === ']') brackets--;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === '[') {
      brackets++;
    } else if (ch === '(') {
      const name = NAME_AT_END.exec(text.slice(Math.max(1, i - LONGEST_NAME), i))?.[0] ?? '';
      calls.push({ name: name ? stripFunctionPrefixes(name).toUpperCase() : undefined, start: i - name.length, arg: 0, braces: 0 });
    } else if (ch === ')') {
      calls.pop();
    } else {
      const call = calls[calls.length - 1];
      if (!call) continue;
      if (ch === '{') call.braces++;
      else if (ch === '}') call.braces--;
      else if (ch === ',' && !call.braces) call.arg++;
    }
  }
  return { literal: !!quote || brackets > 0, calls };
}

export interface Completion {
  /** Where the name being typed starts; it ends at the caret. */
  start: number;
  end: number;
  /** The functions it could be, in alphabetical order. */
  matches: FunctionHelp[];
}

/**
 * The function name being typed at the caret and the functions it could be:
 * a word at the caret after "=", an operator, a comma, a bracket or a space,
 * and not in text, a sheet name or a table reference.
 */
export function completionAt(text: string, caret: number): Completion | undefined {
  if (!text.startsWith('=') || NAME_CHAR.test(text[caret] ?? '')) return undefined;
  const word = NAME_AT_END.exec(text.slice(Math.max(1, caret - LONGEST_NAME), caret))?.[0];
  if (!word) return undefined;
  const start = caret - word.length;
  if (start > 1 && !BEFORE_NAME.test(text[start - 1])) return undefined;
  if (contextAt(text, start).literal) return undefined;
  const matches = catalog?.functionsStartingWith(word) ?? [];
  return matches.length ? { start, end: caret, matches } : undefined;
}

/** The formula with the name being typed replaced by a function's and its opening bracket, and where the caret goes: inside the brackets. */
export function insertFunction(text: string, at: Pick<Completion, 'start' | 'end'>, name: string): { text: string; caret: number } {
  const bracket = text[at.end] === '(' ? '' : '(';
  return { text: text.slice(0, at.start) + name + bracket + text.slice(at.end), caret: at.start + name.length + 1 };
}

/** The function whose brackets the caret is in, where its name starts, and which of its arguments the caret is in (from 0), when Excel has it. */
export function argumentAt(text: string, caret: number): { fn: FunctionHelp; start: number; index: number } | undefined {
  if (!text.startsWith('=')) return undefined;
  const { calls } = contextAt(text, caret);
  // Inside plain brackets, the hint is still for the function around them.
  for (let i = calls.length - 1; i >= 0; i--) {
    const fn = calls[i].name && catalog?.functionHelp(calls[i].name!);
    if (fn) return { fn, start: calls[i].start, index: calls[i].arg };
    if (calls[i].name) return undefined;
  }
  return undefined;
}

/** A piece of a function's signature as shown, the argument being typed marked. */
export interface SignaturePart {
  text: string;
  current?: boolean;
}

/** One argument, or a group of optional ones shown in one pair of brackets. */
interface Param {
  names: string[];
  optional: boolean;
}

function params(args: string): { list: Param[]; repeats: boolean } {
  const list: Param[] = [];
  let repeats = false;
  for (const m of args.matchAll(/\[([^\]]*)\]|([^,\s][^,]*)/g)) {
    if (m[1] !== undefined) list.push({ names: m[1].split(/,\s*/), optional: true });
    else if (m[2].trim() === '...') repeats = true;
    else list.push({ names: [m[2].trim()], optional: false });
  }
  return { list, repeats };
}

/**
 * The function's signature with argument `index` marked, as Excel shows it:
 * VLOOKUP(lookup_value, table_array, …). Arguments that repeat are shown up
 * to the one being typed: SUM(number1, [number2], [number3], ...). A
 * function without listed arguments shows just its name.
 */
export function signature(fn: FunctionHelp, index: number): SignaturePart[] {
  if (fn.args === undefined) return [{ text: fn.name }];
  const { list, repeats } = params(fn.args);
  if (repeats && list.length) {
    // The last argument (or group) repeats, numbered on: criteria_range2 → criteria_range3.
    const unit = list[list.length - 1];
    const before = list.slice(0, -1).reduce((n, p) => n + p.names.length, 0);
    const more = Math.floor((index - before) / unit.names.length);
    for (let k = 1; k <= more; k++) list.push({ ...unit, names: unit.names.map((n) => n.replace(/\d+$/, (d) => String(Number(d) + k))) });
  }
  const parts: SignaturePart[] = [];
  const add = (text: string, current = false) => {
    const last = parts[parts.length - 1];
    if (current) parts.push({ text, current });
    else if (last && !last.current) last.text += text;
    else parts.push({ text });
  };
  add(`${fn.name}(`);
  let n = 0;
  list.forEach((p, i) => {
    if (i) add(', ');
    if (p.optional) add('[');
    p.names.forEach((name, j) => {
      if (j) add(', ');
      add(name, n++ === index);
    });
    if (p.optional) add(']');
  });
  if (repeats) add(', ...');
  add(')');
  return parts;
}

/** A reference in a formula and its colour while the formula is edited. */
export interface FormulaRef {
  start: number;
  end: number;
  /** From 0, by order of first appearance; the same reference again has the same colour. */
  color: number;
  /** The sheet it names, if it names one. */
  sheet?: string;
  /** Its cells, unless it is in another workbook or spans sheets. */
  rect?: Rect;
}

function refTokens(text: string): (Token & { loc: number[] })[] {
  if (!text.startsWith('=')) return [];
  try {
    return tokenize(text, { mergeRefs: true, withLocation: true, allowTernary: false }).filter(
      (t): t is Token & { loc: number[] } => (t.type === 'range' || t.type === 'range_beam') && !!t.loc,
    );
  } catch {
    return [];
  }
}

/** The cell and range references in a formula, in order, with their colours and cells. */
export function formulaReferences(text: string): FormulaRef[] {
  const colors = new Map<string, number>();
  return refTokens(text).map((t) => {
    const key = t.value.replace(/\$/g, '').toUpperCase();
    let color = colors.get(key);
    if (color === undefined) colors.set(key, (color = colors.size % REF_COLORS));
    const ref: FormulaRef = { start: t.loc[0], end: t.loc[1], color };
    const parsed = parseA1Ref(t.value, { allowTernary: false, allowNamed: false });
    const context = parsed && 'range' in parsed ? (parsed.context ?? []) : [];
    if (parsed && 'range' in parsed && context.length <= 1 && !context[0]?.includes(':')) {
      const { top, left, bottom, right } = parsed.range;
      // Whole columns (A:B) have no rows, and whole rows (1:3) no columns.
      ref.rect = {
        r1: top ?? 0,
        c1: left ?? 0,
        r2: top === null ? MAX_ROWS - 1 : (bottom ?? top),
        c2: left === null ? MAX_COLS - 1 : (right ?? left),
      };
      if (context[0] !== undefined) ref.sheet = context[0];
    }
    return ref;
  });
}

const PART = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/;

/**
 * A reference with its $ signs moved on as F4 moves them in Excel:
 * A1 → $A$1 → A$1 → $A1 → A1. A range moves all its corners together, from
 * the state of the first; a whole column or row (A:A, 1:1) toggles.
 */
export function cycleReference(ref: string): string {
  const bang = ref.lastIndexOf('!');
  const parts = ref
    .slice(bang + 1)
    .split(':')
    .map((p) => PART.exec(p));
  if (parts.some((m) => !m || (!m[2] && !m[4]))) return ref;
  const [, colMark, col, rowMark, row] = parts[0]!;
  let absCol: boolean;
  let absRow: boolean;
  if (col && row) {
    const state = (colMark ? 2 : 0) + (rowMark ? 1 : 0);
    // Relative → both absolute → row absolute → column absolute → relative.
    [absCol, absRow] = ([[true, true], [true, false], [false, false], [false, true]] as const)[state];
  } else {
    absCol = absRow = !(colMark || rowMark);
  }
  const moved = parts.map((m) => (m![2] ? (absCol ? '$' : '') + m![2] : '') + (m![4] ? (absRow ? '$' : '') + m![4] : ''));
  return ref.slice(0, bang + 1) + moved.join(':');
}

/**
 * Applies F4 to a formula being edited: the reference the caret is in or
 * just after (or every reference the selection touches) moves on to its next
 * mix of relative and absolute. Returns the new text and selection, or
 * undefined when there is no reference there.
 */
export function toggleAbsolute(text: string, start: number, end: number): { text: string; start: number; end: number } | undefined {
  const refs = refTokens(text);
  const hit = start === end ? refs.filter((t) => t.loc[0] <= start && start <= t.loc[1]).slice(-1) : refs.filter((t) => t.loc[0] < end && start < t.loc[1]);
  if (!hit.length) return undefined;
  let out = '';
  let from = 0;
  for (const t of hit) {
    out += text.slice(from, t.loc[0]) + cycleReference(t.value);
    from = t.loc[1];
  }
  const after = text.length - from;
  out += text.slice(from);
  // A caret goes to the end of its reference; a selection covers what changed.
  if (start === end) return { text: out, start: out.length - after, end: out.length - after };
  const last = hit[hit.length - 1].loc[1];
  return { text: out, start: Math.min(start, hit[0].loc[0]), end: out.length - (text.length - Math.max(end, last)) };
}
