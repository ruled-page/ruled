import { parse, tokenize } from '@borgar/fx';
import { format } from 'numfmt';
import { msg, type MessageKey, type Text } from '../i18n/message';
import { isUnsupportedError, type EngineValue } from './engineValue';
import { makeShifter } from './formula';
import type { CellContent, ListSource, Rect, RuleValidation } from './types';
import { resolveListSource, type ListContext } from './validations';

/** What validation needs from the calculation engine. */
export interface Evaluator {
  /** Evaluates a formula (with "=") on a sheet. */
  calculate(formula: string, sheet: number): EngineValue | EngineValue[][];
  /** Runs fn with `content` temporarily in the cell, then puts the cell back. */
  withValue<T>(sheet: number, row: number, col: number, content: CellContent, fn: () => T): T;
}

type Node = { type: string; loc?: [number, number]; callee?: { name: string }; arguments?: Node[] };

function toNumber(v: EngineValue | EngineValue[][]): number {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  // An empty cell counts as zero, as in Excel.
  if (v === null) return 0;
  return NaN;
}


/** A rule's formulas are written for its top-left cell; move them to `row, col`. */
function shiftTo(formula: string, anchor: Rect, row: number, col: number): string {
  try {
    return makeShifter(formula, anchor.r1, anchor.c1)(row, col);
  } catch {
    return formula;
  }
}

/**
 * Resolves a list rule's formula to a range or a set of values. OFFSET and
 * INDIRECT are handled here because the engine only accepts constant OFFSET
 * arguments and has no INDIRECT; their arguments are evaluated by the engine.
 */
export function evaluateListFormula(formula: string, sheet: number, ctx: ListContext, ev: Evaluator): ListSource | undefined {
  let ast: Node;
  try {
    ast = parse(tokenize(formula, { mergeRefs: true, allowTernary: false, withLocation: true })) as Node;
  } catch {
    return undefined;
  }
  const src = (n: Node) => (n.loc ? formula.slice(n.loc[0], n.loc[1]) : '');
  const num = (n: Node | undefined) => (n ? toNumber(ev.calculate('=' + src(n), sheet)) : NaN);
  const name = ast.type === 'CallExpression' ? ast.callee?.name.toUpperCase() : undefined;
  const args = ast.arguments ?? [];

  if (name === 'OFFSET' && args.length >= 1) {
    const base = resolveListSource(src(args[0]), { ...ctx, ownSheet: sheet });
    if (base.kind !== 'range') return undefined;
    const rows = num(args[1]) || 0;
    const cols = num(args[2]) || 0;
    const height = args[3] ? num(args[3]) : base.rect.r2 - base.rect.r1 + 1;
    const width = args[4] ? num(args[4]) : base.rect.c2 - base.rect.c1 + 1;
    if (!(height >= 1) || !(width >= 1)) return { kind: 'values', values: [] };
    const r1 = base.rect.r1 + rows;
    const c1 = base.rect.c1 + cols;
    if (r1 < 0 || c1 < 0) return undefined;
    return { kind: 'range', sheet: base.sheet, rect: { r1, c1, r2: r1 + height - 1, c2: c1 + width - 1 } };
  }
  if (name === 'INDIRECT' && args.length >= 1) {
    const target = ev.calculate('=' + src(args[0]), sheet);
    if (typeof target !== 'string' || !target) return undefined;
    const source = resolveListSource(target, { ...ctx, ownSheet: sheet });
    return source.kind === 'formula' ? undefined : source;
  }
  const v = ev.calculate('=' + formula, sheet);
  const values = (Array.isArray(v) ? v.flat() : [v])
    .filter((x): x is string | number | boolean => x !== null && typeof x !== 'object' && x !== '')
    .map((x) => (typeof x === 'boolean' ? (x ? 'TRUE' : 'FALSE') : String(x)));
  return { kind: 'values', values: [...new Set(values)] };
}

function compare(op: RuleValidation['operator'], v: number, a: number, b: number): boolean {
  switch (op) {
    case 'between':
      return v >= Math.min(a, b) && v <= Math.max(a, b);
    case 'notBetween':
      return v < Math.min(a, b) || v > Math.max(a, b);
    case 'equal':
      return v === a;
    case 'notEqual':
      return v !== a;
    case 'greaterThan':
      return v > a;
    case 'lessThan':
      return v < a;
    case 'greaterThanOrEqual':
      return v >= a;
    case 'lessThanOrEqual':
      return v <= a;
  }
}

/**
 * Whether typed input satisfies a number, date, time, text-length or custom
 * rule, judged the way Excel judges it when the value is entered.
 */
export function checkRule(
  rule: RuleValidation,
  sheet: number,
  row: number,
  col: number,
  content: CellContent,
  typed: string,
  ev: Evaluator,
): boolean {
  if (content.kind === 'empty' || content.kind === 'formula') return true;
  const anchor = rule.rects[0];
  const bound = (f: string | undefined) => {
    if (f === undefined) return NaN;
    if (/^-?\d+(\.\d+)?$/.test(f.trim())) return Number(f);
    return toNumber(ev.calculate('=' + shiftTo(f, anchor, row, col), sheet));
  };
  if (rule.kind === 'custom') {
    const formula = '=' + shiftTo(rule.formula1, anchor, row, col);
    const v = ev.withValue(sheet, row, col, content, () => ev.calculate(formula, sheet));
    return v === true || (typeof v === 'number' && v !== 0) || (!Array.isArray(v) && isUnsupportedError(v));
  }
  const a = bound(rule.formula1);
  const b = bound(rule.formula2);
  // A bound that can't be worked out here is left for Excel to judge.
  if (Number.isNaN(a) || (Number.isNaN(b) && (rule.operator === 'between' || rule.operator === 'notBetween'))) return true;
  if (rule.kind === 'textLength') {
    const text = content.kind === 'string' ? content.value : typed.trim();
    return compare(rule.operator, text.length, a, b);
  }
  if (content.kind !== 'number') return false;
  if (rule.kind === 'whole' && !Number.isInteger(content.value)) return false;
  return compare(rule.operator, content.value, a, b);
}

/** A rule's bound as people type it: date and time serials as dates and times. */
function boundText(rule: RuleValidation, formula: string): string {
  if (!/^-?\d+(\.\d+)?$/.test(formula)) return formula;
  const n = Number(formula);
  if (rule.kind === 'date') return format('yyyy-mm-dd', n);
  if (rule.kind === 'time') return format('hh:mm', n);
  return formula;
}

const KINDS: Record<Exclude<RuleValidation['kind'], 'custom'>, MessageKey> = {
  whole: 'core.validation.kind.whole',
  decimal: 'core.validation.kind.decimal',
  date: 'core.validation.kind.date',
  time: 'core.validation.kind.time',
  textLength: 'core.validation.kind.textLength',
};

const OPERATORS: Record<RuleValidation['operator'], MessageKey> = {
  between: 'core.validation.between',
  notBetween: 'core.validation.notBetween',
  equal: 'core.validation.equal',
  notEqual: 'core.validation.notEqual',
  greaterThan: 'core.validation.greaterThan',
  lessThan: 'core.validation.lessThan',
  greaterThanOrEqual: 'core.validation.greaterThanOrEqual',
  lessThanOrEqual: 'core.validation.lessThanOrEqual',
};

/** One-line description of a rule, for the formula bar: "Whole number between 1 and 10". */
export function describeRule(rule: RuleValidation): Text {
  if (rule.kind === 'custom') return msg('core.validation.custom', { formula: rule.formula1 });
  return msg(OPERATORS[rule.operator], {
    what: msg(KINDS[rule.kind]),
    a: boundText(rule, rule.formula1),
    b: boundText(rule, rule.formula2 ?? ''),
  });
}
