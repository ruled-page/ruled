import { displayFormula, shiftFormula } from '@ruled/formula/formula';
import { parseInput } from '../core/input';
import { listSeparator } from '../core/locale';
import type { DisplayStyle, ListValidation, Rect, RuleValidation } from '../core/types';
import { resolveListSource, type ListContext } from '../core/validationRules';
import { VT } from '../core/types';
import type { MessageKey } from '../i18n/message';
import { editText } from './format';

/** What Excel's Data Validation dialog lets a cell take. */
export type ValidationAllow = 'any' | 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';

/** The dialog's boxes, as typed: bounds, a list's source or a formula as the locale writes them. */
export interface ValidationDraft {
  allow: ValidationAllow;
  operator: RuleValidation['operator'];
  /** The minimum (or the value), the list's source, or the custom formula. */
  first: string;
  /** The maximum, for between and not between. */
  second: string;
  allowBlank: boolean;
  inCellDropdown: boolean;
  showPrompt: boolean;
  promptTitle: string;
  prompt: string;
  showError: boolean;
  errorStyle: RuleValidation['errorStyle'];
  errorTitle: string;
  error: string;
}

/** Where the dialog reads and writes: the active cell, which formulas are typed for, and the rule's top-left cell, which they're kept for. */
export interface DraftPlace {
  active: { row: number; col: number };
  rects: Rect[];
  locale: string;
  date1904: boolean;
}

const anchorOf = (rects: Rect[]) => ({ row: Math.min(...rects.map((r) => r.r1)), col: Math.min(...rects.map((r) => r.c1)) });

/** A new rule's settings, as Excel's dialog starts them. */
export const NEW_DRAFT: ValidationDraft = {
  allow: 'any',
  operator: 'between',
  first: '',
  second: '',
  allowBlank: true,
  inCellDropdown: true,
  showPrompt: true,
  promptTitle: '',
  prompt: '',
  showError: true,
  errorStyle: 'stop',
  errorTitle: '',
  error: '',
};

/** A kept bound or formula as the dialog shows it: a number, date or time as the locale writes it, or a formula for the active cell. */
function boundText(formula: string | undefined, allow: ValidationAllow, place: DraftPlace, base: DisplayStyle): string {
  if (!formula) return '';
  if (/^-?\d+(\.\d+)?$/.test(formula)) {
    const n = Number(formula);
    if (allow === 'date' || allow === 'time') {
      const numFmt = allow === 'date' ? 'm/d/yyyy' : 'h:mm:ss';
      return editText({ t: VT.Number, n, style: 0, flags: 0 }, { ...base, numFmt }, place.date1904, place.locale);
    }
    return n.toLocaleString(place.locale, { useGrouping: false, maximumFractionDigits: 15 });
  }
  const a = anchorOf(place.rects);
  return displayFormula(shiftFormula(formula, a.row, a.col, place.active.row, place.active.col) ?? formula);
}

/** The dialog's boxes for a cell's rule, or for a new one. */
export function draftFor(rule: ListValidation | RuleValidation | undefined, place: DraftPlace, base: DisplayStyle): ValidationDraft {
  if (!rule) return NEW_DRAFT;
  const alert = {
    allowBlank: rule.allowBlank,
    showPrompt: rule.showPrompt ?? !!rule.prompt,
    promptTitle: rule.promptTitle ?? '',
    prompt: rule.prompt ?? '',
    showError: rule.showError,
    errorStyle: rule.errorStyle,
    errorTitle: rule.errorTitle ?? '',
    error: rule.error ?? '',
  };
  if ('source' in rule) {
    const formula = rule.formula ?? '';
    // An inline list is typed without its quotes, separated as the locale separates lists.
    const first = formula.startsWith('"')
      ? formula.slice(1, -1).replace(/""/g, '"').split(',').join(listSeparator(place.locale))
      : formula
        ? boundText(formula, 'list', place, base)
        : rule.source.kind === 'values'
          ? rule.source.values.join(listSeparator(place.locale))
          : '';
    return { ...NEW_DRAFT, ...alert, allow: 'list', first, inCellDropdown: rule.inCellDropdown };
  }
  return {
    ...NEW_DRAFT,
    ...alert,
    allow: rule.kind,
    operator: rule.operator,
    first: boundText(rule.formula1, rule.kind, place, base),
    second: boundText(rule.formula2, rule.kind, place, base),
  };
}

/** A formula typed into a box, kept for the rule's top-left cell; undefined when it isn't one. */
function formulaFrom(text: string, place: DraftPlace): string | undefined {
  const parsed = parseInput(text.startsWith('=') ? text : `=${text}`, 'General', place.locale, place.date1904);
  if (parsed.error || parsed.content.kind !== 'formula') return undefined;
  const a = anchorOf(place.rects);
  return shiftFormula(parsed.content.formula, place.active.row, place.active.col, a.row, a.col);
}

/** A bound as the rule keeps it: a number (a date's or time's serial number) or a formula. */
function boundFrom(text: string, allow: ValidationAllow, place: DraftPlace): string | number | undefined {
  const t = text.trim();
  if (!t) return undefined;
  if (t.startsWith('=')) return formulaFrom(t, place);
  const parsed = parseInput(t, 'General', place.locale, place.date1904);
  if (parsed.error || parsed.content.kind !== 'number') return undefined;
  const n = parsed.content.value;
  return (allow === 'whole' || allow === 'textLength') && !Number.isInteger(n) ? undefined : n;
}

/** Why the dialog's boxes don't make a rule, as Excel says it, and which box. */
export interface DraftProblem {
  message: MessageKey;
  field: 'first' | 'second';
}

/**
 * The rule the dialog's boxes make for `rects`, or the problem with them.
 * "Any value" without an input message makes no rule at all, taking the
 * cells' rule away, as in Excel.
 */
export function ruleFromDraft(d: ValidationDraft, place: DraftPlace, lists: ListContext): { rule?: ListValidation | RuleValidation } | { problem: DraftProblem } {
  const text = (s: string) => s.trim() || undefined;
  const alert = {
    allowBlank: d.allowBlank,
    showPrompt: d.showPrompt,
    showError: d.showError,
    errorStyle: d.errorStyle,
    errorTitle: text(d.errorTitle),
    error: text(d.error),
    promptTitle: text(d.promptTitle),
    prompt: text(d.prompt),
  };
  const rects = place.rects;
  if (d.allow === 'any') return d.prompt.trim() || d.promptTitle.trim() ? { rule: { ...alert, rects, kind: 'any', operator: 'between', formula1: '' } } : {};
  if (d.allow === 'list') {
    const source = d.first.trim();
    if (!source) return { problem: { message: 'validation.problem.source', field: 'first' } };
    let formula: string | undefined;
    if (source.startsWith('=')) formula = formulaFrom(source, place);
    else {
      const values = source
        .split(listSeparator(place.locale))
        .map((v) => v.trim())
        .filter(Boolean);
      formula = `"${values.map((v) => v.replace(/"/g, '""')).join(',')}"`;
    }
    if (!formula) return { problem: { message: 'validation.problem.source', field: 'first' } };
    return { rule: { ...alert, rects, source: resolveListSource(formula, lists), formula, inCellDropdown: d.inCellDropdown } };
  }
  if (d.allow === 'custom') {
    const formula = d.first.trim() ? formulaFrom(d.first.trim(), place) : undefined;
    if (!formula) return { problem: { message: 'validation.problem.formula', field: 'first' } };
    return { rule: { ...alert, rects, kind: 'custom', operator: 'between', formula1: formula } };
  }
  const two = d.operator === 'between' || d.operator === 'notBetween';
  const needs: MessageKey =
    d.allow === 'whole' || d.allow === 'textLength' ? 'validation.problem.whole' : d.allow === 'date' ? 'validation.problem.date' : d.allow === 'time' ? 'validation.problem.time' : 'validation.problem.number';
  const a = boundFrom(d.first, d.allow, place);
  if (a === undefined) return { problem: { message: needs, field: 'first' } };
  const b = two ? boundFrom(d.second, d.allow, place) : undefined;
  if (two && b === undefined) return { problem: { message: needs, field: 'second' } };
  if (typeof a === 'number' && typeof b === 'number' && b < a) return { problem: { message: 'validation.problem.order', field: 'second' } };
  return { rule: { ...alert, rects, kind: d.allow, operator: d.operator, formula1: String(a), ...(two && { formula2: String(b) }) } };
}
