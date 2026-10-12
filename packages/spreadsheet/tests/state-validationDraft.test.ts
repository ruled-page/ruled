import { describe, expect, it } from 'vitest';
import type { ListValidation, RuleValidation } from '../src/core/types';
import { draftFor, NEW_DRAFT, ruleFromDraft, type DraftPlace, type ValidationDraft } from '../src/state/validationDraft';
import { PLAIN } from './state-fixtures';

const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
/** The selection B2:B5 with its active cell B3, in the US or in Germany. */
const place = (locale = 'en-US'): DraftPlace => ({ active: { row: 2, col: 1 }, rects: [rect(1, 1, 4, 1)], locale, date1904: false });
const lists = { sheetIndex: new Map([['data', 1]]), ownSheet: 0, names: [{ name: 'Sizes', formula: 'Data!$A$1:$A$3' }] };
const draft = (d: Partial<ValidationDraft>): ValidationDraft => ({ ...NEW_DRAFT, ...d });
const ruleOf = (d: Partial<ValidationDraft>, p = place()) => {
  const r = ruleFromDraft(draft(d), p, lists);
  if ('problem' in r) throw new Error(r.problem.message);
  return r.rule;
};
const problem = (d: Partial<ValidationDraft>, p = place()) => {
  const r = ruleFromDraft(draft(d), p, lists);
  return 'problem' in r ? r.problem : undefined;
};

describe('the Data Validation dialog makes rules', () => {
  it('of numbers, dates and times typed as the locale writes them', () => {
    expect(ruleOf({ allow: 'decimal', operator: 'between', first: '1,5', second: '10' }, place('de-DE'))).toMatchObject({ kind: 'decimal', formula1: '1.5', formula2: '10' });
    expect(ruleOf({ allow: 'date', operator: 'greaterThan', first: '2026-01-05' })).toMatchObject({ kind: 'date', operator: 'greaterThan', formula1: '46027' });
    expect(ruleOf({ allow: 'time', operator: 'lessThan', first: '9:30' })).toMatchObject({ formula1: '0.3958333333333333' });
    expect(ruleOf({ allow: 'textLength', operator: 'lessThanOrEqual', first: '20' })).not.toHaveProperty('formula2');
  });

  it("with formulas typed for the active cell, kept for the rule's top-left cell", () => {
    expect(ruleOf({ allow: 'whole', operator: 'lessThan', first: '=C3*2' })).toMatchObject({ formula1: 'C2*2' });
    expect(ruleOf({ allow: 'custom', first: '=ISNUMBER(B3)' })).toMatchObject({ kind: 'custom', formula1: 'ISNUMBER(B2)' });
  });

  it('of lists, from items separated as the locale separates them, or a range or a name', () => {
    expect(ruleOf({ allow: 'list', first: 'Yes; No ;"Maybe"' }, place('de-DE'))).toMatchObject({ formula: '"Yes,No,""Maybe"""', source: { kind: 'values', values: ['Yes', 'No', '"Maybe"'] } });
    expect(ruleOf({ allow: 'list', first: '=Data!$A$1:$A$9', inCellDropdown: false })).toMatchObject({ formula: 'Data!$A$1:$A$9', source: { kind: 'range', sheet: 1 }, inCellDropdown: false });
    expect(ruleOf({ allow: 'list', first: '=Sizes' })).toMatchObject({ formula: 'Sizes', source: { kind: 'range', sheet: 1, rect: rect(0, 0, 2, 0) } });
  });

  it('with their messages, and none at all for any value without one', () => {
    const r = ruleOf({ allow: 'whole', operator: 'equal', first: '3', promptTitle: ' Qty ', prompt: 'Three', errorStyle: 'warning', error: 'Not three' });
    expect(r).toMatchObject({ promptTitle: 'Qty', prompt: 'Three', errorStyle: 'warning', error: 'Not three', errorTitle: undefined, rects: [rect(1, 1, 4, 1)] });
    expect(ruleOf({ allow: 'any' })).toBeUndefined();
    expect(ruleOf({ allow: 'any', prompt: 'Anything goes' })).toMatchObject({ kind: 'any', formula1: '' });
  });

  it('say what is wrong, and with which box', () => {
    expect(problem({ allow: 'whole', operator: 'between', first: '1.5', second: '3' })).toEqual({ message: 'validation.problem.whole', field: 'first' });
    expect(problem({ allow: 'decimal', operator: 'between', first: '1', second: '' })).toEqual({ message: 'validation.problem.number', field: 'second' });
    expect(problem({ allow: 'date', operator: 'between', first: '2026-02-01', second: '2026-01-01' })).toEqual({ message: 'validation.problem.order', field: 'second' });
    expect(problem({ allow: 'time', operator: 'equal', first: 'noon' })?.message).toBe('validation.problem.time');
    expect(problem({ allow: 'date', operator: 'equal', first: '' })?.message).toBe('validation.problem.date');
    expect(problem({ allow: 'list', first: ' ' })?.message).toBe('validation.problem.source');
    expect(problem({ allow: 'list', first: '=SUM(' })?.message).toBe('validation.problem.source');
    expect(problem({ allow: 'custom', first: '' })?.message).toBe('validation.problem.formula');
  });
});

describe('the Data Validation dialog shows a rule', () => {
  const alert = { allowBlank: false, showError: true, errorStyle: 'stop' as const };
  it('with its bounds as typed: numbers, dates and times in the locale, and formulas for the active cell', () => {
    const rule: RuleValidation = { ...alert, rects: [rect(1, 1, 4, 1)], kind: 'date', operator: 'between', formula1: '46027', formula2: 'C2+7', prompt: 'When' };
    expect(draftFor(rule, place('de-DE'), PLAIN)).toMatchObject({ allow: 'date', first: '05.01.2026', second: '=C3+7', showPrompt: true, prompt: 'When', allowBlank: false });
    const time: RuleValidation = { ...alert, rects: [rect(1, 1)], kind: 'time', operator: 'lessThan', formula1: '0.5', showPrompt: false };
    expect(draftFor(time, place(), PLAIN)).toMatchObject({ first: '12:00:00', showPrompt: false });
    expect(draftFor({ ...rule, kind: 'decimal', formula1: '1.5', formula2: undefined }, place('de-DE'), PLAIN)).toMatchObject({ first: '1,5', second: '' });
  });

  it('with a list as typed: its items separated as the locale separates them, or its range', () => {
    const list: ListValidation = { ...alert, rects: [rect(1, 1)], source: { kind: 'values', values: ['a', 'b'] }, formula: '"a,""b"""', inCellDropdown: false };
    expect(draftFor(list, place('de-DE'), PLAIN)).toMatchObject({ allow: 'list', first: 'a;"b"', inCellDropdown: false });
    expect(draftFor({ ...list, formula: 'Data!$A$1:$A$3' }, place(), PLAIN).first).toBe('=Data!$A$1:$A$3');
    expect(draftFor({ ...list, formula: undefined }, place(), PLAIN).first).toBe('a,b');
    expect(draftFor({ ...list, formula: undefined, source: { kind: 'range', sheet: 0, rect: rect(0, 0) } }, place(), PLAIN).first).toBe('');
    expect(draftFor(undefined, place(), PLAIN)).toBe(NEW_DRAFT);
  });
});
