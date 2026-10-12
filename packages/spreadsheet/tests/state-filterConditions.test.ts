import { describe, expect, it } from 'vitest';
import type { FilterCondition, FilterOperator } from '../src/core/types';
import { averageOf, conditionPasses, RULE_OPS, ruleFor, ruleOp, top10Threshold, typedRuleValue, type FilterCell } from '../src/state/filterConditions';

const cells: FilterCell[] = [
  { text: 'Apple' },
  { text: 'apricot' },
  { text: 'Banana' },
  { text: '' },
  { text: '5', n: 5 },
  { text: '12', n: 12 },
  { text: '$7.50', n: 7.5 },
];
const shown = (c: FilterCondition) => cells.filter((_, i) => conditionPasses(c, cells)[i]).map((x) => x.text);
const rule = (operator: FilterOperator, value: string, more?: { operator: FilterOperator; value: string }, and = false): FilterCondition => ({
  kind: 'custom',
  and,
  rules: more ? [{ operator, value }, more] : [{ operator, value }],
});

describe("Excel's custom filter rules", () => {
  it('match text ignoring case, and numbers by value, so "equals 7.5" finds $7.50', () => {
    expect(shown(rule('equal', 'apple'))).toEqual(['Apple']);
    expect(shown(rule('equal', '7.5'))).toEqual(['$7.50']);
    expect(shown(rule('equal', ''))).toEqual(['']);
  });

  it('read wildcards: begins with, ends with and contains', () => {
    expect(shown(rule('equal', 'ap*'))).toEqual(['Apple', 'apricot']);
    expect(shown(rule('equal', '*a'))).toEqual(['Banana']);
    expect(shown(rule('equal', '*an*'))).toEqual(['Banana']);
    expect(shown(rule('notEqual', '*p*'))).toEqual(['Banana', '', '5', '12', '$7.50']);
    expect(shown(rule('equal', '?????'))).toEqual(['Apple', '$7.50']);
  });

  it('compare numbers only with numbers, and text only with text, leaving out empty cells', () => {
    expect(shown(rule('greaterThan', '6'))).toEqual(['12', '$7.50']);
    expect(shown(rule('lessThanOrEqual', '7.5'))).toEqual(['5', '$7.50']);
    expect(shown(rule('greaterThanOrEqual', 'b'))).toEqual(['Banana']);
    expect(shown(rule('lessThan', 'b'))).toEqual(['Apple', 'apricot']);
  });

  it('hold both rules with And, either with Or', () => {
    expect(shown(rule('greaterThanOrEqual', '5', { operator: 'lessThanOrEqual', value: '8' }, true))).toEqual(['5', '$7.50']);
    expect(shown(rule('equal', 'apple', { operator: 'equal', value: '12' }))).toEqual(['Apple', '12']);
  });
});

describe('Top 10 and the average', () => {
  it('keep the largest or smallest numbers, or a percentage of them, ties included', () => {
    const top2: FilterCondition = { kind: 'top10', top: true, percent: false, count: 2 };
    expect(top10Threshold(top2, cells)).toBe(7.5);
    expect(shown(top2)).toEqual(['12', '$7.50']);
    expect(shown({ kind: 'top10', top: false, percent: false, count: 1 })).toEqual(['5']);
    // A percentage of three numbers rounds down, but keeps at least one.
    expect(shown({ kind: 'top10', top: true, percent: true, count: 10 })).toEqual(['12']);
    expect(top10Threshold(top2, [{ text: 'none' }])).toBeUndefined();
    expect(shown({ kind: 'top10', top: true, percent: false, count: 1, threshold: 5 })).toEqual(['5', '12', '$7.50']);
  });

  it('keep the numbers above or below their average, or the one it was applied with', () => {
    expect(averageOf(cells)).toBeCloseTo(8.1667, 4);
    expect(shown({ kind: 'average', above: true })).toEqual(['12']);
    expect(shown({ kind: 'average', above: false })).toEqual(['5', '$7.50']);
    expect(shown({ kind: 'average', above: true, average: 6 })).toEqual(['12', '$7.50']);
    expect(averageOf([{ text: 'x' }])).toBeUndefined();
    expect(conditionPasses({ kind: 'average', above: true }, [{ text: 'x' }])).toEqual([false]);
  });
});

describe("the Custom AutoFilter dialog's rules", () => {
  it('write "begins with", "ends with" and "contains" as Excel does, with wildcards, and read them back', () => {
    for (const op of RULE_OPS) {
      const rule = ruleFor(op, 'abc');
      expect(ruleOp(rule)).toEqual({ op, value: 'abc' });
    }
    expect(ruleFor('notContains', 'x')).toEqual({ operator: 'notEqual', value: '*x*' });
    // Wildcards typed by hand stay as typed.
    expect(ruleOp({ operator: 'equal', value: 'a?c*' })).toEqual({ op: 'equal', value: 'a?c*' });
    expect(ruleOp({ operator: 'equal', value: '*' })).toEqual({ op: 'equal', value: '*' });
  });

  it('read a typed value as the number or date it means in the locale, where the column holds them', () => {
    expect(typedRuleValue('1,5', 'greaterThan', 'number', 'de-DE', false)).toBe('1.5');
    expect(typedRuleValue('2026-01-05', 'lessThan', 'date', 'en-US', false)).toBe('46027');
    expect(typedRuleValue('05.01.2026', 'equal', 'date', 'de-DE', false)).toBe('46027');
    // A number in a date column isn't a date; text in a number column can still be looked for.
    expect(typedRuleValue('12', 'greaterThan', 'date', 'en-US', false)).toBeUndefined();
    expect(typedRuleValue('n/a', 'equal', 'number', 'en-US', false)).toBe('n/a');
    expect(typedRuleValue('n/a', 'lessThan', 'number', 'en-US', false)).toBeUndefined();
    expect(typedRuleValue('12', 'contains', 'number', 'en-US', false)).toBe('12');
    expect(typedRuleValue('Pear', 'greaterThan', 'text', 'en-US', false)).toBe('Pear');
  });
});
