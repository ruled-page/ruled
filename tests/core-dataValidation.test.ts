import { describe, expect, it } from 'vitest';
import type { EngineValue } from '../src/core/calc';
import { checkRule, describeRule, evaluateListFormula, type Evaluator } from '../src/core/dataValidation';
import type { CellContent, RuleValidation } from '../src/core/types';
import type { ListContext } from '../src/core/validations';
import { english } from '../src/i18n';

/** An engine stand-in: formulas (with "=") it knows the answer to, and a record of what it was asked. */
function engine(answers: Record<string, EngineValue | EngineValue[][]> = {}) {
  const asked: string[] = [];
  const held: CellContent[] = [];
  const ev: Evaluator = {
    calculate: (f) => {
      asked.push(f);
      return f in answers ? answers[f] : { error: '#NAME?' };
    },
    withValue: (_s, _r, _c, content, fn) => {
      held.push(content);
      return fn();
    },
  };
  return { ev, asked, held };
}

const rule = (kind: RuleValidation['kind'], operator: RuleValidation['operator'], formula1: string, formula2?: string): RuleValidation => ({
  rects: [{ r1: 1, c1: 0, r2: 9, c2: 0 }],
  kind,
  operator,
  formula1,
  formula2,
  allowBlank: true,
  showError: true,
  errorStyle: 'stop',
});

const num = (value: number): CellContent => ({ kind: 'number', value });
const str = (value: string): CellContent => ({ kind: 'string', value });

describe('number rules', () => {
  const check = (r: RuleValidation, content: CellContent, ev = engine().ev) => checkRule(r, 0, 1, 0, content, '', ev);

  it('compare typed numbers with each operator', () => {
    const cases: [RuleValidation['operator'], number, boolean][] = [
      ['notBetween', 5, false],
      ['notBetween', 11, true],
      ['equal', 1, true],
      ['notEqual', 1, false],
      ['greaterThan', 2, true],
      ['lessThan', 1, false],
      ['greaterThanOrEqual', 1, true],
      ['lessThanOrEqual', 0, true],
    ];
    for (const [op, v, ok] of cases) expect(check(rule('decimal', op, '1', '10'), num(v)), `${op} ${v}`).toBe(ok);
    // Bounds given the wrong way round still make a range.
    expect(check(rule('decimal', 'between', '10', '1'), num(5))).toBe(true);
  });

  it('turn down fractions for whole numbers, and text for any number rule', () => {
    expect(check(rule('whole', 'between', '1', '10'), num(5.5))).toBe(false);
    expect(check(rule('whole', 'between', '1', '10'), num(5))).toBe(true);
    expect(check(rule('decimal', 'greaterThan', '0'), str('abc'))).toBe(false);
  });

  it('accept empty cells and formulas, as Excel checks only typed values', () => {
    expect(check(rule('whole', 'equal', '1'), { kind: 'empty' })).toBe(true);
    expect(check(rule('whole', 'equal', '1'), { kind: 'formula', formula: 'A1' })).toBe(true);
  });

  it('work out bounds given as formulas, moved to the cell being checked', () => {
    const { ev, asked } = engine({ '=B5': 3, '=$C$1': true, '=$C$2': null });
    expect(checkRule(rule('decimal', 'greaterThan', 'B2'), 0, 4, 0, num(4), '4', ev)).toBe(true);
    expect(asked).toEqual(['=B5']);
    // TRUE counts as 1 and an empty cell as 0.
    expect(check(rule('decimal', 'equal', '$C$1'), num(1), ev)).toBe(true);
    expect(check(rule('decimal', 'equal', '$C$2'), num(0), ev)).toBe(true);
  });

  it('leave values to Excel when a bound cannot be worked out here', () => {
    expect(check(rule('decimal', 'lessThan', 'MYSTERY()'), num(1e9))).toBe(true);
    expect(check(rule('decimal', 'between', '1'), num(1e9))).toBe(true);
    // A one-bound operator doesn't need the second.
    expect(check(rule('decimal', 'lessThan', '5'), num(1e9))).toBe(false);
  });
});

describe('text length and custom rules', () => {
  it('measure typed text, or the typed form of a number', () => {
    const r = rule('textLength', 'lessThanOrEqual', '3');
    expect(checkRule(r, 0, 1, 0, str('abcd'), 'abcd', engine().ev)).toBe(false);
    expect(checkRule(r, 0, 1, 0, str('abc'), 'abc', engine().ev)).toBe(true);
    expect(checkRule(r, 0, 1, 0, num(12345), ' 12345 ', engine().ev)).toBe(false);
  });

  it('run a custom formula with the typed value in place', () => {
    const r = rule('custom', 'between', 'A2>0');
    const results: [EngineValue, boolean][] = [
      [true, true],
      [2, true],
      [0, false],
      [false, false],
      [{ error: '#ERROR!' }, true],
      [{ error: '#NAME?', message: 'Function MYFN is not recognized' }, true],
      [{ error: '#VALUE!' }, false],
    ];
    for (const [answer, ok] of results) {
      const { ev, held } = engine({ '=A6>0': answer });
      expect(checkRule(r, 0, 5, 0, num(7), '7', ev), JSON.stringify(answer)).toBe(ok);
      expect(held).toEqual([num(7)]);
    }
  });
});

describe('list formulas', () => {
  const ctx: ListContext = { sheetIndex: new Map([['data', 1]]), ownSheet: 0, names: [] };

  it('resolve OFFSET to the range it points at, with the base size by default', () => {
    const { ev } = engine({ '=1': 1, '=0': 0, '=3': 3, '=2': 2 });
    expect(evaluateListFormula('OFFSET(Data!$A$1,1,0,3,1)', 0, ctx, ev)).toEqual({ kind: 'range', sheet: 1, rect: { r1: 1, c1: 0, r2: 3, c2: 0 } });
    expect(evaluateListFormula('OFFSET($B$2:$C$3,1,2)', 0, ctx, ev)).toEqual({ kind: 'range', sheet: 0, rect: { r1: 2, c1: 3, r2: 3, c2: 4 } });
  });

  it('give no choices for an empty OFFSET, and none at all for one off the sheet or not from a range', () => {
    const { ev } = engine({ '=0': 0, '=-5': -5, '=1': 1 });
    expect(evaluateListFormula('OFFSET($A$1,0,0,0,1)', 0, ctx, ev)).toEqual({ kind: 'values', values: [] });
    expect(evaluateListFormula('OFFSET($A$1,-5,0)', 0, ctx, ev)).toBeUndefined();
    expect(evaluateListFormula('OFFSET("x",1,0)', 0, ctx, ev)).toBeUndefined();
  });

  it('resolve INDIRECT through the text its argument works out to', () => {
    const { ev } = engine({ '="Data!B1:B3"': 'Data!B1:B3', '=$Z$1': 7, '="nonsense!!"': 'nonsense!!' });
    expect(evaluateListFormula('INDIRECT("Data!B1:B3")', 0, ctx, ev)).toEqual({ kind: 'range', sheet: 1, rect: { r1: 0, c1: 1, r2: 2, c2: 1 } });
    expect(evaluateListFormula('INDIRECT($Z$1)', 0, ctx, ev)).toBeUndefined();
    expect(evaluateListFormula('INDIRECT("nonsense!!")', 0, ctx, ev)).toBeUndefined();
  });

  it('list the distinct values of any other formula, leaving out blanks and errors', () => {
    const { ev } = engine({ '=CHOOSE(1,X1:X5)': [['a', 2], [null, 'a'], [true, ''], [{ error: '#N/A' }, false]] });
    expect(evaluateListFormula('CHOOSE(1,X1:X5)', 0, ctx, ev)).toEqual({ kind: 'values', values: ['a', '2', 'TRUE', 'FALSE'] });
    const one = engine({ '=$A$1': 'only' });
    expect(evaluateListFormula('$A$1', 0, ctx, one.ev)).toEqual({ kind: 'values', values: ['only'] });
  });

  it('give up on formulas it cannot read', () => {
    expect(evaluateListFormula('SUM(1,', 0, ctx, engine().ev)).toBeUndefined();
  });
});

describe('describing rules', () => {
  it('writes date and time bounds as people type them', () => {
    expect(english.text(describeRule(rule('date', 'between', '46023', '46053')))).toBe('Date between 2026-01-01 and 2026-01-31');
    expect(english.text(describeRule(rule('time', 'greaterThanOrEqual', '0.375')))).toBe('Time at least 09:00');
    expect(english.text(describeRule(rule('decimal', 'lessThan', '$B$1')))).toBe('Number less than $B$1');
    expect(english.text(describeRule(rule('textLength', 'notEqual', '5')))).toBe('Text length not equal to 5');
    expect(english.text(describeRule(rule('custom', 'between', 'ISNUMBER(A2)')))).toBe('Excel checks this cell with =ISNUMBER(A2)');
  });
});
