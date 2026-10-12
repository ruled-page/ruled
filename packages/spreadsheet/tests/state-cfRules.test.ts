import { describe, expect, it } from 'vitest';
import type { CfRule } from '../src/core/types';
import { english } from '../src/i18n';
import { CF_GALLERY, CF_PRESETS, describeCfRule, galleryRule, quickOf, quickRule, type CfQuick } from '../src/state/cfRules';
import type { DraftPlace } from '../src/state/validationDraft';

const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
/** B2:B10, with B4 active. */
const place = (locale = 'en-US'): DraftPlace => ({ active: { row: 3, col: 1 }, rects: [rect(1, 1, 9, 1)], locale, date1904: false });
const red = CF_PRESETS[0].style;
const made = (kind: CfQuick, values: string[], p = place()) => {
  const r = quickRule(kind, values, red, p);
  if ('problem' in r) throw new Error(r.problem.message);
  return r.rule;
};

describe("Excel's quick conditional formatting rules", () => {
  it('compare cells with numbers in the locale, text, or formulas kept for the top-left cell', () => {
    expect(made('greaterThan', ['1,5'], place('de-DE'))).toMatchObject({ type: 'cellIs', operator: 'greaterThan', formulas: ['1.5'], style: red });
    expect(made('equal', ['Say "hi"'])).toMatchObject({ formulas: ['"Say ""hi"""'] });
    expect(made('between', ['=C4', '10'])).toMatchObject({ operator: 'between', formulas: ['C2', '10'] });
  });

  it('find text, duplicates, the top and bottom, the average, and formulas', () => {
    expect(made('textContains', ['a"b'])).toMatchObject({ type: 'containsText', operator: 'containsText', text: 'a"b', formulas: ['NOT(ISERROR(SEARCH("a""b",B2)))'] });
    expect(made('duplicate', [])).toMatchObject({ type: 'duplicateValues' });
    expect(made('bottomPercent', ['15'])).toMatchObject({ type: 'top10', rank: 15, percent: true, bottom: true });
    expect(made('top', ['3'])).not.toHaveProperty('percent');
    expect(made('belowAverage', [])).toMatchObject({ type: 'aboveAverage', aboveAverage: false });
    expect(made('formula', ['=MOD(B4,2)=0'])).toMatchObject({ type: 'expression', formulas: ['MOD(B2,2)=0'] });
  });

  it('say which box is wrong', () => {
    const problem = (kind: CfQuick, values: string[]) => {
      const r = quickRule(kind, values, red, place());
      return 'problem' in r ? r.problem : undefined;
    };
    expect(problem('between', ['1', ''])).toEqual({ message: 'cf.problem.value', box: 1 });
    expect(problem('textContains', [' '])).toEqual({ message: 'cf.problem.text', box: 0 });
    expect(problem('topPercent', ['101'])?.message).toBe('cf.problem.percent');
    expect(problem('top', ['2.5'])?.message).toBe('cf.problem.rank');
    expect(problem('formula', ['=SUM('])?.message).toBe('cf.problem.formula');
  });

  it('open again as they were made, values as typed, or not at all when made another way', () => {
    for (const [kind, values] of [
      ['greaterThan', ['1.5']],
      ['equal', ['Pear']],
      ['between', ['=C4', '10']],
      ['textContains', ['an']],
      ['duplicate', []],
      ['topPercent', ['15']],
      ['bottom', ['3']],
      ['aboveAverage', []],
      ['formula', ['=MOD(B4,2)=0']],
    ] as [CfQuick, string[]][]) {
      expect(quickOf(made(kind, values), place()), kind).toEqual({ kind, values });
    }
    expect(quickOf({ ...made('greaterThan', ['1']), operator: 'notEqual' }, place())).toBeUndefined();
    expect(quickOf(galleryRule([rect(0, 0)], CF_GALLERY.dataBars[0]), place())).toBeUndefined();
    expect(quickOf({ ...made('aboveAverage', []), stdDev: 1 }, place())).toBeUndefined();
  });
});

describe("Excel's galleries and Rules Manager", () => {
  it('make data bars, color scales and icon sets with its usual thresholds', () => {
    expect(galleryRule([rect(0, 0)], CF_GALLERY.dataBars[1])).toMatchObject({ type: 'dataBar', dataBar: { color: '#63C384', minLength: 10, maxLength: 90 } });
    expect(galleryRule([rect(0, 0)], CF_GALLERY.colorScales[0]).colorScale?.cfvo.map((c) => c.type)).toEqual(['min', 'percentile', 'max']);
    expect(galleryRule([rect(0, 0)], CF_GALLERY.colorScales[2]).colorScale?.cfvo.map((c) => c.type)).toEqual(['min', 'max']);
    expect(galleryRule([rect(0, 0)], CF_GALLERY.iconSets[0])).toMatchObject({ type: 'iconSet', iconSet: { set: '3Arrows' } });
  });

  it('describe each rule in a line', () => {
    const rule = (r: Partial<CfRule>): CfRule => ({ rects: [], type: 'cellIs', priority: 1, stopIfTrue: false, formulas: [], ...r });
    const said = (r: Partial<CfRule>) => english.text(describeCfRule(rule(r)));
    expect(said({ operator: 'between', formulas: ['1', '5'] })).toBe('Cell value between 1 and 5');
    expect(said({ operator: 'weird' })).toBe('Cell value');
    expect(said({ type: 'containsText', text: 'an' })).toBe('Cell value contains “an”');
    expect(said({ type: 'top10', rank: 5, percent: true, bottom: true })).toBe('Bottom 5%');
    expect(said({ type: 'top10' })).toBe('Top 10');
    expect(said({ type: 'expression', formulas: ['A1>0'] })).toBe('Formula: =A1>0');
    expect(['uniqueValues', 'duplicateValues', 'colorScale', 'dataBar', 'iconSet', 'timePeriod'].map((type) => said({ type }))).toEqual([
      'Unique values',
      'Duplicate values',
      'Graded color scale',
      'Data bar',
      'Icon set',
      'Rule of type timePeriod',
    ]);
    expect(said({ type: 'aboveAverage', aboveAverage: false })).toBe('Below average');
  });
});
