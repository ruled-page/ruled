import { describe, expect, it } from 'vitest';
import type { EngineValue } from '../src/core/calc';
import { evaluateConditionalFormats, parseConditionalFormats, type CfContext } from '../src/core/conditionalFormat';
import { parseA1, parseRange, toA1 } from '../src/core/refs';
import { StyleBook } from '../src/core/styles';
import type { BorderEdge, CfResult, CfRule } from '../src/core/types';
import { parseXml } from '../src/core/xml';

const BOLD = { bold: true };

/** A rule on `ref` that sets bold, unless the options say otherwise. */
function rule(type: string, ref: string, options: Partial<CfRule> = {}): CfRule {
  return { rects: ref.split(' ').map((r) => parseRange(r)!), type, priority: 1, stopIfTrue: false, formulas: [], style: BOLD, ...options };
}

/**
 * A sheet of values by address. Formulas are worked out by a stand-in that
 * reads "ref>n", "ref", "$ref" and numbers, and records what it was asked.
 */
function sheet(values: Record<string, EngineValue>) {
  const cells = new Map(Object.entries(values).map(([ref, v]) => [ref, v]));
  const asked: string[] = [];
  let maxRow = -1;
  let maxCol = -1;
  for (const ref of cells.keys()) {
    const { row, col } = parseA1(ref)!;
    maxRow = Math.max(maxRow, row);
    maxCol = Math.max(maxCol, col);
  }
  const ctx: CfContext = {
    value: (row, col) => cells.get(toA1(row, col)) ?? null,
    calculate: (f) => {
      asked.push(f);
      const gt = /^\$?([A-Z]+)\$?(\d+)>(\d+)$/.exec(f);
      if (gt) {
        const v = cells.get(gt[1] + gt[2]);
        return typeof v === 'number' && v > Number(gt[3]);
      }
      const ref = /^\$?([A-Z]+)\$?(\d+)$/.exec(f);
      if (ref) return [[cells.get(ref[1] + ref[2]) ?? null]];
      return { error: '#NAME?' };
    },
    maxRow,
    maxCol,
  };
  return { ctx, asked };
}

/** The cells each format lands on, by address. */
function formatted(r: CfResult) {
  return Object.fromEntries(
    r.cells.map((c) => [toA1(c.row, c.col), { ...(c.style !== undefined ? r.styles[c.style] : {}), ...(c.bar ? { bar: c.bar } : {}), ...(c.icon ? { icon: c.icon.index } : {}), ...(c.hideValue ? { hidden: true } : {}) }]),
  );
}

const matching = (rules: CfRule[], values: Record<string, EngineValue>) => Object.keys(formatted(evaluateConditionalFormats(rules, 0, sheet(values).ctx))).sort();

describe('cell value rules', () => {
  const values: Record<string, EngineValue> = { A1: 1, A2: 5, A3: 10, A4: 'abc', A6: { error: '#N/A' }, A7: true, A8: 'ABC' };
  const cellIs = (operator: string, ...formulas: string[]) => matching([rule('cellIs', 'A1:A8', { operator, formulas })], values);

  it('compare as Excel does: blanks as zero, text above numbers, logicals above text, errors never', () => {
    expect(cellIs('between', '2', '10')).toEqual(['A2', 'A3']);
    expect(cellIs('notBetween', '2', '10')).toEqual(['A1', 'A4', 'A5', 'A7', 'A8']);
    expect(cellIs('greaterThan', '5')).toEqual(['A3', 'A4', 'A7', 'A8']);
    expect(cellIs('lessThan', '5')).toEqual(['A1', 'A5']);
    expect(cellIs('greaterThanOrEqual', '5')).toEqual(['A2', 'A3', 'A4', 'A7', 'A8']);
    expect(cellIs('lessThanOrEqual', '5')).toEqual(['A1', 'A2', 'A5']);
    expect(cellIs('notEqual', '5')).toEqual(['A1', 'A3', 'A4', 'A5', 'A7', 'A8']);
  });

  it('match text ignoring case, and treat a blank as empty text beside text', () => {
    expect(cellIs('equal', '"abc"')).toEqual(['A4', 'A8']);
    expect(cellIs('lessThan', '"abd"')).toEqual(['A1', 'A2', 'A3', 'A4', 'A5', 'A8']);
    expect(cellIs('equal', '""')).toEqual(['A5']);
    expect(cellIs('equal', '"say ""hi"""')).toEqual([]);
  });

  it('read an empty bound as zero, empty text or FALSE, whichever the cell holds', () => {
    expect(matching([rule('cellIs', 'A1:A4', { operator: 'equal', formulas: ['$Z$1'] })], { A1: 0, A2: '', A3: 'x', A4: false })).toEqual(['A1', 'A2', 'A4']);
  });

  it('need their bounds: one missing, an unknown operator, or an error bound matches nothing', () => {
    expect(cellIs('between', '2')).toEqual([]);
    expect(cellIs('greaterThan')).toEqual([]);
    expect(cellIs('greaterThan', '')).toEqual([]);
    expect(cellIs('approximately', '5')).toEqual([]);
    expect(cellIs('greaterThan', 'NOPE()')).toEqual([]);
  });

  it('work out an absolute bound once, and a relative one for each cell', () => {
    const values = { A1: 1, A2: 5, A3: 9, B1: 2, B2: 6, B3: 1, C1: 4 };
    const once = sheet(values);
    const r1 = evaluateConditionalFormats([rule('cellIs', 'A1:A3', { operator: 'greaterThan', formulas: ['$C$1'] })], 0, once.ctx);
    expect(Object.keys(formatted(r1))).toEqual(['A2', 'A3']);
    expect(once.asked).toEqual(['$C$1']);
    const each = sheet(values);
    const r2 = evaluateConditionalFormats([rule('cellIs', 'A1:A3', { operator: 'lessThan', formulas: ['B1'] })], 0, each.ctx);
    expect(Object.keys(formatted(r2))).toEqual(['A1', 'A2']);
    expect(each.asked).toEqual(['B1', 'B2', 'B3']);
  });
});

describe('text, blank and error rules', () => {
  const values: Record<string, EngineValue> = { A1: 'Overdue', A2: 'due soon', A3: '  ', A5: { error: '#DIV/0!' }, A6: 42 };
  const textRule = (type: string, text?: string) => matching([rule(type, 'A1:A6', { text })], values);

  it('read the rule text when the file has no formula for them', () => {
    expect(textRule('containsText', 'DUE')).toEqual(['A1', 'A2']);
    expect(textRule('notContainsText', 'due')).toEqual(['A3', 'A4', 'A5', 'A6']);
    expect(textRule('beginsWith', 'due')).toEqual(['A2']);
    expect(textRule('endsWith', 'due')).toEqual(['A1']);
    expect(textRule('containsText', '4')).toEqual(['A6']);
    expect(textRule('containsBlanks')).toEqual(['A3', 'A4']);
    expect(textRule('notContainsBlanks')).toEqual(['A1', 'A2', 'A5', 'A6']);
    expect(textRule('containsErrors')).toEqual(['A5']);
    expect(textRule('notContainsErrors')).toEqual(['A1', 'A2', 'A3', 'A4', 'A6']);
    // Dates need the formula Excel writes with them.
    expect(textRule('timePeriod')).toEqual([]);
  });

  it('use the formula when there is one, true for any non-zero number', () => {
    const { ctx, asked } = sheet({ A1: 1, A2: 3, A3: 5, B1: 0, B2: 2, B3: 'x', C3: 'c' });
    const r = evaluateConditionalFormats([rule('containsText', 'A1:A3', { text: 'ignored', formulas: ['A1>2'] }), rule('expression', 'C1:C3', { formulas: ['B1'] })], 0, ctx);
    expect(Object.keys(formatted(r)).sort()).toEqual(['A2', 'A3', 'C2']);
    expect(asked).toEqual(['A1>2', 'A2>2', 'A3>2', 'B1', 'B2', 'B3']);
  });

  it('take a constant formula as it is', () => {
    expect(matching([rule('expression', 'A1:A2', { formulas: ['1'] })], { A1: 'x', A2: 'y' })).toEqual(['A1', 'A2']);
    expect(matching([rule('expression', 'A1:A2', { formulas: ['0'] })], { A1: 'x', A2: 'y' })).toEqual([]);
  });
});

describe('ranking rules', () => {
  const values = Object.fromEntries([3, 9, 1, 7, 5, 'x', null, 8, 2, 6].map((v, i) => [`A${i + 1}`, v as EngineValue]));
  const one = (r: CfRule) => matching([r], values);

  it('pick the top or bottom N, or N percent', () => {
    expect(one(rule('top10', 'A1:A10', { rank: 2 }))).toEqual(['A2', 'A8']);
    expect(one(rule('top10', 'A1:A10', { rank: 3, bottom: true }))).toEqual(['A1', 'A3', 'A9']);
    // 25% of 8 numbers is 2; a percentage too small for one cell still picks one.
    expect(one(rule('top10', 'A1:A10', { rank: 25, percent: true }))).toEqual(['A2', 'A8']);
    expect(one(rule('top10', 'A1:A10', { rank: 1, percent: true, bottom: true }))).toEqual(['A3']);
    // Without a rank, Excel's default of ten takes every number here.
    expect(one(rule('top10', 'A1:A10')).length).toBe(8);
    expect(matching([rule('top10', 'A1:A3', { rank: 1 })], { A1: 'a', A2: 'b' })).toEqual([]);
  });

  it('compare with the average, optionally including it or a number of standard deviations away', () => {
    // The numbers 1, 2, 3, 5, 6, 7, 8, 9 average 5.125.
    expect(one(rule('aboveAverage', 'A1:A10', { aboveAverage: true }))).toEqual(['A10', 'A2', 'A4', 'A8']);
    expect(one(rule('aboveAverage', 'A1:A10', { aboveAverage: false }))).toEqual(['A1', 'A3', 'A5', 'A9']);
    expect(matching([rule('aboveAverage', 'A1:A3', { aboveAverage: true, equalAverage: true })], { A1: 2, A2: 4, A3: 6 })).toEqual(['A2', 'A3']);
    expect(matching([rule('aboveAverage', 'A1:A3', { aboveAverage: false, equalAverage: true })], { A1: 2, A2: 4, A3: 6 })).toEqual(['A1', 'A2']);
    // Two standard deviations above or below: only the outliers.
    const spread = { A1: 0, A2: 0, A3: 0, A4: 0, A5: 0, A6: 0, A7: 0, A8: 0, A9: 0, A10: 100 };
    expect(matching([rule('aboveAverage', 'A1:A10', { aboveAverage: true, stdDev: 2 })], spread)).toEqual(['A10']);
    expect(matching([rule('aboveAverage', 'A1:A10', { aboveAverage: false, stdDev: 2 })], spread)).toEqual([]);
    expect(matching([rule('aboveAverage', 'A1:A2', { aboveAverage: true })], { A1: 'only text' })).toEqual([]);
  });

  it('find duplicate and unique values, matching text without case and ignoring blanks and errors', () => {
    const v: Record<string, EngineValue> = { A1: 'Apple', A2: 'apple', A3: 1, A4: true, A5: true, A6: '', A8: { error: '#N/A' }, A9: { error: '#N/A' }, A10: 2 };
    expect(matching([rule('duplicateValues', 'A1:A10')], v)).toEqual(['A1', 'A2', 'A4', 'A5']);
    expect(matching([rule('uniqueValues', 'A1:A10')], v)).toEqual(['A10', 'A3']);
  });
});

describe('colour scales, data bars and icon sets', () => {
  const values = { A1: 0, A2: 50, A3: 100, A4: 'n/a' };
  const fills = (r: CfRule, v: Record<string, EngineValue> = values) =>
    Object.fromEntries(Object.entries(formatted(evaluateConditionalFormats([r], 0, sheet(v).ctx))).map(([k, f]) => [k, f.fill]));
  const scale = (cfvo: NonNullable<CfRule['colorScale']>['cfvo'], colors: string[]) => rule('colorScale', 'A1:A5', { style: undefined, colorScale: { cfvo, colors } });

  it('blend two colours from the lowest number to the highest, leaving text alone', () => {
    expect(fills(scale([{ type: 'min', gte: true }, { type: 'max', gte: true }], ['#000000', '#FFFFFF']))).toEqual({ A1: '#000000', A2: '#808080', A3: '#FFFFFF' });
  });

  it('blend three colours through a midpoint', () => {
    const r = scale(
      [
        { type: 'num', val: '0', gte: true },
        { type: 'percentile', val: '50', gte: true },
        { type: 'num', val: '100', gte: true },
      ],
      ['#FF0000', '#FFFF00', '#00FF00'],
    );
    expect(fills(r, { A1: 0, A2: 25, A3: 50, A4: 75, A5: 100 })).toEqual({ A1: '#FF0000', A2: '#FF8000', A3: '#FFFF00', A4: '#80FF00', A5: '#00FF00' });
  });

  it('read thresholds from formulas, falling back to the range when one is not a number', () => {
    const r = scale(
      [
        { type: 'formula', val: '$B$1', gte: true },
        { type: 'formula', val: 'oops', gte: true },
      ],
      ['#000000', '#FFFFFF'],
    );
    // B1 holds 50: below it is all black, and the top is the highest number.
    expect(fills(r, { A1: 0, A2: 50, A3: 75, A4: 100, B1: 50 })).toMatchObject({ A1: '#000000', A2: '#000000', A3: '#808080', A4: '#FFFFFF' });
  });

  it('colour every cell with the last colour when all the numbers are the same', () => {
    expect(fills(scale([{ type: 'min', gte: true }, { type: 'max', gte: true }], ['#000000', '#FFFFFF']), { A1: 7, A2: 7 })).toEqual({ A1: '#FFFFFF', A2: '#FFFFFF' });
  });

  it('do nothing without two colours or any numbers', () => {
    expect(fills(scale([{ type: 'min', gte: true }], ['#000000']))).toEqual({});
    expect(fills(scale([{ type: 'min', gte: true }, { type: 'max', gte: true }], ['#000000', '#FFFFFF']), { A1: 'a' })).toEqual({});
  });

  const bar = (opts: Partial<NonNullable<CfRule['dataBar']>>, v: Record<string, EngineValue> = values) => {
    const r = rule('dataBar', 'A1:A4', { style: undefined, dataBar: { cfvo: [], color: '#638EC6', showValue: true, gradient: true, minLength: 10, maxLength: 90, ...opts } });
    return formatted(evaluateConditionalFormats([r], 0, sheet(v).ctx));
  };

  it('size bars between the shortest and longest lengths', () => {
    const f = bar({});
    expect([f.A1.bar?.to, f.A2.bar?.to, f.A3.bar?.to]).toEqual([0.1, 0.5, 0.9]);
    expect(f.A4).toBeUndefined();
  });

  it('measure bars from zero with automatic ends, colour negative ones and can hide the value', () => {
    const f = bar(
      { cfvo: [{ type: 'autoMin', gte: true }, { type: 'autoMax', gte: true }], minLength: 0, maxLength: 100, negativeColor: '#FF0000', showValue: false, gradient: false },
      { A1: -50, A2: 0, A3: 50 },
    );
    expect(f.A1).toEqual({ bar: { from: 0, to: 0, color: '#FF0000', gradient: false }, hidden: true });
    expect(f.A3.bar).toEqual({ from: 0, to: 1, color: '#638EC6', gradient: false });
    // Positive numbers alone: the automatic minimum is still zero.
    expect(bar({ cfvo: [{ type: 'autoMin', gte: true }, { type: 'max', gte: true }], minLength: 0, maxLength: 100 }, { A1: 20, A2: 40 }).A1.bar?.to).toBe(0.5);
  });

  it('place bars between percentage points of the range', () => {
    const f = bar({ cfvo: [{ type: 'percent', val: '50', gte: true }, { type: 'max', gte: true }], minLength: 0, maxLength: 100 });
    expect([f.A1.bar?.to, f.A2.bar?.to, f.A3.bar?.to]).toEqual([0, 0, 1]);
    expect(bar({}, { A1: 'x' })).toEqual({});
  });

  const icons = (s: Partial<NonNullable<CfRule['iconSet']>>, v: Record<string, EngineValue> = { A1: 0, A2: 33, A3: 34, A4: 67, A5: 100, A6: 'x' }) =>
    formatted(
      evaluateConditionalFormats(
        [
          rule('iconSet', 'A1:A6', {
            style: undefined,
            iconSet: { set: '3Arrows', reverse: false, showValue: true, cfvo: [{ type: 'percent', val: '0', gte: true }, { type: 'percent', val: '33', gte: true }, { type: 'percent', val: '67', gte: true }], ...s },
          }),
        ],
        0,
        sheet(v).ctx,
      ),
    );

  it('pick icons by threshold, at or above it unless the threshold says above only', () => {
    expect(Object.values(icons({})).map((c) => c.icon)).toEqual([0, 1, 1, 2, 2]);
    const above = icons({ cfvo: [{ type: 'num', val: '0', gte: true }, { type: 'num', val: '33', gte: false }, { type: 'num', val: '67', gte: false }] });
    expect(Object.values(above).map((c) => c.icon)).toEqual([0, 0, 1, 1, 2]);
  });

  it('reverse the icons, and can show them without the value', () => {
    const f = icons({ reverse: true, showValue: false });
    expect(Object.values(f).map((c) => c.icon)).toEqual([2, 1, 1, 0, 0]);
    expect(f.A1.hidden).toBe(true);
    expect(icons({}, { A1: 'x' })).toEqual({});
  });
});

describe('several rules', () => {
  it('keep what a higher-priority rule set, filling in what it left, including border edges', () => {
    const thin: BorderEdge = { width: 1, pattern: 'solid', color: '#000000' };
    const thick: BorderEdge = { width: 3, pattern: 'solid', color: '#FF0000' };
    const rules = [
      rule('expression', 'A1', { formulas: ['1'], style: { fill: '#FFFF00', border: { top: thin } } }),
      rule('expression', 'A1', { formulas: ['1'], style: { fill: '#00FF00', bold: true, border: { top: thick, bottom: thick } } }),
    ];
    const f = formatted(evaluateConditionalFormats(rules, 0, sheet({ A1: 1 }).ctx));
    expect(f.A1).toEqual({ fill: '#FFFF00', bold: true, border: { top: thin, bottom: thick } });
  });

  it('stop at a matching rule marked "stop if true", for that cell only', () => {
    const rules = [
      rule('cellIs', 'A1:A2', { operator: 'greaterThan', formulas: ['5'], stopIfTrue: true, style: { fill: '#FF0000' } }),
      rule('expression', 'A1:A2', { formulas: ['1'], style: { italic: true } }),
      rule('dataBar', 'A1:A2', { style: undefined, dataBar: { cfvo: [], color: '#638EC6', showValue: true, gradient: true, minLength: 0, maxLength: 100 } }),
    ];
    const f = formatted(evaluateConditionalFormats(rules, 0, sheet({ A1: 9, A2: 1 }).ctx));
    expect(f.A1).toEqual({ fill: '#FF0000' });
    expect(f.A2).toEqual({ italic: true, bar: { from: 0, to: 0, color: '#638EC6', gradient: true } });
  });

  it('let a rule with no format of its own shield cells from the rules after it', () => {
    const rules = [
      rule('cellIs', 'A1:A2', { operator: 'equal', formulas: ['0'], stopIfTrue: true, style: undefined }),
      rule('expression', 'A1:A2', { formulas: ['1'] }),
    ];
    expect(formatted(evaluateConditionalFormats(rules, 0, sheet({ A1: 0, A2: 1 }).ctx))).toEqual({ A2: BOLD });
  });

  it('share identical formats between cells and skip ranges past the used area', () => {
    const r = evaluateConditionalFormats([rule('expression', 'A1:B2 Z100:Z200', { formulas: ['1'] })], 0, sheet({ B2: 1 }).ctx);
    expect(r.styles).toEqual([BOLD]);
    expect(r.cells.map((c) => [toA1(c.row, c.col), c.style])).toEqual([
      ['A1', 0],
      ['B1', 0],
      ['A2', 0],
      ['B2', 0],
    ]);
    expect(evaluateConditionalFormats([rule('expression', 'C3:D4', { formulas: ['1'] })], 0, sheet({ A1: 1 }).ctx).cells).toEqual([]);
  });

  it('skip formula rules over more cells than the evaluation limit, and count them', () => {
    const big = 'A1:A200001';
    const ctx: CfContext = { value: (row) => (row === 7 ? 'flag' : null), calculate: () => true, maxRow: 200_000, maxCol: 0 };
    const r = evaluateConditionalFormats(
      [rule('expression', big, { formulas: ['$B$1'] }), rule('containsText', big, { text: 'flag', priority: 2, style: { italic: true } })],
      0,
      ctx,
    );
    expect(r.skipped).toBe(1);
    // Rules that need no formula still apply.
    expect(r.cells).toEqual([{ row: 7, col: 0, style: 0 }]);
    expect(r.styles).toEqual([{ italic: true }]);
  });
});

describe('reading rules', () => {
  const styles = new StyleBook(
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>',
    undefined,
    'en-US',
  );
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="x14" xmlns:xm="xm"';
  const read = (body: string) => parseConditionalFormats(parseXml(`<worksheet ${NS}>${body}</worksheet>`), styles);

  it('reads options, defaults and inline formats, in priority order', () => {
    const rules = read(
      '<conditionalFormatting sqref="B1:B5 D1:D5">' +
        '<cfRule type="top10" priority="5" rank="20" percent="1" bottom="true" dxfId="0"/>' +
        '<cfRule type="aboveAverage" priority="6" aboveAverage="0" equalAverage="1" stdDev="2" dxfId="0"/>' +
        '<cfRule priority="2" stopIfTrue="1"><dxf><font><i/></font></dxf><formula> A1&gt;1 </formula></cfRule>' +
        '<cfRule type="colorScale"><colorScale><cfvo type="min"/><cfvo val="5"/><color rgb="FFFF0000"/><color/></colorScale></cfRule>' +
        '<cfRule type="iconSet" priority="9"><iconSet><cfvo type="percent" val="0"/><cfvo type="percent" val="50" gte="0"/></iconSet></cfRule>' +
        '</conditionalFormatting>' +
        // A block without a range has nothing to format.
        '<conditionalFormatting sqref=""><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>',
    );
    expect(rules.map((r) => [r.type, r.priority])).toEqual([
      ['colorScale', 0],
      ['expression', 2],
      ['top10', 5],
      ['aboveAverage', 6],
      ['iconSet', 9],
    ]);
    const [scale, expr, top, avg, icons] = rules;
    expect(scale.colorScale).toEqual({ cfvo: [{ type: 'min', val: undefined, gte: true }, { type: 'num', val: '5', gte: true }], colors: ['#FF0000', '#FFFFFF'] });
    expect(expr).toMatchObject({ rects: [{ r1: 0, c1: 1, r2: 4, c2: 1 }, { r1: 0, c1: 3, r2: 4, c2: 3 }], stopIfTrue: true, style: { italic: true }, formulas: ['A1>1'] });
    expect(top).toMatchObject({ rank: 20, percent: true, bottom: true, style: { bold: true } });
    expect(avg).toMatchObject({ aboveAverage: false, equalAverage: true, stdDev: 2 });
    expect(icons.iconSet).toEqual({ set: '3TrafficLights1', reverse: false, showValue: true, cfvo: [{ type: 'percent', val: '0', gte: true }, { type: 'percent', val: '50', gte: false }] });
  });

  it('merges Excel 2010 options into the rule they extend, and reads rules that exist only there', () => {
    const rules = read(
      '<conditionalFormatting sqref="A1:A10">' +
        '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext><x14:id>{A}</x14:id></ext></extLst></cfRule>' +
        '<cfRule type="iconSet" priority="2"><iconSet><cfvo type="percent" val="0"/><cfvo type="percent" val="50"/></iconSet><extLst><ext><x14:id>{B}</x14:id></ext></extLst></cfRule>' +
        '</conditionalFormatting>' +
        '<extLst><ext><x14:conditionalFormattings>' +
        '<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{A}"><x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>' +
        '<x14:negativeFillColor rgb="FFFF0000"/></x14:dataBar></x14:cfRule><xm:sqref>A1:A10</xm:sqref></x14:conditionalFormatting>' +
        '<x14:conditionalFormatting><x14:cfRule type="iconSet" id="{B}"><x14:iconSet iconSet="3Arrows" reverse="1" showValue="0"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>' +
        '<x14:cfvo type="num" gte="0"><xm:f> 5 </xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>A1:A10</xm:sqref></x14:conditionalFormatting>' +
        '<x14:conditionalFormatting><x14:cfRule type="expression" priority="3" id="{C}"><xm:f>Other!$A$1&gt;0</xm:f><x14:dxf><font><u/></font></x14:dxf></x14:cfRule><xm:sqref>C1:C3</xm:sqref></x14:conditionalFormatting>' +
        '<x14:conditionalFormatting><x14:cfRule type="dataBar" priority="4" id="{D}"><x14:dataBar><x14:cfvo type="min"/><x14:cfvo type="max"/><x14:fillColor rgb="FF00FF00"/></x14:dataBar></x14:cfRule><xm:sqref>E1:E3</xm:sqref></x14:conditionalFormatting>' +
        '</x14:conditionalFormattings></ext></extLst>',
    );
    expect(rules.map((r) => r.type)).toEqual(['dataBar', 'iconSet', 'expression', 'dataBar']);
    const [bar, icons, expr, own] = rules;
    // The 2010 options win, except the colour, which the main rule keeps.
    expect(bar.dataBar).toEqual({
      cfvo: [{ type: 'autoMin', val: undefined, gte: true }, { type: 'autoMax', val: undefined, gte: true }],
      color: '#638EC6',
      negativeColor: '#FF0000',
      showValue: true,
      gradient: false,
      minLength: 0,
      maxLength: 100,
    });
    expect(icons.iconSet).toEqual({ set: '3Arrows', reverse: true, showValue: false, cfvo: [{ type: 'percent', val: '0', gte: true }, { type: 'num', val: '5', gte: false }] });
    expect(expr).toMatchObject({ rects: [{ r1: 0, c1: 2, r2: 2, c2: 2 }], formulas: ['Other!$A$1>0'], style: { underline: true } });
    // A bar defined only in 2010 form uses the full width and its fill colour.
    expect(own.dataBar).toMatchObject({ color: '#00FF00', minLength: 0, maxLength: 100 });
  });
});
