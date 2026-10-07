import { describe, expect, it } from 'vitest';
import { checkRule, describeRule } from '../src/core/dataValidation';
import { WorkbookSession } from '../src/core/session';
import type { RuleValidation } from '../src/core/types';
import { parseValidations } from '../src/core/validations';
import { parseXml } from '../src/core/xml';
import { generateSchema } from '../src/schema/generate';
import { SheetView } from '../src/state/sheetView';
import { listAllows, listOptions, listValidationAt } from '../src/state/validation';
import { english } from '../src/i18n';
import { sample } from './helpers';

const ctx = {
  sheetIndex: new Map([
    ['data', 0],
    ['my lists', 1],
  ]),
  ownSheet: 0,
  names: [
    { name: 'Regions', formula: "'My Lists'!$C$2:$C$5" },
    { name: 'Regions', formula: '$Z$1:$Z$2', localSheet: 0 },
  ],
};

describe('parsing list validations', () => {
  it('reads every source kind, the dropdown flag and both XML forms', () => {
    const xml = `<worksheet xmlns="main" xmlns:x14="x14" xmlns:xm="xm">
      <dataValidations count="6">
        <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="B2:B10 D2"><formula1>"North,South,""Far"" East"</formula1></dataValidation>
        <dataValidation type="list" sqref="C2:C10" errorStyle="warning" error="Not listed"><formula1>$H$1:$H$4</formula1></dataValidation>
        <dataValidation type="list" sqref="E2" showDropDown="1"><formula1>'My Lists'!$A$1:$A$9</formula1></dataValidation>
        <dataValidation type="list" sqref="F2"><formula1>Regions</formula1></dataValidation>
        <dataValidation type="list" sqref="G2"><formula1>OFFSET($A$1,0,0,5)</formula1></dataValidation>
        <dataValidation type="whole" sqref="K2"><formula1>1</formula1></dataValidation>
      </dataValidations>
      <extLst><ext><x14:dataValidations count="1"><x14:dataValidation type="list" showErrorMessage="1">
        <x14:formula1><xm:f>'My Lists'!$C$2:$C$5</xm:f></x14:formula1><xm:sqref>J2:J20</xm:sqref>
      </x14:dataValidation></x14:dataValidations></ext></extLst>
    </worksheet>`;
    const { lists, rules, other } = parseValidations(parseXml(xml), ctx);
    expect(other).toBe(0);
    expect(rules).toMatchObject([{ kind: 'whole', operator: 'between', formula1: '1' }]);
    expect(lists.map((l) => l.source)).toEqual([
      { kind: 'values', values: ['North', 'South', '"Far" East'] },
      { kind: 'range', sheet: 0, rect: { r1: 0, c1: 7, r2: 3, c2: 7 } },
      { kind: 'range', sheet: 1, rect: { r1: 0, c1: 0, r2: 8, c2: 0 } },
      // The sheet-scoped name wins over the workbook-wide one.
      { kind: 'range', sheet: 0, rect: { r1: 0, c1: 25, r2: 1, c2: 25 } },
      { kind: 'formula', formula: 'OFFSET($A$1,0,0,5)' },
      { kind: 'range', sheet: 1, rect: { r1: 1, c1: 2, r2: 4, c2: 2 } },
    ]);
    expect(lists[0].rects).toHaveLength(2);
    expect(lists[0]).toMatchObject({ allowBlank: true, showError: true, errorStyle: 'stop', inCellDropdown: true });
    expect(lists[1]).toMatchObject({ errorStyle: 'warning', error: 'Not listed' });
    expect(lists[2].inCellDropdown).toBe(false);
  });
});

function load(file: string) {
  const s = WorkbookSession.open(sample(file), file);
  const r = s.loadResult();
  const views = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
  return { r, views };
}

describe('orders.xlsx', () => {
  it('resolves inline, cross-sheet and named lists to their values', () => {
    const { r, views } = load('orders.xlsx');
    const orders = r.sheets[0];
    const at = (row: number, col: number) => listOptions(listValidationAt(orders, row, col)!, views, r.styles, false);
    expect(at(1, 1)).toEqual(['Acme Ltd', 'Globex', 'Initech', 'Umbrella']);
    expect(at(1, 2)).toEqual(['Widget', 'Gadget', 'Gizmo', 'Doohickey', 'Thingamajig']);
    expect(at(1, 3)).toEqual(['North', 'South', 'East', 'West']);
    expect(listValidationAt(orders, 1, 4)).toMatchObject({ errorStyle: 'stop', prompt: 'Where the order is up to.' });
    expect(listValidationAt(orders, 1, 5)).toBeUndefined();
    expect(listValidationAt(orders, 0, 1)).toBeUndefined();
    const finding = r.findings.find((f) => f.kind === 'dataValidation');
    expect(english.text(finding?.detail)).toMatch(/6 list rules show a dropdown.*5 number, date, text-length or custom rules check/);
    expect(english.text(finding?.detail)).not.toMatch(/not enforced/);
  });

  it('turns list columns into enum fields when generating a schema', () => {
    const { r, views } = load('orders.xlsx');
    const fields = generateSchema(views, r.styles, 'orders.xlsx').sheets[0].tables![0].fields;
    expect(fields.filter((f) => f.type === 'enum').map((f) => [f.key, f.values?.length])).toEqual([
      ['customer', 4],
      ['product', 5],
      ['region', 4],
      ['status', 4],
    ]);
  });

  it('matches typed values like Excel', () => {
    expect(listAllows(['North', 'South'], ' north ')).toBe(true);
    expect(listAllows(['1,000', '2'], '1000')).toBe(true);
    expect(listAllows(['North'], 'Nort')).toBe(false);
  });
});

describe('Excel validation rules', () => {
  const session = () => WorkbookSession.open(sample('orders.xlsx'), 'orders.xlsx');
  /** Checks typed text in a column of the order log against the rule covering it. */
  function checker(s: WorkbookSession) {
    const rules = s.loadResult().sheets[0].ruleValidations;
    return (col: number, text: string, row = 1) => {
      const index = rules.findIndex((r) => r.rects[0].c1 === col);
      expect(index).toBeGreaterThanOrEqual(0);
      return s.checkValidation(0, row, col, text, index);
    };
  }

  it('checks whole numbers, decimals and text length', () => {
    const check = checker(session());
    expect([check(5, '10'), check(5, '1000'), check(5, '0'), check(5, '1001'), check(5, '2.5'), check(5, 'ten')]).toEqual([
      true, true, false, false, false, false,
    ]);
    expect([check(10, 'Call first'), check(10, 'x'.repeat(21)), check(10, '12345')]).toEqual([true, false, true]);
  });

  it('checks dates, and bounds that are formulas', () => {
    const s = session();
    const check = checker(s);
    expect([check(6, '2026-03-01'), check(6, '2025-12-31')]).toEqual([true, false]);
    // Discount is capped by the MaxDiscount cell, so changing the cell moves the cap.
    expect([check(8, '0.2'), check(8, '20%'), check(8, '0.6')]).toEqual([true, true, false]);
    s.apply({ type: 'input', sheet: 1, row: 1, col: 8, text: '0.75' });
    expect(check(8, '0.6')).toBe(true);
  });

  it('runs custom formulas against the typed value at the right cell, then restores it', () => {
    const s = session();
    const check = checker(s);
    expect([check(7, 'PO-12', 4), check(7, 'PO12', 4), check(7, 'PO-123456', 4)]).toEqual([true, false, false]);
    // The trial value is not left behind in the engine.
    expect(s.listSource(0, 'UPPER(H5)')).toEqual({ kind: 'values', values: ['PO-103'] });
  });

  it('accepts formulas and blanks, which Excel does not check on entry', () => {
    const check = checker(session());
    expect([check(5, '=1/2'), check(5, '')]).toEqual([true, true]);
  });

  it('leaves rules the engine cannot evaluate to Excel', () => {
    const rule = (formula1: string, kind: RuleValidation['kind'] = 'custom'): RuleValidation => ({
      kind, operator: 'greaterThan', formula1, rects: [{ r1: 0, c1: 0, r2: 9, c2: 0 }], allowBlank: true, showError: true, errorStyle: 'stop',
    });
    const ev = {
      calculate: (f: string) => (f.includes('LAMBDA') ? { error: '#NAME?', message: 'Function name LAMBDA not recognized.' } : f.includes('1/0') ? { error: '#DIV/0!' } : 0),
      withValue: <T,>(_s: number, _r: number, _c: number, _v: unknown, fn: () => T) => fn(),
    };
    const typed = { kind: 'number', value: 5 } as const;
    expect(checkRule(rule('LAMBDA(x,x)(A1)'), 0, 0, 0, typed, '5', ev)).toBe(true);
    expect(checkRule(rule('1/0'), 0, 0, 0, typed, '5', ev)).toBe(false);
    expect(checkRule(rule('LAMBDA(x,x)(1)', 'whole'), 0, 0, 0, typed, '5', ev)).toBe(true);
    // An empty bound cell is zero.
    expect(checkRule(rule('B1', 'whole'), 0, 0, 0, typed, '5', { ...ev, calculate: () => null })).toBe(true);
  });

  it('resolves OFFSET and INDIRECT lists, following changes to their source', () => {
    const s = session();
    const lists = 1;
    const offset = 'OFFSET(Lists!$A$2,0,0,COUNTA(Lists!$A:$A)-1,1)';
    expect(s.listSource(0, offset)).toEqual({ kind: 'range', sheet: lists, rect: { r1: 1, c1: 0, r2: 5, c2: 0 } });
    expect(s.listSource(0, 'INDIRECT("Gadget")')).toEqual({ kind: 'range', sheet: lists, rect: { r1: 1, c1: 5, r2: 2, c2: 5 } });
    expect(s.listSource(0, 'INDIRECT("Nothing")')).toBeUndefined();
    s.apply({ type: 'input', sheet: lists, row: 6, col: 0, text: 'Sprocket' });
    expect(s.listSource(0, offset)).toMatchObject({ rect: { r1: 1, r2: 6 } });
  });

  it('describes rules for the formula bar', () => {
    const rules = session().loadResult().sheets[0].ruleValidations;
    expect(rules.map((rule) => english.text(describeRule(rule)))).toEqual([
      'Whole number between 1 and 1000',
      'Date at least 2026-01-01',
      'Excel checks this cell with =AND(LEN(H2)<=8,ISNUMBER(SEARCH("-",H2)))',
      'Number between 0 and MaxDiscount',
      'Text length at most 20',
    ]);
  });
});

describe('kitchen-sink.xlsx', () => {
  it('finds the region list on the Data sheet', () => {
    const { r, views } = load('kitchen-sink.xlsx');
    const data = r.sheets[1];
    expect(listOptions(listValidationAt(data, 5, 1)!, views, r.styles, false)).toEqual(['North', 'South', 'East', 'West']);
  });
});
