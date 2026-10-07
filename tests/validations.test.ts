import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { parseListValidations } from '../src/core/validations';
import { parseXml } from '../src/core/xml';
import { generateSchema } from '../src/schema/generate';
import { SheetView } from '../src/state/sheetView';
import { listAllows, listOptions, listValidationAt } from '../src/state/validation';
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
    const { lists, other } = parseListValidations(parseXml(xml), ctx);
    expect(other).toBe(1);
    expect(lists.map((l) => l.source)).toEqual([
      { kind: 'values', values: ['North', 'South', '"Far" East'] },
      { kind: 'range', sheet: 0, rect: { r1: 0, c1: 7, r2: 3, c2: 7 } },
      { kind: 'range', sheet: 1, rect: { r1: 0, c1: 0, r2: 8, c2: 0 } },
      // The sheet-scoped name wins over the workbook-wide one.
      { kind: 'range', sheet: 0, rect: { r1: 0, c1: 25, r2: 1, c2: 25 } },
      { kind: 'unsupported', formula: 'OFFSET($A$1,0,0,5)' },
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
    expect(finding?.detail).toMatch(/4 list rules show a dropdown.*1 other rule/);
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

describe('kitchen-sink.xlsx', () => {
  it('finds the region list on the Data sheet', () => {
    const { r, views } = load('kitchen-sink.xlsx');
    const data = r.sheets[1];
    expect(listOptions(listValidationAt(data, 5, 1)!, views, r.styles, false)).toEqual(['North', 'South', 'East', 'West']);
  });
});
