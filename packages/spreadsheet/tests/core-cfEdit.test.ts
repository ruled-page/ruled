import ExcelJS from 'exceljs';
import { strToU8, zipSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import type { CfRule } from '../src/core/types';
import { english } from '../src/i18n';
import { parts } from './helpers';

const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

/** Fmt!A1:A10 has 1..10, with a "greater than 5" rule (red fill) and a formula rule; B1:B10 has a data bar. */
let bytes: Uint8Array;
beforeAll(async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Fmt');
  for (let r = 1; r <= 10; r++) ws.getCell(`A${r}`).value = r;
  ws.addConditionalFormatting({
    ref: 'A1:A10',
    rules: [
      { type: 'cellIs', operator: 'greaterThan', formulae: ['5'], priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } },
      { type: 'expression', formulae: ['MOD(A1,2)=0'], priority: 2, style: { font: { bold: true } } },
    ],
  });
  ws.addConditionalFormatting({ ref: 'B1:B10', rules: [{ type: 'dataBar', priority: 3, cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } } as ExcelJS.ConditionalFormattingRule] });
  bytes = new Uint8Array(await wb.xlsx.writeBuffer());
});

const open = (b = bytes) => WorkbookSession.open(b, 'fmt.xlsx', 'en-US');
const rulesOf = (s: WorkbookSession) => s.model.sheets[0].info.conditionalFormats;
const set = (s: WorkbookSession, rules: CfRule[]) => s.apply({ type: 'layout', sheet: 0, op: { kind: 'conditionalFormats', rules } });
const sheetXml = (s: WorkbookSession) => parts(s.save())['xl/worksheets/sheet1.xml'];

describe('changing conditional formatting', () => {
  it('writes the rules in their order, a new format joining the workbook’s, and reads them back the same', () => {
    const s = open();
    const [greater, formula] = rulesOf(s);
    const added: CfRule = { rects: [rect(0, 2, 9, 2)], type: 'top10', priority: 0, stopIfTrue: true, formulas: [], rank: 3, bottom: true, style: { color: '#9C0006', fill: '#FFC7CE' } };
    const r = set(s, [added, { ...formula, rects: [rect(0, 0, 4, 0)] }, greater]);
    expect(r.layout?.undo[0].op).toMatchObject({ kind: 'conditionalFormats', rules: [{ type: 'cellIs' }, { type: 'expression' }, { type: 'dataBar' }] });
    const files = parts(s.save());
    const xml = files['xl/worksheets/sheet1.xml'];
    const dxfs = Number(/<dxfs count="(\d+)"/.exec(files['xl/styles.xml'])?.[1]);
    expect(files['xl/styles.xml']).toContain('<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>');
    expect(xml).toContain(`<conditionalFormatting sqref="C1:C10"><cfRule type="top10" dxfId="${dxfs - 1}" priority="1" stopIfTrue="1" rank="3" bottom="1"></cfRule></conditionalFormatting>`);
    expect(xml).toContain('<conditionalFormatting sqref="A1:A5"><cfRule type="expression" dxfId="1" priority="2"><formula>MOD(A1,2)=0</formula></cfRule></conditionalFormatting>');
    expect(xml).toMatch(/<cfRule type="cellIs" dxfId="0" priority="3" operator="greaterThan"><formula>5<\/formula><\/cfRule>/);
    // The data bar, and Excel 2010's part of it, are gone.
    expect(xml).not.toContain('dataBar');
    const back = rulesOf(open(s.save()));
    expect(back.map((x) => [x.type, x.priority, x.rects, x.style])).toEqual([
      ['top10', 1, [rect(0, 2, 9, 2)], { color: '#9C0006', fill: '#FFC7CE' }],
      ['expression', 2, [rect(0, 0, 4, 0)], { bold: true }],
      ['cellIs', 3, [rect(0, 0, 9, 0)], greater.style],
    ]);
    expect(s.changes().map((c) => english.text(c.note!))).toEqual(['Conditional formatting changed']);
  });

  it('writes an average rule’s side, whether it counts the average, its deviations, and a text rule’s text', () => {
    const s = open();
    const below: CfRule = { rects: [rect(0, 2, 9, 2)], type: 'aboveAverage', priority: 0, stopIfTrue: false, formulas: [], aboveAverage: false, equalAverage: true, stdDev: 2, style: { bold: true } };
    const text: CfRule = { rects: [rect(0, 3, 9, 3)], type: 'containsText', priority: 0, stopIfTrue: false, formulas: ['NOT(ISERROR(SEARCH("a<b",D1)))'], operator: 'containsText', text: 'a<"b"', percent: true, style: { bold: true } };
    set(s, [below, text]);
    const xml = sheetXml(s);
    expect(xml).toMatch(/<cfRule type="aboveAverage" dxfId="\d+" priority="1" aboveAverage="0" equalAverage="1" stdDev="2"\/?>/);
    expect(xml).toMatch(/<cfRule type="containsText" dxfId="\d+" priority="2" percent="1" operator="containsText" text="a&lt;&quot;b&quot;"><formula>NOT\(ISERROR\(SEARCH\("a&lt;b",D1\)\)\)<\/formula>/);
    expect(rulesOf(open(s.save())).map((r) => [r.type, r.aboveAverage, r.equalAverage, r.stdDev, r.text])).toEqual([
      ['aboveAverage', false, true, 2, undefined],
      ['containsText', undefined, undefined, undefined, 'a<"b"'],
    ]);
  });

  it("starts the workbook's formats where it has none, writing every part a rule's format sets", async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('Plain').getCell('A1').value = 1;
    const files = parts(new Uint8Array(await wb.xlsx.writeBuffer()));
    const styles = files['xl/styles.xml'].replace(/<dxfs[^>]*\/>|<dxfs[\s\S]*?<\/dxfs>/, '');
    const s = open(zipSync(Object.fromEntries(Object.entries({ ...files, 'xl/styles.xml': styles }).filter(([k]) => !k.endsWith('/')).map(([k, v]) => [k, strToU8(v)]))));
    const thin = { width: 1 as const, pattern: 'solid' as const, color: '#FF0000' };
    const style = { bold: true, italic: true, underline: true, strike: true, border: { left: thin, right: { width: 2 as const, pattern: 'dashed' as const }, top: thin, bottom: thin } };
    set(s, [{ rects: [rect(0, 0)], type: 'duplicateValues', priority: 0, stopIfTrue: false, formulas: [], style }]);
    const out = parts(s.save())['xl/styles.xml'];
    expect(out).toContain(
      '<dxfs count="1"><dxf><font><b/><i/><strike/><u/></font><border><left style="thin"><color rgb="FFFF0000"/></left><right style="mediumDashed"><color auto="1"/></right>' +
        '<top style="thin"><color rgb="FFFF0000"/></top><bottom style="thin"><color rgb="FFFF0000"/></bottom></border></dxf></dxfs>',
    );
  });

  it('writes colour scales, data bars and icon sets as Excel 2007 reads them', () => {
    const s = open();
    const scale: CfRule = { rects: [rect(0, 3)], type: 'colorScale', priority: 0, stopIfTrue: false, formulas: [], colorScale: { cfvo: [{ type: 'min', gte: true }, { type: 'percentile', val: '50', gte: true }, { type: 'max', gte: true }], colors: ['#F8696B', '#FFEB84', '#63BE7B'] } };
    const bar: CfRule = { rects: [rect(0, 4)], type: 'dataBar', priority: 0, stopIfTrue: false, formulas: [], dataBar: { cfvo: [{ type: 'autoMin', gte: true }, { type: 'autoMax', gte: true }], color: '#638EC6', showValue: false, gradient: true, minLength: 0, maxLength: 100 } };
    const icons: CfRule = { rects: [rect(0, 5)], type: 'iconSet', priority: 0, stopIfTrue: false, formulas: [], iconSet: { set: '3Stars', cfvo: [{ type: 'percent', val: '0', gte: true }, { type: 'percent', val: '33', gte: false }, { type: 'percent', val: '67', gte: true }], reverse: true, showValue: true } };
    set(s, [scale, bar, icons]);
    const xml = sheetXml(s);
    expect(xml).toContain('<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale>');
    expect(xml).toContain('<dataBar minLength="0" maxLength="100" showValue="0"><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>');
    // Stars are Excel 2010's: Excel 2007 gets traffic lights.
    expect(xml).toContain('<iconSet reverse="1"><cfvo type="percent" val="0"/><cfvo type="percent" val="33" gte="0"/><cfvo type="percent" val="67"/></iconSet>');
  });

  it("keeps Excel 2010's part of a rule in step with it: moved, re-ranged or gone", () => {
    const files = parts(bytes);
    const id = '{00000000-0000-0000-0000-00000000AAAA}';
    const own = '{00000000-0000-0000-0000-00000000BBBB}';
    const x14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"';
    const xm = 'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';
    const sheet =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1"><c r="B1"><v>1</v></c></row></sheetData>' +
      '<conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>' +
      `<conditionalFormatting sqref="B1:B10"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ${x14}><x14:id>${id}</x14:id></ext></extLst></cfRule></conditionalFormatting>` +
      `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${x14}><x14:conditionalFormattings>` +
      `<x14:conditionalFormatting ${xm}><x14:cfRule type="dataBar" id="${id}"><x14:dataBar minLength="0" maxLength="100" negativeBarColorSameAsPositive="1"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>B1:B10</xm:sqref></x14:conditionalFormatting>` +
      `<x14:conditionalFormatting ${xm}><x14:cfRule type="iconSet" priority="9" id="${own}"><x14:iconSet iconSet="3Stars"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>C1:C10</xm:sqref></x14:conditionalFormatting>` +
      '</x14:conditionalFormattings></ext></extLst></worksheet>';
    const all = { ...files, 'xl/worksheets/sheet1.xml': sheet };
    const s = open(zipSync(Object.fromEntries(Object.entries(all).filter(([k]) => !k.endsWith('/')).map(([k, v]) => [k, strToU8(v)]))));
    const rules = rulesOf(s);
    const bar = rules.find((r) => r.type === 'dataBar')!;
    const stars = rules.find((r) => r.x14Only)!;
    expect([bar.extId, stars.extId, stars.iconSet?.set]).toEqual([id, own, '3Stars']);
    // The stars first, the bar on fewer cells, the rest as they were.
    set(s, [stars, ...rules.filter((r) => r !== stars && r !== bar), { ...bar, rects: [rect(0, 1, 4, 1)] }]);
    let xml = sheetXml(s);
    expect(xml).toContain(`priority="1" id="${own}"`);
    expect(xml).toMatch(/<x14:cfRule type="dataBar" id="\{[^}]*AAAA\}">[\s\S]*?<xm:sqref>B1:B5<\/xm:sqref>/);
    expect(xml).toContain('<cfRule type="dataBar" priority="3"><dataBar minLength="0" maxLength="100">');
    // Gone, both rules take their Excel 2010 parts with them.
    set(s, rulesOf(s).filter((r) => !r.extId));
    xml = sheetXml(s);
    expect(xml).not.toContain('x14:');
    expect(xml).not.toContain('extLst');
  });
});
