import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession, type OpenOptions } from '../src/core/session';
import { loadedFormula, loadedValue, parts } from './helpers';

function patch(bytes: Uint8Array, edits: Record<string, string | Uint8Array | ((xml: string) => string)>): Uint8Array {
  const files = unzipSync(bytes);
  for (const [path, edit] of Object.entries(edits)) {
    files[path] = edit instanceof Uint8Array ? new Uint8Array(edit) : strToU8(typeof edit === 'string' ? edit : edit(strFromU8(files[path])));
  }
  return zipSync(files);
}

/**
 * Sheet S:
 *   A1 = 1; C1 = IF(A1>5,"big","small"); D1 = A1>5; E1 = 1/(A1-7); I1 = "x" in bold; J1 = "y"
 *   A3:A7 one shared formula, A1+1 copied down, whose followers saved a text, a
 *   logical, an error and no value; A9:A10 followers whose master is missing;
 *   row 15 empty but for its height.
 * The workbook has a calculation chain, and its calcPr asks Excel not to recalculate.
 */
async function build(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 1;
  ws.getCell('I1').value = 'x';
  ws.getCell('I1').font = { bold: true };
  const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
  const row1 = (xml: string) => /<row r="1"[\s\S]*?<\/row>/.exec(xml)![0];
  return patch(bytes, {
    'xl/worksheets/sheet1.xml': (xml) =>
      xml.replace(
        /<sheetData>[\s\S]*<\/sheetData>/,
        '<sheetData>' +
          row1(xml).replace(
            '</row>',
            '<c r="C1" t="str"><f>IF(A1&gt;5,"big","small")</f><v>small</v></c><c r="D1" t="b"><f>A1&gt;5</f><v>0</v></c>' +
              '<c r="E1"><f>1/(A1-7)</f><v>-0.166666666666667</v></c><c r="J1" t="inlineStr"><is><t>y</t></is></c></row>',
          ) +
          '<row r="3"><c r="A3"><f t="shared" ref="A3:A7" si="0">A1+1</f><v>2</v></c></row>' +
          '<row r="4"><c r="A4" t="str"><f t="shared" si="0"/><v>x</v></c></row>' +
          '<row r="5"><c r="A5" t="b"><f t="shared" si="0"/><v>1</v></c></row>' +
          '<row r="6"><c r="A6" t="e"><f t="shared" si="0"/><v>#N/A</v></c></row>' +
          '<row r="7"><c r="A7"><f t="shared" si="0"/></c></row>' +
          '<row r="9"><c r="A9"><f t="shared" si="1"/><v>5</v></c></row>' +
          '<row r="10"><c r="A10"><f t="shared" si="1"/><v>6</v></c></row>' +
          '<row r="15" ht="30" customHeight="1"/>' +
          '</sheetData>',
      ),
    'xl/workbook.xml': (xml) => xml.replace('<calcPr calcId="171027"/>', '<calcPr calcId="171027" fullCalcOnLoad="0"/>'),
    'xl/calcChain.xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C1" i="1"/></calcChain>',
    'xl/_rels/workbook.xml.rels': (xml) =>
      xml.replace('</Relationships>', '<Relationship Id="rIdChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>'),
    '[Content_Types].xml': (xml) =>
      xml.replace('</Types>', '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>'),
  });
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await build();
});

function edited(edits: [string, string][], options: OpenOptions = {}, from = bytes) {
  const s = WorkbookSession.open(from, 'w.xlsx', 'en-US', options);
  for (const [ref, text] of edits) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref)!;
    s.apply({ type: 'input', sheet: 0, row: Number(m[2]) - 1, col: m[1].charCodeAt(0) - 65, text });
  }
  const saved = s.save();
  return { xml: parts(saved), file: WorkbookSession.open(saved, 'w.xlsx') };
}

const cell = (xml: Record<string, string>, ref: string) => new RegExp(`<c r="${ref}"[^>]*?(?:/>|>[\\s\\S]*?</c>)`).exec(xml['xl/worksheets/sheet1.xml'])?.[0];

describe('saving recalculated formulas', () => {
  it('stores each kind of new result with the formula left as it was', () => {
    const { xml } = edited([['A1', '7']]);
    expect(cell(xml, 'C1')).toBe('<c r="C1" t="str"><f>IF(A1&gt;5,"big","small")</f><v>big</v></c>');
    expect(cell(xml, 'D1')).toBe('<c r="D1" t="b"><f>A1&gt;5</f><v>1</v></c>');
    expect(cell(xml, 'E1')).toBe('<c r="E1" t="e"><f>1/(A1-7)</f><v>#DIV/0!</v></c>');
  });
});

describe('shared formulas', () => {
  it("write each follower out in full when the master is edited, keeping the follower's saved result", () => {
    // Without recalculation, the results Excel saved are all there is.
    const { xml, file } = edited([['A3', '=A1+10']], { recalc: 'off' });
    expect(cell(xml, 'A3')).toBe('<c r="A3"><f>A1+10</f></c>');
    expect(cell(xml, 'A4')).toBe('<c r="A4" t="str"><f>A2+1</f><v>x</v></c>');
    expect(cell(xml, 'A5')).toBe('<c r="A5" t="b"><f>A3+1</f><v>1</v></c>');
    expect(cell(xml, 'A6')).toBe('<c r="A6" t="e"><f>A4+1</f><v>#N/A</v></c>');
    expect(cell(xml, 'A7')).toBe('<c r="A7"><f>A5+1</f></c>');
    // Followers with no master are left alone.
    expect(cell(xml, 'A9')).toBe('<c r="A9"><f t="shared" si="1"/><v>5</v></c>');
    expect(loadedFormula(file, 'S', 'A6')).toBe('A4+1');
  });

  it('keep a follower that was edited itself as edited', () => {
    const { xml } = edited([['A3', '=A1+10'], ['A5', '42']], { recalc: 'off' });
    expect(cell(xml, 'A5')).toBe('<c r="A5"><v>42</v></c>');
    expect(cell(xml, 'A4')).toContain('<f>A2+1</f>');
  });
});

describe('edited cells', () => {
  it('write errors, padded text and new rows, and keep a row element that had no cells', () => {
    const { xml, file } = edited([['G1', '#N/A'], ['H1', ' padded '], ['A20', 'new'], ['A15', 'tall']]);
    expect(cell(xml, 'G1')).toBe('<c r="G1" t="e"><v>#N/A</v></c>');
    expect(cell(xml, 'H1')).toBe('<c r="H1" t="inlineStr"><is><t xml:space="preserve"> padded </t></is></c>');
    expect(xml['xl/worksheets/sheet1.xml']).toContain('<row r="20"><c r="A20" t="inlineStr">');
    expect(xml['xl/worksheets/sheet1.xml']).toContain('<row r="15" ht="30" customHeight="1"><c r="A15" t="inlineStr">');
    expect(xml['xl/worksheets/sheet1.xml']).toContain('<dimension ref="A1:J20"/>');
    expect(loadedValue(file, 'S', 'H1')).toBe(' padded ');
  });

  it('keep the format of a cleared cell, drop a cleared plain one, and leave cleared cells out of the used range', () => {
    const { xml } = edited([['I1', ''], ['J1', ''], ['A30', 'x'], ['A30', '']]);
    expect(cell(xml, 'I1')).toMatch(/^<c r="I1" s="\d+"\/>$/);
    expect(cell(xml, 'J1')).toBeUndefined();
    expect(cell(xml, 'A30')).toBeUndefined();
    expect(xml['xl/worksheets/sheet1.xml']).toContain('<dimension ref="A1:J10"/>');
  });

  it('add the number formats new values need to a workbook that had none of its own', () => {
    const { xml, file } = edited([['B2', '2026-01-05']]);
    expect(xml['xl/styles.xml']).toMatch(/<numFmts count="1"><numFmt numFmtId="164" formatCode="[^"]+"\/><\/numFmts><fonts/);
    expect(file.model.styles.get(file.current(0, 1, 1).style).numFmt).not.toBe('General');
  });
});

describe('the workbook parts', () => {
  it('drop the calculation chain and ask Excel to recalculate on open', () => {
    const { xml } = edited([['A1', '2']]);
    expect(xml['xl/calcChain.xml']).toBeUndefined();
    expect(xml['xl/_rels/workbook.xml.rels']).not.toContain('calcChain');
    expect(xml['[Content_Types].xml']).not.toContain('calcChain');
    expect(xml['xl/workbook.xml']).toContain('<calcPr calcId="171027" fullCalcOnLoad="1"/>');
  });

  it('add calculation settings where a workbook has none, after its names', () => {
    const none = patch(bytes, { 'xl/workbook.xml': (x) => x.replace(/<calcPr[^>]*\/>/, '') });
    expect(edited([['A1', '2']], {}, none).xml['xl/workbook.xml']).toMatch(/<\/sheets><calcPr fullCalcOnLoad="1"\/><\/workbook>/);
    const named = patch(none, { 'xl/workbook.xml': (x) => x.replace('</sheets>', '</sheets><definedNames><definedName name="N">S!$A$1</definedName></definedNames>') });
    expect(edited([['A1', '2']], {}, named).xml['xl/workbook.xml']).toContain('</definedNames><calcPr fullCalcOnLoad="1"/>');
  });

  it('remove a VBA project with its signature when macros are to go', () => {
    const xlsm = patch(bytes, {
      'xl/vbaProject.bin': new Uint8Array([1, 2, 3]),
      'xl/vbaProjectSignature.bin': new Uint8Array([4, 5, 6]),
      'xl/_rels/vbaProject.bin.rels':
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature" Target="vbaProjectSignature.bin"/></Relationships>',
      'xl/_rels/workbook.xml.rels': (x) =>
        x.replace('</Relationships>', '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'),
      '[Content_Types].xml': (x) =>
        x
          .replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
          .replace(
            '</Types>',
            '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/vbaProjectSignature.bin" ContentType="application/vnd.ms-office.vbaProjectSignature"/></Types>',
          ),
    });
    const { xml } = edited([['A1', '2']], { macros: 'remove' }, xlsm);
    expect(Object.keys(xml).filter((k) => /vba/i.test(k))).toEqual([]);
    expect(xml['[Content_Types].xml']).not.toMatch(/vbaProject|macroEnabled/);
    expect(xml['[Content_Types].xml']).toContain('spreadsheetml.sheet.main+xml');
  });
});
