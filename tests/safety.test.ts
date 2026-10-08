import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { english, errorText } from '../src/i18n';
import { parts, sample } from './helpers';

/** What a call throws, in English. */
function thrown(fn: () => unknown): string | undefined {
  try {
    fn();
  } catch (e) {
    return english.text(errorText(e));
  }
  return undefined;
}

/** kitchen-sink.xlsx turned into a macro-enabled workbook with a stand-in VBA project. */
function xlsm(): Uint8Array {
  const files = unzipSync(sample('kitchen-sink.xlsx'));
  const text = (p: string) => strFromU8(files[p]);
  files['xl/vbaProject.bin'] = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3, 4]);
  files['[Content_Types].xml'] = strToU8(
    text('[Content_Types].xml')
      .replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
      .replace('<Default Extension="xml"', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default Extension="xml"'),
  );
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    text('xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
    ),
  );
  return zipSync(files);
}

describe('limits on untrusted files', () => {
  it('refuses files over the size, unzipped size or cell limits', () => {
    const bytes = sample('kitchen-sink.xlsx');
    expect(thrown(() => WorkbookSession.open(bytes, 'k.xlsx', 'en-US', { limits: { maxFileBytes: 1000 } }))).toMatch(/most this editor accepts/);
    expect(thrown(() => WorkbookSession.open(bytes, 'k.xlsx', 'en-US', { limits: { maxUnzippedBytes: 50_000 } }))).toMatch(/unpacks to more than/);
    expect(thrown(() => WorkbookSession.open(bytes, 'k.xlsx', 'en-US', { limits: { maxCells: 100 } }))).toBe('This workbook has more than 100 cells, the most this editor can hold.');
  });

  it('refuses a zip bomb before inflating it', () => {
    const bomb = zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      'xl/worksheets/sheet1.xml': new Uint8Array(30 * 1024 * 1024),
    });
    expect(bomb.length).toBeLessThan(200_000);
    expect(thrown(() => WorkbookSession.open(bomb, 'bomb.xlsx'))).toMatch(/compressed suspiciously well/);
  });

  it('refuses files that are not workbooks', () => {
    expect(thrown(() => WorkbookSession.open(strToU8('hello'), 'x.xlsx'))).toMatch(/not an Excel workbook/);
    expect(thrown(() => WorkbookSession.open(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0]), 'x.xls'))).toMatch(/password-protected or in the old .xls format/);
  });
});

describe('macro policy', () => {
  it('keeps macros by default', () => {
    const s = WorkbookSession.open(xlsm(), 'm.xlsm');
    s.apply({ type: 'input', sheet: 3, row: 25, col: 0, text: 'x' });
    const p = parts(s.save());
    expect(p['xl/vbaProject.bin']).toBe('<binary>');
    expect(p['[Content_Types].xml']).toContain('macroEnabled.main+xml');
    expect(s.saveFileName).toBe('m.xlsm');
  });

  it('removes them on save, even without edits, and saves as .xlsx', () => {
    const s = WorkbookSession.open(xlsm(), 'm.xlsm', 'en-US', { macros: 'remove' });
    expect(english.text(s.model.findings.find((f) => english.text(f.title) === 'Macros (VBA)')?.detail)).toMatch(/removed when you save/);
    const p = parts(s.save());
    expect(p['xl/vbaProject.bin']).toBeUndefined();
    expect(p['[Content_Types].xml']).not.toMatch(/macroEnabled|vbaProject/);
    expect(p['xl/_rels/workbook.xml.rels']).not.toContain('vbaProject');
    expect(s.saveFileName).toBe('m.xlsx');
    expect(() => WorkbookSession.open(s.save(), 'm.xlsx')).not.toThrow();
  });

  it('refuses macro workbooks when blocked', () => {
    expect(thrown(() => WorkbookSession.open(xlsm(), 'm.xlsm', 'en-US', { macros: 'block' }))).toMatch(/contains macros/);
    expect(() => WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx', 'en-US', { macros: 'block' })).not.toThrow();
  });
});

/**
 * kitchen-sink.xlsx with a hidden Excel 4.0 macro sheet that runs on opening
 * (Auto_Open), an ActiveX control and an embedded object: content Excel can
 * run with or without a VBA project.
 */
function xlm(): Uint8Array {
  const files = unzipSync(sample('kitchen-sink.xlsx'));
  const text = (p: string) => strFromU8(files[p]);
  files['xl/macrosheets/sheet1.xml'] = strToU8(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xm:macrosheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><sheetData><row r="1"><c r="A1"><f>EXEC("calc.exe")</f></c></row><row r="2"><c r="A2"><f>HALT()</f></c></row></sheetData></xm:macrosheet>',
  );
  files['xl/activeX/activeX1.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8" standalone="no"?><ax:ocx xmlns:ax="http://schemas.microsoft.com/office/2006/activeX" ax:classid="{8BD21D40-EC42-11CE-9E0D-00AA006002F3}" ax:persistence="persistStreamInit"/>');
  files['xl/embeddings/oleObject1.bin'] = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 9, 9]);
  files['[Content_Types].xml'] = strToU8(
    text('[Content_Types].xml').replace(
      '</Types>',
      '<Override PartName="/xl/macrosheets/sheet1.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"/><Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/><Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/></Types>',
    ),
  );
  files['xl/_rels/workbook.xml.rels'] = strToU8(
    text('xl/_rels/workbook.xml.rels').replace(
      '</Relationships>',
      '<Relationship Id="rIdXlm" Type="http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="macrosheets/sheet1.xml"/></Relationships>',
    ),
  );
  files['xl/workbook.xml'] = strToU8(
    text('xl/workbook.xml')
      .replace('</sheets>', '<sheet name="Macro1" sheetId="99" state="hidden" r:id="rIdXlm"/></sheets>')
      .replace(/<definedNames>|<calcPr/, (m) => (m === '<definedNames>' ? '<definedNames><definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName>' : `<definedNames><definedName name="_xlnm.Auto_Open">Macro1!$A$1</definedName></definedNames>${m}`)),
  );
  return zipSync(files);
}

describe('Excel 4.0 macros, ActiveX controls and embedded objects', () => {
  it('reports them, and keeps them by default', () => {
    const s = WorkbookSession.open(xlm(), 'x.xlsx');
    const titles = s.model.findings.map((f) => english.text(f.title));
    expect(titles).toContain('Excel 4.0 macro sheet "Macro1"');
    expect(titles).toContain('ActiveX controls and embedded objects');
    expect(english.text(s.model.findings.find((f) => f.kind === 'activeContent')?.detail)).toMatch(/^ActiveX controls: 1\. Embedded objects: 1\./);
    s.apply({ type: 'input', sheet: 0, row: 1, col: 1, text: '1' });
    expect(parts(s.save())['xl/macrosheets/sheet1.xml']).toContain('EXEC');
  });

  it('refuses a workbook with macro sheets when macros are blocked', () => {
    expect(thrown(() => WorkbookSession.open(xlm(), 'x.xlsx', 'en-US', { macros: 'block' }))).toBe(
      "This workbook contains macros (VBA or Excel 4.0 macro sheets), which aren't allowed here. Save it as .xlsx in Excel and open that instead.",
    );
  });

  it('empties macro sheets and drops the names that run them when macros are removed', () => {
    const s = WorkbookSession.open(xlm(), 'x.xlsm', 'en-US', { macros: 'remove' });
    expect(english.text(s.model.findings.find((f) => english.text(f.title) === 'Excel 4.0 macro sheet "Macro1"')?.detail)).toMatch(/emptied when you save/);
    const saved = s.save();
    const p = parts(saved);
    expect(p['xl/macrosheets/sheet1.xml']).not.toContain('EXEC');
    expect(p['xl/macrosheets/sheet1.xml']).toContain('<sheetData/>');
    expect(p['xl/_rels/workbook.xml.rels']).not.toMatch(/xlMacrosheet/);
    expect(p['[Content_Types].xml']).not.toMatch(/macrosheet\+xml/);
    expect(p['xl/workbook.xml']).not.toContain('Auto_Open');
    expect(s.saveFileName).toBe('x.xlsx');
    if (process.env.ROUNDTRIP_OUT) writeFileSync(join(process.env.ROUNDTRIP_OUT, 'macros-removed.xlsx'), saved);
    const reopened = WorkbookSession.open(saved, 'x.xlsx', 'en-US', { macros: 'block' });
    expect(reopened.model.sheets.find((m) => m.info.name === 'Macro1')?.info.kind).toBe('worksheet');
  });

  it('refuses ActiveX controls and embedded objects when asked to', () => {
    expect(thrown(() => WorkbookSession.open(xlm(), 'x.xlsx', 'en-US', { activeContent: 'block' }))).toBe(
      "This workbook contains ActiveX controls or embedded objects, which aren't allowed here.",
    );
    expect(() => WorkbookSession.open(sample('orders.xlsx'), 'o.xlsx', 'en-US', { activeContent: 'block' })).not.toThrow();
  });
});
