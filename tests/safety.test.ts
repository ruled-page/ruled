import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { parts, sample } from './helpers';

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
    expect(() => WorkbookSession.open(bytes, 'k.xlsx', 'en-US', { limits: { maxFileBytes: 1000 } })).toThrow(/most this editor accepts/);
    expect(() => WorkbookSession.open(bytes, 'k.xlsx', 'en-US', { limits: { maxUnzippedBytes: 50_000 } })).toThrow(/unpacks to more than/);
    expect(() => WorkbookSession.open(bytes, 'k.xlsx', 'en-US', { limits: { maxCells: 100 } })).toThrow(/has 1,\d{3} cells/);
  });

  it('refuses a zip bomb before inflating it', () => {
    const bomb = zipSync({
      '[Content_Types].xml': strToU8('<Types/>'),
      'xl/worksheets/sheet1.xml': new Uint8Array(30 * 1024 * 1024),
    });
    expect(bomb.length).toBeLessThan(200_000);
    expect(() => WorkbookSession.open(bomb, 'bomb.xlsx')).toThrow(/compressed suspiciously well/);
  });

  it('refuses files that are not workbooks', () => {
    expect(() => WorkbookSession.open(strToU8('hello'), 'x.xlsx')).toThrow(/not an Excel workbook/);
    expect(() => WorkbookSession.open(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0]), 'x.xls')).toThrow(/password-protected or in the old .xls format/);
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
    expect(s.model.findings.find((f) => f.title === 'Macros (VBA)')?.detail).toMatch(/removed when you save/);
    const p = parts(s.save());
    expect(p['xl/vbaProject.bin']).toBeUndefined();
    expect(p['[Content_Types].xml']).not.toMatch(/macroEnabled|vbaProject/);
    expect(p['xl/_rels/workbook.xml.rels']).not.toContain('vbaProject');
    expect(s.saveFileName).toBe('m.xlsx');
    expect(() => WorkbookSession.open(s.save(), 'm.xlsx')).not.toThrow();
  });

  it('refuses macro workbooks when blocked', () => {
    expect(() => WorkbookSession.open(xlsm(), 'm.xlsm', 'en-US', { macros: 'block' })).toThrow(/contains macros/);
    expect(() => WorkbookSession.open(sample('kitchen-sink.xlsx'), 'k.xlsx', 'en-US', { macros: 'block' })).not.toThrow();
  });
});
