import { readdirSync } from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { english, errorText } from '../src/i18n';
import { CellFlag, has, sample, SAMPLES } from './helpers';

/**
 * Small files built to hang or exhaust the editor: tags that never end, huge
 * row numbers, scans that backtrack, formulas that build enormous arrays.
 * Each must be refused or handled quickly.
 */

const DAMAGED = "This file is damaged, so it wasn't opened. Excel may be able to repair it.";

const sheetXml = (data: string, after = '') =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${data}</sheetData>${after}</worksheet>`;

/** A one-sheet workbook with the given parts replaced. */
function xlsx(parts: { sheet?: string; sst?: string; rels?: string; workbook?: string; styles?: string } = {}): Uint8Array {
  const sst = parts.sst !== undefined;
  const files: Record<string, string> = {
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml':
      parts.workbook ??
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels':
      parts.rels ??
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>${sst ? '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' : ''}</Relationships>`,
    'xl/worksheets/sheet1.xml': parts.sheet ?? sheetXml('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row>'),
  };
  if (sst) files['xl/sharedStrings.xml'] = parts.sst!;
  if (parts.styles !== undefined) files['xl/styles.xml'] = parts.styles;
  return zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));
}

/** Opens the bytes and builds the engine; what was thrown, in English, and how long it took. */
function attempt(bytes: Uint8Array): { error?: string; ms: number; session?: WorkbookSession } {
  const t0 = performance.now();
  try {
    const session = WorkbookSession.open(bytes, 'hostile.xlsx', 'en-US', { recalc: 'on' });
    session.ensureEngine();
    return { ms: performance.now() - t0, session };
  } catch (e) {
    return { error: english.text(errorText(e)), ms: performance.now() - t0 };
  }
}

describe('tags that never end', () => {
  it('refuses a part that stops inside a comment, a tag or a closing tag', () => {
    const rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    for (const cut of [`${rels}<!--`, `${rels}</Relationships`, `${rels}<Relationship Id="rId1"`, `${rels}<![CDATA[x`]) {
      const r = attempt(xlsx({ rels: cut }));
      expect(r.error).toBe(DAMAGED);
      expect(r.ms).toBeLessThan(1000);
    }
  });

  it('refuses a row or a cell that never closes', () => {
    expect(attempt(xlsx({ sheet: sheetXml('<row r="1"><c r="A1"><v>1</v></c>') })).error).toBe(DAMAGED);
    expect(attempt(xlsx({ sheet: sheetXml('<row r="1"><c r="A1"><v>1</v></row>') })).error).toBe(DAMAGED);
    expect(attempt(xlsx({ sheet: sheetXml('<row r="1"><c r="A1"><v>1</v></c></row>').replace('</sheetData>', '') })).error).toBe(DAMAGED);
  });

  it('refuses a tag that runs into the next one', () => {
    expect(attempt(xlsx({ sheet: sheetXml('<row r="1"><c r="A1"<v>1</v></c></row>') })).error).toBe(DAMAGED);
  });
});

describe('sizes beyond Excel’s', () => {
  it('refuses row and column numbers past the end of a sheet, rather than sizing the grid by them', () => {
    const t0 = performance.now();
    expect(attempt(xlsx({ sheet: sheetXml('<row r="2000000000"><c><v>1</v></c></row>') })).error).toBe(DAMAGED);
    expect(attempt(xlsx({ sheet: sheetXml('<row r="1"><c r="XFE1"><v>1</v></c></row>') })).error).toBe(DAMAGED);
    expect(performance.now() - t0).toBeLessThan(1000);
    // The last row and column are fine.
    const last = attempt(xlsx({ sheet: sheetXml('<row r="1048576"><c r="XFD1048576"><v>7</v></c></row>') }));
    expect(last.error).toBeUndefined();
  });
});

describe('scans that could take quadratic time', () => {
  it('reads an attribute with no value, or thousands of unfinished strings, in linear time', () => {
    const t0 = performance.now();
    attempt(xlsx({ sheet: sheetXml(`<row r="1" ${'x'.repeat(400_000)}><c r="A1"><v>1</v></c></row>`) }));
    attempt(xlsx({ sst: `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${'<si><t>a</t>'.repeat(100_000)}</sst>` }));
    attempt(xlsx({ styles: `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs>${'<xf>'.repeat(100_000)}</cellXfs></styleSheet>` }));
    expect(performance.now() - t0).toBeLessThan(3000);
  });

  it('copes with deeply nested elements', () => {
    const deep = `${'<extLst>'.repeat(200_000)}${'</extLst>'.repeat(200_000)}`;
    const workbook = `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>${deep}</workbook>`;
    expect(attempt(xlsx({ workbook })).error).toBeUndefined();
  });
});

describe('formulas that would hang the engine or run it out of memory', () => {
  const formulaSheet = (cells: string) => sheetXml(`<row r="1">${cells}</row>`);

  it('matches wildcards in linear time', () => {
    const text = 'a'.repeat(40);
    const pattern = `${'*a'.repeat(14)}*b`;
    const sheet = formulaSheet(
      `<c r="A1" t="inlineStr"><is><t>${text}</t></is></c><c r="B1"><f>COUNTIF(A1,"${pattern}")</f><v>0</v></c><c r="C1"><f>COUNTIF(A1,"*a?a*")</f><v>1</v></c><c r="D1"><f>SEARCH("a*a",A1)</f><v>1</v></c><c r="E1"><f>COUNTIF(A1,"~*")</f><v>0</v></c>`,
    );
    const r = attempt(xlsx({ sheet }));
    expect(r.error).toBeUndefined();
    expect(r.ms).toBeLessThan(2000);
    // Each saved value matched the editor's own, so none was kept as a difference.
    expect(r.session!.calcReport!.mismatched).toBe(0);
    const s = r.session!.apply({ type: 'input', sheet: 0, row: 0, col: 0, text: `${text}b` });
    const b1 = s.updates.find((u) => u.row === 0 && u.col === 1);
    expect(b1?.n).toBe(1);
  });

  it('leaves ranges and arrays too large to calculate at Excel’s values', () => {
    const sheet = formulaSheet(
      `<c r="A1"><f>SUMPRODUCT(B2:XFD1048576,B2:XFD1048576)</f><v>5</v></c><c r="B1"><f>SUM(SEQUENCE(10000,10000))</f><v>6</v></c><c r="C1"><f>SUM(A:Z)</f><v>7</v></c>`,
    );
    const r = attempt(xlsx({ sheet }));
    expect(r.error).toBeUndefined();
    expect(r.ms).toBeLessThan(2000);
    const report = r.session!.calcReport!;
    expect(report.unsupported).toBe(3);
    expect(report.unsupportedFunctions.map((t) => english.text(t))).toEqual(['Ranges or arrays too large to calculate here']);
  });

  it('caps arrays whose size comes from cells, and never hands typed huge ranges to the engine', () => {
    // No saved value, so the editor's own result shows.
    const r = attempt(xlsx({ sheet: formulaSheet('<c r="A1"><v>10000</v></c><c r="B1"><f>SUM(SEQUENCE(A1,A1))</f></c>') }));
    expect(r.ms).toBeLessThan(2000);
    const b1 = () => r.session!.stateUpdates().find((u) => u.row === 0 && u.col === 1);
    expect(b1()?.s).toBe('#NUM!');
    r.session!.apply({ type: 'input', sheet: 0, row: 0, col: 0, text: '3' });
    expect(b1()?.n).toBe(45);

    const t0 = performance.now();
    const typed = r.session!.apply({ type: 'input', sheet: 0, row: 1, col: 0, text: '=SUMPRODUCT(B3:XFD1048576,B3:XFD1048576)' });
    expect(performance.now() - t0).toBeLessThan(2000);
    expect(has(typed.updates.find((u) => u.row === 1 && u.col === 0)?.flags, CellFlag.Unsupported)).toBe(true);
  });
});

/** A small deterministic random number generator, so a failure can be replayed. */
function random(seed: number): () => number {
  let s = seed >>> 0;
  return () => {
    s = (s + 0x6d2b79f5) >>> 0;
    let t = s;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

describe('damaged copies of the sample workbooks', () => {
  const names = readdirSync(SAMPLES).filter((n) => n.endsWith('.xlsx') && n !== 'big.xlsx');

  it('open, or are refused, quickly whatever part is cut short or garbled', () => {
    const rand = random(20261008);
    let opened = 0;
    let refused = 0;
    for (let round = 0; round < 120; round++) {
      const name = names[Math.floor(rand() * names.length)];
      const files = unzipSync(sample(name));
      const xmlParts = Object.keys(files).filter((p) => /\.(xml|rels|vml)$/.test(p));
      const part = xmlParts[Math.floor(rand() * xmlParts.length)];
      let text = strFromU8(files[part]);
      const at = Math.floor(rand() * text.length);
      const how = Math.floor(rand() * 4);
      if (how === 0) text = text.slice(0, at);
      else if (how === 1) text = text.slice(0, at) + '<' + text.slice(at);
      else if (how === 2) text = text.slice(0, at) + text.slice(at + 1 + Math.floor(rand() * 40));
      else text = text.slice(0, at) + '"' + text.slice(at);
      files[part] = strToU8(text);
      const bytes = zipSync(files);
      const t0 = performance.now();
      try {
        const s = WorkbookSession.open(bytes, name, 'en-US', { recalc: 'on' });
        s.ensureEngine();
        s.save();
        opened++;
      } catch {
        refused++;
      }
      const ms = performance.now() - t0;
      expect(ms, `${name} ${part} (${how} at ${at})`).toBeLessThan(5000);
    }
    expect(opened + refused).toBe(120);
  });
});
