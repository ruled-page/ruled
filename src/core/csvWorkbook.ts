import { msg, TextError } from '../i18n/message';
import { CSV_FILE, decodeText, parseCsv, sepLine, sniffDelimiter } from './csv';
import { parseInput } from './input';
import { listSeparator } from './locale';
import { rangeToA1, toA1 } from './refs';
import { builtinIdFor } from './styles';
import type { CellContent } from './types';
import { encodeOoxmlText, escapeXml } from './xml';
import { DEFAULT_LIMITS, OoxmlPackage, type PackageLimits } from './zip';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** Whether a file is opened as CSV, by its name. */
export function isCsvFile(fileName: string): boolean {
  return CSV_FILE.test(fileName);
}

/**
 * A CSV field's value as Excel reads it when it opens a CSV file in the
 * user's locale: numbers, dates, times, percentages, currencies, TRUE and
 * FALSE and error values become values (leading zeros and digits past the
 * fifteenth are lost, as in Excel). Two differences: nothing becomes a
 * formula, since a CSV file is untrusted data, and a leading apostrophe is
 * kept as part of the text (Excel keeps it too when opening a CSV, unlike
 * when one is typed).
 */
export function csvValue(text: string, locale: string): { content: CellContent; numFmt?: string } {
  if (text[0] === '=' || text[0] === "'") return { content: { kind: 'string', value: text } };
  // Without a digit, text can only be a logical or an error value: the slower number and date parsers are skipped.
  if (!/\d/.test(text) && !/^\s*(true|false)\s*$/i.test(text) && !text.includes('#')) return { content: { kind: 'string', value: text } };
  const parsed = parseInput(text, 'General', locale);
  return { content: parsed.content, numFmt: parsed.numFmt };
}

/** A sheet name Excel accepts, from the file's name: without the extension, characters sheet names can't hold replaced, at most 31 characters. */
export function sheetNameFor(fileName: string): string {
  const base = fileName.replace(/^.*[/\\]/, '').replace(CSV_FILE, '');
  const name = base.replace(/[[\]:*?/\\]/g, '_').replace(/^'+|'+$/g, '').slice(0, 31).trim();
  return name || 'Sheet1';
}

const preserve = (s: string) => (/^\s|\s$|\n/.test(s) ? ' xml:space="preserve"' : '');

/**
 * Opens a CSV file as a one-sheet workbook package, named after the file, so
 * that it is edited, checked and saved like any workbook. The file's size and
 * its number of values are held to the same limits as a workbook's.
 */
export function csvPackage(bytes: Uint8Array, fileName: string, locale: string, maxCells: number, limits: PackageLimits = {}): OoxmlPackage {
  const maxBytes = limits.maxFileBytes ?? DEFAULT_LIMITS.maxFileBytes;
  if (bytes.length > maxBytes) {
    throw new TextError(msg('core.open.tooBig', { size: Math.round(bytes.length / 1024 / 1024), max: Math.round(maxBytes / 1024 / 1024) }));
  }
  const { text } = decodeText(bytes);
  const sep = sepLine(text);
  const delimiter = sep?.delimiter ?? (/\.tsv$/i.test(fileName) ? '\t' : sniffDelimiter(text, /\.txt$/i.test(fileName) ? '\t' : listSeparator(locale)));

  const strings = new Map<string, number>();
  const sst: string[] = [];
  let stringCells = 0;
  /** Number formats the values implied, as style indexes; style 0 is General. */
  const formats = new Map<string, number>();
  const rows: string[] = [];
  let row = -1;
  let cells = '';
  let maxRow = -1;
  let maxCol = -1;
  const endRow = () => {
    if (row >= 0) rows.push(`<row r="${row + 1}">${cells}</row>`);
  };
  parseCsv(
    sep ? text.slice(sep.start) : text,
    delimiter,
    (r, c, value) => {
      if (r !== row) {
        endRow();
        row = r;
        cells = '';
      }
      maxRow = r;
      if (c > maxCol) maxCol = c;
      const ref = toA1(r, c);
      const { content, numFmt } = csvValue(value, locale);
      switch (content.kind) {
        case 'number': {
          let style = 0;
          if (numFmt) {
            style = formats.get(numFmt) ?? formats.size + 1;
            formats.set(numFmt, style);
          }
          cells += `<c r="${ref}"${style ? ` s="${style}"` : ''}><v>${content.value}</v></c>`;
          break;
        }
        case 'bool':
          cells += `<c r="${ref}" t="b"><v>${content.value ? 1 : 0}</v></c>`;
          break;
        case 'error':
          cells += `<c r="${ref}" t="e"><v>${escapeXml(content.value)}</v></c>`;
          break;
        case 'string': {
          let index = strings.get(content.value);
          if (index === undefined) {
            index = sst.length;
            strings.set(content.value, index);
            sst.push(`<si><t${preserve(content.value)}>${escapeXml(encodeOoxmlText(content.value))}</t></si>`);
          }
          stringCells++;
          cells += `<c r="${ref}" t="s"><v>${index}</v></c>`;
          break;
        }
      }
    },
    maxCells,
  );
  endRow();

  const dimension = maxRow >= 0 ? rangeToA1({ r1: 0, c1: 0, r2: maxRow, c2: maxCol }) : 'A1';
  const custom: string[] = [];
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
  for (const code of formats.keys()) {
    let id = builtinIdFor(code);
    if (id === undefined) {
      id = 164 + custom.length;
      custom.push(`<numFmt numFmtId="${id}" formatCode="${escapeXml(code)}"/>`);
    }
    xfs.push(`<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`);
  }
  const contentType = (part: string, type: string) => `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${type}+xml"/>`;

  return OoxmlPackage.fromParts({
    '[Content_Types].xml':
      `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      contentType('xl/workbook.xml', 'sheet.main') +
      contentType('xl/worksheets/sheet1.xml', 'worksheet') +
      contentType('xl/styles.xml', 'styles') +
      contentType('xl/sharedStrings.xml', 'sharedStrings') +
      '</Types>',
    '_rels/.rels': `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml':
      `${XML_HEAD}<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><bookViews><workbookView/></bookViews>` +
      `<sheets><sheet name="${escapeXml(sheetNameFor(fileName))}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels':
      `${XML_HEAD}<Relationships xmlns="${PKG_REL_NS}">` +
      `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rId3" Type="${REL_NS}/sharedStrings" Target="sharedStrings.xml"/>` +
      '</Relationships>',
    'xl/worksheets/sheet1.xml':
      `${XML_HEAD}<worksheet xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><dimension ref="${dimension}"/>` +
      '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
      `<sheetData>${rows.join('')}</sheetData></worksheet>`,
    'xl/styles.xml':
      `${XML_HEAD}<styleSheet xmlns="${MAIN_NS}">` +
      (custom.length ? `<numFmts count="${custom.length}">${custom.join('')}</numFmts>` : '') +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>',
    'xl/sharedStrings.xml': `${XML_HEAD}<sst xmlns="${MAIN_NS}" count="${stringCells}" uniqueCount="${sst.length}">${sst.join('')}</sst>`,
  });
}
