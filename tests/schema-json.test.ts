import ExcelJS from 'exceljs';
import { beforeAll, describe, expect, it } from 'vitest';
import { WorkbookSession } from '../src/core/session';
import { english, errorText } from '../src/i18n';
import { bindSchema, cellText } from '../src/schema/binding';
import { describeMatch, matchLabel, schemaFromJSON, schemaToJSON, textMatches } from '../src/schema/json';
import { parseStamp, stampedFirst, stampFor, versionNote } from '../src/schema/stamp';
import type { WorkbookSchema } from '../src/schema/types';
import { SheetView } from '../src/state/sheetView';

/** Why a schema file is turned down, in English. */
function refusal(text: string): string | undefined {
  try {
    schemaFromJSON(text);
  } catch (e) {
    return english.text(errorText(e));
  }
  return undefined;
}

const valid = {
  id: 'x',
  name: 'X',
  sheets: [{ name: 'S', tables: [{ key: 't', at: 'A1', fields: [{ key: 'a', label: 'A', type: 'text' }] }], cells: [{ key: 'c', ref: 'B9', type: 'number' }] }],
};
/** The valid schema with one thing changed, as JSON. */
// Deliberately loose: these edits break the schema's types.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const json = (change: (s: any) => unknown) => {
  const copy = JSON.parse(JSON.stringify(valid));
  change(copy);
  return JSON.stringify(copy);
};

describe('importing schemas', () => {
  it('accepts a well-formed schema as it is', () => {
    expect(schemaFromJSON(JSON.stringify(valid))).toEqual(valid);
  });

  it('says what is wrong with the file as a whole', () => {
    expect(refusal('{')).toMatch(/^The schema file isn't valid JSON: /);
    expect(refusal('null')).toBe('The schema file is not valid: expected an object.');
    expect(refusal(json((s) => delete (s).id))).toBe('The schema file is not valid: "id" is missing.');
    expect(refusal(json((s) => (s.name = { en: 1 })))).toBe('The schema file is not valid: "name" is missing.');
    expect(refusal(json((s) => (s.name = [])))).toBe('The schema file is not valid: "name" is missing.');
    expect(refusal(json((s) => (s.sheets = [])))).toBe('The schema file is not valid: "sheets" must list at least one sheet.');
    expect(refusal(json((s) => (s.version = 2)))).toBe('The schema file is not valid: "version" must be text, e.g. "2".');
    expect(refusal(json((s) => (s.hookTimeoutMs = 0)))).toBe('The schema file is not valid: "hookTimeoutMs" must be a positive number of milliseconds.');
  });

  it('says which sheet, table, field or cell is wrong', () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const t = (s: any) => s.sheets[0].tables[0];
    expect(refusal(json((s) => (s.sheets[0].name = '')))).toBe('The schema file is not valid: sheet 1 has no name.');
    expect(refusal(json((s) => (t(s).key = '')))).toBe('The schema file is not valid: a table on sheet 1 has no key.');
    expect(refusal(json((s) => (t(s).at = 'top left')))).toBe('The schema file is not valid: table "t" has an invalid "at" cell.');
    expect(refusal(json((s) => (t(s).fields = [])))).toBe('The schema file is not valid: table "t" has no fields.');
    expect(refusal(json((s) => (t(s).fields[0].key = '')))).toBe('The schema file is not valid: a field in table "t" has no key.');
    expect(refusal(json((s) => t(s).fields.push({ key: 'a', label: 'A2', type: 'text' })))).toBe('The schema file is not valid: table "t" has two fields called "a".');
    expect(refusal(json((s) => (t(s).fields[0] = { key: 'a', label: 'A', type: 'enum' })))).toBe('The schema file is not valid: enum field "a" lists no values.');
    expect(refusal(json((s) => (t(s).fields[0].pattern = '(')))).toBe('The schema file is not valid: field "a" has an invalid pattern.');
    expect(refusal(json((s) => (t(s).checks = [{ rule: 'requiredIf', field: 'a', when: 'b' }])))).toBe(
      'The schema file is not valid: a "requiredIf" check in table "t" names "b", which isn\'t one of its fields.',
    );
    expect(refusal(json((s) => (t(s).checks = [{ rule: 'uniqueTogether', field: 'a', fields: ['a', 'z'] }])))).toMatch(/"uniqueTogether" check .* names "z"/);
    expect(refusal(json((s) => (t(s).checks = [{ rule: 'sum', field: '*', max: 1 }])))).toBeUndefined();
    expect(refusal(json((s) => (s.sheets[0].cells[0].ref = '')))).toBe('The schema file is not valid: a cell on sheet 1 is missing its key or ref.');
    expect(refusal(json((s) => (s.sheets[0].cells[0].type = 'money')))).toBe('The schema file is not valid: cell "c" has unknown type "money".');
  });
});

describe('text matches', () => {
  it('match plain text ignoring case and spaces, and patterns as regular expressions', () => {
    expect(textMatches(' total ', 'TOTAL')).toBe(true);
    expect(textMatches(/^Q\d$/, ' Q3 ')).toBe(true);
    expect(textMatches({ pattern: '^q\\d$', flags: 'i' }, 'Q3')).toBe(true);
    expect(textMatches({ pattern: '^q\\d$' }, 'Q3')).toBe(false);
  });

  it('are shown and exported in a readable form', () => {
    expect(english.text(describeMatch('Total'))).toBe('"Total"');
    expect(english.text(describeMatch(/^Q\d$/))).toBe('text matching /^Q\\d$/');
    expect(english.text(describeMatch({ pattern: 'x+' }))).toBe('text matching /x+/');
    expect([matchLabel('Total'), matchLabel(/^Q\d$/i), matchLabel({ pattern: 'x+' })]).toEqual(['Total', '/^Q\\d$/', '/x+/']);
    const schema: WorkbookSchema = { id: 'r', name: 'R', sheets: [{ name: /^Data$/i, tables: [{ key: 't', at: 'A1', fields: [{ key: 'q', label: /^Q\d$/, type: 'text' }] }] }] };
    const back = schemaFromJSON(schemaToJSON(schema));
    expect(back.sheets[0].name).toEqual({ pattern: '^Data$', flags: 'i' });
    expect(back.sheets[0].tables![0].fields[0].label).toEqual({ pattern: '^Q\\d$' });
  });
});

describe('binding to a workbook', () => {
  let views: SheetView[];
  beforeAll(async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    ws.addRow([2026, true, null, 'Name']);
    ws.addRow(['Owner', 'Ann']);
    const s = WorkbookSession.open(new Uint8Array(await wb.xlsx.writeBuffer()), 'b.xlsx');
    const r = s.loadResult();
    views = r.sheets.map((info, i) => new SheetView(info, r.grids[i]));
  });

  it('reads labels from numbers and logicals as they show', () => {
    expect([0, 1, 2, 3].map((c) => cellText(views[0], 0, c))).toEqual(['2026', 'TRUE', '', 'Name']);
  });

  it('reports a corner label or single-cell label that is missing, and skips optional sheets', () => {
    const schema: WorkbookSchema = {
      id: 'b',
      name: 'B',
      sheets: [
        {
          name: 'Data',
          tables: [{ key: 't', at: 'C1', corner: 'Metric', records: { labels: ['Owner'] }, fields: [{ key: 'n', label: 'Name', type: 'text' }] }],
          cells: [
            { key: 'who', ref: 'B2', label: { ref: 'A2', text: 'Owner' }, type: 'text' },
            { key: 'gone', ref: 'B5', label: { ref: 'A5', text: 'Approved' }, type: 'text' },
            { key: 'wrong', ref: 'B1', label: { ref: 'A1', text: 'Year' }, type: 'text' },
          ],
        },
        { name: 'Extra', optional: true },
      ],
    };
    const r = bindSchema(schema, views);
    expect(r.found).toBe(1);
    expect(r.problems.map((p) => english.text(p))).toEqual([
      'Data!C2 should be "Owner" (record label) but is empty.',
      'Data!C1 should be "Metric" (corner label) but is empty.',
      'Data!A5 should be "Approved" but is empty.',
      'Data!A1 should be "Year" but is "2026".',
    ]);
  });

  it('accepts any text where the schema says "*"', () => {
    const schema: WorkbookSchema = { id: 'b', name: 'B', sheets: [{ name: 'Data', tables: [{ key: 't', at: 'A1', fields: [{ key: 'y', label: '*', type: 'number' }] }] }] };
    expect(bindSchema(schema, views).problems).toEqual([]);
  });
});

describe('schema versions', () => {
  const schema = (version?: string): WorkbookSchema => ({ id: 'tasks', name: { en: 'Tasks', de: 'Aufgaben' }, version, sheets: [{ name: 'S' }] });
  const note = (s: WorkbookSchema, stamp: string | undefined, language?: string) => {
    const n = versionNote(s, stamp ? { WorkbookSchema: stamp } : {}, language);
    return n && english.text(n);
  };

  it('read the stamp a file was saved with', () => {
    expect(parseStamp(undefined)).toBeUndefined();
    expect(parseStamp('tasks')).toEqual({ id: 'tasks', version: '' });
    expect(parseStamp('a@b@3')).toEqual({ id: 'a@b', version: '3' });
    expect(stampFor(schema('2'))).toEqual({ WorkbookSchema: 'tasks@2' });
    expect(stampFor(schema())).toBeUndefined();
    expect(stampFor(undefined)).toBeUndefined();
  });

  it('try the schema a file was stamped with first', () => {
    const other: WorkbookSchema = { ...schema('1'), id: 'other' };
    expect(stampedFirst([other, schema('2')], { WorkbookSchema: 'tasks@1' }).map((s) => s.id)).toEqual(['tasks', 'other']);
    expect(stampedFirst([other, schema('2')], {}).map((s) => s.id)).toEqual(['other', 'tasks']);
  });

  it('say when a file was saved under another version, in the editor language', () => {
    expect(note(schema('2'), 'tasks@2')).toBeUndefined();
    expect(note(schema('2'), 'other@1')).toBeUndefined();
    expect(note(schema('2'), 'tasks@1')).toBe('This file was last saved with version 1 of “Tasks”; it is checked against version 2.');
    expect(note(schema('2'), 'tasks', 'de')).toBe('This file was last saved with an unversioned “Aufgaben” schema; it is checked against version 2.');
    expect(note(schema(), 'tasks@1')).toBe('This file was last saved with version 1 of “Tasks”; it is checked against the current version.');
  });
});
