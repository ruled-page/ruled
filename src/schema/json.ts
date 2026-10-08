import { english } from '../i18n';
import { msg, TextError, type Message, type MessageKey, type MessageParams } from '../i18n/message';
import { isUnsafePattern } from './patterns';
import type { FieldType, TextMatch, WorkbookSchema } from './types';

/** Whether `text` satisfies the expectation: plain strings ignore case and surrounding spaces. */
export function textMatches(match: TextMatch, text: string): boolean {
  if (typeof match === 'string') return match.trim().toLowerCase() === text.trim().toLowerCase();
  const re = match instanceof RegExp ? match : new RegExp(match.pattern, match.flags);
  return re.test(text.trim());
}

/** How an expectation is shown to people. */
export function describeMatch(match: TextMatch): Message {
  if (typeof match === 'string') return msg('schema.match.text', { text: match });
  return msg('schema.match.pattern', { pattern: match instanceof RegExp ? match.source : match.pattern });
}

/** Display form of an expectation for labels and headers. */
export function matchLabel(match: TextMatch): string {
  if (typeof match === 'string') return match;
  return match instanceof RegExp ? `/${match.source}/` : `/${match.pattern}/`;
}

/** The schema as JSON. Code hooks are dropped; regular expressions become { pattern, flags }. */
export function schemaToJSON(schema: WorkbookSchema): string {
  return JSON.stringify(
    schema,
    (_key, value: unknown) => {
      if (typeof value === 'function') return undefined;
      // JSON writes functions in arrays (such as table rules) as null, so leave them out first.
      if (Array.isArray(value) && value.some((v) => typeof v === 'function')) {
        const rest = value.filter((v) => typeof v !== 'function');
        return rest.length ? rest : undefined;
      }
      if (value instanceof RegExp) return { pattern: value.source, flags: value.flags || undefined };
      return value;
    },
    2,
  );
}

const FIELD_TYPES: FieldType[] = ['text', 'number', 'integer', 'percent', 'date', 'boolean', 'enum'];

/** Text in one language or several: a non-empty string, or an object of strings. */
function isLocalized(x: unknown): boolean {
  if (typeof x === 'string') return x !== '';
  return !!x && typeof x === 'object' && !Array.isArray(x) && Object.keys(x).length > 0 && Object.values(x).every((v) => typeof v === 'string');
}

/**
 * Parses and sanity-checks a schema exported as JSON. Throws a TextError that
 * says what is wrong: its `text` to show in the editor's language, its
 * `message` in English.
 */
export function schemaFromJSON(text: string): WorkbookSchema {
  const fail = (key: MessageKey, params?: MessageParams): never => {
    const error = new TextError(msg(key, params));
    error.message = english.text(error.text);
    throw error;
  };
  const checkPattern = (pattern: unknown, field: string, flags?: string) => {
    try {
      if (typeof pattern !== 'string') throw new Error();
      new RegExp(pattern, flags);
    } catch {
      fail('schema.json.badPattern', { field });
    }
    if (isUnsafePattern(pattern as string)) fail('schema.json.unsafePattern', { field });
  };
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch (e) {
    fail('schema.json.notJson', { error: (e as Error).message });
  }
  const s = data as WorkbookSchema;
  if (!s || typeof s !== 'object') fail('schema.json.notObject');
  if (typeof s.id !== 'string' || !s.id) fail('schema.json.noId');
  if (!isLocalized(s.name)) fail('schema.json.noName');
  if (!Array.isArray(s.sheets) || !s.sheets.length) fail('schema.json.noSheets');
  if (s.version !== undefined && typeof s.version !== 'string') fail('schema.json.badVersion');
  if (s.hookTimeoutMs !== undefined && !(typeof s.hookTimeoutMs === 'number' && s.hookTimeoutMs > 0)) fail('schema.json.badTimeout');
  s.sheets.forEach((sheet, i) => {
    if (!sheet.name) fail('schema.json.sheetNoName', { sheet: i + 1 });
    for (const t of sheet.tables ?? []) {
      if (!t.key) fail('schema.json.tableNoKey', { sheet: i + 1 });
      if (typeof t.at !== 'string' || !/^\$?[A-Z]{1,3}\$?\d+$/i.test(t.at)) fail('schema.json.tableBadAt', { table: t.key });
      if (!Array.isArray(t.fields) || !t.fields.length) fail('schema.json.tableNoFields', { table: t.key });
      const keys = new Set<string>();
      for (const f of t.fields) {
        if (!f.key) fail('schema.json.fieldNoKey', { table: t.key });
        if (keys.has(f.key)) fail('schema.json.duplicateField', { table: t.key, field: f.key });
        keys.add(f.key);
        if (!FIELD_TYPES.includes(f.type)) fail('schema.json.unknownType', { field: f.key, type: String(f.type) });
        if (f.type === 'enum' && !f.values?.length) fail('schema.json.enumNoValues', { field: f.key });
        if (f.label === undefined || f.label === null || f.label === '') fail('schema.json.fieldNoLabel', { table: t.key, field: f.key });
        if (f.pattern !== undefined) checkPattern(f.pattern, f.key);
      }
      for (const c of t.checks ?? []) {
        const named = [c.field, ...(c.rule === 'compare' ? [c.other] : c.rule === 'requiredIf' ? [c.when] : c.rule === 'uniqueTogether' ? c.fields : [])];
        for (const k of named) if (k !== '*' && !keys.has(k)) fail('schema.json.unknownCheckField', { rule: c.rule, table: t.key, field: k });
      }
    }
    for (const c of sheet.cells ?? []) {
      if (!c.key || !c.ref) fail('schema.json.cellNoKey', { sheet: i + 1 });
      if (!FIELD_TYPES.includes(c.type)) fail('schema.json.cellUnknownType', { cell: c.key, type: String(c.type) });
      if (c.pattern !== undefined) checkPattern(c.pattern, c.key);
    }
  });
  // Names and labels can be patterns too: { pattern, flags }.
  JSON.stringify(s, (_k, v: unknown) => {
    if (v && typeof v === 'object' && 'pattern' in v && typeof (v as { pattern: unknown }).pattern === 'string' && !('key' in v)) {
      const { pattern, flags } = v as { pattern: string; flags?: unknown };
      checkPattern(pattern, pattern, typeof flags === 'string' ? flags : undefined);
    }
    return v;
  });
  return s;
}
