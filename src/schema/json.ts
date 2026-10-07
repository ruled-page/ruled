import type { FieldType, TextMatch, WorkbookSchema } from './types';

/** Whether `text` satisfies the expectation: plain strings ignore case and surrounding spaces. */
export function textMatches(match: TextMatch, text: string): boolean {
  if (typeof match === 'string') return match.trim().toLowerCase() === text.trim().toLowerCase();
  const re = match instanceof RegExp ? match : new RegExp(match.pattern, match.flags);
  return re.test(text.trim());
}

/** How an expectation is shown to people. */
export function describeMatch(match: TextMatch): string {
  if (typeof match === 'string') return `"${match}"`;
  const source = match instanceof RegExp ? match.source : match.pattern;
  return `text matching /${source}/`;
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
      if (value instanceof RegExp) return { pattern: value.source, flags: value.flags || undefined };
      return value;
    },
    2,
  );
}

const FIELD_TYPES: FieldType[] = ['text', 'number', 'integer', 'percent', 'date', 'boolean', 'enum'];

/** Parses and sanity-checks a schema exported as JSON. Throws with a readable message. */
export function schemaFromJSON(text: string): WorkbookSchema {
  let data: unknown;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error(`The schema file isn't valid JSON: ${(e as Error).message}`);
  }
  const fail = (msg: string): never => {
    throw new Error(`The schema file is not valid: ${msg}`);
  };
  const s = data as WorkbookSchema;
  if (!s || typeof s !== 'object') fail('expected an object.');
  if (typeof s.id !== 'string' || !s.id) fail('"id" is missing.');
  if (typeof s.name !== 'string' || !s.name) fail('"name" is missing.');
  if (!Array.isArray(s.sheets) || !s.sheets.length) fail('"sheets" must list at least one sheet.');
  if (s.version !== undefined && typeof s.version !== 'string') fail('"version" must be text, e.g. "2".');
  if (s.hookTimeoutMs !== undefined && !(typeof s.hookTimeoutMs === 'number' && s.hookTimeoutMs > 0)) fail('"hookTimeoutMs" must be a positive number of milliseconds.');
  s.sheets.forEach((sheet, i) => {
    if (!sheet.name) fail(`sheet ${i + 1} has no name.`);
    for (const t of sheet.tables ?? []) {
      if (!t.key) fail(`a table on sheet ${i + 1} has no key.`);
      if (typeof t.at !== 'string' || !/^\$?[A-Z]{1,3}\$?\d+$/i.test(t.at)) fail(`table "${t.key}" has an invalid "at" cell.`);
      if (!Array.isArray(t.fields) || !t.fields.length) fail(`table "${t.key}" has no fields.`);
      const keys = new Set<string>();
      for (const f of t.fields) {
        if (!f.key) fail(`a field in table "${t.key}" has no key.`);
        if (keys.has(f.key)) fail(`table "${t.key}" has two fields called "${f.key}".`);
        keys.add(f.key);
        if (!FIELD_TYPES.includes(f.type)) fail(`field "${f.key}" has unknown type "${String(f.type)}".`);
        if (f.type === 'enum' && !f.values?.length) fail(`enum field "${f.key}" lists no values.`);
        if (f.pattern) {
          try {
            new RegExp(f.pattern);
          } catch {
            fail(`field "${f.key}" has an invalid pattern.`);
          }
        }
      }
      for (const c of t.checks ?? []) {
        const named = [c.field, ...(c.rule === 'compare' ? [c.other] : c.rule === 'requiredIf' ? [c.when] : c.rule === 'uniqueTogether' ? c.fields : [])];
        for (const k of named) if (k !== '*' && !keys.has(k)) fail(`a "${c.rule}" check in table "${t.key}" names "${k}", which isn't one of its fields.`);
      }
    }
    for (const c of sheet.cells ?? []) {
      if (!c.key || !c.ref) fail(`a cell on sheet ${i + 1} is missing its key or ref.`);
      if (!FIELD_TYPES.includes(c.type)) fail(`cell "${c.key}" has unknown type "${String(c.type)}".`);
    }
  });
  return s;
}
