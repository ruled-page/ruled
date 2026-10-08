import { describe, expect, it } from 'vitest';
import { english, errorText } from '../src/i18n';
import { schemaFromJSON } from '../src/schema/json';
import { isUnsafePattern } from '../src/schema/patterns';

describe('patterns that could take minutes to check', () => {
  it('are told apart from ordinary ones', () => {
    for (const p of ['(a+)+$', '^(a*)*$', '(a|aa)+', '((ab)*c)+', '^(\\d+)*$', '(.*a){2,}', '(x(a+))*']) expect(isUnsafePattern(p), p).toBe(true);
    for (const p of ['^[A-Z]{2}-\\d{4}$', '^(\\d{3}-)?\\d{4}$', '(abc)+', '[(a+)+]', '\\(a+\\)+', '^SKU-\\d+$', '(ab){2,3}']) expect(isUnsafePattern(p), p).toBe(false);
  });
});

describe('schemas loaded from JSON', () => {
  const thrown = (schema: unknown) => {
    try {
      schemaFromJSON(JSON.stringify(schema));
    } catch (e) {
      return english.text(errorText(e));
    }
    return undefined;
  };
  const table = (field: Record<string, unknown>) => ({ id: 's', name: 'S', sheets: [{ name: 'Sheet1', tables: [{ key: 't', at: 'A1', fields: [field] }] }] });

  it('refuse patterns that could hang the page, fields without labels, and bad patterns for names', () => {
    expect(thrown(table({ key: 'code', label: 'Code', type: 'text', pattern: '^(a+)+$' }))).toBe(
      'The pattern for "code" could take minutes to check on some values: a group that repeats holds something that repeats too, like (a+)+. Simplify it.',
    );
    expect(thrown(table({ key: 'code', type: 'text' }))).toBe('The schema file is not valid: field "code" in table "t" has no label.');
    expect(thrown({ id: 's', name: 'S', sheets: [{ name: { pattern: '([a-z]+)*' } }] })).toMatch(/could take minutes/);
    expect(thrown({ id: 's', name: 'S', sheets: [{ name: { pattern: '(' } }] })).toMatch(/invalid pattern/);
    expect(thrown(table({ key: 'code', label: 'Code', type: 'text', pattern: '^SKU-\\d+$' }))).toBeUndefined();
  });
});
