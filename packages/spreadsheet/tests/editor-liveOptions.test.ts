import { describe, expect, it } from 'vitest';
import type { EditorHandle } from '../src/editor/handle';
import { applyLiveOptions, changedOptions } from '../src/editor/liveOptions';
import { defineSchema } from '../src/schema/types';

const schema = (id: string) => defineSchema({ id, name: id, sheets: [] });
const a = schema('a');
const b = schema('b');

/** A handle that records what it was told. */
function recorder() {
  const calls: string[] = [];
  const handle = {
    setTheme: (t: unknown) => calls.push(`theme ${JSON.stringify(t)}`),
    setLanguage: (l: string) => calls.push(`language ${l}`),
    setUser: (u?: string) => calls.push(`user ${u}`),
    setCustomLists: (l?: string[][]) => calls.push(`lists ${JSON.stringify(l)}`),
    registerSchema: (s: { id: string }) => calls.push(`register ${s.id}`),
    removeSchema: (id: string) => calls.push(`remove ${id}`),
  } as unknown as EditorHandle;
  return { handle, calls };
}

describe('options that change while the editor is open', () => {
  it('counts a new literal with the same entries as no change', () => {
    expect(changedOptions({ theme: { accent: '#000' }, schemas: [a], customLists: [['N', 'S']] }, { theme: { accent: '#000' }, schemas: [a], customLists: [['N', 'S']] })).toEqual([]);
    expect(changedOptions({ theme: { accent: '#000' } }, { theme: { accent: '#fff' } })).toEqual(['theme']);
    expect(changedOptions({ schemas: [a] }, { schemas: [a, b] })).toEqual(['schemas']);
    expect(changedOptions({ customLists: [['N', 'S']] }, { customLists: [['N', 'E']] })).toEqual(['customLists']);
    expect(changedOptions({ language: 'en' }, { language: 'en', user: 'Ana' })).toEqual(['user']);
    // A new callback is a change of an option the editor reads when made.
    expect(changedOptions({ onSave: () => {} }, { onSave: () => {} })).toEqual(['onSave']);
  });

  it('applies the theme, language, user, lists and schemas, removing schemas no longer listed', () => {
    const { handle, calls } = recorder();
    applyLiveOptions(handle, { theme: {}, language: 'en', user: 'Ana', schemas: [a, b] }, { theme: { accent: '#fff' }, language: 'de', user: 'Ben', customLists: [['N']], schemas: [b, a] });
    expect(calls).toEqual(['theme {"accent":"#fff"}', 'language de', 'user Ben', 'lists [["N"]]', 'register a', 'register b']);
    const removed = recorder();
    applyLiveOptions(removed.handle, { schemas: [a, b] }, { schemas: [b], language: undefined });
    expect(removed.calls).toEqual(['remove a', 'register b']);
  });
});
