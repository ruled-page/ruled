import { createElement } from 'react';
import { renderToString } from 'react-dom/server';
import { describe, expect, it } from 'vitest';
import { plainSaved, sameBytes } from '../src/editor/plainData';
import type { SavedWorkbook } from '../src/state/store';

describe('rendering on a server', () => {
  it('loads the editor with no browser, and renders <Workbook> as an empty box until the browser has it', async () => {
    // A server has neither; nothing may reach for them as the editor's modules load or first render.
    expect(typeof window).toBe('undefined');
    expect(typeof document).toBe('undefined');
    const { Workbook } = await import('../src/editor/index');
    const html = renderToString(createElement(Workbook, { file: { name: 'orders.xlsx', bytes: new Uint8Array([1, 2, 3]) }, className: 'host-editor', language: 'de' }));
    expect(html).toBe('<div class="wbe wbe-root host-editor" aria-busy="true"></div>');
  });
});

describe('what a server is sent and sends', () => {
  it('sends a saved workbook as plain data, without the functions that read its cells', () => {
    const saved = { bytes: new Uint8Array([1]), fileName: 'a.xlsx', editedCount: 1, changes: [], history: [], outOfDate: [], cell: () => null, range: () => [] } as SavedWorkbook;
    const data = plainSaved(saved);
    expect(data).toEqual({ bytes: saved.bytes, fileName: 'a.xlsx', editedCount: 1, changes: [], history: [], outOfDate: [] });
    expect(Object.values(data).some((v) => typeof v === 'function')).toBe(false);
  });

  it('knows a new copy of the same bytes, as each render of a Server Component sends, from another file', () => {
    const bytes = new Uint8Array([1, 2, 3]);
    expect(sameBytes(bytes, bytes)).toBe(true);
    expect(sameBytes(bytes, new Uint8Array([1, 2, 3]))).toBe(true);
    expect(sameBytes(bytes, new Uint8Array([1, 2, 3]).buffer)).toBe(true);
    expect(sameBytes(bytes.buffer, new Uint8Array([1, 2, 4]))).toBe(false);
    expect(sameBytes(bytes, new Uint8Array([1, 2]))).toBe(false);
  });
});
