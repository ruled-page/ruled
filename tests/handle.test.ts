import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { handleFor, storeOf, type EditorHandle } from '../src/editor/handle';
import type { WorkbookSchema } from '../src/schema/types';
import type { EditorEvent, StoreOptions } from '../src/state/store';
import { testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

let bytes: Uint8Array;
beforeAll(async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Lines').addRow([2, 3, { formula: 'A1*B1', result: 6 }]);
  bytes = new Uint8Array(await wb.xlsx.writeBuffer());
});

let editor: EditorHandle;
afterEach(() => editor?.destroy());

async function create(options: StoreOptions = {}): Promise<EditorHandle> {
  editor = handleFor(await testStore(options));
  return editor;
}

/** Answers the editor's next question. */
async function answer(reply: 'confirm' | 'cancel'): Promise<void> {
  const store = storeOf(editor);
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  store.confirmRequest!.resolve(reply);
}

describe('the editor handle', () => {
  it('opens a workbook, leaving the caller its bytes, and says when the user kept the one open', async () => {
    await create();
    const mine = bytes.slice();
    expect(await editor.open(mine, 'lines.xlsx')).toBe(true);
    expect(mine.length).toBe(bytes.length);
    expect([editor.status, editor.fileName, editor.editedCount, editor.unsaved]).toEqual(['ready', 'lines.xlsx', 0, false]);

    await storeOf(editor).input(0, 0, '5');
    const again = editor.open(bytes, 'other.xlsx');
    await answer('cancel');
    expect(await again).toBe(false);
    expect(editor.fileName).toBe('lines.xlsx');
  });

  it('rejects, saying why, when the file can’t be opened', async () => {
    await create();
    await expect(editor.open(new TextEncoder().encode('not a workbook'), 'x.xlsx')).rejects.toThrow('This file is not an Excel workbook (.xlsx or .xlsm).');
    expect(editor.status).toBe('error');
    await expect(editor.open(bytes, 'notes.txt')).rejects.toThrow(/\.xlsx/);
  });

  it('says what saving came to', async () => {
    const saves: string[] = [];
    await create({ onSave: (f) => void saves.push(f.fileName) });
    await editor.open(bytes, 'lines.xlsx');
    expect(await editor.save()).toBe('saved');
    expect(saves).toEqual(['lines.xlsx']);

    // A schema whose last check stops the save.
    const stop: WorkbookSchema = { id: 'stop', name: 'Stop', sheets: [{ name: 'Lines' }], beforeSave: () => 'Not today.' };
    editor.registerSchema(stop);
    expect(await editor.save()).toBe('blocked');
    editor.removeSchema('stop');

    // The host's onSave fails.
    editor.destroy();
    await create({ onSave: () => Promise.reject(new Error('Offline')) });
    await editor.open(bytes, 'lines.xlsx');
    expect(await editor.save()).toBe('failed');
  });

  it('gives the workbook as it is, goes to cells, and changes user and language while open', async () => {
    await create({ user: 'Ana' });
    await editor.open(bytes, 'lines.xlsx');
    editor.goTo('B1');
    expect(storeOf(editor).selection).toMatchObject({ row: 0, col: 1 });
    editor.setUser('Ben');
    await storeOf(editor).input(0, 0, '4');
    expect(storeOf(editor).history.at(-1)?.user).toBe('Ben');
    editor.setLanguage('de');
    expect(storeOf(editor).t.language).toBe('de');
    const now = await editor.getWorkbook();
    expect(now?.editedCount).toBe(1);
    const listener = vi.fn();
    const stop = editor.subscribe(listener);
    await storeOf(editor).input(0, 0, '6');
    expect(listener).toHaveBeenCalled();
    stop();
    expect(() => storeOf({} as EditorHandle)).toThrow('Not an editor made by createWorkbookEditor().');
  });
});

describe('what the editor tells the host', () => {
  it('reports opening, calculating, saving and errors, in English, with timings', async () => {
    const events: EditorEvent[] = [];
    await create({ onEvent: (e) => events.push(e), language: 'de', onSave: () => undefined });
    await editor.open(bytes, 'lines.xlsx');
    for (let i = 0; i < 100 && !events.some((e) => e.type === 'calculated'); i++) await new Promise((r) => setTimeout(r, 10));
    await editor.save();
    await editor.open(new TextEncoder().encode('nope'), 'x.xlsx').catch(() => undefined);
    expect(events.map((e) => e.type)).toEqual(['opened', 'calculated', 'saved', 'openFailed']);
    expect(events[0]).toMatchObject({ fileName: 'lines.xlsx', sheets: 1, ms: expect.any(Number) });
    expect(events[1]).toMatchObject({ formulas: 1, live: 1, notCalculated: 0 });
    expect(events[2]).toMatchObject({ result: 'saved', editedCount: 0 });
    expect(events[3]).toEqual({ type: 'openFailed', fileName: 'x.xlsx', reason: 'This file is not an Excel workbook (.xlsx or .xlsm).' });

    // Errors shown to the user are reported too, and stay until dismissed.
    vi.useFakeTimers();
    storeOf(editor).toast('error', 'Something broke');
    vi.advanceTimersByTime(60_000);
    vi.useRealTimers();
    expect(events.at(-1)).toEqual({ type: 'error', message: 'Something broke' });
    expect(storeOf(editor).toasts.map((t) => t.text)).toContain('Something broke');
  });
});
