import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { VT } from '../src/core/types';
import { english } from '../src/i18n';
import type { Answer, ConfirmRequest, StoreOptions, WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

// An unexpected question would otherwise hang a test until the suite's long timeout.
vi.setConfig({ testTimeout: 10_000 });

/**
 * Lines!A1:D4 = Item, Qty, Price, Total / Bolts, 2, 3, =B2*C2 / Nuts, 4, 0.5, =B3*C3 / Washers, 1, 1, 1,
 * F1:G1 merged; Other!A1 = "x"; Locked is a protected sheet with A1 = 1.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const lines = wb.addWorksheet('Lines');
  lines.addRow(['Item', 'Qty', 'Price', 'Total']);
  lines.addRow(['Bolts', 2, 3, { formula: 'B2*C2', result: 6 }]);
  lines.addRow(['Nuts', 4, 0.5, { formula: 'B3*C3', result: 2 }]);
  lines.addRow(['Washers', 1, 1, 1]);
  lines.getCell('F1').value = 'Notes';
  lines.mergeCells('F1:G1');
  wb.addWorksheet('Other').getCell('A1').value = 'x';
  const locked = wb.addWorksheet('Locked');
  locked.getCell('A1').value = 1;
  await locked.protect('', {});
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
afterEach(() => {
  store?.dispose();
  vi.unstubAllGlobals();
});

async function open(options: StoreOptions = {}): Promise<WorkbookStore> {
  store = await testStore(options);
  await opened(store, bytes, 'lines.xlsx');
  return store;
}

const cell = (row: number, col: number, sheet = store.active) => store.sheets[sheet].get(row, col);
const toasts = () => store.toasts.map((t) => english.text(t.text));
const lastAction = () => english.text(store.history.at(-1)?.action);
const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });

/** Waits for the store to ask something, gives the answer, and returns what was asked. */
async function answer(reply: Answer): Promise<ConfirmRequest> {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  const asked = store.confirmRequest;
  if (!asked) throw new Error('Nothing was asked');
  asked.resolve(reply);
  return asked;
}

describe('typing into a cell', () => {
  it('stores the value, recalculates what depends on it and records the edit', async () => {
    await open({ user: 'Ana' });
    await store.input(1, 1, '5');
    expect(cell(1, 1)).toMatchObject({ t: VT.Number, n: 5 });
    expect(cell(1, 3)?.n).toBe(15);
    expect(store.editedCount).toBe(1);
    expect(store.unsaved).toBe(true);
    expect(store.history).toMatchObject([{ user: 'Ana', cells: 1, refs: ['Lines!B2'] }]);
    expect(lastAction()).toBe('Edit');
    expect(store.busy).toBe(false);
  });

  it('refuses read-only cells, saying why', async () => {
    await open();
    await store.input(0, 6, 'x');
    expect(toasts()).toEqual(['Part of a merged cell. Edit its top-left cell.']);
    store.setActive(2);
    await store.input(0, 0, '2');
    expect(toasts()[1]).toMatch(/protected in Excel/);
    expect(cell(0, 0)?.n).toBe(1);
    expect(store.history).toHaveLength(0);
  });

  it('asks before replacing a formula with a value, and keeps the formula if told not to', async () => {
    await open();
    const typing = store.input(1, 3, '7');
    const asked = await answer('cancel');
    await typing;
    expect(english.text(asked.title)).toBe('Replace a formula?');
    expect(cell(1, 3)?.f).toBe('B2*C2');

    const again = store.input(1, 3, '7');
    await answer('confirm');
    await again;
    expect(cell(1, 3)).toMatchObject({ n: 7, f: undefined });
  });

  it('replaces a formula with another formula without asking', async () => {
    await open();
    await store.input(1, 3, '=B2+C2');
    expect(store.confirmRequest).toBeUndefined();
    expect(cell(1, 3)).toMatchObject({ f: 'B2+C2', n: 5 });
  });

  it('shows a typed percentage as one, with the format the edit made', async () => {
    await open();
    const styles = store.styles.length;
    await store.input(4, 2, '12.5%');
    expect(cell(4, 2)?.n).toBe(0.125);
    expect(store.styles.length).toBe(styles + 1);
    // As in Excel, a percentage typed with decimals gets the 0.00% format.
    expect(store.shownText(store.view!, 4, 2)).toBe('12.50%');
  });

  it('says what is wrong with a formula it cannot accept, and changes nothing', async () => {
    await open();
    await store.input(4, 0, '=SUM(');
    expect(store.toasts[0].kind).toBe('error');
    expect(toasts()[0]).toMatch(/problem with this formula/);
    expect(store.editedCount).toBe(0);
    expect(store.undoStack).toHaveLength(0);
  });
});

describe('clearing cells', () => {
  it('clears values without asking', async () => {
    await open();
    await store.clear([rect(3, 0, 3, 1)]);
    expect(store.confirmRequest).toBeUndefined();
    expect(cell(3, 0)?.t ?? VT.None).toBe(VT.None);
    expect(cell(3, 1)?.t ?? VT.None).toBe(VT.None);
    expect(lastAction()).toBe('Clear');
    expect(store.history.at(-1)?.cells).toBe(2);
  });

  it('asks once for every formula that would be lost', async () => {
    await open();
    const clearing = store.clear([rect(1, 3, 2, 3), rect(3, 3)]);
    const asked = await answer('cancel');
    await clearing;
    expect(english.text(asked.title)).toBe('Replace 2 formulas?');
    expect(english.text(asked.confirmLabel)).toBe('Replace formulas');
    expect(cell(1, 3)?.f).toBe('B2*C2');
    expect(store.history).toHaveLength(0);

    const again = store.clear([rect(1, 3, 2, 3)]);
    await answer('confirm');
    await again;
    expect(cell(1, 3)?.f).toBeUndefined();
    expect(cell(2, 3)?.f).toBeUndefined();
  });
});

describe('undo and redo of cell edits', () => {
  it('puts cells back and forward again, recording each in the history', async () => {
    await open();
    await store.input(1, 1, '5');
    await store.undo();
    expect(cell(1, 1)?.n).toBe(2);
    expect(cell(1, 3)?.n).toBe(6);
    expect(store.editedCount).toBe(0);
    expect(lastAction()).toBe('Undo');
    expect([store.undoStack.length, store.redoStack.length]).toEqual([0, 1]);

    await store.redo();
    expect(cell(1, 1)?.n).toBe(5);
    expect(cell(1, 3)?.n).toBe(15);
    expect(lastAction()).toBe('Redo');
    expect([store.undoStack.length, store.redoStack.length]).toEqual([1, 0]);
    expect(store.history.map((h) => english.text(h.action))).toEqual(['Edit', 'Undo', 'Redo']);
  });

  it('does nothing with nothing to undo or redo', async () => {
    await open();
    await store.undo();
    await store.redo();
    expect(store.history).toHaveLength(0);
  });

  it('forgets what could be redone once something new is edited', async () => {
    await open();
    await store.input(1, 1, '5');
    await store.undo();
    await store.input(1, 2, '10');
    expect(store.redoStack).toHaveLength(0);
    await store.redo();
    expect(cell(1, 1)?.n).toBe(2);
  });

  it('goes to the sheet whose cells it puts back', async () => {
    await open();
    store.setActive(1);
    await store.input(0, 0, 'y');
    store.setActive(0);
    await store.undo();
    expect(store.active).toBe(1);
    expect(cell(0, 0)?.s).toBe('x');
    store.setActive(0);
    await store.redo();
    expect(store.active).toBe(1);
    expect(cell(0, 0)?.s).toBe('y');
  });

  it('applies edits made in quick succession in order, each undone on its own', async () => {
    await open();
    await Promise.all([store.input(4, 0, 'first'), store.input(4, 0, 'second')]);
    expect(cell(4, 0)?.s).toBe('second');
    expect(store.undoStack).toHaveLength(2);
    await store.undo();
    expect(cell(4, 0)?.s).toBe('first');
  });
});

describe('pasting text from outside the workbook', () => {
  it('parses each value like typed input', async () => {
    await open();
    await store.paste({ row: 4, col: 0 }, [['Screws', '3', '2', '=B5*C5']], rect(4, 0));
    expect(cell(4, 0)?.s).toBe('Screws');
    expect(cell(4, 1)?.n).toBe(3);
    expect(cell(4, 3)).toMatchObject({ f: 'B5*C5', n: 6 });
    expect(lastAction()).toBe('Paste');
    expect(store.history.at(-1)?.refs).toEqual(['Lines!A5', 'Lines!B5', 'Lines!C5', 'Lines!D5']);
  });

  it('fills a selection that is a whole multiple of the pasted block', async () => {
    await open();
    await store.paste({ row: 9, col: 0 }, [['a', 'b']], rect(9, 0, 10, 3));
    const shown = [9, 10].map((r) => [0, 1, 2, 3].map((c) => cell(r, c)?.s));
    expect(shown).toEqual([
      ['a', 'b', 'a', 'b'],
      ['a', 'b', 'a', 'b'],
    ]);
  });

  it('pastes the block once when the selection is not a multiple of it', async () => {
    await open();
    await store.paste({ row: 9, col: 0 }, [['a', 'b']], rect(9, 0, 9, 2));
    expect([cell(9, 0)?.s, cell(9, 1)?.s, cell(9, 2)]).toEqual(['a', 'b', null]);
  });

  it('asks before pasting values over formulas, but not formulas over formulas', async () => {
    await open();
    const pasting = store.paste({ row: 1, col: 3 }, [['1'], ['2']], rect(1, 3));
    expect(english.text((await answer('cancel')).title)).toBe('Replace 2 formulas?');
    await pasting;
    expect(cell(1, 3)?.f).toBe('B2*C2');

    await store.paste({ row: 1, col: 3 }, [['=B2'], ['=B3']], rect(1, 3));
    expect(cell(2, 3)).toMatchObject({ f: 'B3', n: 4 });
  });

  it('pastes what only looks like a formula as text, saying so', async () => {
    await open();
    await store.paste({ row: 4, col: 0 }, [['=SUM(']], rect(4, 0));
    expect(cell(4, 0)).toMatchObject({ t: VT.String, s: '=SUM(' });
    expect(store.toasts.map((t) => t.kind)).toEqual(['warn']);
    expect(toasts()).toEqual(["1 pasted value looked like a formula but wasn't valid, so it was pasted as text."]);
  });

  it('ignores an empty clipboard', async () => {
    await open();
    await store.paste({ row: 0, col: 0 }, [], rect(0, 0));
    expect(store.history).toHaveLength(0);
  });

  it('says how many cells it had to leave alone, and why', async () => {
    await open();
    store.setActive(2);
    await store.paste({ row: 0, col: 0 }, [['5', '6']], rect(0, 0));
    expect(toasts()).toEqual(['2 cells were left unchanged. This sheet is protected in Excel, so it is read-only here.']);
    await store.paste({ row: 0, col: 0 }, [['5']], rect(0, 0));
    expect(toasts()[1]).toBe('This sheet is protected in Excel, so it is read-only here.');
    expect(store.history).toHaveLength(0);
  });
});

describe('strict pasting of numbers', () => {
  const values = [['007', '1234567890123456789', 'plain']];

  it('asks about values that would not read back as pasted, listing them', async () => {
    await open({ numbers: { paste: 'strict' } });
    const pasting = store.paste({ row: 5, col: 0 }, values, rect(5, 0));
    const asked = await answer('cancel');
    await pasting;
    expect(english.text(asked.title)).toBe('Check 2 pasted values');
    expect(asked.details?.rows).toEqual([
      ['007', '7', 'leading zeros'],
      ['1234567890123456789', '1234567890123456800', 'too many digits'],
    ]);
    expect(asked.details?.more).toBe(0);
    expect(cell(5, 0)).toBeFalsy();
  });

  it('keeps the risky values as text when asked to', async () => {
    await open({ numbers: { paste: 'strict' } });
    const pasting = store.paste({ row: 5, col: 0 }, values, rect(5, 0));
    await answer('alternate');
    await pasting;
    expect(cell(5, 0)).toMatchObject({ t: VT.String, s: '007' });
    expect(cell(5, 1)).toMatchObject({ t: VT.String, s: '1234567890123456789' });
    expect(cell(5, 2)?.s).toBe('plain');
  });

  it('pastes them as numbers when confirmed', async () => {
    await open({ numbers: { paste: 'strict' } });
    const pasting = store.paste({ row: 5, col: 0 }, values, rect(5, 0));
    await answer('confirm');
    await pasting;
    expect(cell(5, 0)).toMatchObject({ t: VT.Number, n: 7 });
  });

  it('pastes without asking when nothing is at risk, or when not strict', async () => {
    await open({ numbers: { paste: 'strict' } });
    await store.paste({ row: 5, col: 0 }, [['12', 'abc']], rect(5, 0));
    expect(cell(5, 0)?.n).toBe(12);
    store.dispose();
    await open();
    await store.paste({ row: 5, col: 0 }, values, rect(5, 0));
    expect(cell(5, 0)?.n).toBe(7);
  });
});

describe('copying and pasting within the workbook', () => {
  it('brings formulas along, adjusting their references', async () => {
    await open();
    store.noteCopy(rect(1, 3));
    await store.paste({ row: 3, col: 3 }, [['6']], rect(3, 3));
    expect(cell(3, 3)).toMatchObject({ f: 'B4*C4', n: 1 });
    expect(lastAction()).toBe('Paste');
  });

  it('pastes the clipboard as text when it holds something other than what was copied here', async () => {
    await open();
    store.noteCopy(rect(1, 3));
    await store.paste({ row: 3, col: 3 }, [['42']], rect(3, 3));
    expect(cell(3, 3)).toMatchObject({ n: 42, f: undefined });
  });

  it('tiles a copied cell over a larger selection', async () => {
    await open();
    store.noteCopy(rect(1, 3));
    await store.paste({ row: 4, col: 3 }, [['6']], rect(4, 3, 6, 3));
    expect([4, 5, 6].map((r) => cell(r, 3)?.f)).toEqual(['B5*C5', 'B6*C6', 'B7*C7']);
  });

  it('asks before a copy replaces formulas with values', async () => {
    await open();
    store.noteCopy(rect(1, 2, 2, 2));
    const pasting = store.paste({ row: 1, col: 3 }, [['3'], ['0.5']], rect(1, 3));
    expect(english.text((await answer('cancel')).title)).toBe('Replace 2 formulas?');
    await pasting;
    expect(cell(1, 3)?.f).toBe('B2*C2');
  });

  it('pastes what was copied here when the browser will not share the clipboard', async () => {
    await open();
    store.noteCopy(rect(1, 0, 2, 1));
    await store.paste({ row: 7, col: 0 }, [], rect(7, 0), true);
    expect([cell(7, 0)?.s, cell(7, 1)?.n, cell(8, 0)?.s, cell(8, 1)?.n]).toEqual(['Bolts', 2, 'Nuts', 4]);
  });

  it('copies exact numbers instead of what cells show when asked to', async () => {
    await open();
    await store.input(5, 0, '=1/3');
    expect(store.selectionText(rect(5, 0))).toBe('0.333333333');
    store.dispose();

    await open({ numbers: { copy: 'raw' } });
    await store.input(5, 0, '=1/3');
    expect(store.selectionText(rect(5, 0))).toBe('0.333333333333333');
    store.noteCopy(rect(5, 0));
    // The clipboard holds what was copied here, so the formula comes along.
    await store.paste({ row: 6, col: 0 }, [['0.333333333333333']], rect(6, 0));
    expect(cell(6, 0)?.f).toBe('1/3');
  });
});

describe('cutting and pasting', () => {
  it('moves the cells, with formulas that used them following', async () => {
    await open();
    store.noteCopy(rect(1, 1), true);
    expect(store.cutRange).toEqual({ sheet: 0, rect: rect(1, 1) });
    await store.paste({ row: 1, col: 5 }, [['2']], rect(1, 5));
    expect(cell(1, 5)?.n).toBe(2);
    expect(cell(1, 1)).toBeFalsy();
    expect(cell(1, 3)).toMatchObject({ f: 'F2*C2', n: 6 });
    expect(lastAction()).toBe('Moved B2');
    expect(store.cutRange).toBeUndefined();
    expect(store.selection.rects).toEqual([rect(1, 5)]);
  });

  it('moves cells to another sheet, saying where', async () => {
    await open();
    store.noteCopy(rect(1, 1), true);
    store.setActive(1);
    await store.paste({ row: 4, col: 0 }, [['2']], rect(4, 0));
    expect(cell(4, 0, 1)?.n).toBe(2);
    expect(cell(1, 3, 0)?.f).toBe('Other!A5*C2');
    expect(lastAction()).toBe('Moved B2 to Other');
  });

  it('lets a cut be called off', async () => {
    await open();
    const version = store.version;
    store.cancelCut();
    expect(store.version).toBe(version);
    store.noteCopy(rect(1, 1), true);
    store.cancelCut();
    expect(store.cutRange).toBeUndefined();
    await store.paste({ row: 1, col: 5 }, [['2']], rect(1, 5));
    expect(cell(1, 1)?.n).toBe(2);
  });
});

describe('copying and pasting from the menu', () => {
  it('puts the selection on the clipboard as tab-separated text, quoting where needed', async () => {
    await open();
    await store.input(2, 0, 'say "hi"\tthere');
    const writeText = vi.fn(async () => undefined);
    vi.stubGlobal('navigator', { clipboard: { writeText } });
    store.setSelection({ row: 1, col: 0, rects: [rect(1, 0, 3, 1)] });
    await store.copySelection(false);
    expect(writeText).toHaveBeenCalledWith('Bolts\t2\n"say ""hi""\tthere"\t4\nWashers\t1');
    expect(store.cutRange).toBeUndefined();
    expect(store.toasts).toHaveLength(0);
  });

  it('copies cells typed beyond the part of the sheet the file used', async () => {
    await open();
    await store.input(4, 0, 'Screws');
    await store.input(4, 4, '9');
    expect(store.selectionText(rect(3, 0, 5, 4))).toBe('Washers\t1\t1\t1\t\nScrews\t\t\t\t9');
    // Rows and columns past everything on the sheet stay off the clipboard (row 5 now has Screws).
    expect(store.selectionText(rect(0, 5, 1048575, 16383))).toBe('Notes\t\n\t\n\t\n\t\n\t');
  });

  it('still cuts or copies for pasting here when the browser refuses the clipboard', async () => {
    await open();
    vi.stubGlobal('navigator', { clipboard: { writeText: async () => Promise.reject(new Error('denied')) } });
    store.setSelection({ row: 1, col: 0, rects: [rect(1, 0)] });
    await store.copySelection(true);
    expect(store.cutRange).toEqual({ sheet: 0, rect: rect(1, 0) });
    await store.copySelection(false);
    expect(toasts()).toEqual([expect.stringMatching(/^Cut for pasting here/), expect.stringMatching(/^Copied for pasting here/)]);
  });

  it('pastes what the clipboard holds', async () => {
    await open();
    vi.stubGlobal('navigator', { clipboard: { readText: async () => 'p\t1\r\nq\t2' } });
    store.setSelection({ row: 6, col: 0, rects: [rect(6, 0)] });
    await store.pasteFromClipboard();
    expect([cell(6, 0)?.s, cell(6, 1)?.n, cell(7, 0)?.s, cell(7, 1)?.n]).toEqual(['p', 1, 'q', 2]);
  });

  it('falls back to what was copied here when the clipboard cannot be read', async () => {
    await open();
    vi.stubGlobal('navigator', { clipboard: { readText: async () => Promise.reject(new Error('denied')) } });
    store.setSelection({ row: 6, col: 0, rects: [rect(6, 0)] });
    await store.pasteFromClipboard();
    expect(toasts()).toEqual([expect.stringMatching(/didn't allow reading the clipboard/)]);
    expect(cell(6, 0)).toBeFalsy();

    store.noteCopy(rect(1, 0));
    await store.pasteFromClipboard();
    expect(cell(6, 0)?.s).toBe('Bolts');
  });
});

describe('filling', () => {
  it('extends a series by dragging', async () => {
    await open();
    await store.fill(rect(1, 1, 2, 1), rect(1, 1, 4, 1));
    expect([cell(3, 1)?.n, cell(4, 1)?.n]).toEqual([6, 8]);
    expect(lastAction()).toBe('Fill');
  });

  it('asks before filling over formulas with values', async () => {
    await open();
    const filling = store.fill(rect(1, 2), rect(1, 2, 1, 3));
    expect(english.text((await answer('cancel')).title)).toBe('Replace a formula?');
    await filling;
    expect(cell(1, 3)?.f).toBe('B2*C2');
  });

  it('drags formulas on without asking about the cells it starts from', async () => {
    await open();
    await store.fill(rect(1, 3, 2, 3), rect(1, 3, 4, 3));
    expect(store.confirmRequest).toBeUndefined();
    expect([cell(3, 3)?.f, cell(4, 3)?.f]).toEqual(['B4*C4', 'B5*C5']);
  });

  it('asks before filling right over a formula, and fills when confirmed', async () => {
    await open();
    store.setSelection({ row: 1, col: 2, rects: [rect(1, 2, 1, 3)] });
    const declined = store.fillDown(true);
    await answer('cancel');
    await declined;
    expect(cell(1, 3)?.f).toBe('B2*C2');
    const filling = store.fillDown(true);
    await answer('confirm');
    await filling;
    expect(cell(1, 3)).toMatchObject({ n: 3, f: undefined });
    expect(lastAction()).toBe('Fill right');
  });

  it('does nothing filling right in the first column', async () => {
    await open();
    store.setSelection({ row: 1, col: 0, rects: [rect(1, 0, 3, 0)] });
    await store.fillDown(true);
    expect(store.history).toHaveLength(0);
  });
});
