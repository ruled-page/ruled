import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { linkAddress } from '../src/core/links';
import { WorkbookSession } from '../src/core/session';
import { english } from '../src/i18n';
import type { StoreOptions, WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/**
 * "Data": A1 links to a web page, A2 to Other!B3, A3 to a script, A4 to an
 * email, A5 has a HYPERLINK formula, and a defined name points at Other!C5.
 * "Other" is plain; "Hidden" is hidden; "Locked" is protected.
 */
async function workbook(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const data = wb.addWorksheet('Data');
  data.getCell('A1').value = { text: 'Example', hyperlink: 'https://example.com/a', tooltip: 'Example site' };
  data.getCell('A2').value = { text: 'Other', hyperlink: '#Other!B3' };
  data.getCell('A3').value = { text: 'Script', hyperlink: 'javascript:alert(1)' };
  data.getCell('A4').value = { text: 'Mail', hyperlink: 'mailto:someone@example.com' };
  data.getCell('A5').value = { formula: 'HYPERLINK("#Total","Total")', result: 'Total' };
  data.getCell('A6').value = { text: 'Hidden', hyperlink: '#Hidden!A1' };
  data.getCell('B1').value = 'plain';
  const other = wb.addWorksheet('Other');
  other.getCell('A1').value = 1;
  wb.addWorksheet('Hidden', { state: 'hidden' }).getCell('A1').value = 1;
  const locked = wb.addWorksheet('Locked');
  locked.getCell('A1').value = 1;
  await locked.protect('', {});
  wb.definedNames.add('Other!$C$5', 'Total');
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

let bytes: Uint8Array;
beforeAll(async () => {
  bytes = await workbook();
});

let store: WorkbookStore;
let opens: ReturnType<typeof vi.fn>;
afterEach(() => {
  store?.dispose();
  vi.unstubAllGlobals();
});

async function open(options: StoreOptions = {}): Promise<WorkbookStore> {
  opens = vi.fn();
  vi.stubGlobal('open', opens);
  store = await testStore(options);
  await opened(store, bytes, 'links.xlsx');
  return store;
}

const rect = (r1: number, c1: number, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const select = (r: ReturnType<typeof rect>) => store.setSelection({ row: r.r1, col: r.c1, rects: [r] });
const toasts = () => store.toasts.map((t) => english.text(t.text));
const actions = () => store.history.map((h) => english.text(h.action));
async function answer(reply: 'confirm' | 'cancel'): Promise<void> {
  for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
  store.confirmRequest!.resolve(reply);
}
async function saved() {
  const s = await store.snapshot();
  return WorkbookSession.open(s!.bytes, 'links.xlsx').model.sheets[0].info;
}

describe('following links', () => {
  it('goes to a place in the workbook, by reference or defined name', async () => {
    await open();
    await store.followLink(1, 0);
    expect(store.active).toBe(1);
    expect(store.selection).toMatchObject({ row: 2, col: 1 });
    store.setActive(0);
    // The HYPERLINK formula's address names the defined name.
    await store.followLink(4, 0);
    expect([store.active, store.selection.row, store.selection.col]).toEqual([1, 4, 2]);
  });

  it('asks before opening a web page the first time, then opens it in a new tab with no way back to the editor', async () => {
    await open();
    const following = store.followLink(0, 0);
    await answer('confirm');
    await following;
    expect(opens).toHaveBeenCalledWith('https://example.com/a', '_blank', 'noopener,noreferrer');
    // Asked once per address while the workbook is open.
    await store.followLink(0, 0);
    expect(opens).toHaveBeenCalledTimes(2);
    expect(store.confirmRequest).toBeUndefined();
  });

  it('shows the full address when asking, and opens nothing when the answer is no', async () => {
    await open();
    const following = store.followLink(3, 0);
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    expect(english.text(store.confirmRequest!.title)).toBe('Write an email?');
    expect(store.confirmRequest!.details?.rows).toEqual([['mailto:someone@example.com']]);
    store.confirmRequest!.resolve('cancel');
    await following;
    expect(opens).not.toHaveBeenCalled();
  });

  it('refuses a link to a script, saying why', async () => {
    await open();
    await store.followLink(2, 0);
    expect(opens).not.toHaveBeenCalled();
    expect(store.confirmRequest).toBeUndefined();
    expect(toasts()[0]).toMatch(/^This link can't be opened here: “javascript:alert\(1\)”/);
  });

  it("says when a place isn't there, or a cell has no link", async () => {
    await open();
    await store.followLink(5, 0);
    await store.followLink(0, 1);
    expect(toasts()).toEqual(["The place this link goes to isn't in the workbook: Hidden!A1", 'This cell has no link.']);
  });

  it('hands web pages to the host instead, carrying on when it says so', async () => {
    const onLinkClick = vi.fn().mockReturnValueOnce(undefined).mockResolvedValueOnce(true).mockRejectedValueOnce(new Error('No network'));
    await open({ onLinkClick });
    await store.followLink(0, 0);
    expect(onLinkClick).toHaveBeenCalledWith({ url: 'https://example.com/a', kind: 'web', sheet: 'Data', ref: 'A1' });
    expect(opens).not.toHaveBeenCalled();
    const following = store.followLink(0, 0);
    await answer('confirm');
    await following;
    expect(opens).toHaveBeenCalledTimes(1);
    await store.followLink(0, 0);
    expect(toasts()).toEqual(['No network']);
    // Places in the workbook don't go to the host.
    await store.followLink(1, 0);
    expect(onLinkClick).toHaveBeenCalledTimes(3);
  });

  it('can be turned off', async () => {
    await open({ links: 'off' });
    expect(store.followsLinks).toBe(false);
    await store.followLink(1, 0);
    expect(store.active).toBe(0);
    expect(toasts()).toEqual(["Links aren't followed in this editor."]);
  });
});

describe('editing links', () => {
  it('gives the selected cells a link, with the text to display, undone as one step', async () => {
    await open();
    select(rect(1, 1));
    await store.setLink({ target: 'https://example.com/new', display: 'New', tooltip: 'Tip' }, 'New');
    expect(store.linkAt(1, 1)).toMatchObject({ target: 'https://example.com/new', tooltip: 'Tip' });
    expect(store.view!.get(1, 1)?.s).toBe('New');
    // The cell takes Excel's Hyperlink look, as it does when Excel inserts a link.
    expect(store.styles[store.view!.get(1, 1)!.style]).toMatchObject({ linkFont: true, underline: true });
    expect(actions()).toEqual(['Insert link']);
    expect(store.editedCount).toBe(2);
    const info = await saved();
    expect(info.links.map(linkAddress)).toContain('https://example.com/new');
    await store.undo();
    expect(store.linkAt(1, 1)).toBeUndefined();
    expect(store.view!.get(1, 1)?.s).toBeUndefined();
    expect(store.editedCount).toBe(0);
    await store.redo();
    expect(store.linkAt(1, 1)?.target).toBe('https://example.com/new');
  });

  it("changes a cell's link, keeping text that starts = as text", async () => {
    await open();
    select(rect(0, 0));
    await store.setLink({ location: "'Other'!A1" }, '=not a formula');
    expect(store.linkAt(0, 0)).toMatchObject({ location: "'Other'!A1" });
    expect(store.view!.get(0, 0)?.s).toBe('=not a formula');
    expect(actions()).toEqual(['Edit link']);
    await store.refreshChanges();
    expect(store.changeList.map((c) => english.text(c.note ?? c.after))).toEqual(['=not a formula', "Link changed to 'Other'!A1"]);
  });

  it('removes the links of the selected cells, which keep what they hold', async () => {
    await open();
    select(rect(0, 0, 2, 0));
    expect(store.linksInSelection()).toBe(3);
    await store.removeLinks();
    expect(store.linksInSelection()).toBe(0);
    expect(store.view!.get(0, 0)?.s).toBe('Example');
    expect(actions()).toEqual(['Remove link']);
    // Nothing to remove changes nothing.
    await store.removeLinks();
    expect(actions()).toHaveLength(1);
    expect((await saved()).links.map(linkAddress)).toEqual(['mailto:someone@example.com', 'Hidden!A1']);
  });

  it('takes the Hyperlink look away with the link, as Excel does', async () => {
    await open();
    select(rect(1, 1, 2, 1));
    await store.setLink({ target: 'https://example.com/new' }, 'New');
    select(rect(1, 1));
    await store.removeLinks();
    expect(store.styles[store.view!.get(1, 1)!.style].linkFont).toBeUndefined();
    expect(store.view!.get(1, 1)?.s).toBe('New');
    // The rest of the range keeps its link.
    expect(store.linkAt(2, 1)?.target).toBe('https://example.com/new');
  });

  it("keeps a cell's link when its contents are cleared", async () => {
    await open();
    await store.clear([rect(0, 0)]);
    expect(store.linkAt(0, 0)?.target).toBe('https://example.com/a');
  });

  it("can't change a protected sheet's links", async () => {
    await open();
    store.setActive(3);
    select(rect(0, 0));
    await store.setLink({ target: 'https://example.com' });
    expect(store.linkAt(0, 0)).toBeUndefined();
    expect(toasts()).toEqual(['This sheet is protected in Excel, so it is read-only here.']);
  });
});

describe('HYPERLINK formulas', () => {
  it('get the address the worker works out', async () => {
    await open();
    expect(store.view!.get(4, 0)?.link).toBe('#Total');
    expect(store.linkAt(4, 0)).toEqual({ target: '#Total' });
  });
});
