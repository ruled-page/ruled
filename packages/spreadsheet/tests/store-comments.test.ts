import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import type { Draft } from '../src/state/drafts';
import type { StoreOptions, WorkbookStore } from '../src/state/store';
import type { WorkbookSchema } from '../src/schema/types';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

// Node has no IndexedDB, so drafts are kept in a map here instead.
const kept = vi.hoisted(() => new Map<string, Draft>());
vi.mock('../src/state/drafts', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/state/drafts')>()),
  loadDraft: vi.fn(async (key: string) => kept.get(key)),
  saveDraft: vi.fn(async (draft: Draft) => void kept.set(draft.key, structuredClone(draft))),
  deleteDraft: vi.fn(async (key: string) => void kept.delete(key)),
}));

/** Review (a note on B2, a thread on C3, a resolved one on A6), Plain (none), Locked (protected, with a note on A1), Objects (protected, notes allowed). */
const bytes = new Uint8Array(readFileSync(join(__dirname, 'corpus', 'comments.xlsx')));
const REVIEW = 0;
const PLAIN = 1;
const LOCKED = 2;
const OBJECTS = 3;

let store: WorkbookStore;
beforeEach(() => kept.clear());
afterEach(() => store?.dispose());

async function open(options: StoreOptions = {}): Promise<WorkbookStore> {
  store = await testStore({ user: 'Cara', ...options });
  await opened(store, bytes, 'comments.xlsx');
  return store;
}

const note = (sheet: number, row: number, col: number) => store.sheets[sheet].noteAt(row, col);
const toasts = () => store.toasts.map((t) => english.text(t.text));
const actions = () => store.history.map((h) => [english.text(h.action), h.cells, h.refs.join(' ')]);

describe('notes', () => {
  it('are added, edited and deleted with the user as author, a step of undo each', async () => {
    await open();
    store.setActive(PLAIN);
    expect(await store.setNote(1, 1, 'Counted twice.')).toBe(true);
    expect(note(PLAIN, 1, 1)).toEqual({ row: 1, col: 1, kind: 'note', entries: [{ author: 'Cara', text: 'Counted twice.' }] });
    expect(store.editedCount).toBe(1);
    expect(store.unsaved).toBe(true);
    await store.setNote(1, 1, 'Counted once.');
    await store.deleteComment(1, 1);
    expect(note(PLAIN, 1, 1)).toBeUndefined();
    expect(actions()).toEqual([
      ['Add note', 1, 'Plain!B2'],
      ['Edit note', 1, 'Plain!B2'],
      ['Delete note', 1, 'Plain!B2'],
    ]);
    expect(store.editedCount).toBe(0);

    // Undo goes back to the sheet the note is on.
    store.setActive(REVIEW);
    await store.undo();
    expect(store.active).toBe(PLAIN);
    expect(note(PLAIN, 1, 1)?.entries[0].text).toBe('Counted once.');
    await store.undo();
    await store.undo();
    expect(note(PLAIN, 1, 1)).toBeUndefined();
    expect(store.editedCount).toBe(0);
    store.setActive(REVIEW);
    await store.redo();
    expect(store.active).toBe(PLAIN);
    expect(note(PLAIN, 1, 1)?.entries[0].text).toBe('Counted twice.');
    expect(store.editedCount).toBe(1);
  });

  it('keep the author they had when edited, and have none without a user', async () => {
    await open({ user: undefined });
    await store.setNote(1, 1, 'Changed.');
    expect(note(REVIEW, 1, 1)?.entries).toEqual([{ author: 'Ann Analyst', text: 'Changed.' }]);
    store.setActive(PLAIN);
    await store.setNote(0, 0, 'Mine.');
    expect(note(PLAIN, 0, 0)?.entries).toEqual([{ author: undefined, text: 'Mine.' }]);
  });

  it('are listed in the changes, and in the review before saving', async () => {
    await open({ numbers: { reviewBeforeSave: true }, onSave: () => undefined });
    await store.setNote(1, 1, 'Changed.');
    store.setActive(PLAIN);
    await store.setNote(0, 0, 'New.');
    await store.refreshChanges();
    expect(store.changeList.map((c) => [c.ref, c.before, c.after, english.text(c.comment)])).toEqual([
      ['B2', 'Agreed with finance.\nRevisit in January.', 'Changed.', 'Note edited'],
      ['A1', '', 'New.', 'Note added'],
    ]);
    const saving = store.save();
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    expect(store.confirmRequest?.details?.rows).toEqual([
      ['Review!B2', 'Agreed with finance.\nRevisit in January.', 'Note edited: Changed.'],
      ['Plain!A1', '', 'Note added: New.'],
    ]);
    store.confirmRequest?.resolve('cancel');
    expect(await saving).toBe('cancelled');
  });
});

describe('comment threads', () => {
  it('are replied to, edited, resolved and reopened, and their replies deleted', async () => {
    await open();
    expect(await store.reply(2, 2, 'Thanks.')).toBe(true);
    const mine = note(REVIEW, 2, 2)!.entries[2];
    expect(mine).toMatchObject({ author: 'Cara', text: 'Thanks.' });
    await store.editComment(2, 2, mine.id!, 'Thanks, Bob.');
    await store.resolveThread(2, 2, true);
    expect(note(REVIEW, 2, 2)).toMatchObject({ resolved: true, entries: [{}, {}, { text: 'Thanks, Bob.' }] });
    // Resolved threads take no replies until they are reopened.
    expect(await store.reply(2, 2, 'More.')).toBe(false);
    expect(toasts()).toContain('This thread is resolved. Reopen it to reply.');
    await store.resolveThread(2, 2, false);
    await store.deleteReply(2, 2, mine.id!);
    expect(note(REVIEW, 2, 2)?.entries).toHaveLength(2);
    expect(actions().map(([a]) => a)).toEqual(['Reply to comment', 'Edit comment', 'Resolve comment thread', 'Reopen comment thread', 'Delete reply']);
  });

  it('are started and deleted, by "Anonymous" without a user', async () => {
    await open({ user: undefined });
    store.setActive(PLAIN);
    expect(store.commentAuthor).toBe('Anonymous');
    await store.startThread(0, 1, 'Units?');
    expect(note(PLAIN, 0, 1)).toMatchObject({ kind: 'thread', resolved: false, entries: [{ author: 'Anonymous', text: 'Units?' }] });
    await store.deleteReply(0, 1, note(PLAIN, 0, 1)!.entries[0].id!);
    expect(note(PLAIN, 0, 1)).toBeUndefined();
    store.setActive(REVIEW);
    await store.deleteComment(5, 0);
    expect(actions().map(([a]) => a)).toEqual(['New comment', 'Delete comment thread', 'Delete comment thread']);
  });

  it('come back from a draft as they were made, with the same ids', async () => {
    await open({ drafts: true });
    await store.reply(2, 2, 'Thanks.');
    const made = note(REVIEW, 2, 2);
    await new Promise((r) => setTimeout(r, 450));
    expect([...kept.values()][0]?.ops).toEqual([{ type: 'restore', cells: [], notes: [{ sheet: REVIEW, row: 2, col: 2, note: made }] }]);
    store.dispose();

    store = await testStore({ drafts: true, user: 'Cara' });
    const opening = store.open('comments.xlsx', bytes.slice().buffer as ArrayBuffer);
    for (let i = 0; i < 200 && !store.confirmRequest; i++) await new Promise((r) => setTimeout(r, 5));
    store.confirmRequest?.resolve('confirm');
    await opening;
    for (let i = 0; i < 200 && !note(REVIEW, 2, 2)?.entries[2]; i++) await new Promise((r) => setTimeout(r, 5));
    expect(note(REVIEW, 2, 2)).toEqual(made);
  });
});

describe('where notes and comments are read-only', () => {
  it('on a protected sheet whose protection covers objects, they can be read but not changed', async () => {
    await open();
    store.setActive(LOCKED);
    expect(await store.setNote(0, 0, 'Mine now')).toBe(false);
    expect(toasts()).toEqual(['This sheet is protected in Excel, which keeps its notes and comments as they are.']);
    store.openComment(0, 0, { x: 10, y: 10 });
    expect(store.commentFor).toMatchObject({ sheet: LOCKED, row: 0, col: 0 });
    store.set('commentFor', undefined);
    // A cell with nothing to read says why instead.
    store.openComment(4, 1, { x: 10, y: 10 });
    expect(store.commentFor).toBeUndefined();
    expect(toasts()).toHaveLength(2);

    store.setActive(OBJECTS);
    expect(await store.setNote(0, 0, 'Allowed')).toBe(true);
    expect(store.history).toHaveLength(1);
  });

  it("on a schema's read-only cells", async () => {
    const schema: WorkbookSchema = {
      id: 'review',
      name: 'Review',
      otherSheets: 'readOnly',
      sheets: [{ name: 'Review', tables: [{ key: 'lines', at: 'A1', fields: [{ key: 'line', label: 'Line', type: 'text', readOnly: true }, { key: 'budget', label: 'Budget', type: 'number' }, { key: 'actual', label: 'Actual', type: 'number' }] }] }],
    };
    await open({ schemas: [schema] });
    expect(store.schema).toBeDefined();
    expect(await store.setNote(1, 0, 'No')).toBe(false);
    expect(await store.setNote(1, 1, 'Yes')).toBe(true);
    store.setActive(PLAIN);
    expect(await store.startThread(0, 0, 'No')).toBe(false);
    expect(toasts()).toEqual(['“Line” is read-only.', 'Only the sheets described by the “Review” schema can be edited.']);
  });

  it('nowhere once the workbook is closed', async () => {
    store = await testStore();
    expect(await store.setNote(0, 0, 'x')).toBe(false);
    store.openComment(0, 0, { x: 0, y: 0 });
    expect(store.commentFor).toBeUndefined();
  });
});

describe('the comment popup', () => {
  it('closes when the workbook is rebuilt, as its cell may have moved', async () => {
    await open();
    store.openComment(2, 2, { x: 10, y: 10 }, 'thread');
    expect(store.commentFor).toMatchObject({ start: 'thread' });
    store.setSelection({ row: 0, col: 0, rects: [{ r1: 0, c1: 0, r2: 0, c2: 0 }] });
    await store.insertRows('above');
    expect(store.commentFor).toBeUndefined();
    expect(note(REVIEW, 3, 2)?.kind).toBe('thread');
  });
});
