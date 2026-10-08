import { msg, TextError, type Message, type Text } from '../i18n/message';
import type { CellNote, CommentChange, NoteEntry } from './types';
import { joinTexts } from './workbook';

/** A new id for a threaded comment or a person, as Excel writes them: {8-4-4-4-12 hex digits}, in capitals. */
export function newGuid(): string {
  const b = new Uint8Array(16);
  crypto.getRandomValues(b);
  // A random (version 4) GUID.
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const h = Array.from(b, (x) => x.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
}

/** A moment as Excel dates a threaded comment: UTC, to hundredths of a second, with no zone written. */
export function commentTime(at = new Date()): string {
  return at.toISOString().slice(0, 22);
}

export interface ChangeContext {
  /** The id in the workbook's list of people of whoever is writing. */
  person: (author: string) => string;
  /** A new comment's id. */
  newId?: () => string;
  /** When a new comment is written. */
  now?: () => string;
}

/**
 * A cell's note or thread after a change, or null when the change leaves it
 * with neither. Throws, with the reason, a change that doesn't apply to what
 * the cell has, such as replying where there is no thread.
 */
export function applyCommentChange(note: CellNote | null, row: number, col: number, change: CommentChange, ctx: ChangeContext): CellNote | null {
  const entry = (author: string, text: string): NoteEntry => {
    const time = (ctx.now ?? commentTime)();
    return { author, text, date: time.slice(0, 10), id: (ctx.newId ?? newGuid)(), personId: ctx.person(author), time };
  };
  const thread = (): CellNote => {
    if (note?.kind !== 'thread') throw new TextError(msg('core.comment.noThread'));
    return note;
  };
  const nonEmpty = (text: string) => {
    if (!text.trim()) throw new TextError(msg('core.comment.empty'));
    return text;
  };
  switch (change.action) {
    case 'setNote':
      if (note?.kind === 'thread') throw new TextError(msg('core.comment.isThread'));
      // Editing a note keeps whoever wrote it, as Excel does.
      return { row, col, kind: 'note', entries: [{ author: note ? note.entries[0]?.author : change.author || undefined, text: change.text }] };
    case 'startThread':
      if (note) throw new TextError(msg('core.comment.exists'));
      return { row, col, kind: 'thread', entries: [entry(change.author, nonEmpty(change.text))], resolved: false };
    case 'reply': {
      const t = thread();
      if (t.resolved) throw new TextError(msg('core.comment.resolved'));
      return { ...t, entries: [...t.entries, entry(change.author, nonEmpty(change.text))] };
    }
    case 'editEntry': {
      const t = thread();
      if (!t.entries.some((e) => e.id === change.id)) throw new TextError(msg('core.comment.gone'));
      const text = nonEmpty(change.text);
      return { ...t, entries: t.entries.map((e) => (e.id === change.id ? { ...e, text } : e)) };
    }
    case 'deleteEntry': {
      const t = thread();
      const at = t.entries.findIndex((e) => e.id === change.id);
      if (at < 0) throw new TextError(msg('core.comment.gone'));
      // The first comment is the thread: without it the replies have nothing to answer.
      return at === 0 ? null : { ...t, entries: t.entries.filter((_, i) => i !== at) };
    }
    case 'resolve':
      return { ...thread(), resolved: change.resolved };
    case 'delete':
      return null;
  }
}

/** Whether two states of a cell's note or thread are the same. */
export function sameNote(a: CellNote | null | undefined, b: CellNote | null | undefined): boolean {
  if (!a || !b) return !a && !b;
  if (a.kind !== b.kind || !!a.resolved !== !!b.resolved || a.entries.length !== b.entries.length) return false;
  return a.entries.every((e, i) => {
    const f = b.entries[i];
    return e.text === f.text && e.author === f.author && e.id === f.id && e.personId === f.personId && e.time === f.time;
  });
}

/** A note or thread as text, for the change list: a note's text, or each comment as "Author: text" on its own line. */
export function noteText(note: CellNote | null | undefined): string {
  if (!note) return '';
  if (note.kind === 'note') return note.entries[0]?.text ?? '';
  return note.entries.map((e) => (e.author ? `${e.author}: ${e.text}` : e.text)).join('\n');
}

/** What happened to a cell's note or thread, in words, for the change list. */
export function describeCommentChange(before: CellNote | null | undefined, after: CellNote | null | undefined): Text {
  if (!before) return msg(after?.kind === 'thread' ? 'core.change.threadStarted' : 'core.change.noteAdded');
  if (!after) return msg(before.kind === 'thread' ? 'core.change.threadDeleted' : 'core.change.noteDeleted');
  if (after.kind === 'note' || before.kind === 'note') return msg(after.kind === 'note' ? 'core.change.noteEdited' : 'core.change.threadEdited');
  const now = new Map(after.entries.map((e) => [e.id, e]));
  const then = new Set(before.entries.map((e) => e.id));
  const parts: Message[] = [];
  const added = after.entries.filter((e) => !then.has(e.id)).length;
  const edited = before.entries.filter((e) => now.has(e.id) && now.get(e.id)!.text !== e.text).length;
  const deleted = before.entries.filter((e) => !now.has(e.id)).length;
  if (added) parts.push(msg('core.change.repliesAdded', { count: added }));
  if (edited) parts.push(msg('core.change.commentsEdited', { count: edited }));
  if (deleted) parts.push(msg('core.change.repliesDeleted', { count: deleted }));
  if (!!before.resolved !== !!after.resolved) parts.push(msg(after.resolved ? 'core.change.resolved' : 'core.change.reopened'));
  return parts.length ? msg('core.change.thread', { what: joinTexts(parts) }) : msg('core.change.threadEdited');
}
