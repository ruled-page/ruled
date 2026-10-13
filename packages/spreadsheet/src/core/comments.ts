import { parseRange } from '@ruled/formula/refs';
import type { CellNote, NoteEntry } from './types';
import { attr, child, decodeOoxmlText, descendants, parseXml, textOf } from './xml';

/** Someone who has written a threaded comment in the workbook. */
export interface Person {
  /** {GUID}, which their comments name. */
  id: string;
  displayName: string;
  userId?: string;
  /** Where userId comes from: "AD", "PeoplePicker"…, or "None" for a name with no account behind it. */
  providerId?: string;
}

/** The people in a workbook's threaded comments, from its persons part. */
export function parsePeople(xml: string | undefined): Person[] {
  if (!xml) return [];
  return descendants(parseXml(xml), 'person').flatMap((p) => {
    const id = attr(p, 'id');
    return id ? [{ id, displayName: attr(p, 'displayName') ?? '', userId: attr(p, 'userId'), providerId: attr(p, 'providerId') }] : [];
  });
}

/** Display names of the people in a workbook's threaded comments, by person id. */
export function parsePersons(xml: string | undefined): Map<string, string> {
  return new Map(parsePeople(xml).map((p) => [p.id, p.displayName]));
}

/** Notes (Excel's older "comments") from a comments part. */
export function parseNotes(xml: string): CellNote[] {
  const root = parseXml(xml);
  const authors = descendants(child(root, 'authors'), 'author').map((a) => a.text);
  const out: CellNote[] = [];
  for (const c of descendants(root, 'comment')) {
    const at = parseRange(attr(c, 'ref') ?? '');
    if (!at) continue;
    const author = authors[Number(attr(c, 'authorId') ?? -1)] || undefined;
    let text = decodeOoxmlText(textOf(child(c, 'text')));
    // Excel starts a note with its author's name in bold; it is shown separately.
    if (author && text.startsWith(`${author}:`)) text = text.slice(author.length + 1).replace(/^\r?\n/, '');
    out.push({ row: at.r1, col: at.c1, kind: 'note', entries: [{ author, text }] });
  }
  return out;
}

/** Threaded comments, with their replies, from a threadedComments part. */
export function parseThreads(xml: string, persons: Map<string, string>): CellNote[] {
  const byId = new Map<string, CellNote>();
  const out: CellNote[] = [];
  for (const c of descendants(parseXml(xml), 'threadedComment')) {
    const personId = attr(c, 'personId');
    const time = attr(c, 'dT');
    const entry: NoteEntry = {
      author: persons.get(personId ?? '') || undefined,
      text: decodeOoxmlText(child(c, 'text')?.text ?? ''),
      date: time?.slice(0, 10),
      id: attr(c, 'id'),
      personId,
      time,
    };
    const parent = byId.get(attr(c, 'parentId') ?? '');
    if (parent) {
      parent.entries.push(entry);
      continue;
    }
    const at = parseRange(attr(c, 'ref') ?? '');
    if (!at) continue;
    const note: CellNote = { row: at.r1, col: at.c1, kind: 'thread', entries: [entry], resolved: attr(c, 'done') === '1' };
    byId.set(attr(c, 'id') ?? '', note);
    out.push(note);
  }
  return out;
}

/**
 * A sheet's notes and threads. Excel also writes a placeholder note for every
 * thread, for older versions; the thread replaces it.
 */
export function mergeNotes(notes: CellNote[], threads: CellNote[]): CellNote[] {
  const threaded = new Set(threads.map((t) => `${t.row}:${t.col}`));
  return [...threads, ...notes.filter((n) => !threaded.has(`${n.row}:${n.col}`))];
}
