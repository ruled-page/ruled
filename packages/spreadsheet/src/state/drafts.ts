import type { WorkbookOp } from '../core/sheetRules';
import type { EditOp } from '../core/types';

/** One step of a draft: a cell edit, or a structural change (to where cells are, or to the sheets). */
export type DraftOp = EditOp | { type: 'structure'; op: WorkbookOp };

/**
 * Unsaved edits kept in IndexedDB so they survive a closed tab or a crash.
 * A draft is the list of edits applied since the file was opened, keyed by
 * a hash of the file, so replaying it on the same file recreates the work.
 */
export interface Draft {
  key: string;
  fileName: string;
  updatedAt: number;
  editedCount: number;
  ops: DraftOp[];
}

const DB_NAME = 'workbook-editor-drafts';
const STORE = 'drafts';

function openDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') return reject(new Error('IndexedDB is not available.'));
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'key' });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function run<T>(mode: IDBTransactionMode, fn: (store: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await openDb();
  try {
    return await new Promise<T>((resolve, reject) => {
      const req = fn(db.transaction(STORE, mode).objectStore(STORE));
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  } finally {
    db.close();
  }
}

/** A stable key for a file's contents. */
export async function fileKey(bytes: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, '0')).join('');
}

/** Drafts are a convenience: storage being unavailable (private windows) is not an error. */
export async function loadDraft(key: string): Promise<Draft | undefined> {
  try {
    return (await run<Draft | undefined>('readonly', (s) => s.get(key) as IDBRequest<Draft | undefined>)) ?? undefined;
  } catch {
    return undefined;
  }
}

export async function saveDraft(draft: Draft): Promise<void> {
  try {
    await run('readwrite', (s) => s.put(draft));
  } catch {
    // See loadDraft.
  }
}

export async function deleteDraft(key: string): Promise<void> {
  try {
    await run('readwrite', (s) => s.delete(key));
  } catch {
    // See loadDraft.
  }
}
