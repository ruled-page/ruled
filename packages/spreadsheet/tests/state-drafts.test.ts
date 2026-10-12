import { afterEach, describe, expect, it } from 'vitest';
import { deleteDraft, fileKey, loadDraft, saveDraft, type Draft } from '../src/state/drafts';
import { NARROW, isNarrow } from '../src/state/viewport';

/**
 * Just enough of IndexedDB for drafts.ts: one database, object stores keyed
 * by a key path, and requests that complete on a later tick as real ones do.
 */
function fakeIndexedDB(fail: { open?: boolean; requests?: boolean } = {}) {
  type Req = { result?: unknown; error?: Error; onsuccess?: () => void; onerror?: () => void; onupgradeneeded?: () => void };
  const stores = new Map<string, { keyPath: string; rows: Map<string, unknown> }>();
  const log = { opened: 0, closed: 0, upgrades: 0 };
  let version = 0;
  const later = (req: Req, work: () => unknown) => {
    setTimeout(() => {
      if (fail.requests) {
        req.error = new Error('QuotaExceededError');
        req.onerror?.();
      } else {
        req.result = work();
        req.onsuccess?.();
      }
    });
    return req;
  };
  const db = {
    createObjectStore(name: string, options: { keyPath: string }) {
      stores.set(name, { keyPath: options.keyPath, rows: new Map() });
    },
    transaction(name: string, mode: 'readonly' | 'readwrite') {
      const store = stores.get(name);
      if (!store) throw new Error(`NotFoundError: no store ${name}`);
      const writable = () => {
        if (mode !== 'readwrite') throw new Error('ReadOnlyError');
      };
      return {
        objectStore: () => ({
          get: (key: string) => later({}, () => structuredClone(store.rows.get(key))),
          put: (value: Record<string, unknown>) => {
            writable();
            return later({}, () => store.rows.set(value[store.keyPath] as string, structuredClone(value)) && value[store.keyPath]);
          },
          delete: (key: string) => {
            writable();
            return later({}, () => void store.rows.delete(key));
          },
        }),
      };
    },
    close: () => void log.closed++,
  };
  const factory = {
    open(_name: string, wanted: number) {
      log.opened++;
      const req: Req = {};
      setTimeout(() => {
        if (fail.open) {
          req.error = new Error('InvalidStateError');
          return req.onerror?.();
        }
        req.result = db;
        if (version < wanted) {
          version = wanted;
          log.upgrades++;
          req.onupgradeneeded?.();
        }
        req.onsuccess?.();
      });
      return req;
    },
  };
  (globalThis as { indexedDB?: unknown }).indexedDB = factory;
  return { log, stores };
}

afterEach(() => {
  delete (globalThis as { indexedDB?: unknown }).indexedDB;
});

const draft = (key: string, editedCount = 1): Draft => ({
  key,
  fileName: 'orders.xlsx',
  updatedAt: 1_760_000_000_000,
  editedCount,
  ops: [{ type: 'input', sheet: 0, row: 1, col: 2, text: '42' }],
});

describe('keys for drafts', () => {
  const bytes = (text: string) => new TextEncoder().encode(text).buffer as ArrayBuffer;

  it("are the SHA-256 of the file's contents, in hex", async () => {
    expect(await fileKey(bytes('abc'))).toBe('ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
    expect(await fileKey(new ArrayBuffer(0))).toBe('e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855');
  });

  it('match for the same contents, whatever the file is called, and differ otherwise', async () => {
    expect(await fileKey(bytes('same bytes'))).toBe(await fileKey(bytes('same bytes')));
    expect(await fileKey(bytes('same bytes'))).not.toBe(await fileKey(bytes('same bytes!')));
  });
});

describe('keeping drafts in the browser', () => {
  it('saves a draft, finds it again by its key, and forgets it', async () => {
    fakeIndexedDB();
    await saveDraft(draft('k1'));
    await saveDraft(draft('k2', 5));
    expect(await loadDraft('k1')).toEqual(draft('k1'));
    expect((await loadDraft('k2'))?.editedCount).toBe(5);
    await deleteDraft('k1');
    expect(await loadDraft('k1')).toBeUndefined();
    expect(await loadDraft('k2')).toBeDefined();
  });

  it('replaces the draft for a file when saved again', async () => {
    const { stores } = fakeIndexedDB();
    await saveDraft(draft('k1', 1));
    await saveDraft(draft('k1', 2));
    expect((await loadDraft('k1'))?.editedCount).toBe(2);
    expect(stores.get('drafts')?.rows.size).toBe(1);
  });

  it('has nothing for a file without a draft', async () => {
    fakeIndexedDB();
    expect(await loadDraft('never saved')).toBeUndefined();
  });

  it('creates its store once and closes the database after every operation', async () => {
    const { log } = fakeIndexedDB();
    await saveDraft(draft('k1'));
    await loadDraft('k1');
    await deleteDraft('k1');
    expect(log).toEqual({ opened: 3, closed: 3, upgrades: 1 });
  });

  it('carries on without drafts where there is no IndexedDB, as in Node', async () => {
    expect(typeof indexedDB).toBe('undefined');
    await expect(loadDraft('k1')).resolves.toBeUndefined();
    await expect(saveDraft(draft('k1'))).resolves.toBeUndefined();
    await expect(deleteDraft('k1')).resolves.toBeUndefined();
  });

  it('carries on when the browser refuses to open storage, as private windows do', async () => {
    fakeIndexedDB({ open: true });
    await expect(saveDraft(draft('k1'))).resolves.toBeUndefined();
    await expect(loadDraft('k1')).resolves.toBeUndefined();
    await expect(deleteDraft('k1')).resolves.toBeUndefined();
  });

  it('carries on when a read or write fails, and still closes the database', async () => {
    const { log } = fakeIndexedDB({ requests: true });
    await expect(saveDraft(draft('k1'))).resolves.toBeUndefined();
    await expect(loadDraft('k1')).resolves.toBeUndefined();
    await expect(deleteDraft('k1')).resolves.toBeUndefined();
    expect(log.closed).toBe(3);
  });
});

describe('phone-sized screens', () => {
  const original = globalThis.matchMedia;
  afterEach(() => {
    globalThis.matchMedia = original;
  });

  it('are never assumed where media queries are unavailable', () => {
    delete (globalThis as { matchMedia?: unknown }).matchMedia;
    expect(isNarrow()).toBe(false);
  });

  it("follow the stylesheet's narrow-screen query", () => {
    const asked: string[] = [];
    const screen = (width: number) =>
      ((query: string) => {
        asked.push(query);
        return { matches: width <= 720 } as MediaQueryList;
      }) as typeof matchMedia;
    globalThis.matchMedia = screen(390);
    expect(isNarrow()).toBe(true);
    globalThis.matchMedia = screen(1280);
    expect(isNarrow()).toBe(false);
    expect(asked).toEqual([NARROW, NARROW]);
    expect(NARROW).toBe('(max-width: 720px)');
  });
});
