import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { PriceLine } from '../src/embedded/parse';
import type * as Services from '../src/embedded/services';

// Each test gets a fresh module, so the in-memory submissions don't carry over.
let services: typeof Services;
beforeEach(async () => {
  vi.resetModules();
  services = await import('../src/embedded/services');
});
afterEach(() => {
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

const line = (sku: string, price: number, currency = 'USD'): PriceLine => ({
  sku,
  description: '',
  category: 'Hardware',
  price,
  currency,
  minimumOrder: 1,
  validFrom: '2026-11-01',
  orderValue: price,
});

const submission = (id: string, submittedAt: string, lines: PriceLine[] = []): Services.Submission => ({
  id,
  fileName: `${id}.xlsx`,
  submittedAt,
  status: 'submitted',
  bytes: new Uint8Array([1, 2, 3]),
  lines,
  changedCells: 0,
  warnings: 0,
});

/** localStorage as a browser has it, over a Map. */
function fakeStorage(options: { full?: boolean } = {}) {
  const items = new Map<string, string>();
  return {
    items,
    getItem: (k: string) => items.get(k) ?? null,
    setItem: (k: string, v: string) => {
      if (options.full) throw new DOMException('The quota has been exceeded.', 'QuotaExceededError');
      items.set(k, String(v));
    },
    removeItem: (k: string) => void items.delete(k),
  };
}

type Request<T> = { result?: T; error?: unknown; onsuccess?: () => void; onerror?: () => void; onupgradeneeded?: () => void };

/** Enough of IndexedDB for the portal: one object store, requests that answer on the next microtask. */
function fakeIndexedDB(options: { openFails?: boolean; writesFail?: boolean } = {}) {
  const rows = new Map<string, Services.Submission>();
  const created: { name: string; options: unknown }[] = [];
  const answer = <T>(fn: () => T, fails = false): Request<T> => {
    const req: Request<T> = {};
    queueMicrotask(() => {
      if (fails) {
        req.error = new DOMException('Disk full', 'QuotaExceededError');
        req.onerror?.();
      } else {
        req.result = fn();
        req.onsuccess?.();
      }
    });
    return req;
  };
  const store = {
    getAll: () => answer(() => [...rows.values()]),
    get: (id: string) => answer(() => rows.get(id)),
    put: (s: Services.Submission) => answer(() => (rows.set(s.id, s), s.id), options.writesFail),
    delete: (id: string) => answer(() => void rows.delete(id), options.writesFail),
  };
  const db = {
    transaction: () => ({ objectStore: () => store }),
    createObjectStore: (name: string, opts: unknown) => (created.push({ name, options: opts }), store),
  };
  return {
    rows,
    created,
    open: () => {
      const req: Request<typeof db> = {};
      queueMicrotask(() => {
        if (options.openFails) return req.onerror?.();
        req.result = db;
        if (!created.length) req.onupgradeneeded?.();
        req.onsuccess?.();
      });
      return req;
    },
  };
}

describe('without browser storage', () => {
  beforeEach(() => {
    vi.stubGlobal('localStorage', undefined);
    vi.stubGlobal('indexedDB', undefined);
  });

  it('knows the prices the portal approved before the demo was opened', () => {
    const approved = services.approvedPrices();
    expect(Object.keys(approved)).toHaveLength(11);
    expect(approved['SKU-1002']).toEqual({ price: 7.25, currency: 'USD', approvedAt: '2026-07-01' });
    // In the catalogue, never approved.
    expect(approved['SKU-1012']).toBeUndefined();
  });

  it('hands out a copy, so adding or removing products doesn’t change the next answer', () => {
    const first = services.approvedPrices();
    first['SKU-1002'] = { price: 1, currency: 'USD', approvedAt: '2026-10-07' };
    delete first['SKU-1001'];
    expect(services.approvedPrices()['SKU-1002'].price).toBe(7.25);
    expect(services.approvedPrices()['SKU-1001']).toBeDefined();
  });

  it('keeps submissions in memory, newest first', async () => {
    await services.putSubmission(submission('a', '2026-10-01T09:00:00Z'));
    await services.putSubmission(submission('b', '2026-10-03T09:00:00Z'));
    await services.putSubmission(submission('c', '2026-10-02T09:00:00Z'));
    expect((await services.listSubmissions()).map((s) => s.id)).toEqual(['b', 'c', 'a']);
    expect((await services.getSubmission('c'))?.fileName).toBe('c.xlsx');

    await services.putSubmission({ ...submission('c', '2026-10-04T09:00:00Z'), fileName: 'again.xlsx' });
    expect((await services.listSubmissions()).map((s) => s.id)).toEqual(['c', 'b', 'a']);
    expect((await services.getSubmission('c'))?.fileName).toBe('again.xlsx');

    await services.deleteSubmission('b');
    expect((await services.listSubmissions()).map((s) => s.id)).toEqual(['c', 'a']);
    expect(await services.getSubmission('b')).toBeUndefined();
  });

  it('approves a submission for the session, though its prices can’t be kept', async () => {
    const s = submission('a', '2026-10-01T09:00:00Z', [line('SKU-1002', 9.4)]);
    await services.putSubmission(s);
    const approved = await services.approveSubmission(s);
    expect(approved).toEqual({ ...s, status: 'approved' });
    expect(s.status).toBe('submitted');
    expect((await services.getSubmission('a'))?.status).toBe('approved');
    expect(services.approvedPrices()['SKU-1002'].price).toBe(7.25);
  });

  it('makes short ids that differ', () => {
    const ids = new Set(Array.from({ length: 50 }, () => services.newId()));
    expect(ids.size).toBe(50);
    for (const id of ids) expect(id).toMatch(/^[a-z0-9]{1,8}$/);
  });
});

describe('with localStorage', () => {
  beforeEach(() => vi.stubGlobal('indexedDB', undefined));

  it('records approved prices, so the next submission is compared with them', async () => {
    const storage = fakeStorage();
    vi.stubGlobal('localStorage', storage);
    vi.useFakeTimers({ toFake: ['Date'] });
    vi.setSystemTime(new Date('2026-10-07T09:00:00Z'));
    await services.approveSubmission(submission('a', '2026-10-06T09:00:00Z', [line('SKU-1002', 9.4), line('SKU-1099', 55, 'EUR')]));

    const approved = services.approvedPrices();
    expect(approved['SKU-1002']).toEqual({ price: 9.4, currency: 'USD', approvedAt: '2026-10-07' });
    expect(approved['SKU-1099']).toEqual({ price: 55, currency: 'EUR', approvedAt: '2026-10-07' });
    expect(approved['SKU-1001']).toEqual({ price: 4.5, currency: 'USD', approvedAt: '2026-07-01' });
    expect(JSON.parse(storage.items.get('portal-approved-prices')!)['SKU-1002'].price).toBe(9.4);

    // A later approval keeps what the earlier one recorded.
    await services.approveSubmission(submission('b', '2026-10-07T09:00:00Z', [line('SKU-1001', 5)]));
    expect(services.approvedPrices()['SKU-1099'].price).toBe(55);
    expect(services.approvedPrices()['SKU-1001'].price).toBe(5);
  });

  it('falls back to the seeded prices when what is stored can’t be read', () => {
    const storage = fakeStorage();
    storage.items.set('portal-approved-prices', '{not json');
    vi.stubGlobal('localStorage', storage);
    expect(services.approvedPrices()['SKU-1002'].price).toBe(7.25);
  });

  it('still approves when storage is full', async () => {
    vi.stubGlobal('localStorage', fakeStorage({ full: true }));
    const s = submission('a', '2026-10-06T09:00:00Z', [line('SKU-1002', 9.4)]);
    await expect(services.approveSubmission(s)).resolves.toMatchObject({ status: 'approved' });
    expect(services.approvedPrices()['SKU-1002'].price).toBe(7.25);
  });
});

describe('the portal’s APIs', () => {
  beforeEach(() => {
    vi.stubGlobal('localStorage', undefined);
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
  });

  it('look a product up after 400 ms, ignoring case and spaces', async () => {
    let answer: Services.Product | undefined | 'pending' = 'pending';
    void services.lookupProduct(' sku-1004 ').then((p) => (answer = p));
    await vi.advanceTimersByTimeAsync(399);
    expect(answer).toBe('pending');
    await vi.advanceTimersByTimeAsync(1);
    expect(answer).toEqual({ description: 'Inventory app, annual licence', category: 'Software' });

    const unknown = services.lookupProduct('SKU-9999');
    await vi.advanceTimersByTimeAsync(400);
    expect(await unknown).toBeUndefined();
  });

  it('fetch an approved price after 250 ms', async () => {
    const found = services.lookupApprovedPrice('sku-1006');
    const missing = services.lookupApprovedPrice('SKU-1012');
    await vi.advanceTimersByTimeAsync(250);
    expect(await found).toEqual({ price: 85, currency: 'USD', approvedAt: '2026-07-01' });
    expect(await missing).toBeUndefined();
  });

  it('stop waiting when the caller aborts', async () => {
    const controller = new AbortController();
    const product = services.lookupProduct('SKU-1001', controller.signal);
    const price = services.lookupApprovedPrice('SKU-1001', controller.signal);
    expect(vi.getTimerCount()).toBe(2);
    controller.abort();
    await expect(product).rejects.toMatchObject({ name: 'AbortError' });
    await expect(price).rejects.toMatchObject({ name: 'AbortError' });
    expect(vi.getTimerCount()).toBe(0);
  });
});

describe('with IndexedDB', () => {
  beforeEach(() => vi.stubGlobal('localStorage', undefined));

  it('creates the store the first time, and keeps submissions there', async () => {
    const idb = fakeIndexedDB();
    vi.stubGlobal('indexedDB', idb);
    await services.putSubmission(submission('a', '2026-10-01T09:00:00Z'));
    await services.putSubmission(submission('b', '2026-10-02T09:00:00Z'));
    expect(idb.created).toEqual([{ name: 'submissions', options: { keyPath: 'id' } }]);
    expect([...idb.rows.keys()]).toEqual(['a', 'b']);

    // A page opened later (a fresh module, nothing in memory) finds them.
    vi.resetModules();
    const later = await import('../src/embedded/services');
    expect((await later.listSubmissions()).map((s) => s.id)).toEqual(['b', 'a']);
    expect((await later.getSubmission('a'))?.fileName).toBe('a.xlsx');
    await later.deleteSubmission('a');
    expect([...idb.rows.keys()]).toEqual(['b']);
    expect(await later.getSubmission('a')).toBeUndefined();
  });

  it('keeps submissions in memory when the database won’t open', async () => {
    vi.stubGlobal('indexedDB', fakeIndexedDB({ openFails: true }));
    await services.putSubmission(submission('a', '2026-10-01T09:00:00Z'));
    expect((await services.listSubmissions()).map((s) => s.id)).toEqual(['a']);
    expect((await services.getSubmission('a'))?.id).toBe('a');
  });

  it('reports a write the database refuses', async () => {
    const idb = fakeIndexedDB({ writesFail: true });
    vi.stubGlobal('indexedDB', idb);
    await expect(services.putSubmission(submission('a', '2026-10-01T09:00:00Z'))).rejects.toMatchObject({ name: 'QuotaExceededError' });
    expect(idb.rows.size).toBe(0);
    // The page still has it from memory.
    expect((await services.getSubmission('a'))?.id).toBe('a');
  });
});
