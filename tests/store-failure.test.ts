import ExcelJS from 'exceljs';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { english } from '../src/i18n';
import type { StoreOptions, WorkbookStore } from '../src/state/store';
import { opened, testStore } from './storeHarness';

vi.setConfig({ testTimeout: 10_000 });

/**
 * The worker holds the workbook. When it fails to load, crashes, runs out of
 * memory or stops answering, the editor must say so and offer the file again,
 * rather than wait for ever.
 */

let bytes: Uint8Array;
beforeAll(async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('Lines').addRow([2, 3, { formula: 'A1*B1', result: 6 }]);
  bytes = new Uint8Array(await wb.xlsx.writeBuffer());
});

let store: WorkbookStore;
afterEach(() => {
  store?.dispose();
  vi.restoreAllMocks();
});

type FakeWorker = { onerror: (e: Partial<ErrorEvent>) => void; postMessage: (m: unknown) => void };
const workerOf = (s: WorkbookStore) => (s as unknown as { worker?: FakeWorker }).worker;
const until = async (ok: () => boolean) => {
  for (let i = 0; i < 400 && !ok(); i++) await new Promise((r) => setTimeout(r, 5));
};

async function open(options: StoreOptions = {}): Promise<Error[]> {
  const errors: Error[] = [];
  store = await testStore({ ...options, onError: (e) => errors.push(e) });
  await opened(store, bytes, 'lines.xlsx');
  return errors;
}

describe('when the worker fails', () => {
  it('says so, tells the host, fails what was waiting, and reopens the file', async () => {
    const errors = await open();
    const worker = workerOf(store)!;
    const send = worker.postMessage;
    // An edit is on its way when the worker dies.
    worker.postMessage = () => undefined;
    const editing = store.input(0, 0, '5');
    worker.onerror({ preventDefault: () => undefined });
    await editing;
    expect(store.status).toBe('error');
    const said = "The editor stopped working: its background worker crashed, ran out of memory or couldn't load.";
    expect(english.text(store.error)).toBe(said);
    expect(errors.map((e) => e.message)).toEqual([said]);
    expect(store.busy).toBe(false);
    expect(store.canReopen).toBe(true);
    // Nothing more goes to the dead worker.
    await expect(store.snapshot()).resolves.toBeUndefined();

    // The test's stand-in worker is reused for the new one.
    worker.postMessage = send;
    await store.reopen();
    await until(() => !!store.calc);
    expect(store.status).toBe('ready');
    expect(store.fileName).toBe('lines.xlsx');
    expect(store.canReopen).toBe(false);
    await store.input(0, 0, '5');
    expect(store.sheets[0].get(0, 2)?.n).toBe(15);
  });

  it('gives up on a worker that stops answering', async () => {
    const errors = await open({ workerTimeout: 300 });
    workerOf(store)!.postMessage = () => undefined;
    await store.input(0, 0, '5');
    expect(store.status).toBe('error');
    expect(english.text(store.error)).toBe('The editor stopped working: it spent more than 1 second on one task.');
    expect(errors).toHaveLength(1);
  });

  it('reopens without live recalculation when building the engine is what failed', async () => {
    const errors: Error[] = [];
    store = await testStore({ recalc: 'on', onError: (e) => errors.push(e) });
    // The worker module was just loaded afresh; this is its copy of the session.
    const { WorkbookSession } = await import('../src/core/session');
    vi.spyOn(WorkbookSession.prototype, 'ensureEngine').mockImplementationOnce(() => {
      throw new RangeError('Maximum call stack size exceeded');
    });
    await store.open('lines.xlsx', bytes.slice().buffer as ArrayBuffer);
    await until(() => store.status === 'error');
    expect(english.text(store.error)).toBe(
      "This workbook couldn't be calculated (Maximum call stack size exceeded). Reopen it to work without live recalculation.",
    );
    await store.reopen();
    await until(() => !!store.calc);
    expect(store.status).toBe('ready');
    expect(store.calc?.off).toBe(true);
    expect(errors).toHaveLength(1);
  });

  it('leaves the buffer the host passed in usable', async () => {
    store = await testStore();
    const buffer = bytes.slice().buffer as ArrayBuffer;
    await store.open('lines.xlsx', buffer);
    expect(buffer.byteLength).toBe(bytes.length);
  });
});
