import { vi } from 'vitest';
import { WorkbookStore, type StoreOptions } from '../src/state/store';
import type { WorkerRequest, WorkerResponse } from '../src/worker/protocol';

/**
 * A store wired to the real worker module, run in this process: the module
 * talks to `self`, so it gets a stand-in that hands messages across.
 */
export async function testStore(options: StoreOptions = {}): Promise<WorkbookStore> {
  const client = {
    onmessage: null as ((e: MessageEvent<WorkerResponse>) => void) | null,
    // Copied as a worker's messages are, so neither side shares objects with the other.
    postMessage: (req: WorkerRequest) => {
      const data = structuredClone(req);
      setTimeout(() => scope.onmessage?.({ data } as MessageEvent<WorkerRequest>));
    },
    terminate: () => undefined,
  };
  const scope = {
    onmessage: null as ((e: MessageEvent<WorkerRequest>) => void) | null,
    postMessage: (msg: WorkerResponse) => {
      const data = structuredClone(msg);
      setTimeout(() => client.onmessage?.({ data } as MessageEvent<WorkerResponse>));
    },
  };
  (globalThis as { self?: unknown }).self = scope;
  vi.resetModules();
  await import('../src/worker/workbook.worker');
  // The worker's modules are fresh, so the store names the engine ENGINE asks for; by default the worker's own default.
  const engine = process.env.ENGINE === 'hyperformula' ? ({ engine: 'hyperformula' } as const) : {};
  return new WorkbookStore({ drafts: false, warnOnUnload: false, locale: 'en-US', ...engine, ...options, createWorker: () => client as unknown as Worker });
}

/** Resolves once the store has opened a workbook and its engine report has arrived. */
export async function opened(store: WorkbookStore, bytes: Uint8Array, name: string): Promise<void> {
  await store.open(name, bytes.slice().buffer as ArrayBuffer);
  for (let i = 0; i < 200 && !store.calc; i++) await new Promise((r) => setTimeout(r, 5));
}
