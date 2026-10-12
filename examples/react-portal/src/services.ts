/**
 * The portal's back end, simulated in the browser: a product catalogue that
 * answers after a delay like a real API, the prices the portal last
 * approved, and the price lists suppliers have submitted. Nothing leaves the
 * browser; submissions are kept in IndexedDB (in memory where that isn't
 * available, such as some private windows).
 */
import type { PriceLine } from './parse';

const wait = (ms: number, signal?: AbortSignal) =>
  new Promise<void>((resolve, reject) => {
    const timer = setTimeout(resolve, ms);
    signal?.addEventListener('abort', () => {
      clearTimeout(timer);
      reject(new DOMException('Aborted', 'AbortError'));
    });
  });

export interface Product {
  description: string;
  category: 'Hardware' | 'Software' | 'Services';
}

const CATALOGUE: Record<string, Product> = {
  'SKU-1001': { description: 'Steel bracket, small', category: 'Hardware' },
  'SKU-1002': { description: 'Steel bracket, large', category: 'Hardware' },
  'SKU-1003': { description: 'Mounting kit', category: 'Hardware' },
  'SKU-1004': { description: 'Inventory app, annual licence', category: 'Software' },
  'SKU-1005': { description: 'Inventory app, monthly licence', category: 'Software' },
  'SKU-1006': { description: 'Installation (per hour)', category: 'Services' },
  'SKU-1007': { description: 'Cable ties, pack of 100', category: 'Hardware' },
  'SKU-1008': { description: 'Training session', category: 'Services' },
  'SKU-1009': { description: 'Wall anchors, pack of 50', category: 'Hardware' },
  'SKU-1010': { description: 'Support plan, gold', category: 'Services' },
  'SKU-1011': { description: 'Reporting add-on', category: 'Software' },
  'SKU-1012': { description: 'Barcode scanner', category: 'Hardware' },
  'SKU-1013': { description: 'Label printer', category: 'Hardware' },
};

/** Looks a product up in the catalogue, as a call to the portal's API would. */
export async function lookupProduct(sku: string, signal?: AbortSignal): Promise<Product | undefined> {
  await wait(400, signal);
  return CATALOGUE[sku.trim().toUpperCase()];
}

export interface ApprovedPrice {
  price: number;
  currency: string;
  approvedAt: string;
}

/** What the portal had approved before this demo was first opened. */
const SEEDED: Record<string, ApprovedPrice> = Object.fromEntries(
  (
    [
      ['SKU-1001', 4.5],
      ['SKU-1002', 7.25],
      ['SKU-1003', 19.99],
      ['SKU-1004', 240],
      ['SKU-1005', 24],
      ['SKU-1006', 85],
      ['SKU-1007', 6.5],
      ['SKU-1008', 450],
      ['SKU-1009', 8.75],
      ['SKU-1010', 1200],
      ['SKU-1011', 60],
    ] as const
  ).map(([sku, price]) => [sku, { price, currency: 'USD', approvedAt: '2026-07-01' }]),
);

const APPROVED_KEY = 'portal-approved-prices';

/** The last approved price of every product, with what this browser has approved since. */
export function approvedPrices(): Record<string, ApprovedPrice> {
  try {
    return { ...SEEDED, ...JSON.parse(localStorage.getItem(APPROVED_KEY) ?? '{}') };
  } catch {
    return { ...SEEDED };
  }
}

/** Fetches a product's last approved price, as a call to the portal's API would. */
export async function lookupApprovedPrice(sku: string, signal?: AbortSignal): Promise<ApprovedPrice | undefined> {
  await wait(250, signal);
  return approvedPrices()[sku.trim().toUpperCase()];
}

function recordApproved(lines: PriceLine[]): void {
  const approvedAt = new Date().toISOString().slice(0, 10);
  const next = approvedPrices();
  for (const l of lines) next[l.sku] = { price: l.price, currency: l.currency, approvedAt };
  try {
    localStorage.setItem(APPROVED_KEY, JSON.stringify(next));
  } catch {
    // Without storage the approval lasts until the page closes.
  }
}

export interface Submission {
  id: string;
  fileName: string;
  submittedAt: string;
  status: 'submitted' | 'approved';
  /** The workbook as the supplier saved it, to download or edit again. */
  bytes: Uint8Array;
  /** The rows, as the portal reads them from the saved values. */
  lines: PriceLine[];
  /** Cells the supplier changed in the editor before submitting. */
  changedCells: number;
  /** Warnings the supplier submitted anyway. */
  warnings: number;
}

// ---------------------------------------------------------------------------
// Submissions, in IndexedDB.

const DB = 'portal';
const STORE = 'submissions';
const memory = new Map<string, Submission>();

function database(): Promise<IDBDatabase | undefined> {
  return new Promise((resolve) => {
    try {
      const req = indexedDB.open(DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: 'id' });
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(undefined);
    } catch {
      resolve(undefined);
    }
  });
}

async function run<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T | undefined> {
  const db = await database();
  if (!db) return undefined;
  return new Promise((resolve, reject) => {
    const req = fn(db.transaction(STORE, mode).objectStore(STORE));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function listSubmissions(): Promise<Submission[]> {
  const all = (await run('readonly', (s) => s.getAll() as IDBRequest<Submission[]>)) ?? [...memory.values()];
  return all.sort((a, b) => b.submittedAt.localeCompare(a.submittedAt));
}

export async function getSubmission(id: string): Promise<Submission | undefined> {
  return (await run('readonly', (s) => s.get(id) as IDBRequest<Submission | undefined>)) ?? memory.get(id);
}

export async function putSubmission(submission: Submission): Promise<void> {
  memory.set(submission.id, submission);
  await run('readwrite', (s) => s.put(submission));
}

export async function deleteSubmission(id: string): Promise<void> {
  memory.delete(id);
  await run('readwrite', (s) => s.delete(id));
}

/** Approves a submission: its prices become the ones the next submission is compared with. */
export async function approveSubmission(submission: Submission): Promise<Submission> {
  const approved: Submission = { ...submission, status: 'approved' };
  recordApproved(approved.lines);
  await putSubmission(approved);
  return approved;
}

export function newId(): string {
  return Math.random().toString(36).slice(2, 10);
}
