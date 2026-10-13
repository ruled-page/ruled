/**
 * Pretend services for the examples: a product catalogue and a budget
 * service, each answering after a delay, as a real API would. They stop
 * early when the editor aborts a check because the value changed again.
 */
export function wait(ms: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(resolve, ms);
    signal?.addEventListener('abort', () => {
      clearTimeout(timer);
      reject(new DOMException('Aborted', 'AbortError'));
    });
  });
}

const CATALOGUE: Record<string, { description: string; category: string; discontinued?: boolean }> = {
  'SKU-1013': { description: 'Legacy adaptor', category: 'Hardware', discontinued: true },
  'SKU-2001': { description: 'Barcode scanner', category: 'Hardware' },
  'SKU-2002': { description: 'Label printer', category: 'Hardware' },
  'SKU-2003': { description: 'Analytics add-on', category: 'Software' },
  'SKU-2004': { description: 'On-site audit', category: 'Services' },
};

/** Looks a product up in the catalogue. */
export async function lookupSku(sku: string, signal?: AbortSignal) {
  await wait(500, signal);
  return CATALOGUE[sku.toUpperCase()];
}

/** Asks the budget service whether a total needs review. */
export async function needsReview(total: number, signal?: AbortSignal): Promise<boolean> {
  await wait(800, signal);
  return total > 400_000;
}
