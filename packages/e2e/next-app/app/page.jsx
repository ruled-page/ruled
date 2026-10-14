import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { Workbook } from '@ruled/sheet';
import { save } from './actions';

// The workbook is read from disk on each request, as one the app stores would be.
export const dynamic = 'force-dynamic';

/** A Server Component: the workbook's bytes and a Server Action go straight to the editor. */
export default async function Page() {
  const bytes = await readFile(join(process.cwd(), 'workbooks', 'orders.xlsx')).catch(() => null);
  return (
    <main style={{ height: '100vh' }}>
      <Workbook file={bytes && { name: 'orders.xlsx', bytes: new Uint8Array(bytes) }} onSave={save} drafts={false} warnOnUnload={false} />
    </main>
  );
}
