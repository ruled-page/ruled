/**
 * Packs the library as npm would publish it, installs the tarball into the
 * Next.js app in packages/e2e/next-app, and builds that app with its App
 * Router, as a host project would. next.spec.ts then uses it from a browser:
 * a Server Component renders the editor with a workbook read on the server,
 * and a Server Action receives what is saved.
 *
 * Needs `npm run build:lib` first.
 */
import { execFileSync } from 'node:child_process';
import { join } from 'node:path';
import { installPackedLibrary, root } from './packed.mjs';

const app = join(root, 'packages', 'e2e', 'next-app');
installPackedLibrary(app);
execFileSync('npx', ['next', 'build'], { cwd: app, stdio: 'inherit' });
console.log('Built packages/e2e/next-app with the packed library.');
