/**
 * Packs the library as npm would publish it, installs the tarball into the
 * React app in packages/e2e/host-app, and builds that app with Vite, as a host
 * project would. That app has its own React 18, while this repository
 * develops on React 19, so both versions the library supports are tested. package.spec.ts then uses it from a browser, which
 * catches a library build that only works from this repository.
 *
 * Needs `npm run build:lib` first.
 */
import { execFileSync } from 'node:child_process';
import { join } from 'node:path';
import { installPackedLibrary, root } from './packed.mjs';

const app = join(root, 'packages', 'e2e', 'host-app');
installPackedLibrary(app);
execFileSync('npx', ['vite', 'build', '-c', join(app, 'vite.config.ts')], { cwd: root, stdio: 'inherit' });
console.log('Built packages/e2e/host-app with the packed library.');
