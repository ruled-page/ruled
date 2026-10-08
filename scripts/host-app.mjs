/**
 * Packs the library as npm would publish it, installs the tarball into the
 * React app in e2e/host-app, and builds that app with Vite, as a host
 * project would. That app has its own React 18, while this repository
 * develops on React 19, so both versions the library supports are tested. e2e/package.spec.ts then uses it from a browser, which
 * catches a library build that only works from this repository.
 *
 * Needs `npm run build:lib` first.
 */
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const app = join(root, 'e2e', 'host-app');
if (!existsSync(join(root, 'dist-lib', 'react', 'workbook-editor.js'))) {
  console.error('Build the library first: npm run build:lib');
  process.exit(1);
}

// The app's own dependencies first: installing them would remove the library unpacked below.
execFileSync('npm', ['install', '--no-audit', '--no-fund'], { cwd: app, stdio: 'inherit' });

const packed = mkdtempSync(join(tmpdir(), 'workbook-editor-pack-'));
const [{ filename }] = JSON.parse(execFileSync('npm', ['pack', '--json', '--pack-destination', packed], { cwd: root, encoding: 'utf8' }));
const installed = join(app, 'node_modules', 'fake-spreadsheet');
rmSync(installed, { recursive: true, force: true });
mkdirSync(installed, { recursive: true });
execFileSync('tar', ['xzf', join(packed, filename), '-C', installed, '--strip-components=1']);
rmSync(packed, { recursive: true, force: true });

execFileSync('npx', ['vite', 'build', '-c', join(app, 'vite.config.ts')], { cwd: root, stdio: 'inherit' });
console.log('Built e2e/host-app with the packed library.');
