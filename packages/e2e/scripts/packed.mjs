/**
 * Packs the library as npm would publish it and unpacks it into an app's
 * node_modules, as installing the package would. Its dependencies are all
 * bundled, so nothing else is installed with it.
 */
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

export const root = fileURLToPath(new URL('../../..', import.meta.url));
const library = join(root, 'packages', 'spreadsheet');

/** Installs the packed library into `app`, after the app's own dependencies, as installing those would remove it. */
export function installPackedLibrary(app) {
  if (!existsSync(join(library, 'dist', 'react', 'workbook-editor.js'))) {
    console.error('Build the library first: npm run build:lib');
    process.exit(1);
  }
  execFileSync('npm', ['install', '--no-audit', '--no-fund'], { cwd: app, stdio: 'inherit' });
  const packed = mkdtempSync(join(tmpdir(), 'workbook-editor-pack-'));
  const [{ filename }] = JSON.parse(execFileSync('npm', ['pack', '--json', '--pack-destination', packed], { cwd: library, encoding: 'utf8' }));
  const installed = join(app, 'node_modules', '@ruled', 'sheet');
  rmSync(installed, { recursive: true, force: true });
  mkdirSync(installed, { recursive: true });
  execFileSync('tar', ['xzf', join(packed, filename), '-C', installed, '--strip-components=1']);
  rmSync(packed, { recursive: true, force: true });
}
