/**
 * Builds the example sites, each its own Cloudflare Pages project, into
 * dist-sites/ (each example's wrangler.jsonc names its project and folder):
 *
 *   demo     the editor demo with the sample workbooks  (examples/demo, project workbook-editor-poc)
 *   portal   the editor embedded in a React supplier portal  (examples/react-portal, workbook-editor-portal)
 *   angular  the editor embedded in an Angular app  (examples/angular-portal, workbook-editor-angular)
 *
 *   node scripts/build-sites.mjs            all three
 *   node scripts/build-sites.mjs portal     just one (or several)
 *
 * The sites link to each other at their *.pages.dev addresses; set DEMO_URL,
 * PORTAL_URL or ANGULAR_URL to use other addresses, such as custom domains.
 * The Angular site has its own install (made here if missing) and needs Node
 * 22.22.3 or later, which Angular's CLI requires.
 */
import { execSync } from 'node:child_process';
import { copyFileSync, cpSync, existsSync, readdirSync, rmSync } from 'node:fs';

const ALL = ['demo', 'portal', 'angular'];
const wanted = process.argv.slice(2);
for (const site of wanted) {
  if (!ALL.includes(site)) {
    console.error(`Unknown site "${site}". Sites: ${ALL.join(', ')}.`);
    process.exit(2);
  }
}
const sites = wanted.length ? wanted : ALL;

const urls = {
  VITE_DEMO_URL: process.env.DEMO_URL ?? 'https://workbook-editor-poc.pages.dev/',
  VITE_PORTAL_URL: process.env.PORTAL_URL ?? 'https://workbook-editor-portal.pages.dev/',
  VITE_ANGULAR_URL: process.env.ANGULAR_URL ?? 'https://workbook-editor-angular.pages.dev/',
};
const run = (command, options = {}) => execSync(command, { stdio: 'inherit', ...options, env: { ...process.env, ...urls, ...options.env } });

function nodeIsNewEnoughForAngular() {
  const [major, minor, patch] = process.versions.node.split('.').map(Number);
  if (major === 22) return minor > 22 || (minor === 22 && patch >= 3);
  return major >= 24;
}

if (sites.includes('angular') && !nodeIsNewEnoughForAngular()) {
  console.error(
    `The Angular site needs Node 22.22.3 or later (this is ${process.versions.node}).\n` +
      'Switch Node (with fnm: fnm use 22.23), or build the others with: node scripts/build-sites.mjs demo portal',
  );
  process.exit(1);
}

if (sites.includes('demo')) run('npm run build -w @sitepen/demo', { env: { SITE: '1' } });

if (sites.includes('portal')) {
  run('npm run build -w @sitepen/react-portal', { env: { SITE: '1' } });
  // The portal shares the demo's public folder, and needs only its own workbooks from it.
  for (const file of readdirSync('dist-sites/portal/samples')) {
    if (!file.startsWith('supplier-')) rmSync(`dist-sites/portal/samples/${file}`);
  }
}

if (sites.includes('angular')) {
  // The Angular app loads the editor's standalone build.
  run('npm run build:lib');
  if (!existsSync('examples/angular-portal/node_modules')) run('npm ci', { cwd: 'examples/angular-portal' });
  run('npx ng build', { cwd: 'examples/angular-portal' });
  rmSync('dist-sites/angular', { recursive: true, force: true });
  cpSync('examples/angular-portal/dist/browser', 'dist-sites/angular', { recursive: true });
  copyFileSync('examples/demo/public/_headers', 'dist-sites/angular/_headers');
}

console.log(`\nBuilt ${sites.map((s) => `dist-sites/${s}`).join(', ')}`);
