/**
 * Puts what the live examples load into public/: the editor's standalone
 * build (from `npm run build:lib` at the repository's root) and the sample
 * workbooks the demo has. Both are copies, not committed.
 */
import { copyFileSync, cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';

const root = new URL('../../', import.meta.url);
const editor = new URL('packages/spreadsheet/dist/standalone/workbook-editor.js', root);
if (!existsSync(editor)) {
  console.error('The editor is not built: run `npm run build:lib` at the repository root first (`npm run docs:dev` does both).');
  process.exit(1);
}
const out = (path) => {
  const url = new URL(`../public/${path}`, import.meta.url);
  mkdirSync(new URL('./', url), { recursive: true });
  return url;
};
// The whole build: its entry loads the rest in chunks as they're needed. Source maps stay behind.
const editorDir = out('editor/');
rmSync(editorDir, { recursive: true, force: true });
cpSync(new URL('./', editor), editorDir, { recursive: true, filter: (src) => !src.endsWith('.map') });
for (const name of ['price-list.xlsx', 'budget.xlsx', 'forecast.xlsx', 'kitchen-sink.xlsx', 'chart-gallery.xlsx', 'pivot.xlsx', 'orders.xlsx']) {
  copyFileSync(new URL(`examples/demo/public/samples/${name}`, root), out(`samples/${name}`));
}

// Ruled's brand, from brand/: the lockups for the site's logo, and the icons and social card it serves.
const brand = (name) => new URL(`brand/${name}`, root);
mkdirSync(new URL('../src/assets/brand/', import.meta.url), { recursive: true });
for (const name of ['ruled-lockup.svg', 'ruled-lockup-dark.svg']) copyFileSync(brand(name), new URL(`../src/assets/brand/${name}`, import.meta.url));
copyFileSync(brand('ruled-favicon.svg'), out('favicon.svg'));
copyFileSync(brand('ruled-apple-touch-icon.png'), out('apple-touch-icon.png'));
copyFileSync(brand('ruled-social.png'), out('social.png'));

rmSync(new URL('../public/agents/', import.meta.url), { recursive: true, force: true });
// The agent skill, from one source: as a Claude Code skill, and as a Cursor rule.
const site = (process.env.DOCS_URL ?? 'https://ruled.page').replace(/\/$/, '');
const skill = readFileSync(new URL('../agents/ruled.md', import.meta.url), 'utf8').replaceAll('{{SITE}}', site);
const description =
  'Embed Ruled (@ruled/sheet), the spreadsheet editor, in a web app: write its schemas with sync and async validation hooks and business logic, and handle opening and saving. Use when adding an Excel or spreadsheet editor to an app, or writing defineSchema code.';
writeFileSync(out('agents/ruled/SKILL.md'), `---\nname: ruled\ndescription: ${description}\n---\n\n${skill}`);
writeFileSync(out('agents/ruled.mdc'), `---\ndescription: ${description}\nglobs:\nalwaysApply: false\n---\n\n${skill}`);
console.log('Copied the editor and the sample workbooks into public/, and wrote the agent skill.');
