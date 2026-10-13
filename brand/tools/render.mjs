/**
 * Renders the brand's PNGs with Playwright's Chromium (from the repository's
 * e2e tooling): the social card and the app icons.
 *
 *   node brand/tools/render.mjs
 */
import { chromium } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const here = (path) => new URL(path, import.meta.url);
const shots = [
  ['social.html', '../ruled-social.png', 1200, 630],
  ['icon.html', '../ruled-icon-512.png', 512, 512],
  ['icon.html', '../ruled-apple-touch-icon.png', 180, 180],
];
const browser = await chromium.launch();
for (const [page, out, width, height] of shots) {
  const tab = await browser.newPage({ viewport: { width, height } });
  await tab.goto(here(page).href);
  await tab.evaluate(() => document.fonts.ready);
  await tab.screenshot({ path: fileURLToPath(here(out)), omitBackground: true });
  await tab.close();
}
await browser.close();
console.log('Rendered the social card and icons into brand/.');
