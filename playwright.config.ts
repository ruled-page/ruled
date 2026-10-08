import { defineConfig, devices } from '@playwright/test';

const CI = !!process.env.CI;
/** The Angular example has its own install and a newer Node; its test runs when asked for. */
const ANGULAR = !!process.env.E2E_ANGULAR;
/** The packed library installed into a fresh React app; its test runs when asked for. */
const HOST_APP = !!process.env.E2E_HOST_APP;
/** With E2E_BROWSERS=all, the suite runs in Firefox and WebKit (Safari's engine) as well as Chromium. */
const ALL_BROWSERS = process.env.E2E_BROWSERS === 'all';
/**
 * The app's port, with the library's, the Angular example's and the host
 * app's after it. E2E_PORT moves them all, so checkouts can run side by side.
 */
const PORT = Number(process.env.E2E_PORT ?? 4173);
process.env.E2E_PORT = String(PORT);
const testIgnore = [...(ANGULAR ? [] : ['**/angular.spec.ts']), ...(HOST_APP ? [] : ['**/package.spec.ts'])];

/**
 * End-to-end tests in a real browser. The app is tested as it ships, from a
 * production build (React's development-only double mounting stops the
 * grid's screen-reader table from updating, and the tests read cells from
 * it). The HTML host example loads the standalone library build, so a dev
 * server serves it from the project root. With E2E_ANGULAR=1, the Angular
 * example (examples/angular, installed with `npm ci` there) is built and
 * served too, and with E2E_HOST_APP=1 the React app in e2e/host-app, with the
 * library installed from its npm package. Both use the library as already
 * built (`npm run build:lib`), which is then not rebuilt underneath them.
 */
export default defineConfig({
  testDir: 'e2e',
  testIgnore,
  fullyParallel: true,
  // Every test builds a workbook in a worker; with the whole suite running at once, a busy machine needs longer.
  expect: { timeout: 10_000 },
  forbidOnly: CI,
  retries: CI ? 1 : 0,
  reporter: CI ? [['github'], ['html', { open: 'never' }]] : 'list',
  use: {
    baseURL: `http://localhost:${PORT}`,
    trace: 'retain-on-failure',
  },
  // Never another server already on a port: it may be an older build, or another checkout's. A busy port fails the run.
  webServer: [
    {
      command: `npx vite build && npx vite preview --port ${PORT} --strictPort`,
      url: `http://localhost:${PORT}`,
      reuseExistingServer: false,
      timeout: 180_000,
    },
    {
      command: `${ANGULAR || HOST_APP ? '' : 'npm run build:lib && '}npx vite --port ${PORT + 1} --strictPort`,
      url: `http://localhost:${PORT + 1}/examples/html/`,
      reuseExistingServer: false,
      timeout: 180_000,
    },
    ...(ANGULAR
      ? [
          {
            command: `npm --prefix examples/angular run build && cd examples/angular && npx vite preview --outDir dist/browser --port ${PORT + 2} --strictPort`,
            url: `http://localhost:${PORT + 2}`,
            reuseExistingServer: false,
            timeout: 300_000,
          },
        ]
      : []),
    ...(HOST_APP
      ? [
          {
            command: `node scripts/host-app.mjs && npx vite preview -c e2e/host-app/vite.config.ts --port ${PORT + 3} --strictPort`,
            url: `http://localhost:${PORT + 3}/app/`,
            reuseExistingServer: false,
            timeout: 180_000,
          },
        ]
      : []),
  ],
  projects: [
    { name: 'chromium', use: { ...devices['Desktop Chrome'] } },
    ...(ALL_BROWSERS
      ? [
          // Firefox can't emulate a phone (no isMobile or touch), so the phone tests stay out. A
          // project's testIgnore replaces the top-level one, so that list is carried over.
          { name: 'firefox', use: { ...devices['Desktop Firefox'] }, testIgnore: [...testIgnore, '**/mobile.spec.ts'] },
          // The phone tests ask for a Pixel 7, whose settings would start Chromium; naming the
          // browser keeps them in WebKit, at a Pixel 7's size and with touch.
          { name: 'webkit', use: { ...devices['Desktop Safari'], browserName: 'webkit' as const } },
        ]
      : []),
  ],
});
