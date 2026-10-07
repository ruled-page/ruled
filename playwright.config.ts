import { defineConfig, devices } from '@playwright/test';

const CI = !!process.env.CI;
/** The Angular example has its own install and a newer Node; its test runs when asked for. */
const ANGULAR = !!process.env.E2E_ANGULAR;

/**
 * End-to-end tests in a real browser. The app is tested as it ships, from a
 * production build (React's development-only double mounting stops the
 * grid's screen-reader table from updating, and the tests read cells from
 * it). The HTML host example loads the standalone library build, so a dev
 * server serves it from the project root. With E2E_ANGULAR=1, the Angular
 * example (examples/angular, installed with `npm ci` there, and the library
 * built) is built and served too.
 */
export default defineConfig({
  testDir: 'e2e',
  testIgnore: ANGULAR ? [] : ['**/angular.spec.ts'],
  fullyParallel: true,
  forbidOnly: CI,
  retries: CI ? 1 : 0,
  reporter: CI ? [['github'], ['html', { open: 'never' }]] : 'list',
  use: {
    baseURL: 'http://localhost:4173',
    trace: 'retain-on-failure',
  },
  webServer: [
    {
      command: 'npx vite build && npx vite preview --port 4173 --strictPort',
      url: 'http://localhost:4173',
      reuseExistingServer: !CI,
      timeout: 180_000,
    },
    {
      command: 'npm run build:lib && npx vite --port 4174 --strictPort',
      url: 'http://localhost:4174/examples/html/',
      reuseExistingServer: !CI,
      timeout: 180_000,
    },
    ...(ANGULAR
      ? [
          {
            command: 'npm --prefix examples/angular run build && npm --prefix examples/angular run preview',
            url: 'http://localhost:4175',
            reuseExistingServer: !CI,
            timeout: 300_000,
          },
        ]
      : []),
  ],
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
});
