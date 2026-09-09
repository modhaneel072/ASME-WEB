import { defineConfig, devices } from '@playwright/test';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));

// Playwright drives the *built* app served by a throwaway seeded Flask server
// (python manage.py serve-e2e), so it exercises the same path production uses.
const repoRoot = path.resolve(here, '..', '..');
const python =
  process.platform === 'win32'
    ? path.join(repoRoot, '.venv', 'Scripts', 'python.exe')
    : path.join(repoRoot, '.venv', 'bin', 'python');
const port = Number(process.env.ASME_E2E_PORT || 5055);

export default defineConfig({
  testDir: './e2e',
  timeout: 60_000,
  expect: { timeout: 10_000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [['list'], ['html', { open: 'never' }]],
  use: {
    baseURL: `http://127.0.0.1:${port}`,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  webServer: {
    command: `"${python}" manage.py serve-e2e`,
    cwd: repoRoot,
    url: `http://127.0.0.1:${port}/healthz`,
    reuseExistingServer: false,
    timeout: 120_000,
    env: { ASME_E2E_PORT: String(port), ASME_OPS_WEB_DIST: path.join(here, 'dist') },
  },
  projects: [
    { name: 'desktop-1440', use: { ...devices['Desktop Chrome'], viewport: { width: 1440, height: 900 } } },
    {
      name: 'desktop-1366',
      use: { ...devices['Desktop Chrome'], viewport: { width: 1366, height: 768 } },
      grep: /@responsive/,
    },
    {
      name: 'tablet-768',
      use: { ...devices['Desktop Chrome'], viewport: { width: 768, height: 1024 }, hasTouch: true },
      grep: /@responsive/,
    },
    {
      name: 'mobile-390',
      use: {
        ...devices['Desktop Chrome'],
        viewport: { width: 390, height: 844 },
        hasTouch: true,
        isMobile: true,
      },
      grep: /@responsive/,
    },
  ],
});
