// Browser tests for the task pane pages. They need a running stack — see
// README.md "Tests" — and reach it through tests/e2e/lib.sh.
import process from 'node:process';
import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: './tests/browser',
  // Tests share the seeded subject and its token count; serial keeps the DB assertions exact.
  workers: 1,
  timeout: 60_000,
  forbidOnly: !!process.env.CI,
  reporter: process.env.CI ? [['github'], ['html', { open: 'never' }]] : [['list'], ['html', { open: 'never' }]],
  use: {
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure',
  },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
});
