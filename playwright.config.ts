// SPDX-License-Identifier: MIT
import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: 'test',
  fullyParallel: true,
  reporter: process.env.CI ? 'github' : 'list',
  use: {
    ...devices['Desktop Chrome'],
    baseURL: 'http://localhost:5180',
    // PW_CHANNEL=msedge or chrome uses an installed browser instead of `npx playwright install chromium`.
    channel: process.env.PW_CHANNEL,
  },
  webServer: { command: 'npx vite', url: 'http://localhost:5180/test.html', reuseExistingServer: !process.env.CI },
});
