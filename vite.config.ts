// SPDX-License-Identifier: MIT
// Demo page (demo/) with the test files as its public folder; also serves the page the Playwright tests drive.
import { defineConfig } from 'vite';

export default defineConfig({
  root: 'demo',
  publicDir: '../test/fixtures',
  // GitHub Pages serves the demo below /<repository>/.
  base: process.env.DEMO_BASE ?? '/',
  server: { port: 5180, strictPort: true },
  build: { outDir: '../dist-demo', emptyOutDir: true },
});
