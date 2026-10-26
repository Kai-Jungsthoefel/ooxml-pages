// SPDX-License-Identifier: MIT
// Renders every test file completely and compares the page count with what Word and PowerPoint produce.
import { expect, test } from '@playwright/test';

/** Pages (Word 365) and slides (PowerPoint 365) of the files in test/fixtures. */
const EXPECTED: Record<string, number> = {
  'document.docx': 2,
  'floating-pictures.docx': 5,
  'forms-shapes-charts.docx': 4,
  'sections-tables.docx': 7,
  'text-layout.docx': 7,
  'vml-legacy.docx': 2,
  'page-effects.docx': 4,
  'presentation.pptx': 3,
  'shapes.pptx': 4,
  'text-tables-charts.pptx': 3,
};

for (const [file, count] of Object.entries(EXPECTED)) {
  test(file, async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    page.on('console', (message) => {
      if (message.type() === 'error' || message.type() === 'warning') errors.push(message.text());
    });
    await page.goto(`/test.html?file=${file}`);
    await page.waitForFunction(() => window.result !== undefined, null, { timeout: 60_000 });
    const result = await page.evaluate(() => window.result!);
    if ('error' in result) throw new Error(result.error);
    expect(result.count).toBe(count);
    expect(result.states.every((state) => state === 'done')).toBe(true);
    expect(errors).toEqual([]);
  });
}
