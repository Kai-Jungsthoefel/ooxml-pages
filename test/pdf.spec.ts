// SPDX-License-Identifier: MIT
// Turns every test file into a PDF (demo/pdf.html, with the open fonts from @fontsource) and checks the page count,
// that the text comes out once and in order, and that no character went missing.
import { expect, test } from '@playwright/test';

/**
 * Pages, a piece of text the first page must contain exactly once, and pages without text (left blank before an
 * odd-page section). Hidden slides are left out.
 */
const EXPECTED: Record<string, { pages: number; text: string; blank?: number[] }> = {
  'document.docx': { pages: 2, text: 'Sample document' },
  'floating-pictures.docx': { pages: 5, text: 'header with a logo' },
  'forms-shapes-charts.docx': { pages: 4, text: 'Forms, shapes and charts' },
  'sections-tables.docx': { pages: 7, text: 'Sections and tables', blank: [4] },
  'text-layout.docx': { pages: 7, text: 'Fonts, line spacing' },
  'vml-legacy.docx': { pages: 2, text: 'VML and legacy layout' },
  // Gradient fill, outline and shadow: still one piece of text.
  'page-effects.docx': { pages: 4, text: 'Page effects' },
  'presentation.pptx': { pages: 3, text: 'Sample presentation' },
  'shapes.pptx': { pages: 3, text: 'Basic shapes' },
  'text-tables-charts.pptx': { pages: 3, text: 'Text formatting' },
};

for (const [file, expected] of Object.entries(EXPECTED)) {
  test(`${file} to PDF`, async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    page.on('console', (message) => {
      if (message.type() === 'error') errors.push(message.text());
    });
    await page.goto(`/pdf.html?file=${file}`);
    await page.waitForFunction(() => window.pdfResult !== undefined, null, { timeout: 120_000 });
    const result = await page.evaluate(() => window.pdfResult!);
    if ('error' in result) throw new Error(result.error);
    expect(result.pages).toBe(expected.pages);
    expect(result.text.flatMap((text, i) => (text.trim() ? [] : [i + 1]))).toEqual(expected.blank ?? []);
    if (expected.text) expect(result.text[0].split(expected.text).length - 1).toBe(1);
    expect(result.notes.filter((note) => note.startsWith('Some characters'))).toEqual([]);
    expect(errors).toEqual([]);
  });
}
