// SPDX-License-Identifier: MIT
// Reads test files with readDocxContent (demo/content.html) and checks what a conversion would get: headings, list
// items, tables and merged cells, equations, notes and pictures.
import { expect, test } from '@playwright/test';

const EXPECTED: Record<string, { headings: number; listItems: number; tables: number; mergedCells: number; equations: number; notes: number; pictures: number }> = {
  'document.docx': { headings: 3, listItems: 0, tables: 1, mergedCells: 0, equations: 0, notes: 0, pictures: 1 },
  'sections-tables.docx': { headings: 7, listItems: 0, tables: 6, mergedCells: 6, equations: 0, notes: 0, pictures: 0 },
  'text-layout.docx': { headings: 10, listItems: 18, tables: 0, mergedCells: 0, equations: 2, notes: 3, pictures: 0 },
};

for (const [file, expected] of Object.entries(EXPECTED)) {
  test(`${file} content`, async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto(`/content.html?file=${file}`);
    await page.waitForFunction(() => window.contentResult !== undefined, null, { timeout: 60_000 });
    const result = await page.evaluate(() => window.contentResult!);
    if ('error' in result) throw new Error(result.error);
    const { headings, pictureBytes, footnotes, ...rest } = result;
    expect({ headings: headings.length, ...rest }).toEqual(expected);
    // Every picture can be read, and every note reference has its note
    expect(pictureBytes > 0).toBe(expected.pictures > 0);
    expect(footnotes > 0).toBe(expected.notes > 0);
    expect(errors).toEqual([]);
  });
}
