// SPDX-License-Identifier: MIT
// Page the Playwright tests drive: reads ?file= with readDocxContent and reports what it found on window.contentResult.
import { type ContentBlock, readDocxContent } from '../src/content.js';

type Result =
  | { headings: string[]; listItems: number; tables: number; mergedCells: number; equations: number; notes: number; footnotes: number; pictures: number; pictureBytes: number }
  | { error: string };
declare global {
  interface Window { contentResult?: Result }
}

const name = new URLSearchParams(location.search).get('file') ?? '';
try {
  const content = await readDocxContent(await (await fetch(`/${name}`)).blob());
  const r = { headings: [] as string[], listItems: 0, tables: 0, mergedCells: 0, equations: 0, notes: 0, footnotes: content.footnotes.size, pictures: 0, pictureBytes: 0 };
  const pictures: number[] = [];
  const walk = (blocks: ContentBlock[]) => {
    for (const b of blocks) {
      if (b.type === 'table') {
        r.tables++;
        for (const row of b.rows) for (const c of row.cells) {
          if (c.colSpan > 1 || c.rowSpan > 1) r.mergedCells++;
          walk(c.blocks);
        }
        continue;
      }
      if (b.outlineLevel !== undefined) r.headings.push(b.inlines.map((i) => (i.kind === 'text' ? i.text : '')).join(''));
      if (b.list) r.listItems++;
      for (const i of b.inlines) {
        if (i.kind === 'math') r.equations++;
        else if (i.kind === 'note') r.notes++;
        else if (i.kind === 'picture') pictures.push(i.id);
        else if (i.kind === 'box') walk(i.blocks);
      }
    }
  };
  walk(content.blocks);
  r.pictures = pictures.length;
  for (const id of pictures) r.pictureBytes += (await content.picture(id))?.size ?? 0;
  content.close();
  window.contentResult = r;
} catch (error) {
  window.contentResult = { error: String(error instanceof Error ? error.stack : error) };
}
