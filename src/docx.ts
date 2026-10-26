// SPDX-License-Identifier: MIT
// Word (.docx) entry point. The document is laid out like Word does (lines, pages, columns, tables, floating
// objects, headers and footers, notes) and each page is drawn from absolutely positioned text and shapes.
// DrawingML shapes, charts and SmartArt use the PowerPoint renderer. Pages need `ooxml-pages/docx.css`.
import { h } from './dom.js';
import { loadDocument } from './docx/document.js';
import { loadEmbeddedFonts } from './docx/fonts.js';
import { type FlowContext, PageFlow, type PageBox, loadHeaders, sections } from './docx/flow.js';
import { pageNumberText } from './docx/line.js';
import { createContext, defaultSection, emptyNotes, parseBlocks, parseSection } from './docx/parse.js';
import { type RenderShared, renderPage as drawPage } from './docx/render.js';
import { kid } from './docx/xml.js';
import { type LazyOptions, lazyQueue } from './lazy.js';

export type DocxPhase = 'read' | 'parse' | 'layout' | 'relayout';

export interface DocxProgress {
  /** read: opening the file; parse: reading the content; layout / relayout: first and second layout pass. */
  phase: DocxPhase;
  /** Share done (0–1) during the layout passes. */
  fraction?: number;
}

export interface PageInfo {
  index: number;
  /** Page number as Word displays it (sections may restart numbering). */
  number: number;
  /** Page size in CSS px. */
  width: number;
  height: number;
}

export interface DocxLayout {
  pages: PageInfo[];
  /** Bookmark name → page index and y (CSS px from the top of the page). */
  bookmarks: Map<string, { page: number; y: number }>;
  /**
   * Draws one page into `target`, replacing its content, and resolves once its pictures have loaded.
   * Gives `target` the `docx-page` class and the page size.
   */
  renderPage(index: number, target: HTMLElement): Promise<void>;
  /** Closes the file, revokes picture URLs and removes embedded fonts; rendered pages lose their pictures. */
  close(): void;
}

export interface LayoutDocxOptions {
  onProgress?: (progress: DocxProgress) => void;
}

const tick = () => new Promise<void>((resolve) => setTimeout(resolve, 0));

/** Reads and paginates a Word document without drawing anything. */
export async function layoutDocx(file: Blob, options: LayoutDocxOptions = {}): Promise<DocxLayout> {
  const progress = options.onProgress ?? (() => {});
  progress({ phase: 'read' });
  const doc = await loadDocument(file);
  let fonts: Awaited<ReturnType<typeof loadEmbeddedFonts>> | null = null;
  const close = () => {
    doc.media.revokeAll();
    doc.pkg.close();
    fonts?.release();
  };
  try {
    // Embedded fonts first: every measurement depends on them.
    fonts = await loadEmbeddedFonts(doc.pkg, doc.fontTable);
    progress({ phase: 'parse' });
    await tick();
    const notes = emptyNotes();
    const totals = { used: false };
    const blocks = parseBlocks(doc.body, { ...createContext(doc, doc.source, notes, true), totals });
    const finalSect = kid(doc.body, 'sectPr');
    const runs = sections(blocks, finalSect ? parseSection(finalSect, doc.source) : defaultSection());
    const headers = await loadHeaders(doc, runs);

    const flow: FlowContext = {
      doc, notes, stories: new Map(), pageNumber: { value: 1, text: '1', fmt: 'decimal', total: 99, sectionPages: 9 }, budget: { lines: 3_000_000 },
      bookmarks: new Map(), sectionCounts: [], headers,
    };
    /** Page numbers the next pass (and the drawing) uses for NUMPAGES, SECTIONPAGES and PAGEREF. */
    const learn = (pages: PageBox[]) => {
      flow.pageNumber.total = pages.length;
      flow.sectionCounts = [];
      flow.bookmarks = new Map();
      for (const page of pages) {
        flow.sectionCounts[page.sectionIndex] = (flow.sectionCounts[page.sectionIndex] ?? 0) + 1;
        for (const [name] of page.bookmarks) if (!flow.bookmarks.has(name)) flow.bookmarks.set(name, pageNumberText(page.number, page.section.pgNumFmt));
      }
    };
    const layout = async (pre: Map<number, Map<number, never>>, pass: number) => {
      const late = new Map();
      const pageFlow = new PageFlow(flow, pre, late);
      flow.budget.lines = 3_000_000;
      const pages = await pageFlow.run(runs, async (fraction) => {
        progress({ phase: pass === 1 ? 'layout' : 'relayout', fraction });
        await tick();
      });
      return { pageFlow, pages, late };
    };
    let result = await layout(new Map(), 1);
    learn(result.pages);
    // Floating objects found below text they should push aside, or page totals and page references in the
    // text: lay out once more with them known.
    if (result.late.size || totals.used) {
      result = await layout(result.late as Map<number, Map<number, never>>, 2);
      learn(result.pages);
    }
    const { pages, pageFlow } = result;
    if (!pages.length) throw new Error('The document contains no content that can be displayed.');

    const shared: RenderShared = { doc, flow, pages, pageFlow, parsed: new WeakMap() };
    const bookmarks = new Map<string, { page: number; y: number }>();
    pages.forEach((page, index) => {
      for (const [name, y] of page.bookmarks) if (!bookmarks.has(name)) bookmarks.set(name, { page: index, y });
    });
    return {
      pages: pages.map((p, index) => ({ index, number: p.number, width: p.width, height: p.height })),
      bookmarks,
      async renderPage(index, target) {
        const page = pages[index];
        if (!page) throw new RangeError(`No page ${index}`);
        target.classList.add('docx-page');
        target.style.width = `${page.width}px`;
        target.style.height = `${page.height}px`;
        target.replaceChildren();
        const pending = drawPage(page, target, shared);
        // Objects add the pictures they wait for while they render.
        while (pending.length) await Promise.allSettled(pending.splice(0));
      },
      close,
    };
  } catch (error) {
    close();
    throw error;
  }
}

export interface RenderDocxOptions extends LayoutDocxOptions, LazyOptions {}

export interface DocxView {
  readonly layout: DocxLayout;
  /** One element per page, in order. */
  readonly pages: readonly HTMLElement[];
  /** Scrolls the bookmark into view; false if the document has no such bookmark. */
  scrollToBookmark(name: string): boolean;
  /** Renders a page now instead of waiting for it to scroll into view. */
  renderNow(index: number): void;
  /** Resolves once every page that has started rendering is done. */
  settled(): Promise<void>;
  /** Stops rendering, removes the pages from the container and releases the file. */
  destroy(): void;
}

/** Nearest ancestor that scrolls vertically, or null for the page itself. */
function scrollParent(el: HTMLElement): HTMLElement | null {
  for (let p = el.parentElement; p; p = p.parentElement) {
    const overflow = getComputedStyle(p).overflowY;
    if ((overflow === 'auto' || overflow === 'scroll') && p.scrollHeight > p.clientHeight) return p;
  }
  return null;
}

/**
 * Renders a Word document into `container` (given the `docx-doc` class) as a column of pages.
 * Pages are drawn lazily as they come near the visible area unless `lazy: false`. Links to bookmarks inside the
 * document (tables of contents, cross references) scroll to their target.
 */
export async function renderDocx(file: Blob, container: HTMLElement, options: RenderDocxOptions = {}): Promise<DocxView> {
  const layout = await layoutDocx(file, options);
  container.classList.add('docx-doc');
  const pages = layout.pages.map((info) => {
    const el = h('div', { class: 'docx-page' });
    el.style.width = `${info.width}px`;
    el.style.height = `${info.height}px`;
    container.append(el);
    return el;
  });
  const queue = lazyQueue(pages, (i, el) => layout.renderPage(i, el), options, '1600px 0px');

  const scrollToBookmark = (name: string) => {
    const target = layout.bookmarks.get(name);
    if (!target) return false;
    const el = pages[target.page];
    // Measured on screen, so CSS zoom or transforms of the host are taken into account.
    const rect = el.getBoundingClientRect();
    const scale = el.offsetHeight ? rect.height / el.offsetHeight : 1;
    const scroller = options.root instanceof HTMLElement ? options.root : scrollParent(el);
    const top = rect.top + target.y * scale - 24 - (scroller ? scroller.getBoundingClientRect().top : 0);
    (scroller ?? window).scrollBy({ top, behavior: 'smooth' });
    return true;
  };
  const onClick = (event: MouseEvent) => {
    const link = (event.target as HTMLElement).closest<HTMLElement>('a[data-anchor]');
    if (!link || !container.contains(link)) return;
    event.preventDefault();
    scrollToBookmark(link.dataset.anchor ?? '');
  };
  container.addEventListener('click', onClick);

  return {
    layout,
    pages,
    scrollToBookmark,
    renderNow: queue.request,
    settled: queue.idle,
    destroy() {
      queue.destroy();
      container.removeEventListener('click', onClick);
      for (const el of pages) el.remove();
      container.classList.remove('docx-doc');
      layout.close();
    },
  };
}
