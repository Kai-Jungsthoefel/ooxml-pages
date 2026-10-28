// SPDX-License-Identifier: MIT
// PDF output (ooxml-pages/pdf). The document is laid out and drawn exactly as on screen, then each drawn page is
// turned into PDF operators: real text with embedded fonts, vector shapes, gradients and pictures, links and
// internal references. Runs in the browser (it needs the DOM and `ooxml-pages/docx.css` or `pptx.css` on the page)
// and uses the optional dependencies @cantoo/pdf-lib and @cantoo/fontkit.
import fontkit from '@cantoo/fontkit';
import { PDFDocument, PDFString, type PDFObject } from '@cantoo/pdf-lib';
import { layoutDocx } from './docx.js';
import { loadPresentation } from './pptx.js';
import { type PageJob, type PageLink, boxOf, paintElement, stripTransforms } from './pdf/dom.js';
import { FontSession, type PdfFontSource } from './pdf/fonts.js';
import { ImageCache } from './pdf/images.js';
import { PT, Painter } from './pdf/painter.js';
import { parseColor } from './pdf/color.js';
import { rectPath } from './pdf/path.js';
import { resetAscents } from './pdf/text.js';

export type { FontBytes, PdfFontSource } from './pdf/fonts.js';

export type PdfPhase = 'fonts' | 'layout' | 'pages';

export interface PdfProgress {
  /** fonts: finding and loading the fonts; layout: reading and laying out; pages: drawing the pages into the PDF. */
  phase: PdfPhase;
  /** Share done (0–1) during layout and pages. */
  fraction?: number;
}

export interface PdfOptions {
  /**
   * Font files for the fonts the document uses. Text is measured with exactly these files and they are embedded
   * (reduced to the characters used), so the PDF looks the same everywhere. Fonts embedded in the document come
   * first; a name without files falls back to its usual substitutes (Carlito for Calibri …) and then to the
   * generic families "sans-serif", "serif", "monospace" and "symbol".
   */
  fonts: PdfFontSource;
  onProgress?: (progress: PdfProgress) => void;
  /** Title in the PDF's document properties. */
  title?: string;
  /** Application named as the PDF's creator in its document properties; "ooxml-pages" by default. */
  creator?: string;
}

export interface PptxToPdfOptions extends PdfOptions {
  /** Also slides that are hidden in the slide show; PowerPoint leaves them out of PDFs by default. */
  hiddenSlides?: boolean;
}

export interface PdfResult {
  pdf: Uint8Array;
  /** What could not be carried over exactly, in English sentences for the user (e.g. replaced fonts). */
  notes: string[];
}

/** A document opened for drawing: its pages and how to draw each. */
interface Source {
  pages: { width: number; height: number }[];
  render(index: number, target: HTMLElement): Promise<void>;
  /** Bookmark name → page index and y, for internal links. */
  bookmarks: Map<string, { page: number; y: number }>;
  /** What the PDF leaves out on purpose (hidden slides). */
  notes: string[];
  close(): void;
}

type Opener = (file: Blob, progress: (fraction?: number) => void) => Promise<Source>;

let queue: Promise<unknown> = Promise.resolve();

/** One PDF at a time: while a document is laid out for PDF, font names resolve to its faces everywhere. */
function exclusive<T>(task: () => Promise<T>): Promise<T> {
  const run = queue.then(task, task);
  queue = run.catch(() => {});
  return run;
}

async function openDocx(file: Blob, progress: (fraction?: number) => void): Promise<Source> {
  const layout = await layoutDocx(file, { onProgress: (p) => progress(p.fraction) });
  return { pages: layout.pages, render: (i, target) => layout.renderPage(i, target), bookmarks: layout.bookmarks, notes: [], close: () => layout.close() };
}

function pptxOpener(hiddenSlides: boolean): Opener {
  return async (file, progress) => {
    const presentation = await loadPresentation(file, progress);
    const slides = presentation.slides.filter((s) => hiddenSlides || !s.hidden);
    if (!slides.length) {
      presentation.close();
      throw new Error('The presentation contains no slides to print.');
    }
    const left = presentation.slides.length - slides.length;
    return {
      pages: slides.map(() => ({ width: presentation.width, height: presentation.height })),
      render: (i, target) => presentation.renderSlide(slides[i].index, target),
      bookmarks: new Map(),
      notes: left ? [`${left === 1 ? 'One hidden slide was' : `${left} hidden slides were`} left out, as PowerPoint does.`] : [],
      close: () => presentation.close(),
    };
  };
}

/** Converts a Word document (.docx) to PDF. */
export function docxToPdf(file: Blob, options: PdfOptions): Promise<PdfResult> {
  return exclusive(() => convert(file, options, openDocx, false));
}

/** Converts a PowerPoint presentation (.pptx) to PDF, one page per slide. */
export function pptxToPdf(file: Blob, options: PptxToPdfOptions): Promise<PdfResult> {
  return exclusive(() => convert(file, options, pptxOpener(!!options.hiddenSlides), true));
}

/** Link annotations: web and mail links open, internal references jump to their bookmark. */
function addLinks(doc: PDFDocument, pageLinks: PageLink[][], bookmarks: Source['bookmarks']): void {
  const pages = doc.getPages();
  pageLinks.forEach((links, index) => {
    const page = pages[index];
    const height = page.getHeight();
    for (const link of links) {
      const [l, t, r, b] = link.rect;
      const entry: Record<string, PDFObject | string | number | null | (PDFObject | string | number | null)[]> = {
        Type: 'Annot',
        Subtype: 'Link',
        Rect: [l * PT, height - b * PT, r * PT, height - t * PT],
        Border: [0, 0, 0],
      };
      if (link.href && /^(https?:|mailto:|tel:)/i.test(link.href)) {
        entry.A = doc.context.obj({ Type: 'Action', S: 'URI', URI: PDFString.of(link.href) });
      } else if (link.anchor) {
        const target = bookmarks.get(link.anchor);
        if (!target || !pages[target.page]) continue;
        const targetPage = pages[target.page];
        entry.Dest = [targetPage.ref, 'XYZ', null, targetPage.getHeight() - target.y * PT, null];
      } else continue;
      page.node.addAnnot(doc.context.register(doc.context.obj(entry as never)));
    }
  });
}

async function paintPage(target: HTMLElement, job: PageJob): Promise<void> {
  const restore = stripTransforms(target, job);
  try {
    const rect = target.getBoundingClientRect();
    job.left = rect.left;
    job.top = rect.top;
    // The page itself: its color, not its shadow.
    const bg = parseColor(getComputedStyle(target).backgroundColor);
    if (bg) {
      const b = boxOf(target, job);
      job.painter.path(rectPath(0, 0, b.w, b.h));
      job.painter.fill(bg);
    }
    for (const child of target.children) await paintElement(child, job, new DOMMatrix(), 1);
  } finally {
    restore();
  }
}

async function convert(file: Blob, options: PdfOptions, open: Opener, drawToRecord: boolean): Promise<PdfResult> {
  const progress = options.onProgress ?? (() => {});
  const fonts = new FontSession(options.fonts);
  const host = document.createElement('div');
  host.setAttribute('aria-hidden', 'true');
  // Off screen but laid out: the geometry of the drawn pages is what goes into the PDF.
  host.style.cssText = 'position:fixed;left:-200000px;top:0;pointer-events:none';
  document.body.append(host);
  let source: Source | null = null;
  try {
    progress({ phase: 'fonts' });
    // A first pass with the usual fonts finds the font names the document uses. Word documents name them while
    // they are laid out; slides only while they are drawn.
    fonts.record();
    const probe = await open(file, () => {});
    try {
      if (drawToRecord) {
        for (let i = 0; i < probe.pages.length; i++) {
          const target = document.createElement('div');
          host.append(target);
          await probe.render(i, target);
          target.remove();
        }
      }
    } finally {
      probe.close();
    }
    await fonts.load();
    fonts.activate();
    resetAscents();
    source = await open(file, (fraction) => progress({ phase: 'layout', fraction }));

    const doc = await PDFDocument.create();
    doc.registerFontkit(fontkit);
    if (options.title) doc.setTitle(options.title);
    doc.setCreator(options.creator ?? 'ooxml-pages');
    doc.setProducer('ooxml-pages');
    const images = new ImageCache(doc);
    const notes = new Set<string>();
    const pageLinks: PageLink[][] = [];
    for (let index = 0; index < source.pages.length; index++) {
      const { width, height } = source.pages[index];
      progress({ phase: 'pages', fraction: index / source.pages.length });
      const target = document.createElement('div');
      host.append(target);
      await source.render(index, target);
      const page = doc.addPage([width * PT, height * PT]);
      const painter = new Painter(doc, page, height);
      const job: PageJob = { doc, painter, fonts, images, notes, links: [], left: 0, top: 0, transforms: new Map() };
      await paintPage(target, job);
      painter.finish();
      pageLinks.push(job.links);
      target.remove();
    }
    addLinks(doc, pageLinks, source.bookmarks);
    for (const note of source.notes) notes.add(note);
    if (fonts.unknown.size) notes.add(`No font file for ${[...fonts.unknown].sort().join(', ')}; a similar font was used.`);
    if (fonts.missing.size) {
      // Symbol font characters without a Unicode meaning (private use area) are named by their code.
      const shown = [...fonts.missing].slice(0, 20).map((c) => (/[\p{Co}\p{Cc}]/u.test(c) ? `U+${c.codePointAt(0)!.toString(16).toUpperCase()}` : c));
      notes.add(`Some characters have no font and are left out: ${shown.join(' ')}`);
    }
    const pdf = await doc.save();
    progress({ phase: 'pages', fraction: 1 });
    return { pdf, notes: [...notes] };
  } finally {
    source?.close();
    host.remove();
    fonts.release();
    resetAscents();
  }
}
