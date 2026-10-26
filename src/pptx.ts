// SPDX-License-Identifier: MIT
// PowerPoint (.pptx) entry point. Slides are built from HTML text, SVG shapes and <img> pictures, positioned in
// slide pixels (EMU / 9525); they need the styles in `ooxml-pages/pptx.css`.
import { h } from './dom.js';
import { type LazyOptions, lazyQueue } from './lazy.js';
import { type Deck, loadDeck, renderSlide as drawSlide } from './pptx/slide.js';

export interface SlideInfo {
  index: number;
  /** Slide is hidden in the slide show. */
  hidden: boolean;
}

export interface Presentation {
  /** Slide size in CSS px. */
  width: number;
  height: number;
  slides: SlideInfo[];
  /**
   * Draws one slide into `target`, replacing its content, and resolves once its pictures have loaded.
   * Gives `target` the `pptx-slide` class and the slide size.
   */
  renderSlide(index: number, target: HTMLElement): Promise<void>;
  /** Closes the file and revokes the picture URLs; rendered slides lose their pictures. */
  close(): void;
}

/** Reads a presentation without drawing anything; `onProgress` gets the share of slides read (0–1). */
export async function loadPresentation(file: Blob, onProgress?: (fraction: number) => void): Promise<Presentation> {
  const deck: Deck = await loadDeck(file, onProgress);
  return {
    width: deck.width,
    height: deck.height,
    slides: deck.slides.map((ref, index) => ({ index, hidden: ref.hidden })),
    async renderSlide(index, target) {
      if (!deck.slides[index]) throw new RangeError(`No slide ${index}`);
      target.classList.add('pptx-slide');
      target.style.width = `${deck.width}px`;
      target.style.height = `${deck.height}px`;
      target.replaceChildren();
      const pending = await drawSlide(deck, index, target);
      await Promise.allSettled(pending);
    },
    close() {
      deck.media.revokeAll();
      deck.pkg.close();
    },
  };
}

export interface RenderPptxOptions extends LazyOptions {
  /** Share of slides read (0–1) while the file loads. */
  onProgress?: (fraction: number) => void;
  /**
   * Wraps each slide element before it is added to the container, e.g. in a <figure> with a caption.
   * Return the node to append; it must contain `slide`.
   */
  frame?: (slide: HTMLElement, info: SlideInfo) => Node;
}

export interface PptxView {
  readonly presentation: Presentation;
  /** One element per slide, in order. */
  readonly slides: readonly HTMLElement[];
  /** Renders a slide now instead of waiting for it to scroll into view. */
  renderNow(index: number): void;
  /** Resolves once every slide that has started rendering is done. */
  settled(): Promise<void>;
  /** Stops rendering, removes the slides from the container and releases the file. */
  destroy(): void;
}

/**
 * Renders a presentation into `container` (given the `pptx-deck` class) as a column of slides.
 * Slides are drawn lazily as they come near the visible area unless `lazy: false`.
 */
export async function renderPptx(file: Blob, container: HTMLElement, options: RenderPptxOptions = {}): Promise<PptxView> {
  const presentation = await loadPresentation(file, options.onProgress);
  if (presentation.slides.length === 0) {
    presentation.close();
    throw new Error('The presentation contains no slides.');
  }
  container.classList.add('pptx-deck');
  const added: Node[] = [];
  const slides = presentation.slides.map((info) => {
    const el = h('div', { class: 'pptx-slide' });
    el.style.width = `${presentation.width}px`;
    el.style.height = `${presentation.height}px`;
    const node = options.frame ? options.frame(el, info) : el;
    container.append(node);
    added.push(node);
    return el;
  });
  const queue = lazyQueue(slides, (i, el) => presentation.renderSlide(i, el), { errorText: 'This slide could not be displayed.', ...options }, '1200px 0px');
  return {
    presentation,
    slides,
    renderNow: queue.request,
    settled: queue.idle,
    destroy() {
      queue.destroy();
      for (const node of added) node.parentNode?.removeChild(node);
      container.classList.remove('pptx-deck');
      presentation.close();
    },
  };
}
