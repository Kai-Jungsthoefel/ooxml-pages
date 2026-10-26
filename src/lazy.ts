// SPDX-License-Identifier: MIT
// Renders page or slide elements one at a time, either all in order or only as they come near the visible area.
import { h } from './dom.js';

export interface LazyOptions {
  /** Render pages only when they come near the visible area (default true); false renders all of them in order. */
  lazy?: boolean;
  /** Scroll container for lazy rendering; default: the viewport. */
  root?: Element | Document | null;
  /** How far outside the visible area pages are prepared (IntersectionObserver rootMargin). */
  preloadMargin?: string;
  /** Text shown in place of a page that fails to render. */
  errorText?: string;
}

export interface LazyQueue {
  /** Settles once every page requested so far has rendered (or failed). */
  idle(): Promise<void>;
  /** Renders the page now if it has not started yet. */
  request(index: number): void;
  destroy(): void;
}

/**
 * Pages render one at a time: decoding several large pictures at once would only make every page slower.
 * Each element carries `data-state` (pending, loading, done, error) for styling and tests.
 */
export function lazyQueue(
  elements: HTMLElement[],
  render: (index: number, el: HTMLElement) => Promise<void>,
  options: LazyOptions,
  defaultMargin: string,
): LazyQueue {
  let queue: Promise<void> = Promise.resolve();
  let destroyed = false;
  const errorText = options.errorText ?? 'This page could not be displayed.';

  const request = (index: number) => {
    const el = elements[index];
    if (!el || el.dataset.state !== 'pending') return;
    el.dataset.state = 'loading';
    observer?.unobserve(el);
    queue = queue.then(async () => {
      if (destroyed) return;
      try {
        await render(index, el);
        el.dataset.state = 'done';
      } catch (error) {
        console.error(error);
        el.replaceChildren(h('div', { class: 'pptx-error' }, errorText));
        el.dataset.state = 'error';
      }
    });
  };

  for (const el of elements) el.dataset.state = 'pending';
  let observer: IntersectionObserver | null = null;
  if (options.lazy === false || typeof IntersectionObserver === 'undefined') {
    elements.forEach((_, i) => request(i));
  } else {
    observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) if (entry.isIntersecting) request(elements.indexOf(entry.target as HTMLElement));
      },
      { root: options.root ?? null, rootMargin: options.preloadMargin ?? defaultMargin },
    );
    elements.forEach((el) => observer!.observe(el));
  }

  return {
    idle: async () => {
      // Pages requested while waiting extend the queue, so wait until it stops changing.
      let seen: Promise<void>;
      do {
        seen = queue;
        await seen;
      } while (seen !== queue);
    },
    request,
    destroy: () => {
      destroyed = true;
      observer?.disconnect();
    },
  };
}
