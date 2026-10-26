// SPDX-License-Identifier: MIT
// Page the Playwright tests drive: renders ?file= completely (not lazily) and reports the result on window.result.
import { renderDocx } from '../src/docx.js';
import { renderPptx } from '../src/pptx.js';
import '../src/pptx/pptx.css';
import '../src/docx/docx.css';

type Result = { count: number; states: string[]; ms: number } | { error: string };
declare global {
  interface Window { result?: Result }
}

const name = new URLSearchParams(location.search).get('file') ?? '';
const container = document.getElementById('pages')!;
try {
  const blob = await (await fetch(`/${name}`)).blob();
  const started = performance.now();
  let elements: readonly HTMLElement[];
  if (name.endsWith('.pptx')) {
    const view = await renderPptx(blob, container, { lazy: false });
    await view.settled();
    elements = view.slides;
  } else {
    const view = await renderDocx(blob, container, { lazy: false });
    await view.settled();
    elements = view.pages;
  }
  window.result = { count: elements.length, states: elements.map((el) => el.dataset.state ?? ''), ms: performance.now() - started };
} catch (error) {
  window.result = { error: String(error instanceof Error ? error.stack : error) };
}
