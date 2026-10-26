// SPDX-License-Identifier: MIT
import { type DocxView, renderDocx } from '../src/docx.js';
import { type PptxView, renderPptx } from '../src/pptx.js';
import '../src/pptx/pptx.css';
import '../src/docx/docx.css';

const SAMPLES = [
  'text-layout.docx', 'sections-tables.docx', 'floating-pictures.docx', 'forms-shapes-charts.docx', 'vml-legacy.docx', 'page-effects.docx',
  'shapes.pptx', 'text-tables-charts.pptx',
];

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const scroller = $('scroller');
const pages = $('pages');
const status = $('status');
let current: DocxView | PptxView | null = null;

async function open(file: File) {
  current?.destroy();
  current = null;
  pages.replaceChildren();
  scroller.scrollTop = 0;
  const started = performance.now();
  status.textContent = `Reading ${file.name} …`;
  try {
    if (/\.(pptx|pptm|ppsx|potx)$/i.test(file.name)) {
      const view = await renderPptx(file, pages, { root: scroller });
      current = view;
      status.textContent = `${file.name}: ${view.slides.length} slides`;
    } else {
      const view = await renderDocx(file, pages, {
        root: scroller,
        onProgress: ({ phase, fraction }) => {
          status.textContent = `${phase} ${fraction === undefined ? '' : `${Math.round(fraction * 100)} %`}`;
        },
      });
      current = view;
      status.textContent = `${file.name}: ${view.pages.length} pages`;
    }
    status.textContent += ` in ${Math.round(performance.now() - started)} ms`;
    fit();
  } catch (error) {
    console.error(error);
    status.textContent = `${file.name}: ${error instanceof Error ? error.message : error}`;
  }
}

/** Scales the pages down to the window width on small screens. */
function fit() {
  pages.style.transform = '';
  pages.style.height = '';
  const room = scroller.clientWidth - 32;
  const scale = Math.min(1, room / pages.scrollWidth);
  if (scale < 1) {
    pages.style.transform = `scale(${scale})`;
    pages.style.height = `${pages.scrollHeight * scale}px`;
  }
}
addEventListener('resize', fit);

for (const name of SAMPLES) {
  const button = document.createElement('button');
  button.textContent = name;
  button.onclick = async () => {
    const blob = await (await fetch(`${import.meta.env.BASE_URL}${name}`)).blob();
    await open(new File([blob], name));
  };
  $('samples').append(button);
}
$<HTMLInputElement>('file').onchange = (event) => {
  const file = (event.target as HTMLInputElement).files?.[0];
  if (file) open(file);
};
scroller.addEventListener('dragover', (event) => {
  event.preventDefault();
  scroller.classList.add('drop');
});
scroller.addEventListener('dragleave', () => scroller.classList.remove('drop'));
scroller.addEventListener('drop', (event) => {
  event.preventDefault();
  scroller.classList.remove('drop');
  const file = event.dataTransfer?.files[0];
  if (file) open(file);
});
