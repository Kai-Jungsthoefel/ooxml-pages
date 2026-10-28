// SPDX-License-Identifier: MIT
// PDF output next to the rendered pages, for checking by eye and for tests: ?file= names a test file. The fonts
// are the metric-compatible open fonts from @fontsource (dev dependencies only; apps bring their own).
import * as pdfjs from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { renderDocx } from '../src/docx.js';
import { type PdfFontSource, type PdfProgress, docxToPdf, pptxToPdf } from '../src/pdf.js';
import { renderPptx } from '../src/pptx.js';
import '../src/pptx/pptx.css';
import '../src/docx/docx.css';

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

type Result = { pages: number; bytes: number; notes: string[]; text: string[]; ms: number; url: string } | { error: string };
declare global {
  interface Window {
    pdfResult?: Result;
  }
}

const files = import.meta.glob('../node_modules/@fontsource/{carlito,caladea,arimo,tinos,cousine,noto-sans-symbols,noto-sans-symbols-2,noto-sans-math,stix-two-math}/files/*.woff', {
  query: '?url',
  import: 'default',
}) as Record<string, () => Promise<string>>;

const SYMBOLS = ['noto-sans-symbols-2', 'noto-sans-symbols', 'noto-sans-math'];

/** Office fonts and generic families → open fonts with the same widths (or at least the same characters). */
const FAMILIES: Record<string, string[]> = {
  calibri: ['carlito'], carlito: ['carlito'], cambria: ['caladea'], caladea: ['caladea'], arial: ['arimo'], helvetica: ['arimo'], arimo: ['arimo'],
  'liberation sans': ['arimo'], 'times new roman': ['tinos'], times: ['tinos'], tinos: ['tinos'], 'liberation serif': ['tinos'], 'courier new': ['cousine'],
  courier: ['cousine'], cousine: ['cousine'], 'liberation mono': ['cousine'], 'sans-serif': ['carlito'], serif: ['tinos'], monospace: ['cousine'],
  'cambria math': ['stix-two-math', ...SYMBOLS], 'segoe ui symbol': SYMBOLS, 'segoe ui emoji': SYMBOLS, symbol: SYMBOLS,
};

const fonts: PdfFontSource = async (name, { bold, italic }) => {
  const packages = FAMILIES[name.toLowerCase()];
  if (!packages) return null;
  const suffix = `-${bold ? 700 : 400}-${italic ? 'italic' : 'normal'}.woff`;
  const keys = packages.flatMap((p) => Object.keys(files).filter((k) => k.includes(`/@fontsource/${p}/files/`) && k.endsWith(suffix)));
  return Promise.all(keys.map(async (k) => new Uint8Array(await (await fetch(await files[k]())).arrayBuffer())));
};

const name = new URLSearchParams(location.search).get('file') ?? 'document.docx';
const status = document.getElementById('status')!;
try {
  const blob = await (await fetch(`/${name}`)).blob();
  const slides = name.endsWith('.pptx');
  const view = slides ? await renderPptx(blob, document.getElementById('pages')!, { lazy: false }) : await renderDocx(blob, document.getElementById('pages')!, { lazy: false });
  await view.settled();
  const started = performance.now();
  const options = {
    fonts,
    title: name,
    onProgress: (p: PdfProgress) => (status.textContent = `${name}: ${p.phase}${p.fraction !== undefined ? ` ${Math.round(p.fraction * 100)} %` : ''}`),
  };
  const { pdf, notes } = slides ? await pptxToPdf(blob, options) : await docxToPdf(blob, options);
  const ms = performance.now() - started;
  const url = URL.createObjectURL(new Blob([pdf as Uint8Array<ArrayBuffer>], { type: 'application/pdf' }));
  const doc = await pdfjs.getDocument({ data: pdf.slice() }).promise;
  const text: string[] = [];
  const target = document.getElementById('pdf')!;
  const width = document.getElementById('pages')!.firstElementChild?.getBoundingClientRect().width ?? 816;
  for (let i = 1; i <= doc.numPages; i++) {
    const page = await doc.getPage(i);
    const viewport = page.getViewport({ scale: (width / page.getViewport({ scale: 1 }).width) * devicePixelRatio });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    canvas.style.width = `${viewport.width / devicePixelRatio}px`;
    target.append(canvas);
    await page.render({ canvas, viewport }).promise;
    const content = await page.getTextContent();
    text.push(content.items.map((item) => ('str' in item ? item.str + (item.hasEOL ? '\n' : '') : '')).join(''));
  }
  status.innerHTML = '';
  status.append(`${name}: ${doc.numPages} pages, ${Math.round(pdf.length / 1024)} KB in ${Math.round(ms)} ms. `, Object.assign(document.createElement('a'), { href: url, download: name.replace(/\.(docx|pptx)$/, '.pdf'), textContent: 'Download PDF' }), notes.length ? ` Notes: ${notes.join(' ')}` : '');
  window.pdfResult = { pages: doc.numPages, bytes: pdf.length, notes, text, ms, url };
} catch (error) {
  status.textContent = String(error);
  window.pdfResult = { error: String(error instanceof Error ? error.stack : error) };
}
