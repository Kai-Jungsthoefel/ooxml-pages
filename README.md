# ooxml-pages

Render Word (`.docx`) and PowerPoint (`.pptx`) files in the browser as real pages and slides.

- **Laid out like Word:** line breaking, page breaks, columns, sections, headers and footers, footnotes and endnotes, tables that split across pages, floating pictures with every wrap type, keep-with-next, widow control.
- **View only, and small:** about 160 KB (min + gzip) for Word including the PowerPoint drawing code, about 100 KB for PowerPoint alone. No framework, one dependency ([zip.js](https://github.com/gildas-lormeau/zip.js)).
- **Fast:** a typical document lays out in a few hundred milliseconds; pages are drawn lazily while scrolling.
- **Private:** everything happens in the browser. Nothing is uploaded, and no script from the document is ever run.

[Live demo](https://kai-jungsthoefel.github.io/ooxml-pages/) · [Changelog](CHANGELOG.md)

## Install

```sh
npm install ooxml-pages
```

## Usage

```ts
import { renderDocx } from 'ooxml-pages/docx';
import 'ooxml-pages/docx.css';

const view = await renderDocx(file, document.getElementById('pages')!, {
  root: document.getElementById('scroller'), // the element that scrolls, for lazy rendering
  onProgress: ({ phase, fraction }) => console.log(phase, fraction),
});
console.log(view.pages.length, 'pages');
// later
view.destroy();
```

```ts
import { renderPptx } from 'ooxml-pages/pptx';
import 'ooxml-pages/pptx.css';

const view = await renderPptx(file, container, {
  // Optional: wrap each slide, e.g. with a caption.
  frame: (slide, { index, hidden }) => {
    const figure = document.createElement('figure');
    const caption = document.createElement('figcaption');
    caption.textContent = `Slide ${index + 1}${hidden ? ' (hidden)' : ''}`;
    figure.append(slide, caption);
    return figure;
  },
});
```

`file` is any `Blob` (a `File` from an `<input>`, a `fetch` response's `blob()` …). Pages and slides are laid out at 96 dpi in CSS pixels; scale them with CSS (`zoom` or `transform`) to fit your layout.

### Options

| Option | Default | |
|---|---|---|
| `lazy` | `true` | Draw pages only when they come near the visible area. `false` draws all of them in order. |
| `root` | viewport | The scrolling element, used for lazy rendering (and for scrolling to bookmarks in Word documents). |
| `preloadMargin` | `'1600px 0px'` (Word), `'1200px 0px'` (PowerPoint) | How far outside the visible area pages are prepared. |
| `errorText` | `'This page could not be displayed.'` | Shown in place of a page that fails to render. |
| `onProgress` | | Word: `{ phase: 'read' \| 'parse' \| 'layout' \| 'relayout', fraction? }`. PowerPoint: share of slides read (0–1). |
| `frame` | | PowerPoint only: wraps each slide element before it is added. |

Both views have `renderNow(index)`, `settled()` (resolves once everything started so far is drawn) and `destroy()`. A Word view also has `scrollToBookmark(name)`; links inside the document (tables of contents, cross references) already use it.

### Lower level

To draw pages yourself (thumbnails, printing, your own virtual scrolling), load the document without any DOM output and render single pages into elements you own:

```ts
import { layoutDocx } from 'ooxml-pages/docx';
import { loadPresentation } from 'ooxml-pages/pptx';

const doc = await layoutDocx(file);            // { pages: [{ width, height, number }], bookmarks, renderPage, close }
await doc.renderPage(0, element);

const deck = await loadPresentation(file);     // { width, height, slides: [{ hidden }], renderSlide, close }
await deck.renderSlide(0, element);
```

Call `close()` when done: it releases the file, the picture URLs and embedded fonts.

## What is supported

**Word:** styles and direct formatting, numbering and lists, tab stops with leaders, borders and shading, columns, sections with different page sizes and orientations, headers and footers (first page, even and odd), footnotes and endnotes, fields (page numbers, NUMPAGES, PAGEREF, table of contents), equations (Office Math), tables (merged cells, header rows that repeat, rows that split, cell spacing, vertical text), floating pictures and shapes with every wrap type and anchor, rotation, flips and cropping, text boxes, groups, charts and SmartArt, VML shapes, watermarks and horizontal rules, legacy form fields and content controls with check boxes, symbol fonts, drop caps and text frames, embedded fonts, page color and background pictures, page borders, Word 2010 text effects (gradient fill, outline, shadow), EMF and WMF pictures.

**PowerPoint:** slide → layout → master inheritance, themes and color maps, preset and custom shapes, fills (solid, gradient, picture, pattern), lines and arrowheads, text with bullets, autofit and columns, tables with table styles, charts (bar, line, pie, area, scatter …), SmartArt (from the drawing PowerPoint saved), groups, hidden slides.

## Limits

- **View only.** There is no editing and no saving. If you need an editor, look at [folio](https://github.com/stella/folio) or [ooxml](https://github.com/ChristopherVR/ooxml).
- **Browser only.** Layout measures text with the browser's canvas, so it needs a DOM (no Node.js or server-side rendering).
- **Fonts decide the line breaks.** Text is measured in the fonts the browser has. Where a document's fonts (Calibri, Cambria …) are not installed, the browser falls back to metric-compatible ones if present (Carlito, Caladea, Liberation) and otherwise to similar fonts, and lines may break differently from Word. Fonts embedded in a document are used.
- **Approximations.** Art page borders (apples, stars …) are pictures built into Word; they are drawn as a dotted band of the same width. Text that Word sets slightly tighter can wrap a word earlier, so long documents may end up with a page more or less.
- **Not supported:** the binary formats `.doc` and `.ppt`, password-protected files, tracked changes (the document is shown as it was saved), comments, animations and transitions, embedded videos.

## Security

Documents are untrusted input. The renderers never run scripts from a document, never load anything from the network, and show SVG only through `<img>` (where browsers run no scripts). Sizes, nesting and counts are capped so that a hostile file cannot exhaust memory. If you find a way around this, please report it privately (see [SECURITY.md](SECURITY.md)).

## Development

```sh
npm install
npm run dev          # demo page with the test files on http://localhost:5180
npm test             # renders every test file in a browser and checks the page counts against Word
npm run build        # dist/: ES modules, type declarations, stylesheets
npm run samples      # regenerates the test files in test/fixtures/
```

Tests use Playwright. Run `npx playwright install chromium` once, or use an installed browser with `PW_CHANNEL=msedge npm test` (or `chrome`).

The files in `test/fixtures/` are written from scratch by the scripts in `scripts/` (except `document.docx` and `presentation.pptx`), so they can be shared freely. The expected page counts come from Word 365 and PowerPoint 365.

## License

[MIT](LICENSE)
