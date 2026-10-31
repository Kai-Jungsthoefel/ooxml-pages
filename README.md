# ooxml-pages

Render Word (`.docx`) and PowerPoint (`.pptx`) files in the browser as real pages and slides.

- **Laid out like Word:** line breaking, page breaks, columns, sections, headers and footers, footnotes and endnotes, tables that split across pages, floating pictures with every wrap type, keep-with-next, widow control.
- **View only, and small:** about 160 KB (min + gzip) for Word including the PowerPoint drawing code, about 100 KB for PowerPoint alone. No framework, one dependency ([zip.js](https://github.com/gildas-lormeau/zip.js)).
- **Fast:** a typical document lays out in a few hundred milliseconds; pages are drawn lazily while scrolling.
- **PDF export:** the same pages as real PDFs, with selectable text, embedded fonts and vector graphics (optional entry point).
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

### Metafiles

The EMF and WMF renderer that Word pictures use is available on its own, for showing `.emf`, `.wmf`, `.emz` and `.wmz` files:

```ts
import { metafileToPng } from 'ooxml-pages/metafile';

const png = await metafileToPng(new Uint8Array(await file.arrayBuffer())); // Blob, or null if it cannot be drawn
```

It replays the GDI records on a canvas (at most 2400 px on the longer side), so it needs a DOM.

### Content

`ooxml-pages/content` reads what a Word document says without laying it out, for converting it into another markup (LaTeX, Markdown, HTML …):

```ts
import { readDocxContent } from 'ooxml-pages/content';

const content = await readDocxContent(file);
for (const block of content.blocks) {
  if (block.type === 'paragraph') console.log(block.outlineLevel, block.list?.marker, block.inlines);
  else console.log(block.rows.length, 'rows'); // a table
}
const picture = await content.picture(id); // Blob of a picture inline (PNG for EMF/WMF)
content.close();
```

Paragraphs come with their style, outline level (0 for Heading 1), alignment and list marker as Word numbers it; text runs with bold, italic, underline, strike, super- and subscript, caps, monospace and links (runs that look the same are joined). Tables give their grid widths and cells with column and row spans (vertical merges resolved), notes are separate maps referenced from the text, equations stay Office MathML elements, and text frames and boxes appear where they are anchored. Field results and symbol fonts come as Unicode text; hidden text is left out.

### PDF

`ooxml-pages/pdf` turns documents and presentations into PDFs that look like the rendered pages: real text (selectable, searchable, in reading order), embedded fonts reduced to the characters used, vector shapes, charts and gradients, pictures, equations, and working links (web, mail and bookmarks).

```sh
npm install @cantoo/pdf-lib @cantoo/fontkit
```

```ts
import { docxToPdf, pptxToPdf } from 'ooxml-pages/pdf';
import 'ooxml-pages/docx.css'; // and/or pptx.css: the pages are drawn before they become PDF

const { pdf, notes } = await docxToPdf(file, {
  // Font files for a font name and style: one file, several (one font split by Unicode ranges), or null.
  fonts: async (name, { bold, italic }) => loadFontFiles(name, bold, italic),
  onProgress: ({ phase, fraction }) => console.log(phase, fraction),
  title: file.name,      // document properties of the PDF
  creator: 'My app',     // "ooxml-pages" when left out
});
const blob = new Blob([pdf], { type: 'application/pdf' });
```

`pptxToPdf` takes the same options and makes one page per slide; like PowerPoint it leaves out hidden slides unless `hiddenSlides: true`. `notes` lists, in English sentences for the user, what could not be carried over exactly (a font without a file, blurred shadows …).

**Fonts.** A PDF has to embed its fonts, and text is measured with exactly the files that get embedded, so the app supplies them through `fonts`. It is asked for every font the document uses, then for that font's usual substitutes (Carlito for Calibri, Caladea for Cambria, Liberation for Arial, Times New Roman and Courier New …), and finally for the generic families `sans-serif`, `serif`, `monospace` and `symbol` (bullets, Wingdings and Symbol characters, which are translated to Unicode). Fonts embedded in the document are used first. Office's own fonts may not be redistributed, but the metric-compatible open ones may; all of these are on npm as [@fontsource](https://fontsource.org) packages: Carlito, Caladea, Arimo, Tinos, Cousine, plus Noto Sans Symbols, Noto Sans Symbols 2, Noto Sans Math and STIX Two Math for symbols and equations. [`demo/pdf.ts`](demo/pdf.ts) has a complete font source built from them. TTF, OTF and WOFF files work.

The two packages are optional peer dependencies (MIT), so apps that only display documents do not install them. Like the renderers, PDF output runs in the browser: it lays the document out again off-screen with the PDF's fonts and converts the drawn pages. One PDF is made at a time; further calls wait. The font files from `fonts` are registered with the page under private names and stay there for later PDFs, so pages a viewer draws meanwhile keep their fonts.

## What is supported

**Word:** styles and direct formatting, numbering and lists, tab stops with leaders, borders and shading, columns, sections with different page sizes and orientations, headers and footers (first page, even and odd), footnotes and endnotes, fields (page numbers, NUMPAGES, PAGEREF, table of contents), equations (Office Math), tables (merged cells, header rows that repeat, rows that split, cell spacing, vertical text), floating pictures and shapes with every wrap type and anchor, rotation, flips and cropping, text boxes, groups, charts and SmartArt, VML shapes, watermarks and horizontal rules, legacy form fields and content controls with check boxes, symbol fonts, drop caps and text frames, embedded fonts, page color and background pictures, page borders, Word 2010 text effects (gradient fill, outline, shadow), EMF and WMF pictures.

**PowerPoint:** slide → layout → master inheritance, themes and color maps, preset and custom shapes, fills (solid, gradient, picture, pattern), lines and arrowheads, text with bullets, autofit and columns, equations, date fields, tables with table styles, charts (bar, line, pie, area, scatter …), SmartArt (from the drawing PowerPoint saved), groups, hidden slides.

## Limits

- **View only.** There is no editing and no saving. If you need an editor, look at [folio](https://github.com/stella/folio) or [ooxml](https://github.com/ChristopherVR/ooxml).
- **Browser only.** Layout measures text with the browser's canvas, so it needs a DOM (no Node.js or server-side rendering).
- **Fonts decide the line breaks.** Text is measured in the fonts the browser has. Where a document's fonts (Calibri, Cambria …) are not installed, the browser falls back to metric-compatible ones if present (Carlito, Caladea, Liberation) and otherwise to similar fonts, and lines may break differently from Word. Fonts embedded in a document are used.
- **Approximations.** Art page borders (apples, stars …) are pictures built into Word; they are drawn as a dotted band of the same width. Text that Word sets slightly tighter can wrap a word earlier, so long documents may end up with a page more or less.
- **PDF output** leaves out shape shadows, glow, soft edges and picture color effects (recoloring, brightness), and draws soft text shadows from a few faint copies. `notes` says when something was left out.
- **Not supported:** the binary formats `.doc` and `.ppt`, password-protected files, tracked changes (the document is shown as it was saved), comments, animations and transitions, embedded videos.

## Security

Documents are untrusted input. The renderers never run scripts from a document, never load anything from the network, and show SVG only through `<img>` (where browsers run no scripts). Sizes, nesting and counts are capped so that a hostile file cannot exhaust memory. If you find a way around this, please report it privately (see [SECURITY.md](SECURITY.md)).

## Development

```sh
npm install
npm run dev          # demo page with the test files on http://localhost:5180 (pdf.html?file=… compares PDF output)
npm test             # renders every test file in a browser, checks the page counts against Word, makes PDFs
npm run build        # dist/: ES modules, type declarations, stylesheets
npm run samples      # regenerates the test files in test/fixtures/
```

Tests use Playwright. Run `npx playwright install chromium` once, or use an installed browser with `PW_CHANNEL=msedge npm test` (or `chrome`).

The files in `test/fixtures/` are written from scratch by the scripts in `scripts/`, so they can be shared freely. The expected page counts come from Word 365 and PowerPoint 365.

To release, set the version in `package.json` and `CHANGELOG.md`, commit, and push a tag `v<version>` on that commit. The release workflow publishes it to npm.

## License

[MIT](LICENSE)
