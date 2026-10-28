# Changelog

## 0.3.0 – 2026-10-28

- New entry point `ooxml-pages/pdf` with `docxToPdf` and `pptxToPdf`: PDFs that look like the rendered pages, with real text, embedded font subsets, vector shapes, charts, gradients, pictures, equations and links. The app supplies the font files (`PdfFontSource`); `@cantoo/pdf-lib` and `@cantoo/fontkit` are optional peer dependencies, needed only for this entry point.
- `npm test` also converts every test file to PDF and checks pages and text.
- PowerPoint: equations are laid out as equations (as in Word) instead of the picture PowerPoint stores for older readers, which often shows other fonts.
- PowerPoint: date and time fields show the current date or time in the field's language, as PowerPoint updates them.
- PowerPoint: slides no longer depend on the page's `box-sizing`; on pages without `border-box`, centered and right-aligned text was shifted by the text box insets.
- Equations: accents (dot, bar, arrow …) are drawn; Office stores them as combining marks, which browsers show without width or not at all.

## 0.2.0 – 2026-10-27

- New entry point `ooxml-pages/metafile` with `metafileToPng`: the EMF/WMF renderer of the Word pictures, for showing metafiles (also EMZ/WMZ) on their own.

## 0.1.1 – 2026-10-27

Packaging only, the code is the same as in 0.1.0.

- Published from GitHub Actions with a provenance statement, which links the package on npm to the commit it was built from.
- The package includes this changelog.

## 0.1.0 – 2026-10-27

First release as a library of its own. The renderers come from the file viewer OPSIS, where they were written.

- `renderDocx`, `layoutDocx` (`ooxml-pages/docx`) and `renderPptx`, `loadPresentation` (`ooxml-pages/pptx`), with lazy rendering, progress reports and bookmark links.
- zip.js is configured per file instead of globally, so the library does not change zip.js for the host application.
- Word: page color and background pictures (`w:background`), art page borders (drawn as a dotted band of the same width), Word 2010 text effects (gradient fill, outline, shadow), VML horizontal rules, tables with cell spacing (separate borders).
- Word: a table whose first row does not fit at the foot of a page now starts on the next page, and keep-with-next chains continue through tables whose last row is kept with the next paragraph, as in Word.
