# Changelog

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
