# Changelog

## 0.1.0 – unreleased

First release as a library of its own. The renderers come from the file viewer OPSIS, where they were written.

- `renderDocx`, `layoutDocx` (`ooxml-pages/docx`) and `renderPptx`, `loadPresentation` (`ooxml-pages/pptx`), with lazy rendering, progress reports and bookmark links.
- zip.js is configured per file instead of globally, so the library does not change zip.js for the host application.
- Word: page color and background pictures (`w:background`), art page borders (drawn as a dotted band of the same width), Word 2010 text effects (gradient fill, outline, shadow), VML horizontal rules, tables with cell spacing (separate borders).
- Word: a table whose first row does not fit at the foot of a page now starts on the next page, and keep-with-next chains continue through tables whose last row is kept with the next paragraph, as in Word.
