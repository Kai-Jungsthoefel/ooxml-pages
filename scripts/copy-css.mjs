// SPDX-License-Identifier: MIT
// Copies the stylesheets into dist/. Word pages draw shapes, charts and SmartArt with the PowerPoint renderer,
// so docx.css contains the PowerPoint styles too.
import { readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..');
const read = (file) => readFileSync(path.join(root, 'src', file), 'utf8');
writeFileSync(path.join(root, 'dist', 'pptx.css'), read('pptx/pptx.css'));
writeFileSync(path.join(root, 'dist', 'docx.css'), `${read('pptx/pptx.css')}\n${read('docx/docx.css')}`);
console.log('dist/pptx.css, dist/docx.css');
