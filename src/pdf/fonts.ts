// SPDX-License-Identifier: MIT
// Fonts for PDF output. Text has to be measured with exactly the font files the PDF embeds, so while a PDF is
// made every Office font name maps to a face of its own (opdf-1 …) built from files the app supplies, and
// installed fonts with the same name never mix in. A family may consist of several files, each covering part of
// Unicode (Latin, Cyrillic …); like the browser, a character the first family lacks comes from the next one.
import fontkit from '@cantoo/fontkit';
import type { Font, FontCollection } from '@cantoo/fontkit';
import type { PDFDocument, PDFFont } from '@cantoo/pdf-lib';
import { resetFontCaches, setInstalledHook } from '../docx/measure.js';
import { type EmbeddedFontFile, setEmbeddedFontSink } from '../docx/fonts.js';
import { type GenericFamily, isSymbolFont, setFamilyHook, substitutes } from '../pptx/fonts.js';

export type FontBytes = ArrayBuffer | Uint8Array;

/**
 * Font files for an Office font name (or a generic family: "sans-serif", "serif", "monospace", "symbol") in a style;
 * null when the app has none. Several files are one font split by Unicode ranges.
 */
export type PdfFontSource = (name: string, style: { bold: boolean; italic: boolean }) => Promise<FontBytes | FontBytes[] | null | undefined>;

export interface Face {
  bold: boolean;
  italic: boolean;
  bytes: Uint8Array<ArrayBuffer>;
  font: Font;
  /** Code point ranges the file covers, for the CSS face and the choice per character. */
  ranges: [number, number][];
}

export interface Family {
  /** Private CSS family name. */
  css: string;
  faces: Face[];
}

/** The face that draws a character, and what the PDF has to imitate because the font has no such style. */
export interface FaceChoice {
  face: Face;
  fakeBold: boolean;
  fakeItalic: boolean;
}

const STYLES = [
  { bold: false, italic: false },
  { bold: true, italic: false },
  { bold: false, italic: true },
  { bold: true, italic: true },
];

/** Code point ranges of a font's character map, merged. */
function rangesOf(font: Font): [number, number][] {
  const codes = [...font.characterSet].sort((a, b) => a - b);
  const ranges: [number, number][] = [];
  for (const code of codes) {
    const last = ranges[ranges.length - 1];
    if (last && code <= last[1] + 1) last[1] = Math.max(last[1], code);
    else ranges.push([code, code]);
  }
  return ranges;
}

const hex = (n: number) => n.toString(16).toUpperCase();

let count = 0;

/**
 * Faces from the app's font source, by source and name, and their families. They stay registered as long as the
 * page lives: a viewer that lays out pages while a PDF is made gets the PDF's families, which must not disappear
 * afterwards. Later PDFs reuse them.
 */
const shared = new WeakMap<PdfFontSource, { files: Map<string, Promise<Face[] | null>>; families: WeakMap<Face[], Family> }>();

function covers(face: Face, code: number): boolean {
  // Ranges are sorted: binary search.
  let lo = 0, hi = face.ranges.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const [a, b] = face.ranges[mid];
    if (code < a) hi = mid - 1;
    else if (code > b) lo = mid + 1;
    else return face.font.hasGlyphForCodePoint(code);
  }
  return false;
}

function addFace(faces: Face[], bytes: FontBytes, bold: boolean, italic: boolean): void {
  const data = new Uint8Array(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  try {
    let font = fontkit.create(data) as Font | FontCollection;
    if ('fonts' in font) font = font.fonts[0];
    const ranges = rangesOf(font as Font);
    if (ranges.length) faces.push({ bold, italic, bytes: data, font: font as Font, ranges });
  } catch {
    // A file fontkit cannot read is left out; its text falls back like a missing font.
  }
}

/** Characters that need no glyph of their own: zero-width and direction marks, the soft hyphen. */
function invisible(code: number): boolean {
  return (code >= 0x200b && code <= 0x200f) || (code >= 0x2028 && code <= 0x202e) || (code >= 0x2060 && code <= 0x2064) || code === 0xfeff || code === 0xad;
}

export class FontSession {
  private readonly names = new Set<string>();
  private readonly embedded: EmbeddedFontFile[] = [];
  /** Faces of fonts embedded in the document, by lower-case name; only for this PDF. */
  private readonly own = new Map<string, Face[]>();
  private readonly families = new Map<string, Family>();
  private readonly byCss = new Map<string, Family>();
  private readonly generic = new Map<GenericFamily, Family | null>();
  /** CSS faces of the document's own fonts, removed again when done. */
  private readonly cssFaces: FontFace[] = [];
  private readonly pdfFonts = new Map<Face, PDFFont>();
  private readonly cache: { files: Map<string, Promise<Face[] | null>>; families: WeakMap<Face[], Family> };
  /** Office names no file was found for; their text uses a generic family. */
  readonly unknown = new Set<string>();
  /** Characters no font had a glyph for. */
  readonly missing = new Set<string>();

  constructor(private readonly source: PdfFontSource) {
    let cache = shared.get(source);
    if (!cache) {
      cache = { files: new Map(), families: new WeakMap() };
      shared.set(source, cache);
    }
    this.cache = cache;
  }

  /** Notes the font names the document asks for while it is laid out once with the usual fonts. */
  record(): void {
    setFamilyHook((names) => {
      for (const name of names) this.names.add(name);
      return null;
    });
    setEmbeddedFontSink((file) => this.embedded.push(file));
    // No font on this computer counts: symbol fonts are translated to Unicode, as the PDF will need them.
    setInstalledHook(() => false);
    // Fonts measured before (a document on screen) would otherwise skip the hook.
    resetFontCaches();
  }

  /** Gets the files for all names noted and registers them as faces with private names. */
  async load(): Promise<void> {
    setFamilyHook(null);
    setEmbeddedFontSink(null);
    for (const name of this.names) await this.family(name);
    for (const generic of ['sans-serif', 'serif', 'monospace', 'symbol'] as GenericFamily[]) {
      const faces = await this.files(generic);
      this.generic.set(generic, faces ? this.register(faces) : null);
    }
    const used = [...this.byCss.keys()];
    await Promise.all([...document.fonts].filter((face) => used.includes(face.family.replace(/^["']|["']$/g, ''))).map((face) => face.load().catch(() => {})));
  }

  /** From now on the layout measures with the registered faces. */
  activate(): void {
    setFamilyHook((names, generic) => this.cssList(names, generic));
    setInstalledHook((name) => {
      const family = this.families.get(name.toLowerCase());
      if (!family) return false;
      // A symbol font (Wingdings, Symbol) counts only with its own encoding (U+F020 …); a Unicode symbol font
      // standing in for it needs the text translated to Unicode first.
      return !isSymbolFont(name) || family.faces.some((face) => covers(face, 0xf06c) || covers(face, 0xf0b7));
    });
    resetFontCaches();
  }

  release(): void {
    setFamilyHook(null);
    setInstalledHook(null);
    setEmbeddedFontSink(null);
    for (const face of this.cssFaces) document.fonts.delete(face);
    this.cssFaces.length = 0;
    resetFontCaches();
  }

  /** Faces for a name: the document's own font, else the app's files (shared between PDFs); null when neither has it. */
  private async files(name: string): Promise<Face[] | null> {
    const key = name.toLowerCase();
    let own = this.own.get(key);
    if (!own) {
      own = [];
      for (const file of this.embedded) {
        if (file.name.toLowerCase() === key) addFace(own, file.bytes, file.bold, file.italic);
      }
      this.own.set(key, own);
    }
    if (own.length) return own;
    let files = this.cache.files.get(key);
    if (!files) {
      files = this.fetchFiles(name);
      this.cache.files.set(key, files);
    }
    return files;
  }

  private async fetchFiles(name: string): Promise<Face[] | null> {
    const faces: Face[] = [];
    for (const style of STYLES) {
      let result: FontBytes | FontBytes[] | null | undefined = null;
      try {
        result = await this.source(name, style);
      } catch {
        result = null;
      }
      for (const bytes of result ? (Array.isArray(result) ? result : [result]) : []) addFace(faces, bytes, style.bold, style.italic);
    }
    return faces.length ? faces : null;
  }

  /** The family for an Office name: its own files, else those of a substitute (Carlito for Calibri …). */
  private async family(name: string): Promise<Family | null> {
    const key = name.toLowerCase();
    if (this.families.has(key)) return this.families.get(key)!;
    let faces: Face[] | null = null;
    for (const candidate of [name, ...substitutes(name)]) {
      faces = await this.files(candidate);
      if (faces) break;
    }
    if (!faces) {
      this.unknown.add(name);
      return null;
    }
    const family = this.register(faces);
    this.families.set(key, family);
    return family;
  }

  /** A family with a private CSS name; the app's files keep theirs for later PDFs, the document's own go when done. */
  private register(faces: Face[]): Family {
    const own = [...this.own.values()].includes(faces);
    let family = own ? undefined : this.cache.families.get(faces);
    if (!family) {
      family = { css: `opdf-${++count}`, faces };
      for (const face of faces) {
        const unicodeRange = face.ranges.map(([a, b]) => (a === b ? `U+${hex(a)}` : `U+${hex(a)}-${hex(b)}`)).join(', ');
        const css = new FontFace(family.css, face.bytes, { weight: face.bold ? '700' : '400', style: face.italic ? 'italic' : 'normal', unicodeRange });
        document.fonts.add(css);
        if (own) this.cssFaces.push(css);
      }
      if (!own) this.cache.families.set(faces, family);
    }
    this.byCss.set(family.css, family);
    return family;
  }

  /** CSS family list for Office names: their families, then the generic one and the symbol fallback. */
  private cssList(names: string[], generic: GenericFamily): string | null {
    const list: string[] = [];
    for (const name of names) {
      const family = this.families.get(name.toLowerCase());
      if (family) list.push(family.css);
      // A name first met while drawing (chart and shape text) gets the generic family.
      else if (!this.families.has(name.toLowerCase())) this.unknown.add(name);
    }
    for (const g of [generic, 'sans-serif', 'symbol'] as GenericFamily[]) {
      const family = this.generic.get(g);
      if (family && !list.includes(family.css)) list.push(family.css);
    }
    return list.length ? list.map((css) => `"${css}"`).join(', ') : null;
  }

  /** The private families in a CSS font-family list, in order; other names are ignored. */
  familiesOf(cssFamily: string): Family[] {
    const out: Family[] = [];
    for (const part of cssFamily.split(',')) {
      const family = this.byCss.get(part.trim().replace(/^["']|["']$/g, ''));
      if (family && !out.includes(family)) out.push(family);
    }
    // Text outside the document's fonts (a generic CSS family) still needs a font in the PDF.
    if (!out.length) {
      const fallback = this.generic.get('sans-serif') ?? [...this.generic.values()].find(Boolean);
      if (fallback) out.push(fallback);
    }
    return out;
  }

  /** The face for a character in a family list and style, as the browser picks it. */
  faceFor(families: Family[], code: number, bold: boolean, italic: boolean): FaceChoice | null {
    for (const family of families) {
      const candidates = family.faces.filter((face) => covers(face, code));
      if (!candidates.length) continue;
      const face =
        candidates.find((f) => f.bold === bold && f.italic === italic) ??
        candidates.find((f) => f.italic === italic) ??
        candidates.find((f) => f.bold === bold) ??
        candidates[0];
      return { face, fakeBold: bold && !face.bold, fakeItalic: italic && !face.italic };
    }
    return null;
  }

  /** Whether every visible character of a text has a face; otherwise it is drawn character by character. */
  coversAll(text: string, families: Family[], bold: boolean, italic: boolean): boolean {
    for (const char of text) {
      const code = char.codePointAt(0)!;
      if (!invisible(code) && !/\s/.test(char) && !this.faceFor(families, code, bold, italic)) return false;
    }
    return true;
  }

  /** Splits text into runs of one face each. Characters no font has are left out and noted. */
  runs(text: string, families: Family[], bold: boolean, italic: boolean): { text: string; choice: FaceChoice }[] {
    const out: { text: string; choice: FaceChoice }[] = [];
    for (const char of text) {
      const code = char.codePointAt(0)!;
      const choice = this.faceFor(families, code, bold, italic);
      if (!choice) {
        if (!invisible(code) && !/\s/.test(char)) this.missing.add(char);
        continue;
      }
      const last = out[out.length - 1];
      if (last && last.choice.face === choice.face && last.choice.fakeBold === choice.fakeBold && last.choice.fakeItalic === choice.fakeItalic) last.text += char;
      else out.push({ text: char, choice });
    }
    return out;
  }

  /** The face embedded in the PDF, reduced to the glyphs used. */
  async pdfFont(doc: PDFDocument, face: Face): Promise<PDFFont> {
    let font = this.pdfFonts.get(face);
    if (!font) {
      // Ligatures stay off, as when the text was measured; kerning, where the text has it, comes as TJ adjustments.
      font = await doc.embedFont(face.bytes, { subset: true, features: { liga: false, kern: false } });
      this.pdfFonts.set(face, font);
    }
    return font;
  }
}
