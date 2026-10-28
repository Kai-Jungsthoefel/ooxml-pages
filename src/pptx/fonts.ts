// SPDX-License-Identifier: MIT
// Maps Office font names to CSS font-family lists. Nothing is downloaded: when a font is not installed, a
// metric-compatible or at least similar system font takes its place, the way PowerPoint substitutes too.

// Lowercase name → substitutes, tried in order after the font itself.
const SUBSTITUTES: Record<string, string[]> = {
  // Aptos is a cloud font that is often missing; PowerPoint then falls back to Calibri.
  aptos: ['Calibri', 'Carlito', 'Segoe UI'],
  'aptos display': ['Calibri', 'Carlito', 'Segoe UI'],
  'aptos narrow': ['Arial Narrow', 'Calibri', 'Carlito'],
  calibri: ['Carlito', 'Segoe UI'],
  'calibri light': ['Calibri', 'Carlito', 'Segoe UI Light'],
  cambria: ['Caladea', 'Georgia'],
  'cambria math': ['Cambria', 'STIX Two Math', 'Latin Modern Math'],
  arial: ['Liberation Sans', 'Arimo', 'Helvetica'],
  'arial narrow': ['Liberation Sans Narrow', 'Arial'],
  'arial black': ['Arial', 'Helvetica'],
  helvetica: ['Arial', 'Liberation Sans', 'Arimo'],
  'helvetica neue': ['Helvetica', 'Arial', 'Liberation Sans'],
  'times new roman': ['Liberation Serif', 'Tinos', 'Times'],
  times: ['Times New Roman', 'Liberation Serif'],
  'courier new': ['Liberation Mono', 'Cousine', 'Courier'],
  'segoe ui': ['Segoe UI', 'Noto Sans', 'Helvetica Neue', 'Arial'],
  'segoe ui light': ['Segoe UI', 'Noto Sans'],
  'century gothic': ['URW Gothic', 'Avant Garde', 'Arial'],
  'franklin gothic book': ['Franklin Gothic', 'Arial'],
  verdana: ['DejaVu Sans'],
  tahoma: ['Verdana', 'DejaVu Sans'],
  consolas: ['Menlo', 'Monaco', 'DejaVu Sans Mono'],
  'lucida console': ['Consolas', 'Menlo'],
  georgia: ['Gelasio', 'Times New Roman'],
  garamond: ['EB Garamond', 'Georgia'],
  'trebuchet ms': ['Arial'],
  'source sans pro': ['Source Sans 3', 'Arial'],
  'avenir next lt pro': ['Avenir Next', 'Avenir', 'Segoe UI'],
  // Office cloud fonts that browsers do not get: the closest installed shapes and widths.
  'tw cen mt': ['Calibri', 'Carlito'],
  'gill sans mt': ['Gill Sans', 'Calibri'],
  rockwell: ['Rockwell', 'Georgia'],
  'century schoolbook': ['Century Schoolbook', 'Georgia'],
};

const SERIF = /times|georgia|cambria|garamond|palatino|book antiqua|bookman|century schoolbook|baskerville|didot|bodoni|constantia|rockwell|serif|minion|caslon|charter|calisto|sylfaen|mincho|song|batang|ming/i;
const MONO = /courier|consolas|mono|lucida console|menlo|monaco|fixedsys/i;

/** Font names come from the document; keep them from breaking out of the family list. */
function clean(name: string): string {
  return name.replace(/["'\\;{}<>\p{Cc}]/gu, '').trim().slice(0, 64);
}

export type GenericFamily = 'sans-serif' | 'serif' | 'monospace' | 'symbol';

/**
 * While a PDF is made (src/pdf/fonts.ts), Office font names resolve to faces made from the very files the PDF
 * embeds, so text is measured exactly as it will be drawn. The hook returns null to keep the usual family list.
 */
export type FamilyHook = (names: string[], generic: GenericFamily) => string | null;
let familyHook: FamilyHook | null = null;

export function setFamilyHook(hook: FamilyHook | null): void {
  familyHook = hook;
}

/** Fonts that stand in for an Office font, best first. */
export function substitutes(name: string): string[] {
  return SUBSTITUTES[name.toLowerCase()] ?? [];
}

export function genericOf(name: string): GenericFamily {
  return MONO.test(name) ? 'monospace' : SERIF.test(name) ? 'serif' : 'sans-serif';
}

export function fontFamily(...names: (string | undefined | null)[]): string {
  const list: string[] = [];
  const asked: string[] = [];
  for (const raw of names) {
    if (!raw) continue;
    const name = clean(raw);
    if (!name || list.includes(name)) continue;
    list.push(name);
    asked.push(name);
    for (const sub of SUBSTITUTES[name.toLowerCase()] ?? []) if (!list.includes(sub)) list.push(sub);
  }
  const generic = genericOf(list[0] ?? '');
  const hooked = familyHook?.(asked, generic);
  if (hooked) return hooked;
  return [...list.map((n) => `"${n}"`), generic].join(', ');
}

/** The family list for formulas (MATH_FAMILY in docx/math.ts). */
export function mathFamily(fallback: string): string {
  return familyHook?.(['Cambria Math'], 'serif') ?? fallback;
}

/** The family list for symbol characters drawn with a Unicode font (SYMBOL_FAMILY in docx/measure.ts). */
export function symbolFamily(fallback: string): string {
  return familyHook?.([], 'symbol') ?? fallback;
}

// Symbol fonts map letters to pictograms. Browsers on other systems lack them, so the common bullet glyphs
// are translated to Unicode and drawn with a regular font.
const WINGDINGS: Record<number, string> = {
  0x6c: '●', 0x6e: '■', 0x6f: '□', 0x71: '❑', 0x75: '◆', 0x76: '❖', 0x77: '⬥', 0x9f: '•', 0xa7: '▪',
  0xa8: '◻', 0xd8: '➢', 0xe0: '→', 0xe8: '➔', 0xf0: '⇨', 0xfb: '✗', 0xfc: '✓', 0xfe: '☑', 0x6d: '❍',
  0x70: '◼', 0x73: '⧫', 0x74: '⬧', 0xa1: '○', 0xa4: '◉', 0xab: '★', 0xd7: '⮚', 0xdf: '🡐', 0xe2: '↑',
};
const WINGDINGS_2: Record<number, string> = { 0xa3: '■', 0xa4: '□', 0x97: '●', 0x98: '○', 0xf8: '✓', 0x50: '✓', 0x4f: '✗' };
const SYMBOL: Record<number, string> = { 0xb7: '•', 0xa8: '♣', 0xa9: '♦', 0xaa: '♥', 0xab: '♠', 0xae: '→', 0xde: '⇒', 0x2d: '−' };

/** Returns a Unicode replacement for a character in a symbol font, or null when the font is a normal one. */
export function symbolChar(font: string | undefined, char: string): string | null {
  if (!font) return null;
  const name = font.toLowerCase();
  const table = name === 'wingdings' ? WINGDINGS : name === 'wingdings 2' ? WINGDINGS_2 : name === 'symbol' ? SYMBOL : null;
  if (!table) return null;
  let code = char.codePointAt(0) ?? 0;
  // Symbol fonts are often addressed through the private use area (U+F000 + code).
  if (code >= 0xf000 && code <= 0xf0ff) code -= 0xf000;
  return table[code] ?? null;
}

export interface FontMetrics {
  /** Ascent and descent per em of the font the browser actually uses for a family list. */
  ascent: number;
  descent: number;
}

const metricsCache = new Map<string, FontMetrics>();
let metricsContext: CanvasRenderingContext2D | null | undefined;

/** Measured with a canvas, which resolves the family list exactly like the text it is used for. */
export function fontMetrics(family: string): FontMetrics {
  let metrics = metricsCache.get(family);
  if (metrics) return metrics;
  metrics = { ascent: 0.9, descent: 0.22 };
  if (metricsContext === undefined) metricsContext = document.createElement('canvas').getContext('2d');
  if (metricsContext) {
    metricsContext.font = `100px ${family}`;
    const m = metricsContext.measureText('Hg');
    if (m.fontBoundingBoxAscent > 0 && m.fontBoundingBoxDescent > 0) {
      metrics = { ascent: m.fontBoundingBoxAscent / 100, descent: m.fontBoundingBoxDescent / 100 };
    }
  }
  metricsCache.set(family, metrics);
  return metrics;
}

/** Advance width of a text in a CSS font shorthand (e.g. "700 24px "Calibri", sans-serif"). */
export function textWidth(text: string, font: string): number {
  if (metricsContext === undefined) metricsContext = document.createElement('canvas').getContext('2d');
  if (!metricsContext || !text) return 0;
  metricsContext.font = font;
  return metricsContext.measureText(text).width;
}

export function isSymbolFont(font: string | undefined): boolean {
  return !!font && /^(wingdings|webdings|symbol|marlett)/i.test(font);
}
