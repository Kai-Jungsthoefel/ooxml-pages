// SPDX-License-Identifier: MIT
// Text into the PDF: split into runs by the face each character comes from (as the browser falls back per
// character), each run embedded with its font's glyphs. Positions come from the browser's layout. Copies that
// only decorate (shadows) and glyph areas a gradient fills are drawn as outlines, so the text a reader selects,
// copies or searches is there exactly once.
import type { Font } from '@cantoo/fontkit';
import { PDFHexString, type PDFDocument, type PDFFont } from '@cantoo/pdf-lib';
import type { Family, FontSession } from './fonts.js';
import type { Painter, Rgba } from './painter.js';
import type { Segment } from './path.js';

export interface TextStyle {
  families: Family[];
  /** Font size in px. */
  size: number;
  bold: boolean;
  italic: boolean;
  /** Fill color; null for outlined text without fill. */
  color: Rgba | null;
  outline: { color: Rgba; width: number } | null;
  /** Paints the glyph area in place of the color (gradient text). */
  fillWith?: (() => void) | null;
  /** Drawn as shapes, not as text (shadow copies). */
  shape?: boolean;
  /** Extra space after each character, px (CSS letter-spacing). */
  charSpacing: number;
  scaleX?: number;
  /** Pair kerning, where the browser kerned the text (CSS font-kerning). */
  kern?: boolean;
}

export interface TextTarget {
  doc: PDFDocument;
  painter: Painter;
  fonts: FontSession;
}

const NO_COLOR: Rgba = { r: 0, g: 0, b: 0, a: 0 };
/** Slant of made-up italics. */
const SLANT = 0.21;

/**
 * Glyphs of a run and their advances in font units. Ligatures stay off, as when the text was measured and as the
 * embedded font encodes it (same glyphs); kerning only moves them, which the PDF does with TJ adjustments.
 */
function layout(font: Font, text: string, kern: boolean) {
  const { glyphs, positions } = font.layout(text, { liga: false, kern });
  return { glyphs, advances: glyphs.map((g, i) => (kern ? (positions?.[i]?.xAdvance ?? g.advanceWidth) : g.advanceWidth)) };
}

/** Advance width of a run in px, as the PDF sets it (glyph advances plus letter spacing). */
function runWidth(font: Font, text: string, size: number, charSpacing: number, kern: boolean): number {
  const { advances } = layout(font, text, kern);
  return (advances.reduce((sum, a) => sum + a, 0) * size) / font.unitsPerEm + charSpacing * advances.length;
}

/** Glyph outlines of a run, baseline at (x, y), in page space (y down). */
function runOutline(font: Font, text: string, size: number, x: number, y: number, charSpacing: number, scaleX: number, slant: number, kern: boolean): Segment[] {
  const out: Segment[] = [];
  const k = size / font.unitsPerEm;
  let pen = 0;
  const { glyphs, advances } = layout(font, text, kern);
  for (const [i, glyph] of glyphs.entries()) {
    // As the text matrix maps them: the horizontal scale applies to advances and glyphs, not to the slant.
    const map = (gx: number, gy: number): [number, number] => [x + (pen + gx * k) * scaleX + slant * gy * k, y - gy * k];
    let cx = 0, cy = 0;
    for (const { command, args } of glyph.path.commands) {
      if (command === 'moveTo' || command === 'lineTo') {
        const [px, py] = map(args[0], args[1]);
        out.push({ k: command === 'moveTo' ? 'M' : 'L', x: px, y: py });
        [cx, cy] = [args[0], args[1]];
      } else if (command === 'quadraticCurveTo') {
        // As a cubic: control points two thirds of the way to the quadratic one.
        const [qx, qy, ex, ey] = args;
        const [x1, y1] = map(cx + ((qx - cx) * 2) / 3, cy + ((qy - cy) * 2) / 3);
        const [x2, y2] = map(ex + ((qx - ex) * 2) / 3, ey + ((qy - ey) * 2) / 3);
        const [px, py] = map(ex, ey);
        out.push({ k: 'C', x1, y1, x2, y2, x: px, y: py });
        [cx, cy] = [ex, ey];
      } else if (command === 'bezierCurveTo') {
        const [x1, y1] = map(args[0], args[1]);
        const [x2, y2] = map(args[2], args[3]);
        const [px, py] = map(args[4], args[5]);
        out.push({ k: 'C', x1, y1, x2, y2, x: px, y: py });
        [cx, cy] = [args[4], args[5]];
      } else out.push({ k: 'Z' });
    }
    pen += advances[i] * k + charSpacing;
  }
  return out;
}

/**
 * The run as TJ operands: the encoded glyphs (two bytes each), split where kerning moves the next glyph, with the
 * shift in thousandths of the font size (negative moves right).
 */
function textParts(pdfFont: PDFFont, font: Font, text: string, kern: boolean): (PDFHexString | number)[] {
  const encoded = pdfFont.encodeText(text);
  if (!kern) return [encoded];
  const { glyphs, advances } = layout(font, text, kern);
  const hex = encoded.asString();
  if (hex.length !== glyphs.length * 4) return [encoded];
  const parts: (PDFHexString | number)[] = [];
  let from = 0;
  for (let i = 0; i < glyphs.length; i++) {
    const shift = advances[i] - glyphs[i].advanceWidth;
    if (!shift || i === glyphs.length - 1) continue;
    parts.push(PDFHexString.of(hex.slice(from * 4, (i + 1) * 4)), (-shift * 1000) / font.unitsPerEm);
    from = i + 1;
  }
  parts.push(PDFHexString.of(hex.slice(from * 4)));
  return parts;
}

/** Advance width of a text at scale 1 in px, as drawText sets it. */
export function measureText(fonts: FontSession, text: string, style: TextStyle): number {
  let width = 0;
  for (const run of fonts.runs(text, style.families, style.bold, style.italic)) width += runWidth(run.choice.face.font, run.text, style.size, style.charSpacing, !!style.kern);
  return width;
}

/** Draws text with its baseline starting at (x, y); returns its advance width in px. */
export async function drawText(target: TextTarget, text: string, x: number, y: number, style: TextStyle): Promise<number> {
  if (!text || (!style.color && !style.outline && !style.fillWith)) return 0;
  const { painter } = target;
  const scale = style.scaleX ?? 1;
  let advance = 0;
  for (const run of target.fonts.runs(text, style.families, style.bold, style.italic)) {
    const { face, fakeBold, fakeItalic } = run.choice;
    const rx = x + advance;
    const kern = !!style.kern;
    advance += runWidth(face.font, run.text, style.size, style.charSpacing, kern) * scale;
    if (style.shape || style.fillWith) {
      const outline = runOutline(face.font, run.text, style.size, rx, y, style.charSpacing, scale, fakeItalic ? SLANT : 0, kern);
      if (style.shape) {
        if (!style.color || !outline.length) continue;
        painter.path(outline);
        painter.fill(style.color);
        if (fakeBold) {
          painter.path(outline);
          painter.stroke(style.color, { width: style.size * 0.03, join: 'round' });
        }
        continue;
      }
      if (outline.length) {
        painter.save();
        painter.path(outline);
        painter.clip();
        style.fillWith!();
        painter.restore();
      }
    }
    // With a gradient the text itself only carries the outline, or nothing visible.
    const font = await target.fonts.pdfFont(target.doc, face);
    painter.text(font, style.size, rx, y, textParts(font, face.font, run.text, kern), {
      color: style.fillWith ? NO_COLOR : (style.color ?? NO_COLOR),
      charSpacing: style.charSpacing,
      fakeBold,
      fakeItalic,
      outline: style.outline,
      scaleX: scale,
    });
  }
  return advance;
}

const ascents = new Map<string, number>();
let context: CanvasRenderingContext2D | null | undefined;

/**
 * Distance from the top of a text box to its baseline for a CSS font shorthand: the font's ascent as the browser
 * uses it for the content area of inline text (the same numbers canvas reports).
 */
export function ascentOf(font: string): number {
  let a = ascents.get(font);
  if (a !== undefined) return a;
  if (context === undefined) context = document.createElement('canvas').getContext('2d');
  a = 0;
  if (context) {
    context.font = font;
    a = context.measureText('Hg').fontBoundingBoxAscent;
  }
  ascents.set(font, a);
  return a;
}

/** Forgets measured ascents; the faces behind a family name change from one PDF to the next. */
export function resetAscents(): void {
  ascents.clear();
}
