// SPDX-License-Identifier: MIT
// A drawn page (DOM) to PDF operators. The renderers position everything absolutely in page pixels, so the
// browser's geometry is the layout: boxes come from getBoundingClientRect, text word by word from ranges (which
// also gets justified text and letter spacing right), SVG through its transforms (svg.ts). CSS transforms are
// taken off while measuring and put back as PDF transforms, so rotated and scaled content measures correctly.
import { parseColor } from './color.js';
import { paintGradient, parseGradient, splitTop } from './gradient.js';
import type { ImageCache } from './images.js';
import type { Matrix, Rgba, StrokeStyle } from './painter.js';
import { type Segment, parsePath, rectPath } from './path.js';
import { paintSvg } from './svg.js';
import { type TextStyle, type TextTarget, ascentOf, drawText } from './text.js';

export interface PageLink {
  /** Rectangle in page px (left, top, right, bottom). */
  rect: [number, number, number, number];
  href?: string;
  anchor?: string;
}

export interface PageJob extends TextTarget {
  images: ImageCache;
  /** What could not be carried over exactly; shown to the user. */
  notes: Set<string>;
  links: PageLink[];
  /** Page position on screen, subtracted from client rectangles. */
  left: number;
  top: number;
  /** CSS transforms taken off while measuring. */
  transforms: Map<Element, { matrix: DOMMatrix; origin: [number, number] }>;
}

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export const toMatrix = (m: DOMMatrixReadOnly): Matrix => [m.a, m.b, m.c, m.d, m.e, m.f];

export function boxOf(el: Element, job: PageJob): Box {
  const r = el.getBoundingClientRect();
  return { x: r.left - job.left, y: r.top - job.top, w: r.width, h: r.height };
}

const px = (value: string) => parseFloat(value) || 0;

export function offsetPath(segments: Segment[], dx: number, dy: number): Segment[] {
  return segments.map((s) => {
    if (s.k === 'Z') return s;
    if (s.k === 'C') return { ...s, x1: s.x1 + dx, y1: s.y1 + dy, x2: s.x2 + dx, y2: s.y2 + dy, x: s.x + dx, y: s.y + dy };
    return { ...s, x: s.x + dx, y: s.y + dy };
  });
}

/** Takes CSS transforms off (they are measured back in as PDF transforms); returns how to put them back. */
export function stripTransforms(root: HTMLElement, job: PageJob): () => void {
  const found: [HTMLElement | SVGSVGElement, string, string][] = [];
  for (const el of root.querySelectorAll<HTMLElement | SVGSVGElement>('*')) {
    // Inside an SVG, transforms are attributes that getScreenCTM accounts for; only outer SVG boxes count here.
    if (!(el instanceof HTMLElement) && !(el instanceof SVGSVGElement && !(el.parentElement instanceof SVGElement))) continue;
    const cs = getComputedStyle(el);
    if (!cs.transform || cs.transform === 'none') continue;
    const origin = cs.transformOrigin.split(' ').map(px);
    job.transforms.set(el, { matrix: new DOMMatrix(cs.transform), origin: [origin[0] ?? 0, origin[1] ?? 0] });
    found.push([el, el.style.getPropertyValue('transform'), el.style.getPropertyPriority('transform')]);
  }
  for (const [el] of found) el.style.setProperty('transform', 'none', 'important');
  return () => {
    for (const [el, value, priority] of found) {
      if (value) el.style.setProperty('transform', value, priority);
      else el.style.removeProperty('transform');
    }
  };
}

/** The element's own CSS transform around its origin, in page coordinates. */
function ownTransform(el: Element, box: Box, job: PageJob): DOMMatrix | null {
  const t = job.transforms.get(el);
  if (!t) return null;
  const ox = box.x + t.origin[0], oy = box.y + t.origin[1];
  return new DOMMatrix().translate(ox, oy).multiply(t.matrix).translate(-ox, -oy);
}

/** The first color of a CSS gradient, for text whose fill is a gradient (background-clip: text). */
function firstGradientColor(image: string, alpha: number): Rgba | null {
  const m = /(rgba?\([^)]*\)|color\([^)]*\))/.exec(image);
  return m ? parseColor(m[1], alpha) : null;
}

export function textStyleOf(cs: CSSStyleDeclaration, job: PageJob, alpha: number, fill?: string): TextStyle {
  let color = parseColor(fill ?? (cs.webkitTextFillColor || cs.color), alpha);
  if (!color && cs.backgroundClip === 'text') color = firstGradientColor(cs.backgroundImage, alpha);
  const strokeWidth = px(cs.webkitTextStrokeWidth);
  const strokeColor = strokeWidth > 0 ? (parseColor(cs.webkitTextStrokeColor, alpha) ?? color) : null;
  return {
    families: job.fonts.familiesOf(cs.fontFamily),
    size: px(cs.fontSize),
    bold: (parseInt(cs.fontWeight, 10) || 400) >= 600,
    italic: /italic|oblique/.test(cs.fontStyle),
    color,
    outline: strokeColor ? { color: strokeColor, width: strokeWidth } : null,
    charSpacing: cs.letterSpacing === 'normal' ? 0 : px(cs.letterSpacing),
    kern: cs.fontKerning !== 'none',
  };
}

/** A letter as its mathematical italic (text-transform: math-auto on single-letter <mi>), or null. */
function mathItalic(char: string): string | null {
  const c = char.codePointAt(0)!;
  if (c === 0x68) return 'ℎ';
  if (c >= 0x61 && c <= 0x7a) return String.fromCodePoint(0x1d44e + c - 0x61);
  if (c >= 0x41 && c <= 0x5a) return String.fromCodePoint(0x1d434 + c - 0x41);
  if (c >= 0x3b1 && c <= 0x3c9) return String.fromCodePoint(0x1d6fc + c - 0x3b1);
  if (c >= 0x391 && c <= 0x3a9) return String.fromCodePoint(0x1d6e2 + c - 0x391);
  return null;
}

function caseOf(text: string, transform: string): string {
  if (transform === 'uppercase') return text.toUpperCase();
  if (transform === 'lowercase') return text.toLowerCase();
  if (transform === 'capitalize') return text.replace(/(^|\s)(\p{L})/gu, (_, a: string, b: string) => a + b.toUpperCase());
  return text;
}

// ---- Boxes ----------------------------------------------------------------------------------------------------

const DASHED = new Set(['dashed']);
const THREE_D = new Set(['inset', 'outset', 'groove', 'ridge']);

async function paintBorders(cs: CSSStyleDeclaration, b: Box, job: PageJob, alpha: number): Promise<void> {
  const sides = [
    { w: px(cs.borderTopWidth), style: cs.borderTopStyle, color: cs.borderTopColor, side: 'top' },
    { w: px(cs.borderRightWidth), style: cs.borderRightStyle, color: cs.borderRightColor, side: 'right' },
    { w: px(cs.borderBottomWidth), style: cs.borderBottomStyle, color: cs.borderBottomColor, side: 'bottom' },
    { w: px(cs.borderLeftWidth), style: cs.borderLeftStyle, color: cs.borderLeftColor, side: 'left' },
  ] as const;
  const { painter } = job;
  for (const s of sides) {
    if (s.w <= 0 || s.style === 'none' || s.style === 'hidden') continue;
    let color = parseColor(s.color, alpha);
    if (!color) continue;
    // 3D styles in black come out gray in Chrome (Word's inset and outset table borders).
    if (THREE_D.has(s.style) && Math.max(color.r, color.g, color.b) < 0.13) color = { r: 0.659, g: 0.659, b: 0.659, a: color.a };
    const horizontal = s.side === 'top' || s.side === 'bottom';
    // The band the border occupies, along its side.
    const band = (offset: number, width: number): Box =>
      s.side === 'top' ? { x: b.x, y: b.y + offset, w: b.w, h: width }
      : s.side === 'bottom' ? { x: b.x, y: b.y + b.h - offset - width, w: b.w, h: width }
      : s.side === 'left' ? { x: b.x + offset, y: b.y, w: width, h: b.h }
      : { x: b.x + b.w - offset - width, y: b.y, w: width, h: b.h };
    if (s.style === 'double' && s.w >= 3) {
      for (const offset of [0, (s.w * 2) / 3]) {
        const r = band(offset, s.w / 3);
        painter.path(rectPath(r.x, r.y, r.w, r.h));
        painter.fill(color);
      }
    } else if (DASHED.has(s.style) || s.style === 'dotted') {
      const r = band(0, s.w);
      const line: Segment[] = horizontal
        ? [{ k: 'M', x: r.x, y: r.y + s.w / 2 }, { k: 'L', x: r.x + r.w, y: r.y + s.w / 2 }]
        : [{ k: 'M', x: r.x + s.w / 2, y: r.y }, { k: 'L', x: r.x + s.w / 2, y: r.y + r.h }];
      const style: StrokeStyle = s.style === 'dotted' ? { width: s.w, dash: [0, s.w * 2], cap: 'round' } : { width: s.w, dash: [s.w * 3, s.w * 3] };
      painter.path(line);
      painter.stroke(color, style);
    } else {
      const r = band(0, s.w);
      painter.path(rectPath(r.x, r.y, r.w, r.h));
      painter.fill(color);
    }
  }
}

/** Background layers, the last one first (CSS paints the first on top). */
async function paintBackgroundImages(cs: CSSStyleDeclaration, b: Box, job: PageJob, alpha: number): Promise<void> {
  for (const layer of splitTop(cs.backgroundImage).reverse()) {
    if (layer.startsWith('url(')) {
      await paintBackgroundImage(layer, cs, b, job, alpha);
      continue;
    }
    const g = parseGradient(layer, b, alpha);
    // Stripes (repeating gradients, the placeholder of a missing picture) come out in their first color.
    const solid = !g && layer.startsWith('repeating-') ? firstGradientColor(layer, alpha) : null;
    if (!g && !solid) {
      if (layer !== 'none') job.notes.add('Some pattern fills are left out.');
      continue;
    }
    job.painter.save();
    job.painter.path(rectPath(b.x, b.y, b.w, b.h, px(cs.borderTopLeftRadius)));
    if (g) {
      job.painter.clip();
      paintGradient(job.painter, g);
    } else job.painter.fill(solid!);
    job.painter.restore();
  }
}

async function paintBackgroundImage(layer: string, cs: CSSStyleDeclaration, b: Box, job: PageJob, alpha: number): Promise<void> {
  const url = /url\(["']?([^"')]+)["']?\)/.exec(layer)?.[1];
  if (!url) return;
  const image = await job.images.fromUrl(url);
  if (!image) return;
  const { painter } = job;
  if (/100% 100%|cover|contain/.test(cs.backgroundSize)) {
    painter.image(image, b.x, b.y, b.w, b.h, alpha);
    return;
  }
  // Tiled at its own size (page background pictures); a limit keeps tiny tiles from flooding the page.
  const w = Math.max(image.width * 0.75, 4), h = Math.max(image.height * 0.75, 4);
  if ((b.w / w) * (b.h / h) > 2000) {
    painter.image(image, b.x, b.y, b.w, b.h, alpha);
    return;
  }
  painter.save();
  painter.path(rectPath(b.x, b.y, b.w, b.h));
  painter.clip();
  for (let y = b.y; y < b.y + b.h; y += h) for (let x = b.x; x < b.x + b.w; x += w) painter.image(image, x, y, w, h, alpha);
  painter.restore();
}

async function paintBox(cs: CSSStyleDeclaration, b: Box, job: PageJob, alpha: number): Promise<void> {
  if (cs.backgroundClip !== 'text') {
    const bg = parseColor(cs.backgroundColor, alpha);
    if (bg && b.w > 0 && b.h > 0) {
      job.painter.path(rectPath(b.x, b.y, b.w, b.h, px(cs.borderTopLeftRadius)));
      job.painter.fill(bg);
    }
    if (cs.backgroundImage && cs.backgroundImage !== 'none') await paintBackgroundImages(cs, b, job, alpha);
  }
  await paintBorders(cs, b, job, alpha);
  if (cs.boxShadow && cs.boxShadow !== 'none' && !cs.boxShadow.includes(' 0px 0px 0px')) job.notes.add('Shadows are left out.');
}

/** The content box: inside borders and padding. */
function contentBox(cs: CSSStyleDeclaration, b: Box): Box {
  const l = px(cs.borderLeftWidth) + px(cs.paddingLeft), t = px(cs.borderTopWidth) + px(cs.paddingTop);
  const r = px(cs.borderRightWidth) + px(cs.paddingRight), bt = px(cs.borderBottomWidth) + px(cs.paddingBottom);
  return { x: b.x + l, y: b.y + t, w: Math.max(0, b.w - l - r), h: Math.max(0, b.h - t - bt) };
}

async function paintImage(img: HTMLImageElement, cs: CSSStyleDeclaration, b: Box, job: PageJob, alpha: number): Promise<void> {
  const c = contentBox(cs, b);
  if (c.w <= 0 || c.h <= 0) return;
  const image = await job.images.fromUrl(img.currentSrc || img.src, { w: c.w, h: c.h });
  if (!image) {
    if (img.src) job.notes.add('A picture could not be carried over.');
    return;
  }
  let target = c;
  const fit = cs.objectFit;
  if ((fit === 'contain' || fit === 'cover') && image.width && image.height) {
    const scale = (fit === 'contain' ? Math.min : Math.max)(c.w / image.width, c.h / image.height);
    const w = image.width * scale, h = image.height * scale;
    target = { x: c.x + (c.w - w) / 2, y: c.y + (c.h - h) / 2, w, h };
  }
  if (cs.filter && cs.filter !== 'none') job.notes.add('Picture color effects are left out.');
  job.painter.save();
  if (target !== c) {
    job.painter.path(rectPath(c.x, c.y, c.w, c.h));
    job.painter.clip();
  }
  job.painter.image(image, target.x, target.y, target.w, target.h, alpha);
  job.painter.restore();
}

// ---- Text -----------------------------------------------------------------------------------------------------

const range = () => document.createRange();

/** Underline and line-through over each line box of a text node. */
function paintDecorations(node: Text, cs: CSSStyleDeclaration, style: TextStyle, ascent: number, job: PageJob, alpha: number): void {
  const lines = cs.textDecorationLine;
  if (!lines || lines === 'none') return;
  const color = parseColor(cs.textDecorationColor, alpha) ?? style.color;
  if (!color) return;
  const face = style.families[0]?.faces[0]?.font;
  const em = style.size / (face?.unitsPerEm || 1000);
  const thickness = /px$/.test(cs.textDecorationThickness) ? px(cs.textDecorationThickness) : Math.max(0.6, face ? face.underlineThickness * em : style.size / 14);
  const offset = /px$/.test(cs.textUnderlineOffset) ? px(cs.textUnderlineOffset) : face ? -face.underlinePosition * em : style.size * 0.12;
  const r = range();
  r.selectNodeContents(node);
  const kind = cs.textDecorationStyle;
  for (const rect of r.getClientRects()) {
    if (rect.width <= 0) continue;
    const x0 = rect.left - job.left, x1 = rect.right - job.left;
    const baseline = rect.top - job.top + ascent;
    const ys: number[] = [];
    if (lines.includes('underline')) ys.push(baseline + offset + thickness / 2);
    if (lines.includes('line-through')) ys.push(baseline - (face?.xHeight ? (face.xHeight * em) / 2 : style.size * 0.25));
    if (lines.includes('overline')) ys.push(baseline - ascent + thickness / 2);
    for (const y of ys) {
      const doubled = kind === 'double' ? [y - thickness, y + thickness] : [y];
      for (const ly of doubled) {
        if (kind === 'wavy') {
          const segs: Segment[] = [{ k: 'M', x: x0, y: ly }];
          const step = Math.max(thickness * 3, 2);
          for (let x = x0, up = true; x < x1; x += step, up = !up) segs.push({ k: 'L', x: Math.min(x + step, x1), y: ly + (up ? -thickness : thickness) });
          job.painter.path(segs);
          job.painter.stroke(color, { width: thickness, join: 'round' });
          continue;
        }
        job.painter.path([{ k: 'M', x: x0, y: ly }, { k: 'L', x: x1, y: ly }]);
        const dash = kind === 'dotted' ? [thickness, thickness * 2] : kind === 'dashed' ? [thickness * 4, thickness * 3] : undefined;
        job.painter.stroke(color, { width: thickness, dash });
      }
    }
  }
}

/** A word; small capitals are capitals at 70 % for the lower-case letters, as browsers make them up. */
async function drawWord(job: PageJob, text: string, x: number, y: number, style: TextStyle, caps: string): Promise<void> {
  if (caps !== 'small-caps' && caps !== 'all-small-caps') {
    await drawText(job, text, x, y, style);
    return;
  }
  for (const part of text.match(caps === 'all-small-caps' ? /[\s\S]+/gu : /\p{Ll}+|[^\p{Ll}]+/gu) ?? []) {
    const small = caps === 'all-small-caps' || /\p{Ll}/u.test(part);
    x += await drawText(job, part.toUpperCase(), x, y, small ? { ...style, size: style.size * 0.7 } : style);
  }
}

/** A text node of an HTML element, word by word where the browser set it. */
async function paintTextNode(node: Text, cs: CSSStyleDeclaration, job: PageJob, alpha: number): Promise<void> {
  const raw = node.data;
  if (!raw) return;
  const style = textStyleOf(cs, job, alpha);
  if (!style.size || !style.families.length) return;
  const ascent = ascentOf(`${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`);
  const shadow = textShadowOf(cs, style.size, job, alpha);
  const gradient = gradientText(node, cs, job, alpha);
  const stretch = stretchOf(node, job);
  if (stretch) {
    job.painter.save();
    job.painter.transform(stretch);
  }
  const r = range();
  for (const m of raw.matchAll(/\S+\s*/gu)) {
    const start = m.index!;
    const word = m[0].trimEnd();
    r.setStart(node, start);
    r.setEnd(node, start + word.length);
    const rects = r.getClientRects();
    if (!rects.length) continue;
    // A space after the word keeps words apart when text is copied out of the PDF.
    let drawn = caseOf(m[0].length > word.length ? `${word} ` : word, cs.textTransform);
    if (cs.textTransform === 'math-auto' && [...word].length === 1) {
      // A single-letter variable: the math font's italic letter, or the plain letter slanted.
      const italic = mathItalic(word);
      if (italic && job.fonts.coversAll(italic, style.families, style.bold, false)) drawn = italic;
      else style.italic = true;
    }
    if (rects.length === 1 && job.fonts.coversAll(drawn, style.families, style.bold, style.italic)) {
      const x = rects[0].left - job.left, y = rects[0].top - job.top + ascent;
      const shadowStyle = { ...style, outline: null, fillWith: null, shape: true };
      if (shadow) for (const copy of shadow) await drawWord(job, drawn, x + copy.dx, y + copy.dy, { ...shadowStyle, color: copy.color }, cs.fontVariantCaps);
      await drawWord(job, drawn, x, y, { ...style, fillWith: gradient?.(rects[0]) }, cs.fontVariantCaps);
      continue;
    }
    // Broken across lines, or with characters no font has (left out): character by character.
    let i = 0;
    for (const char of word) {
      r.setStart(node, start + i);
      r.setEnd(node, start + i + char.length);
      i += char.length;
      const rect = r.getClientRects()[0];
      if (rect) await drawText(job, caseOf(char, cs.textTransform), rect.left - job.left, rect.top - job.top + ascent, { ...style, fillWith: gradient?.(rect) });
    }
  }
  if (stretch) job.painter.restore();
  paintDecorations(node, cs, style, ascent, job, alpha);
}

/**
 * A stretched MathML operator (parentheses around a fraction): the browser draws a larger variant of the glyph
 * over the operator's box, the text itself measures as the plain glyph. The plain glyph stretched to the box
 * stands in, as a transform for drawing it.
 */
function stretchOf(node: Text, job: PageJob): Matrix | null {
  const mo = node.parentElement;
  if (mo?.localName !== 'mo' || mo.namespaceURI !== MATHML || [...node.data.trim()].length !== 1) return null;
  const r = range();
  r.selectNodeContents(node);
  const glyph = r.getBoundingClientRect(), box = mo.getBoundingClientRect();
  if (!glyph.height || box.height < glyph.height * 1.2) return null;
  const sy = box.height / glyph.height;
  return [1, 0, 0, sy, 0, box.top - job.top - (glyph.top - job.top) * sy];
}

/**
 * Gradient text (background-clip: text, Word's gradient text fill): for a piece of text, what paints its glyphs.
 * The fragments of an inline element share one background laid end to end, as with box-decoration-break: slice.
 */
function gradientText(node: Text, cs: CSSStyleDeclaration, job: PageJob, alpha: number): ((rect: DOMRect) => (() => void) | null) | null {
  if (cs.backgroundClip !== 'text') return null;
  const layer = splitTop(cs.backgroundImage).find((l) => /^(linear|radial)-gradient\(/.test(l));
  const el = node.parentElement;
  if (!layer || !el) return null;
  const fragments = [...el.getClientRects()];
  const total = fragments.reduce((sum, f) => sum + f.width, 0);
  return (rect) => {
    const mid = rect.top + rect.height / 2;
    let before = 0;
    let fragment = fragments.find((f) => {
      const inside = mid >= f.top && mid <= f.bottom && rect.left >= f.left - 1 && rect.right <= f.right + 1;
      if (!inside) before += f.width;
      return inside;
    });
    if (!fragment) {
      fragment = rect;
      before = 0;
    }
    const g = parseGradient(layer, { x: fragment.left - before - job.left, y: fragment.top - job.top, w: total || fragment.width, h: fragment.height }, alpha);
    return g ? () => paintGradient(job.painter, g) : null;
  };
}

/** Where a blurred shadow is drawn several times, around a circle of half its blur radius. */
const SPREAD = [[0, 0], ...Array.from({ length: 8 }, (_, i) => [Math.cos((i * Math.PI) / 4) / 2, Math.sin((i * Math.PI) / 4) / 2])];

/**
 * The copies that make a text shadow (text-shadow, or a drop-shadow filter on text): one for a hard shadow
 * (emboss, simple WordArt); a soft one is drawn several times, slightly apart and faint, which together come
 * close to the blur. Wide blurs (glow) are left out.
 */
function textShadowOf(cs: CSSStyleDeclaration, size: number, job: PageJob, alpha: number): { color: Rgba; dx: number; dy: number }[] | null {
  const value = cs.textShadow && cs.textShadow !== 'none' ? cs.textShadow : /drop-shadow\(((?:[^()]|\([^)]*\))*)\)/.exec(cs.filter ?? '')?.[1];
  if (!value) return null;
  const m = /^(rgba?\([^)]*\)|color\([^)]*\))\s+(-?[\d.]+)px\s+(-?[\d.]+)px(?:\s+(-?[\d.]+)px)?/.exec(value);
  const color = m ? parseColor(m[1], alpha) : null;
  if (!m || !color) return null;
  const dx = parseFloat(m[2]), dy = parseFloat(m[3]), blur = parseFloat(m[4] ?? '0');
  if (blur <= 1) return [{ color, dx, dy }];
  if (blur > Math.max(4, size * 0.15)) {
    job.notes.add('Blurred text shadows and glow are left out.');
    return null;
  }
  // The copies add up to the shadow's own opacity where they all overlap.
  const each = { ...color, a: 1 - Math.pow(1 - color.a, 1 / SPREAD.length) };
  return SPREAD.map(([sx, sy]) => ({ color: each, dx: dx + sx * blur, dy: dy + sy * blur }));
}

// ---- Elements -------------------------------------------------------------------------------------------------

/** Children in painting order: positioned elements by z-index, the rest as they come. */
function ordered(el: Element): ChildNode[] {
  const nodes = [...el.childNodes];
  const z = nodes.map((n) => {
    if (!(n instanceof Element)) return 0;
    const cs = getComputedStyle(n);
    return cs.position !== 'static' && cs.zIndex !== 'auto' ? parseInt(cs.zIndex, 10) || 0 : 0;
  });
  return nodes.map((node, i) => ({ node, z: z[i], i })).sort((a, b) => a.z - b.z || a.i - b.i).map((e) => e.node);
}

function linkOf(a: HTMLAnchorElement, b: Box, ctm: DOMMatrix): PageLink | null {
  const anchor = a.dataset.anchor;
  const href = a.getAttribute('href');
  if (!anchor && !href) return null;
  const pts = [new DOMPoint(b.x, b.y), new DOMPoint(b.x + b.w, b.y), new DOMPoint(b.x, b.y + b.h), new DOMPoint(b.x + b.w, b.y + b.h)].map((p) => ctm.transformPoint(p));
  const xs = pts.map((p) => p.x), ys = pts.map((p) => p.y);
  return { rect: [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)], href: href ?? undefined, anchor };
}

function cssClip(cs: CSSStyleDeclaration, b: Box): { segments: Segment[]; evenOdd: boolean } | null {
  const m = /^path\(\s*(?:(nonzero|evenodd)\s*,\s*)?["'](.*)["']\s*\)$/.exec(cs.clipPath ?? '');
  if (!m) return null;
  return { segments: offsetPath(parsePath(m[2]), b.x, b.y), evenOdd: m[1] === 'evenodd' };
}

const MATHML = 'http://www.w3.org/1998/Math/MathML';

/**
 * Fraction bars and radical signs, which the browser draws for MathML without an element of their own; their
 * geometry follows from the parts around them.
 */
function paintMathLines(el: Element, cs: CSSStyleDeclaration, b: Box, job: PageJob, alpha: number): void {
  const color = parseColor(cs.color, alpha);
  if (!color) return;
  const size = px(cs.fontSize);
  const thickness = Math.max(0.75, size * 0.06);
  const { painter } = job;
  if (el.localName === 'mfrac' && el.children.length >= 2 && el.getAttribute('linethickness') !== '0') {
    const num = boxOf(el.children[0], job), den = boxOf(el.children[1], job);
    const y = (num.y + num.h + den.y) / 2;
    painter.path([{ k: 'M', x: b.x + 1, y }, { k: 'L', x: b.x + b.w - 1, y }]);
    painter.stroke(color, { width: thickness });
  } else if ((el.localName === 'msqrt' || el.localName === 'mroot') && el.children.length) {
    const parts = el.localName === 'mroot' ? [el.children[0]] : [...el.children];
    const boxes = parts.map((p) => boxOf(p, job));
    const left = Math.min(...boxes.map((p) => p.x)), right = Math.max(...boxes.map((p) => p.x + p.w));
    const top = b.y + thickness / 2, bottom = b.y + b.h - thickness / 2;
    const w = Math.max(size * 0.45, Math.min(left - b.x, size * 0.8));
    painter.path([
      { k: 'M', x: left - w, y: b.y + b.h * 0.62 },
      { k: 'L', x: left - w * 0.78, y: b.y + b.h * 0.55 },
      { k: 'L', x: left - w * 0.42, y: bottom },
      { k: 'L', x: left - w * 0.06, y: top },
      { k: 'L', x: right, y: top },
    ]);
    painter.stroke(color, { width: thickness, join: 'round', cap: 'round' });
  }
}

export async function paintElement(el: Element, job: PageJob, ctm: DOMMatrix, alpha: number): Promise<void> {
  const math = el.namespaceURI === MATHML;
  if (!(el instanceof HTMLElement) && !(el instanceof SVGSVGElement) && !math) return;
  const cs = getComputedStyle(el);
  if (cs.display === 'none') return;
  const { painter } = job;
  const b = boxOf(el, job);
  const own = ownTransform(el, b, job);
  let depth = 0;
  if (own) {
    painter.save();
    painter.transform(toMatrix(own));
    ctm = ctm.multiply(own);
    depth++;
  }
  alpha *= Number.isFinite(parseFloat(cs.opacity)) ? parseFloat(cs.opacity) : 1;
  try {
    if (alpha <= 0.002) return;
    const clipPath = cssClip(cs, b);
    if (clipPath) {
      painter.save();
      painter.path(clipPath.segments);
      painter.clip(clipPath.evenOdd);
      depth++;
    }
    const visible = cs.visibility === 'visible';
    if (el instanceof SVGSVGElement) {
      if (cs.overflow !== 'visible') {
        painter.save();
        painter.path(rectPath(b.x, b.y, b.w, b.h));
        painter.clip();
        depth++;
      }
      await paintSvg(el, job, alpha);
      return;
    }
    if (visible) await paintBox(cs, b, job, alpha);
    if (el instanceof HTMLAnchorElement) {
      const link = linkOf(el, b, ctm);
      if (link) job.links.push(link);
    }
    if (el instanceof HTMLImageElement) {
      if (visible) await paintImage(el, cs, b, job, alpha);
      return;
    }
    if (el instanceof HTMLCanvasElement) {
      const image = visible ? await job.images.fromCanvas(el) : null;
      if (image) painter.image(image, b.x, b.y, b.w, b.h, alpha);
      return;
    }
    if (cs.overflowX !== 'visible' || cs.overflowY !== 'visible') {
      const c = { x: b.x + px(cs.borderLeftWidth), y: b.y + px(cs.borderTopWidth), w: b.w - px(cs.borderLeftWidth) - px(cs.borderRightWidth), h: b.h - px(cs.borderTopWidth) - px(cs.borderBottomWidth) };
      painter.save();
      painter.path(rectPath(c.x, c.y, Math.max(0, c.w), Math.max(0, c.h)));
      painter.clip();
      depth++;
    }
    for (const child of ordered(el)) {
      if (child instanceof Text) {
        if (visible) await paintTextNode(child, cs, job, alpha);
      } else if (child instanceof Element) {
        await paintElement(child, job, ctm, alpha);
      }
    }
    if (math && visible) paintMathLines(el, cs, b, job, alpha);
  } finally {
    for (let i = 0; i < depth; i++) painter.restore();
  }
}
