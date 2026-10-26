// SPDX-License-Identifier: MIT
// Fills and outlines: parsing (including theme style references with their placeholder color) and painting
// onto SVG paths or, for backgrounds and table cells, as CSS backgrounds.
import { type Color, type ColorContext, colorIn, colorOf, css, mix, mixLinear, shadeFor } from './color.js';
import { type Env, type Media, type Source, px, r2, svg, uid } from './env.js';
import { attr, bool, child, children, clamp, int, relAttr } from './xml.js';

export interface GradientStop {
  pos: number;
  color: Color;
}

export type Fill =
  | { type: 'none' }
  | { type: 'solid'; color: Color }
  | { type: 'grad'; stops: GradientStop[]; angle: number; scaled: boolean; path: string | null; focus: [number, number, number, number] }
  | { type: 'blip'; el: Element; source: Source }
  | { type: 'patt'; fg: Color; bg: Color; preset: string }
  | { type: 'group' };

export interface LineEnd {
  type: string;
  /** Width and length as multiples of the line width. */
  w: number;
  len: number;
}

export interface Line {
  fill: Fill;
  /** Width in px. */
  width: number;
  dash: string;
  cap: 'butt' | 'round' | 'square';
  join: 'miter' | 'round' | 'bevel';
  head?: LineEnd;
  tail?: LineEnd;
}

const FILLS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

/** The fill element among the children of spPr, bgPr, tcPr …, if any. */
export function fillElement(container: Element | null | undefined): Element | null {
  if (!container) return null;
  for (const el of container.children) if (FILLS.has(el.localName)) return el;
  return null;
}

export function parseFill(el: Element, colors: ColorContext, source: Source): Fill {
  switch (el.localName) {
    case 'solidFill': {
      const color = colorIn(el, colors);
      return color ? { type: 'solid', color } : { type: 'none' };
    }
    case 'gradFill': {
      const stops = children(child(el, 'gsLst'), 'gs')
        .map((gs) => ({ pos: clamp((int(gs, 'pos') ?? 0) / 100000, 0, 1), color: colorIn(gs, colors) }))
        .filter((s): s is GradientStop => s.color !== null)
        .sort((a, b) => a.pos - b.pos);
      if (stops.length === 0) return { type: 'none' };
      const lin = child(el, 'lin');
      const pathEl = child(el, 'path');
      const rect = child(pathEl, 'fillToRect');
      const pct = (name: string) => (int(rect, name) ?? 0) / 100000;
      return {
        type: 'grad',
        stops,
        angle: (int(lin, 'ang') ?? 0) / 60000,
        scaled: bool(lin, 'scaled') ?? false,
        path: pathEl ? (attr(pathEl, 'path') ?? 'circle') : null,
        focus: [pct('l'), pct('t'), pct('r'), pct('b')],
      };
    }
    case 'blipFill':
      return { type: 'blip', el, source };
    case 'pattFill': {
      const fg = colorIn(child(el, 'fgClr'), colors) ?? { r: 0, g: 0, b: 0, a: 1 };
      const bg = colorIn(child(el, 'bgClr'), colors) ?? { r: 255, g: 255, b: 255, a: 1 };
      return { type: 'patt', fg, bg, preset: attr(el, 'prst') ?? 'pct50' };
    }
    case 'grpFill':
      return { type: 'group' };
    default:
      return { type: 'none' };
  }
}

/** Fill given directly in a container (spPr …); undefined when the container does not specify one. */
export function ownFill(container: Element | null | undefined, colors: ColorContext, source: Source): Fill | undefined {
  const el = fillElement(container);
  return el ? parseFill(el, colors, source) : undefined;
}

function withPlaceholder(ref: Element | null, colors: ColorContext): ColorContext {
  const ph = colorIn(ref, colors);
  return ph ? { ...colors, phClr: ph } : colors;
}

/** Fill from a style reference (fillRef of p:style, bgRef of a background). */
export function refFill(ref: Element | null, env: Env, source: Source): Fill | undefined {
  if (!ref) return undefined;
  const idx = int(ref, 'idx') ?? 0;
  if (idx === 0) return { type: 'none' };
  const list = idx >= 1000 ? env.theme.bgFills : env.theme.fills;
  const el = list[(idx >= 1000 ? idx - 1000 : idx) - 1];
  return el ? parseFill(el, withPlaceholder(ref, env.colors), source) : undefined;
}

const DASHES: Record<string, number[]> = {
  dash: [4, 3], dot: [1, 3], lgDash: [8, 3], dashDot: [4, 3, 1, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
};
const END_SIZE: Record<string, number> = { sm: 2, med: 3, lg: 5 };

function lineEnd(el: Element | null): LineEnd | undefined {
  const type = attr(el, 'type');
  if (!el || !type || type === 'none') return undefined;
  return { type, w: END_SIZE[attr(el, 'w') ?? 'med'] ?? 3, len: END_SIZE[attr(el, 'len') ?? 'med'] ?? 3 };
}

/** An a:ln element applied on top of a base line (from the theme style), property by property. */
export function parseLine(ln: Element | null, colors: ColorContext, source: Source, base?: Line): Line | undefined {
  if (!ln) return base;
  const line: Line = base ? { ...base } : { fill: { type: 'none' }, width: 1, dash: 'solid', cap: 'butt', join: 'round' };
  const w = int(ln, 'w');
  if (w !== undefined) line.width = clamp(px(w), 0, 500);
  const fill = ownFill(ln, colors, source);
  if (fill) line.fill = fill;
  else if (!base) line.fill = { type: 'none' };
  const cap = attr(ln, 'cap');
  if (cap) line.cap = cap === 'rnd' ? 'round' : cap === 'sq' ? 'square' : 'butt';
  const dash = attr(child(ln, 'prstDash'), 'val');
  if (dash) line.dash = dash;
  else if (child(ln, 'custDash')) line.dash = 'dash';
  if (child(ln, 'round')) line.join = 'round';
  else if (child(ln, 'bevel')) line.join = 'bevel';
  else if (child(ln, 'miter')) line.join = 'miter';
  const head = child(ln, 'headEnd'), tail = child(ln, 'tailEnd');
  if (head) line.head = lineEnd(head);
  if (tail) line.tail = lineEnd(tail);
  return line;
}

export function refLine(ref: Element | null, env: Env, source: Source): Line | undefined {
  if (!ref) return undefined;
  const idx = int(ref, 'idx') ?? 0;
  if (idx === 0) return undefined;
  const ln = env.theme.lines[idx - 1];
  if (!ln) return undefined;
  const line = parseLine(ln, withPlaceholder(ref, env.colors), source);
  // Without an explicit width the theme lines are 0.75 pt.
  if (line && int(ln, 'w') === undefined) line.width = 1;
  return line;
}

// ---- Painting -------------------------------------------------------------------------------------------------

function patternColor(fill: Extract<Fill, { type: 'patt' }>): Color {
  const pct = /^pct(\d+)$/.exec(fill.preset);
  return pct ? mix(fill.bg, fill.fg, Number(pct[1]) / 100) : fill.fg;
}

/** Representative single color of a fill, e.g. for text or where only a flat color can be drawn. */
export function fillColor(fill: Fill | undefined): Color | null {
  if (!fill) return null;
  switch (fill.type) {
    case 'solid':
      return fill.color;
    case 'grad':
      return fill.stops[0].color;
    case 'patt':
      return patternColor(fill);
    default:
      return null;
  }
}

/** Blip fill crop (srcRect) and stretch insets (fillRect) as fractions. */
export function blipBox(fill: Element): { crop: number[]; inset: number[]; tile: Element | null } {
  const frac = (el: Element | null, name: string) => clamp((int(el, name) ?? 0) / 100000, -10, 0.99);
  const src = child(fill, 'srcRect');
  const fillRect = child(child(fill, 'stretch'), 'fillRect');
  return {
    crop: ['l', 't', 'r', 'b'].map((k) => frac(src, k)),
    inset: ['l', 't', 'r', 'b'].map((k) => frac(fillRect, k)),
    tile: child(fill, 'tile'),
  };
}

/** CSS filter and opacity for the effects PowerPoint applies to pictures (transparency, grayscale …). */
function duotoneFilter(el: Element, env: Env): string | null {
  const colors = [...el.children].map((c) => colorOf(c, env.colors)).filter((c): c is Color => c !== null);
  if (colors.length < 2) return null;
  const [dark, light] = colors;
  // Gray value of the picture, then mapped onto the range between the two colors.
  const row = (a: number, b: number) => {
    const d = (b - a) / 255;
    return `${r2(d * 0.299)} ${r2(d * 0.587)} ${r2(d * 0.114)} 0 ${r2(a / 255)}`;
  };
  const filter = svg('filter', { id: uid('duo'), 'color-interpolation-filters': 'sRGB' });
  filter.append(svg('feColorMatrix', { type: 'matrix', values: `${row(dark.r, light.r)} ${row(dark.g, light.g)} ${row(dark.b, light.b)} 0 0 0 1 0` }));
  env.defs().append(filter);
  return `url(#${filter.id})`;
}

export function blipEffects(blip: Element | null, env: Env): { opacity: number; filter: string } {
  let opacity = 1;
  const filters: string[] = [];
  for (const el of blip ? [...blip.children] : []) {
    switch (el.localName) {
      case 'alphaModFix':
        opacity *= clamp((int(el, 'amt') ?? 100000) / 100000, 0, 1);
        break;
      case 'grayscl':
        filters.push('grayscale(1)');
        break;
      case 'biLevel':
        filters.push('grayscale(1)', 'contrast(10)');
        break;
      case 'duotone': {
        const duo = duotoneFilter(el, env);
        if (duo) filters.push(duo);
        break;
      }
      case 'lum': {
        const bright = (int(el, 'bright') ?? 0) / 100000, contrast = (int(el, 'contrast') ?? 0) / 100000;
        if (bright) filters.push(`brightness(${r2(clamp(1 + bright, 0, 3))})`);
        if (contrast) filters.push(`contrast(${r2(clamp(1 + contrast, 0, 3))})`);
        break;
      }
    }
  }
  return { opacity, filter: filters.join(' ') };
}

/**
 * Image parts of a blip in order of preference. PowerPoint stores its own rendering of an embedded SVG as the
 * main picture; drawing that reproduces what PowerPoint shows (its SVG renderer differs from browsers), so the
 * SVG itself is only used when that picture is missing or cannot be displayed.
 */
export function blipParts(blip: Element | null, source: Source): string[] {
  if (!blip) return [];
  const ids = [relAttr(blip, 'embed')];
  for (const ext of blip.getElementsByTagName('*')) if (ext.localName === 'svgBlip') ids.push(relAttr(ext, 'embed'));
  const parts: string[] = [];
  for (const id of ids) {
    const rel = id ? source.rels.get(id) : undefined;
    // r:link points at a file outside the package; we never load those.
    if (rel && !rel.external && !parts.includes(rel.target)) parts.push(rel.target);
  }
  return parts;
}

export async function loadBlip(env: Env, parts: string[]): Promise<Media> {
  let result: Media = { url: null, unsupported: 'unknown' };
  for (const part of parts) {
    result = await env.media.get(part);
    if (result.url) return result;
  }
  return result;
}

/**
 * PowerPoint blends gradient colors in linear light, browsers in sRGB, which makes the middle of a gradient
 * visibly darker. Extra stops computed in linear light make the browser's curve follow PowerPoint's.
 */
function smoothStops(stops: GradientStop[]): GradientStop[] {
  const out: GradientStop[] = [];
  stops.forEach((stop, i) => {
    const next = stops[i + 1];
    out.push(stop);
    if (!next || next.pos <= stop.pos) return;
    for (let k = 1; k < 8; k++) {
      const t = k / 8;
      out.push({ pos: stop.pos + (next.pos - stop.pos) * t, color: mixLinear(stop.color, next.color, t) });
    }
  });
  return out;
}

/**
 * Center of a path gradient's focus rectangle (fillToRect insets). Office themes also use "inverted"
 * rectangles such as l=50% r=100%, where PowerPoint takes the left/top inset as the point.
 */
function gradientFocus(fill: Extract<Fill, { type: 'grad' }>): [number, number] {
  const [l, t, r, b] = fill.focus;
  return [l + r >= 1 ? l : (l + 1 - r) / 2, t + b >= 1 ? t : (t + 1 - b) / 2];
}

function gradientElement(fill: Extract<Fill, { type: 'grad' }>, w: number, h: number, shade?: string): SVGGradientElement {
  let el: SVGGradientElement;
  if (fill.path) {
    const [cx, cy] = gradientFocus(fill);
    if (fill.path === 'circle') {
      const fx = cx * w, fy = cy * h;
      const radius = Math.max(Math.hypot(fx, fy), Math.hypot(w - fx, fy), Math.hypot(fx, h - fy), Math.hypot(w - fx, h - fy));
      el = svg('radialGradient', { gradientUnits: 'userSpaceOnUse', cx: r2(fx), cy: r2(fy), r: r2(radius || 1) });
    } else {
      el = svg('radialGradient', { cx: r2(cx), cy: r2(cy), r: 0.75 });
    }
  } else {
    const a = (fill.angle * Math.PI) / 180;
    const cos = Math.cos(a), sin = Math.sin(a);
    if (fill.scaled) {
      const half = (Math.abs(cos) + Math.abs(sin)) / 2;
      el = svg('linearGradient', { x1: r2(0.5 - cos * half), y1: r2(0.5 - sin * half), x2: r2(0.5 + cos * half), y2: r2(0.5 + sin * half) });
    } else {
      const half = (Math.abs(w * cos) + Math.abs(h * sin)) / 2;
      el = svg('linearGradient', {
        gradientUnits: 'userSpaceOnUse',
        x1: r2(w / 2 - cos * half), y1: r2(h / 2 - sin * half), x2: r2(w / 2 + cos * half), y2: r2(h / 2 + sin * half),
      });
    }
  }
  el.id = uid('grad');
  // For path gradients stop 0 sits at the focus point, which is also where SVG starts a radial gradient.
  for (const stop of smoothStops(fill.stops)) {
    const color = shade ? shadeFor(stop.color, shade) : stop.color;
    el.append(svg('stop', { offset: r2(stop.pos), 'stop-color': css({ ...color, a: 1 }), 'stop-opacity': r2(color.a) }));
  }
  return el;
}

/** Image pattern for a blip fill; the image itself arrives asynchronously. */
function blipPattern(fill: Extract<Fill, { type: 'blip' }>, w: number, h: number, env: Env): SVGPatternElement | null {
  const blip = child(fill.el, 'blip');
  const parts = blipParts(blip, fill.source);
  if (!parts.length) return null;
  const { crop, inset, tile } = blipBox(fill.el);
  const pattern = svg('pattern', { id: uid('img'), patternUnits: 'userSpaceOnUse', x: 0, y: 0, width: r2(Math.max(w, 1)), height: r2(Math.max(h, 1)) });
  const image = svg('image', { preserveAspectRatio: 'none' });
  const effects = blipEffects(blip, env);
  if (effects.opacity < 1) image.setAttribute('opacity', String(r2(effects.opacity)));
  if (effects.filter) image.style.filter = effects.filter;
  pattern.append(image);
  env.pending.push(
    loadBlip(env, parts).then(async (media) => {
      if (!media.url) return;
      if (tile) {
        // Tiles repeat at the picture's own size, scaled by sx/sy.
        const img = new Image();
        img.src = media.url;
        await img.decode().catch(() => undefined);
        const sx = clamp((int(tile, 'sx') ?? 100000) / 100000, 0.01, 100), sy = clamp((int(tile, 'sy') ?? 100000) / 100000, 0.01, 100);
        const tw = Math.max(2, (img.naturalWidth || 64) * sx), th = Math.max(2, (img.naturalHeight || 64) * sy);
        pattern.setAttribute('width', String(r2(tw)));
        pattern.setAttribute('height', String(r2(th)));
        pattern.setAttribute('x', String(r2(px(int(tile, 'tx') ?? 0))));
        pattern.setAttribute('y', String(r2(px(int(tile, 'ty') ?? 0))));
        Object.entries({ x: 0, y: 0, width: r2(tw), height: r2(th) }).forEach(([k, v]) => image.setAttribute(k, String(v)));
      } else {
        const bw = w * (1 - inset[0] - inset[2]), bh = h * (1 - inset[1] - inset[3]);
        const fullW = bw / Math.max(0.01, 1 - crop[0] - crop[2]), fullH = bh / Math.max(0.01, 1 - crop[1] - crop[3]);
        image.setAttribute('x', String(r2(w * inset[0] - crop[0] * fullW)));
        image.setAttribute('y', String(r2(h * inset[1] - crop[1] * fullH)));
        image.setAttribute('width', String(r2(fullW)));
        image.setAttribute('height', String(r2(fullH)));
      }
      image.setAttribute('href', media.url);
    }),
  );
  return pattern;
}

/** Paints a fill onto an SVG shape element. `defs` is created on demand in the shape's SVG. */
export function paintFill(target: SVGElement, fill: Fill | undefined, defs: () => SVGDefsElement, w: number, h: number, env: Env, shade = 'norm'): void {
  const mode = shade === 'norm' ? undefined : shade;
  switch (fill?.type) {
    case 'solid':
    case 'patt': {
      const color = fillColor(fill)!;
      const c = mode ? shadeFor(color, mode) : color;
      target.setAttribute('fill', css({ ...c, a: 1 }));
      if (c.a < 1) target.setAttribute('fill-opacity', String(r2(c.a)));
      return;
    }
    case 'grad': {
      const gradient = gradientElement(fill, w, h, mode);
      defs().append(gradient);
      target.setAttribute('fill', `url(#${gradient.id})`);
      return;
    }
    case 'blip': {
      const pattern = blipPattern(fill, w, h, env);
      if (!pattern) break;
      defs().append(pattern);
      target.setAttribute('fill', `url(#${pattern.id})`);
      return;
    }
  }
  target.setAttribute('fill', 'none');
}

export function paintStroke(target: SVGElement, line: Line | undefined, defs: () => SVGDefsElement, w: number, h: number): boolean {
  if (!line || line.fill.type === 'none' || line.fill.type === 'group') {
    target.setAttribute('stroke', 'none');
    return false;
  }
  const width = Math.max(line.width, 0.5);
  if (line.fill.type === 'grad') {
    const gradient = gradientElement(line.fill, w, h);
    defs().append(gradient);
    target.setAttribute('stroke', `url(#${gradient.id})`);
  } else {
    const color = fillColor(line.fill) ?? { r: 0, g: 0, b: 0, a: 1 };
    target.setAttribute('stroke', css({ ...color, a: 1 }));
    if (color.a < 1) target.setAttribute('stroke-opacity', String(r2(color.a)));
  }
  target.setAttribute('stroke-width', String(r2(width)));
  target.setAttribute('stroke-linecap', line.cap);
  target.setAttribute('stroke-linejoin', line.join);
  const dash = DASHES[line.dash];
  if (dash) {
    // Round and square caps add half a width to both ends of every dash; PowerPoint keeps the pattern length.
    const capExtra = line.cap === 'butt' ? 0 : 1;
    target.setAttribute('stroke-dasharray', dash.map((d, i) => r2(Math.max(0.1, (d + (i % 2 ? capExtra : -capExtra)) * width))).join(' '));
  }
  return true;
}

/** CSS background for slide backgrounds and table cells. */
export function cssFill(el: HTMLElement, fill: Fill | undefined, env: Env): void {
  switch (fill?.type) {
    case 'solid':
    case 'patt':
      el.style.backgroundColor = css(fillColor(fill)!);
      return;
    case 'grad': {
      const stops = smoothStops(fill.stops).map((s) => `${css(s.color)} ${r2(s.pos * 100)}%`).join(', ');
      if (fill.path) {
        const [x, y] = gradientFocus(fill);
        el.style.backgroundImage = `radial-gradient(circle farthest-corner at ${r2(x * 100)}% ${r2(y * 100)}%, ${stops})`;
      } else {
        el.style.backgroundImage = `linear-gradient(${r2(fill.angle + 90)}deg, ${stops})`;
      }
      return;
    }
    case 'blip': {
      const blip = child(fill.el, 'blip');
      const parts = blipParts(blip, fill.source);
      if (!parts.length) return;
      const { tile } = blipBox(fill.el);
      const effects = blipEffects(blip, env);
      // Transparency and filters only apply to the picture, so it gets a layer of its own.
      const layer = document.createElement('div');
      layer.className = 'pptx-bg-image';
      if (effects.opacity < 1) layer.style.opacity = String(r2(effects.opacity));
      if (effects.filter) layer.style.filter = effects.filter;
      layer.style.backgroundRepeat = tile ? 'repeat' : 'no-repeat';
      layer.style.backgroundSize = tile ? 'auto' : '100% 100%';
      el.append(layer);
      env.pending.push(
        loadBlip(env, parts).then((media) => {
          if (media.url) layer.style.backgroundImage = `url("${media.url}")`;
        }),
      );
      return;
    }
  }
}
