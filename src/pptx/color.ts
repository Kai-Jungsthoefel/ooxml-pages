// SPDX-License-Identifier: MIT
// DrawingML colors: the six color models, the theme scheme with its color map, and the transforms
// (lumMod, tint, alpha …) that PowerPoint applies on top.
import { attr, child, clamp, int } from './xml.js';

/** r, g, b in 0–255 (fractional), a in 0–1. */
export interface Color {
  r: number;
  g: number;
  b: number;
  a: number;
}

export interface FontSet {
  latin: string;
  ea: string;
  cs: string;
}

export interface Theme {
  colors: Map<string, Color>;
  major: FontSet;
  minor: FontSet;
  /** fmtScheme lists, indexed by the idx of fillRef/lnRef/bgRef (1-based). */
  fills: Element[];
  lines: Element[];
  bgFills: Element[];
}

export interface ColorContext {
  theme: Theme;
  /** Slide color map: bg1 → lt1, tx1 → dk1 … */
  clrMap: Map<string, string>;
  /** Placeholder color of a style reference (fillRef, lnRef …). */
  phClr?: Color;
}

export const BLACK: Color = { r: 0, g: 0, b: 0, a: 1 };

const DEFAULT_MAP: Record<string, string> = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };

// The preset colors PowerPoint actually writes; the full list mirrors CSS names.
const PRESET: Record<string, string> = {
  black: '000000', white: 'ffffff', red: 'ff0000', green: '008000', blue: '0000ff', yellow: 'ffff00',
  cyan: '00ffff', magenta: 'ff00ff', gray: '808080', grey: '808080', silver: 'c0c0c0', maroon: '800000',
  navy: '000080', olive: '808000', purple: '800080', teal: '008080', lime: '00ff00', orange: 'ffa500',
  pink: 'ffc0cb', brown: 'a52a2a', gold: 'ffd700', darkgray: 'a9a9a9', darkgrey: 'a9a9a9', lightgray: 'd3d3d3',
  lightgrey: 'd3d3d3', darkblue: '00008b', darkred: '8b0000', darkgreen: '006400', lightblue: 'add8e6',
  lightgreen: '90ee90', lightyellow: 'ffffe0', dimgray: '696969', whitesmoke: 'f5f5f5', beige: 'f5f5dc',
  coral: 'ff7f50', crimson: 'dc143c', indigo: '4b0082', violet: 'ee82ee', tomato: 'ff6347', skyblue: '87ceeb',
  steelblue: '4682b4', royalblue: '4169e1', tan: 'd2b48c', salmon: 'fa8072', khaki: 'f0e68c', orchid: 'da70d6',
};

const SYSTEM: Record<string, string> = {
  windowtext: '000000', window: 'ffffff', btnface: 'f0f0f0', btntext: '000000', highlight: '0078d7',
  highlighttext: 'ffffff', graytext: '6d6d6d', menu: 'f0f0f0', menutext: '000000', infobk: 'ffffe1',
  infotext: '000000', '3ddkshadow': '696969', '3dlight': 'e3e3e3', btnshadow: 'a0a0a0', btnhighlight: 'ffffff',
};

export function hex(value: string | null | undefined): Color | null {
  if (!value || !/^[0-9a-f]{6}$/i.test(value)) return null;
  const n = parseInt(value, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: 1 };
}

const toLinear = (c: number) => {
  const v = c / 255;
  return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
};
const fromLinear = (v: number) => {
  const c = v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055;
  return clamp(c, 0, 1) * 255;
};

function toHsl({ r, g, b }: Color): [number, number, number] {
  const R = r / 255, G = g / 255, B = b / 255;
  const max = Math.max(R, G, B), min = Math.min(R, G, B);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = max === R ? (G - B) / d + (G < B ? 6 : 0) : max === G ? (B - R) / d + 2 : (R - G) / d + 4;
  h *= 60;
  return [h, s, l];
}

function fromHsl(h: number, s: number, l: number, a: number): Color {
  h = ((h % 360) + 360) % 360;
  s = clamp(s, 0, 1);
  l = clamp(l, 0, 1);
  const c = (1 - Math.abs(2 * l - 1)) * s;
  const x = c * (1 - Math.abs(((h / 60) % 2) - 1));
  const m = l - c / 2;
  const [r, g, b] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
  return { r: (r + m) * 255, g: (g + m) * 255, b: (b + m) * 255, a };
}

/** Applies the transform children of a color element in document order. */
function transform(base: Color, el: Element): Color {
  let c = { ...base };
  for (const mod of [...el.children]) {
    const v = (int(mod, 'val') ?? 0) / 100000;
    switch (mod.localName) {
      case 'alpha':
        c.a = clamp(v, 0, 1);
        break;
      case 'alphaMod':
        c.a = clamp(c.a * v, 0, 1);
        break;
      case 'alphaOff':
        c.a = clamp(c.a + v, 0, 1);
        break;
      case 'tint': {
        // Mixes with white in linear RGB, which is how PowerPoint gets its fairly light tints.
        const f = clamp(v, 0, 1);
        const t = (x: number) => fromLinear(1 - (1 - toLinear(x)) * f);
        c = { r: t(c.r), g: t(c.g), b: t(c.b), a: c.a };
        break;
      }
      case 'shade': {
        const f = clamp(v, 0, 1);
        const s = (x: number) => fromLinear(toLinear(x) * f);
        c = { r: s(c.r), g: s(c.g), b: s(c.b), a: c.a };
        break;
      }
      case 'lumMod':
      case 'lumOff':
      case 'lum':
      case 'satMod':
      case 'satOff':
      case 'sat':
      case 'hueMod':
      case 'hueOff':
      case 'hue': {
        let [h, s, l] = toHsl(c);
        const name = mod.localName;
        if (name === 'lumMod') l *= v;
        else if (name === 'lumOff') l += v;
        else if (name === 'lum') l = v;
        else if (name === 'satMod') s *= v;
        else if (name === 'satOff') s += v;
        else if (name === 'sat') s = v;
        else if (name === 'hueMod') h *= v;
        // Hue values are in 60000ths of a degree, not percent.
        else if (name === 'hueOff') h += (v * 100000) / 60000;
        else h = (v * 100000) / 60000;
        c = fromHsl(h, s, l, c.a);
        break;
      }
      case 'comp': {
        const [h, s, l] = toHsl(c);
        c = fromHsl(h + 180, s, l, c.a);
        break;
      }
      case 'inv':
        c = { r: 255 - c.r, g: 255 - c.g, b: 255 - c.b, a: c.a };
        break;
      case 'gray': {
        const y = 0.299 * c.r + 0.587 * c.g + 0.114 * c.b;
        c = { r: y, g: y, b: y, a: c.a };
        break;
      }
      case 'red':
        c.r = clamp(v, 0, 1) * 255;
        break;
      case 'green':
        c.g = clamp(v, 0, 1) * 255;
        break;
      case 'blue':
        c.b = clamp(v, 0, 1) * 255;
        break;
      case 'redMod':
        c.r = clamp(c.r * v, 0, 255);
        break;
      case 'greenMod':
        c.g = clamp(c.g * v, 0, 255);
        break;
      case 'blueMod':
        c.b = clamp(c.b * v, 0, 255);
        break;
      case 'redOff':
        c.r = clamp(c.r + v * 255, 0, 255);
        break;
      case 'greenOff':
        c.g = clamp(c.g + v * 255, 0, 255);
        break;
      case 'blueOff':
        c.b = clamp(c.b + v * 255, 0, 255);
        break;
      case 'gamma':
        c = { r: fromLinear(c.r / 255), g: fromLinear(c.g / 255), b: fromLinear(c.b / 255), a: c.a };
        break;
      case 'invGamma':
        c = { r: toLinear(c.r) * 255, g: toLinear(c.g) * 255, b: toLinear(c.b) * 255, a: c.a };
        break;
    }
  }
  return c;
}

function schemeColor(name: string, ctx: ColorContext): Color {
  if (name === 'phClr') return ctx.phClr ?? BLACK;
  const mapped = ctx.clrMap.get(name) ?? DEFAULT_MAP[name] ?? name;
  return ctx.theme.colors.get(mapped) ?? ctx.theme.colors.get(name) ?? BLACK;
}

const COLOR_ELEMENTS = new Set(['srgbClr', 'schemeClr', 'scrgbClr', 'hslClr', 'sysClr', 'prstClr']);

/** Base color of a single color element (srgbClr, schemeClr …) with its transforms applied. */
export function colorOf(el: Element, ctx: ColorContext | null): Color | null {
  let base: Color | null = null;
  switch (el.localName) {
    case 'srgbClr':
      base = hex(attr(el, 'val'));
      break;
    case 'scrgbClr': {
      const pct = (name: string) => fromLinear(clamp((int(el, name) ?? 0) / 100000, 0, 1));
      base = { r: pct('r'), g: pct('g'), b: pct('b'), a: 1 };
      break;
    }
    case 'hslClr':
      base = fromHsl((int(el, 'hue') ?? 0) / 60000, (int(el, 'sat') ?? 0) / 100000, (int(el, 'lum') ?? 0) / 100000, 1);
      break;
    case 'sysClr':
      base = hex(attr(el, 'lastClr')) ?? hex(SYSTEM[(attr(el, 'val') ?? '').toLowerCase()]) ?? BLACK;
      break;
    case 'prstClr': {
      const name = (attr(el, 'val') ?? '').toLowerCase().replace(/^dk/, 'dark').replace(/^lt/, 'light').replace(/^med/, 'medium');
      base = hex(PRESET[name]) ?? BLACK;
      break;
    }
    case 'schemeClr':
      base = ctx ? schemeColor(attr(el, 'val') ?? '', ctx) : BLACK;
      break;
  }
  return base ? transform(base, el) : null;
}

/** The color held by a container such as <a:solidFill> or <a:buClr>: its first color child. */
export function colorIn(container: Element | null | undefined, ctx: ColorContext | null): Color | null {
  if (!container) return null;
  for (const el of container.children) if (COLOR_ELEMENTS.has(el.localName)) return colorOf(el, ctx);
  return null;
}

export function css(c: Color): string {
  const r = Math.round(clamp(c.r, 0, 255)), g = Math.round(clamp(c.g, 0, 255)), b = Math.round(clamp(c.b, 0, 255));
  return c.a >= 1 ? `rgb(${r}, ${g}, ${b})` : `rgba(${r}, ${g}, ${b}, ${Math.round(clamp(c.a, 0, 1) * 1000) / 1000})`;
}

/** Mix in linear light, the way PowerPoint blends gradient colors. */
export function mixLinear(a: Color, b: Color, t: number): Color {
  const m = (x: number, y: number) => fromLinear(toLinear(x) + (toLinear(y) - toLinear(x)) * t);
  return { r: m(a.r, b.r), g: m(a.g, b.g), b: m(a.b, b.b), a: a.a + (b.a - a.a) * t };
}

export function mix(a: Color, b: Color, t: number): Color {
  return { r: a.r + (b.r - a.r) * t, g: a.g + (b.g - a.g) * t, b: a.b + (b.b - a.b) * t, a: a.a + (b.a - a.a) * t };
}

/**
 * The fill modes of preset paths (darken, lighten …) used by 3D-looking shapes like cans and cubes. Unlike the
 * tint and shade transforms these mix in plain sRGB, which is what PowerPoint's rendering shows.
 */
export function shadeFor(c: Color, mode: string): Color {
  const toWhite = (f: number) => ({ r: c.r + (255 - c.r) * f, g: c.g + (255 - c.g) * f, b: c.b + (255 - c.b) * f, a: c.a });
  const toBlack = (f: number) => ({ r: c.r * f, g: c.g * f, b: c.b * f, a: c.a });
  switch (mode) {
    case 'darken':
      return toBlack(0.6);
    case 'darkenLess':
      return toBlack(0.8);
    case 'lighten':
      return toWhite(0.4);
    case 'lightenLess':
      return toWhite(0.2);
    default:
      return c;
  }
}

// ---- Theme ---------------------------------------------------------------------------------------------------

const EMPTY_FONTS: FontSet = { latin: 'Calibri', ea: '', cs: '' };

function fontSet(el: Element | null, fallback: FontSet): FontSet {
  if (!el) return fallback;
  return {
    latin: attr(child(el, 'latin'), 'typeface') || fallback.latin,
    ea: attr(child(el, 'ea'), 'typeface') ?? '',
    cs: attr(child(el, 'cs'), 'typeface') ?? '',
  };
}

export function parseColorScheme(scheme: Element | null, into: Map<string, Color>): void {
  for (const el of scheme ? [...scheme.children] : []) {
    const color = colorIn(el, null);
    if (color) into.set(el.localName, color);
  }
}

export function parseTheme(doc: Document | null): Theme {
  const root = doc?.documentElement ?? null;
  const elements = child(root, 'themeElements');
  const colors = new Map<string, Color>();
  parseColorScheme(child(elements, 'clrScheme'), colors);
  const fonts = child(elements, 'fontScheme');
  const fmt = child(elements, 'fmtScheme');
  return {
    colors,
    major: fontSet(child(fonts, 'majorFont'), EMPTY_FONTS),
    minor: fontSet(child(fonts, 'minorFont'), EMPTY_FONTS),
    fills: [...(child(fmt, 'fillStyleLst')?.children ?? [])],
    lines: [...(child(fmt, 'lnStyleLst')?.children ?? [])],
    bgFills: [...(child(fmt, 'bgFillStyleLst')?.children ?? [])],
  };
}

/** A theme override (charts, slides) replaces the color scheme but keeps fonts and formats. */
export function overrideTheme(theme: Theme, doc: Document | null): Theme {
  const scheme = doc ? child(doc.documentElement, 'clrScheme') : null;
  if (!scheme) return theme;
  const colors = new Map(theme.colors);
  parseColorScheme(scheme, colors);
  return { ...theme, colors };
}

/** p:clrMap / a:overrideClrMapping attributes. */
export function parseColorMap(el: Element | null, base?: Map<string, string>): Map<string, string> {
  const map = new Map(base ?? Object.entries(DEFAULT_MAP));
  if (el) for (const a of el.attributes) map.set(a.localName, a.value);
  return map;
}
