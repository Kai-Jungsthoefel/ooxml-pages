// SPDX-License-Identifier: MIT
// Fonts and text measurement. Word lays out lines with the font's own metrics: single line spacing is
// ascent + descent + line gap, which is also what the browser uses for line-height: normal. Widths come from a
// canvas, which resolves fonts exactly like the DOM text drawn later.
import type { Theme } from '../pptx/color.js';
import { fontFamily, symbolFamily } from '../pptx/fonts.js';
import type { RunProps } from './props.js';
import { fullWidthBoxFont } from './symbols.js';
import { halfPt } from './xml.js';

export interface FontInfo {
  /** CSS font-family list. */
  family: string;
  /** The Word font name the run asks for. */
  name: string;
  /** Font size in px (before super/subscript reduction). */
  size: number;
  bold: boolean;
  italic: boolean;
  /** Canvas font shorthand. */
  css: string;
  kern: boolean;
  ligatures: boolean;
  /** Per em. */
  ascent: number;
  descent: number;
  gap: number;
}

export type Script = 'ascii' | 'hAnsi' | 'eastAsia' | 'cs';

let context: CanvasRenderingContext2D | null | undefined;

function ctx(): CanvasRenderingContext2D | null {
  if (context === undefined) context = document.createElement('canvas').getContext('2d');
  return context;
}

const metricsCache = new Map<string, { ascent: number; descent: number; gap: number }>();

/** Ascent, descent and line gap per em of the font the browser picks for a family list. */
export function fontMetrics(family: string, bold: boolean, italic: boolean): { ascent: number; descent: number; gap: number } {
  const key = `${bold ? 'b' : ''}${italic ? 'i' : ''}|${family}`;
  let m = metricsCache.get(key);
  if (m) return m;
  m = { ascent: 0.9, descent: 0.22, gap: 0.03 };
  const c = ctx();
  const style = `${italic ? 'italic ' : ''}${bold ? '700 ' : ''}`;
  if (c) {
    c.font = `${style}1000px ${family}`;
    const t = c.measureText('Hg');
    if (t.fontBoundingBoxAscent > 0) {
      const div = document.createElement('div');
      div.style.cssText = `position:absolute;left:-99999px;top:0;visibility:hidden;white-space:nowrap;line-height:normal;font:${style}1000px ${family}`;
      div.textContent = 'Hg';
      document.body.append(div);
      const normal = div.getBoundingClientRect().height;
      div.remove();
      const ascent = t.fontBoundingBoxAscent / 1000, descent = t.fontBoundingBoxDescent / 1000;
      m = { ascent, descent, gap: Math.max(0, normal / 1000 - ascent - descent) };
    }
  }
  metricsCache.set(key, m);
  return m;
}

const widthCache = new Map<string, Map<string, number>>();
let widthEntries = 0;

/** Advance width of a text in px (character spacing and scaling not included). */
export function measure(text: string, font: FontInfo, size = font.size): number {
  if (!text) return 0;
  const css = size === font.size ? font.css : `${font.italic ? 'italic ' : ''}${font.bold ? '700 ' : ''}${size}px ${font.family}`;
  const key = `${font.kern ? 'k' : ''}${font.ligatures ? 'l' : ''}|${css}`;
  let cache = widthCache.get(key);
  if (!cache) {
    cache = new Map();
    widthCache.set(key, cache);
  }
  let w = cache.get(text);
  if (w !== undefined) return w;
  const c = ctx();
  if (!c) return text.length * size * 0.5;
  c.font = css;
  c.fontKerning = font.kern ? 'normal' : 'none';
  c.textRendering = font.kern || font.ligatures ? 'auto' : 'optimizeSpeed';
  w = c.measureText(text).width;
  // A runaway cache would only cost memory; documents rarely have more distinct words than this.
  if (widthEntries++ > 400_000) {
    widthCache.clear();
    widthEntries = 0;
  }
  cache.set(text, w);
  return w;
}

// ---- Fonts of a run -------------------------------------------------------------------------------------------

export function scriptOf(code: number): Script {
  if (code < 0x80) return 'ascii';
  if (
    (code >= 0x0590 && code <= 0x08ff) || // Hebrew, Arabic, Syriac, Thaana …
    (code >= 0x0900 && code <= 0x0dff) || // Indic scripts
    (code >= 0x0e00 && code <= 0x0eff) || // Thai, Lao
    (code >= 0xfb1d && code <= 0xfdff) ||
    (code >= 0xfe70 && code <= 0xfeff)
  ) return 'cs';
  if (
    (code >= 0x1100 && code <= 0x11ff) ||
    (code >= 0x2e80 && code <= 0x9fff) ||
    (code >= 0xa000 && code <= 0xa4cf) ||
    (code >= 0xac00 && code <= 0xd7af) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xfe30 && code <= 0xfe4f) ||
    (code >= 0xff00 && code <= 0xffef) ||
    (code >= 0x20000 && code <= 0x2fa1f)
  ) return 'eastAsia';
  return 'hAnsi';
}

function themeFont(name: string | undefined, theme: Theme): string | undefined {
  if (!name) return undefined;
  const major = name.startsWith('major');
  const set = major ? theme.major : theme.minor;
  if (/EastAsia$/.test(name)) return set.ea || set.latin;
  if (/Bidi$/.test(name)) return set.cs || set.latin;
  return set.latin;
}

/** The Word font name a run uses for a script. */
export function fontName(run: RunProps, script: Script, theme: Theme): string {
  const f = run.fonts;
  const pick = (slot: Script): string | undefined => themeFont(f[`${slot}Theme`], theme) ?? f[slot];
  let name: string | undefined;
  if (run.rtl || run.cs || script === 'cs') name = pick('cs');
  else if (script === 'eastAsia') name = pick('eastAsia');
  else if (script === 'ascii') name = pick('ascii');
  else name = f.hint === 'eastAsia' && pick('eastAsia') ? pick('eastAsia') : pick('hAnsi');
  return name ?? pick('ascii') ?? pick('hAnsi') ?? 'Times New Roman';
}

const fontCache = new Map<string, FontInfo>();

/** Font of a run for a script; `override` is a CSS family list that replaces the run's fonts. */
export function runFont(run: RunProps, script: Script, theme: Theme, override?: string): FontInfo {
  const complex = run.rtl || run.cs || script === 'cs';
  const name = override ? 'Segoe UI Symbol' : fontName(run, script, theme);
  const halfPoints = (complex ? (run.szCs ?? run.sz) : run.sz) ?? 20;
  const bold = !!(complex ? (run.bCs ?? run.b) : run.b);
  const italic = !!(complex ? (run.iCs ?? run.i) : run.i);
  // Word kerns from the given size on; ligatures only when the run asks for them.
  const size = halfPt(halfPoints);
  const kern = !!run.kern && run.kern > 0 && halfPoints >= run.kern;
  const ligatures = !!run.ligatures && run.ligatures !== 'none';
  const key = `${name}|${override ?? ''}|${size}|${bold}|${italic}|${kern}|${ligatures}`;
  let info = fontCache.get(key);
  if (info) return info;
  const family = override ? symbolFamily(override) : fontFamily(name, script === 'eastAsia' ? undefined : fontName(run, 'eastAsia', theme));
  const m = fontMetrics(family, bold, italic);
  info = {
    family,
    name,
    size,
    bold,
    italic,
    css: `${italic ? 'italic ' : ''}${bold ? '700 ' : ''}${size}px ${family}`,
    kern,
    ligatures,
    ascent: m.ascent,
    descent: m.descent,
    gap: m.gap,
  };
  fontCache.set(key, info);
  return info;
}

const installed = new Map<string, boolean>();
let installedHook: ((name: string) => boolean) | null = null;

/**
 * While a PDF is made (src/pdf/fonts.ts), a font counts as installed only when its file goes into the PDF, so that
 * symbol fonts on this computer are translated to Unicode as on any other.
 */
export function setInstalledHook(hook: ((name: string) => boolean) | null): void {
  installedHook = hook;
  installed.clear();
}

/** Whether a font is installed: text in it measures differently from the generic fallbacks. */
export function fontInstalled(name: string): boolean {
  if (installedHook) return installedHook(name);
  const key = name.toLowerCase();
  let ok = installed.get(key);
  if (ok !== undefined) return ok;
  ok = false;
  const c = ctx();
  if (c) {
    const sample = 'mmmmmmmmmmlli10OWy\uf0fc\uf0a8';
    const clean = name.replace(/["'\\;{}<>]/g, '');
    for (const generic of ['monospace', 'serif', 'sans-serif']) {
      c.font = `72px ${generic}`;
      const base = c.measureText(sample).width;
      c.font = `72px "${clean}", ${generic}`;
      if (Math.abs(c.measureText(sample).width - base) > 0.5) {
        ok = true;
        break;
      }
    }
  }
  installed.set(key, ok);
  return ok;
}

/** Families for symbol characters that were translated to Unicode. */
export const SYMBOL_FAMILY = '"Segoe UI Symbol", "Segoe UI Emoji", "Apple Symbols", "Noto Sans Symbols 2", "Noto Sans Symbols", "DejaVu Sans", sans-serif';

/** Width Word gives a ballot box glyph in a font. */
export function boxWidth(font: FontInfo): number {
  if (fullWidthBoxFont(font.name)) return font.size;
  return Math.max(measure('☐', font), font.size * 0.8);
}

const baselineCache = new Map<string, number>();

/** Forgets everything measured, after fonts were added to or removed from the page (embedded fonts). */
export function resetFontCaches(): void {
  metricsCache.clear();
  widthCache.clear();
  widthEntries = 0;
  fontCache.clear();
  installed.clear();
  baselineCache.clear();
}

/** Where the browser puts the baseline inside a line box of line-height: normal, in px from its top. */
export function domBaseline(css: string): number {
  let b = baselineCache.get(css);
  if (b !== undefined) return b;
  const div = document.createElement('div');
  div.style.cssText = `position:absolute;left:-99999px;top:0;visibility:hidden;white-space:nowrap;line-height:normal;font:${css}`;
  const marker = document.createElement('span');
  marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline';
  div.append('Hg', marker);
  document.body.append(div);
  b = marker.getBoundingClientRect().top - div.getBoundingClientRect().top;
  div.remove();
  baselineCache.set(css, b);
  return b;
}
