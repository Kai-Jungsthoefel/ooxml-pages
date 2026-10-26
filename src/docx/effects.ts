// SPDX-License-Identifier: MIT
// Word 2010 text effects (w14:textFill, w14:textOutline, w14:shadow) as CSS. Their colors are DrawingML colors,
// but with the attributes in the w14 namespace (w14:val), so they are copied into plain DrawingML first.
import { type ColorContext, colorIn, css } from '../pptx/color.js';
import type { RunProps } from './props.js';
import { clampNum, kids, wi } from './xml.js';

let scratch: XMLDocument | null = null;

/** Copy of `el` with every attribute under its local name, as the DrawingML readers expect. */
function plain(el: Element): Element {
  scratch ??= document.implementation.createDocument(null, null, null);
  const copy = scratch.createElement(el.localName);
  for (const a of el.attributes) copy.setAttribute(a.localName, a.value);
  for (const c of el.children) copy.append(plain(c));
  return copy;
}

function colorOfFill(container: Element | null | undefined, ctx: ColorContext): string | null {
  if (!container) return null;
  const c = colorIn(plain(container), ctx);
  return c ? css(c) : null;
}

/** EMU to CSS px. */
const emu = (v: number | undefined) => clampNum(v ?? 0, 0, 5_000_000) / 9525;

export interface TextEffects {
  /** Text color (also 'transparent' for w14:noFill). */
  color?: string;
  /** Gradient painted into the glyphs (background-clip: text). */
  gradient?: string;
  /** -webkit-text-stroke value. */
  stroke?: string;
  /** Shadow as "dx dy blur color". */
  shadow?: string;
}

export function textEffects(run: RunProps, ctx: ColorContext): TextEffects {
  const fx: TextEffects = {};
  const fill = kids(run.textFill)[0];
  if (fill?.localName === 'noFill') fx.color = 'transparent';
  else if (fill?.localName === 'solidFill') fx.color = colorOfFill(fill, ctx) ?? undefined;
  else if (fill?.localName === 'gradFill') {
    const stops = kids(kids(fill).find((c) => c.localName === 'gsLst'))
      .map((gs) => ({ pos: clampNum((wi(gs, 'pos') ?? 0) / 1000, 0, 100), color: colorOfFill(gs, ctx) }))
      .filter((s): s is { pos: number; color: string } => s.color !== null)
      .sort((a, b) => a.pos - b.pos);
    const lin = kids(fill).find((c) => c.localName === 'lin');
    // DrawingML measures from "to the right", clockwise; CSS from "to the top".
    const angle = ((wi(lin, 'ang') ?? 0) / 60000 + 90) % 360;
    if (stops.length === 1) fx.color = stops[0].color;
    else if (stops.length > 1) {
      const list = stops.map((s) => `${s.color} ${Math.round(s.pos * 10) / 10}%`).join(', ');
      fx.gradient = kids(fill).some((c) => c.localName === 'path') ? `radial-gradient(${list})` : `linear-gradient(${angle}deg, ${list})`;
    }
  }
  const outline = run.textOutline;
  if (outline) {
    const paint = kids(outline).find((c) => c.localName === 'solidFill' || c.localName === 'gradFill');
    // A gradient outline is drawn in the color of its first stop.
    const firstStop = kids(kids(paint).find((c) => c.localName === 'gsLst'))[0];
    const color = paint?.localName === 'solidFill' ? colorOfFill(paint, ctx) : colorOfFill(firstStop, ctx);
    if (color) fx.stroke = `${Math.max(0.3, Math.round(emu(wi(outline, 'w') ?? 9525) * 100) / 100)}px ${color}`;
  }
  const shadow = run.textShadow;
  if (shadow) {
    const color = colorOfFill(shadow, ctx);
    if (color) {
      const dist = emu(wi(shadow, 'dist')), dir = ((wi(shadow, 'dir') ?? 0) / 60000) * (Math.PI / 180);
      const r = (v: number) => Math.round(v * 100) / 100;
      // CSS blur radius is twice the standard deviation; DrawingML's blurRad is closer to the CSS value halved.
      fx.shadow = `${r(Math.cos(dir) * dist)}px ${r(Math.sin(dir) * dist)}px ${r(emu(wi(shadow, 'blurRad')) / 2)}px ${color}`;
    }
  }
  return fx;
}
