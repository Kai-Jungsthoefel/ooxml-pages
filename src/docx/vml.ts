// SPDX-License-Identifier: MIT
// VML, the drawing format of Word 97–2003 that Word still writes as a fallback and for watermarks, OLE object
// pictures and old text boxes. This part reads positions, sizes and wrapping for the layout.
import type { Anchor, AnchorPos, Inline, Wrap } from './model.js';
import type { ParseContext } from './parse.js';
import type { RunProps } from './props.js';
import { clampNum, kid, kids, wa } from './xml.js';

let vmlIds = 1_000_000_000;

/** The CSS-like style attribute of a VML element as a map. */
export function vmlStyle(el: Element): Map<string, string> {
  const map = new Map<string, string>();
  for (const part of (wa(el, 'style') ?? '').split(';')) {
    const i = part.indexOf(':');
    if (i <= 0) continue;
    map.set(part.slice(0, i).trim().toLowerCase(), part.slice(i + 1).trim());
  }
  return map;
}

/** A VML length (pt, in, cm, mm, px, emu or plain px) in CSS px. */
export function vmlLength(value: string | null | undefined, fallback = 0): number {
  if (!value) return fallback;
  const m = /^(-?\d*\.?\d+(?:e-?\d+)?)\s*(pt|in|cm|mm|px|pc|emu|em)?$/i.exec(value.trim());
  if (!m) return fallback;
  const n = Number(m[1]);
  if (!Number.isFinite(n)) return fallback;
  const unit = (m[2] ?? 'px').toLowerCase();
  const factor: Record<string, number> = { pt: 4 / 3, in: 96, cm: 96 / 2.54, mm: 96 / 25.4, px: 1, pc: 16, emu: 1 / 9525, em: 16 };
  return clampNum(n * (factor[unit] ?? 1), -100000, 100000);
}

const H_REL: Record<string, string> = {
  margin: 'margin', page: 'page', text: 'column', char: 'character', 'left-margin-area': 'leftMargin',
  'right-margin-area': 'rightMargin', 'inner-margin-area': 'insideMargin', 'outer-margin-area': 'outsideMargin',
};
const V_REL: Record<string, string> = {
  margin: 'margin', page: 'page', text: 'paragraph', line: 'line', 'top-margin-area': 'topMargin',
  'bottom-margin-area': 'bottomMargin', 'inner-margin-area': 'insideMargin', 'outer-margin-area': 'outsideMargin',
};

function position(style: Map<string, string>, axis: 'h' | 'v'): AnchorPos {
  const rel = style.get(`mso-position-${axis === 'h' ? 'horizontal' : 'vertical'}-relative`) ?? 'text';
  const align = style.get(`mso-position-${axis === 'h' ? 'horizontal' : 'vertical'}`) ?? 'absolute';
  const pos: AnchorPos = { relativeFrom: (axis === 'h' ? H_REL : V_REL)[rel] ?? (axis === 'h' ? 'column' : 'paragraph') };
  if (align !== 'absolute') pos.align = align;
  else pos.offset = vmlLength(style.get(axis === 'h' ? 'margin-left' : 'margin-top') ?? style.get(axis === 'h' ? 'left' : 'top'));
  return pos;
}

const BOOL_TRUE = new Set(['t', 'true', '1', 'on']);
export const vmlBool = (value: string | null | undefined, fallback: boolean) => (value === null || value === undefined ? fallback : BOOL_TRUE.has(value.trim().toLowerCase()));

/** Model inline for a VML shape: an inline object, or a floating one when it is positioned absolutely. */
export function parseVml(el: Element, ctx: ParseContext, run: RunProps): Inline | null {
  const style = vmlStyle(el);
  if (style.get('visibility') === 'hidden' || style.get('display') === 'none') return null;
  let w = vmlLength(style.get('width'));
  let h = vmlLength(style.get('height'));
  let lineX = 0, lineY = 0;
  if (el.localName === 'line') {
    // Lines have no box of their own; from/to give it, relative to the anchor position.
    const [x1, y1] = (wa(el, 'from') ?? '0,0').split(',').map((v) => vmlLength(v));
    const [x2, y2] = (wa(el, 'to') ?? '0,0').split(',').map((v) => vmlLength(v));
    w = Math.max(Math.abs(x2 - x1), 1);
    h = Math.max(Math.abs(y2 - y1), 1);
    lineX = Math.min(x1, x2);
    lineY = Math.min(y1, y2);
  }
  // Horizontal rules (Insert > Horizontal Line, HTML <hr>) span a share of the text width: o:hrpct in tenths
  // of a percent, or the whole width when neither it nor a width is given.
  const hrPct = Number(wa(el, 'hrpct') ?? 0);
  const hr = vmlBool(wa(el, 'hr'), false) ? (hrPct > 0 ? clampNum(hrPct / 1000, 0.01, 1) : w > 0 ? undefined : 1) : undefined;
  if (hr !== undefined && !(h > 0)) h = 2;
  if (!(w > 0) && !(h > 0)) return null;
  const obj = { id: vmlIds++, graphic: { kind: 'vml' as const, el, source: ctx.source }, w: Math.max(w, 0), h: Math.max(h, 0), effect: { t: 0, b: 0, l: 0, r: 0 }, hr };
  if (style.get('position') !== 'absolute') return { t: 'obj', obj, run };

  const z = Number(style.get('z-index') ?? '0') || 0;
  const wrapEl = kids(el).find((c) => c.localName === 'wrap');
  const type = wa(wrapEl, 'type');
  const dist = {
    t: vmlLength(style.get('mso-wrap-distance-top')),
    b: vmlLength(style.get('mso-wrap-distance-bottom')),
    l: vmlLength(style.get('mso-wrap-distance-left'), 12),
    r: vmlLength(style.get('mso-wrap-distance-right'), 12),
  };
  const sides: Record<string, string> = { both: 'bothSides', left: 'left', right: 'right', largest: 'largest' };
  const wrap: Wrap = {
    type: type === 'square' ? 'square' : type === 'tight' ? 'tight' : type === 'through' ? 'through' : type === 'topAndBottom' ? 'topAndBottom' : 'none',
    side: sides[wa(wrapEl, 'side') ?? 'both'] ?? 'bothSides',
    dist,
  };
  const posH = position(style, 'h'), posV = position(style, 'v');
  if (posH.offset !== undefined) posH.offset += lineX;
  if (posV.offset !== undefined) posV.offset += lineY;
  const anchor: Anchor = {
    posH,
    posV,
    behindDoc: z < 0,
    // VML z-index values are the same numbers DrawingML calls relativeHeight.
    z: Math.abs(z),
    wrap,
    layoutInCell: true,
    allowOverlap: true,
  };
  const pctW = Number(style.get('mso-width-percent') ?? 0), pctH = Number(style.get('mso-height-percent') ?? 0);
  const relName = (v: string | undefined) => (v === 'page' ? 'page' : v === 'left-margin-area' ? 'leftMargin' : v === 'right-margin-area' ? 'rightMargin' : 'margin');
  if (pctW > 0) anchor.sizeRelH = [relName(style.get('mso-width-relative')), pctW / 1000];
  if (pctH > 0) anchor.sizeRelV = [relName(style.get('mso-height-relative')), pctH / 1000];
  return { t: 'float', obj: { ...obj, anchor }, run };
}

/** The w:txbxContent of a VML text box, if the shape has one. */
export function vmlTextbox(el: Element): Element | null {
  const box = kid(el, 'textbox');
  return box ? kid(box, 'txbxContent') : null;
}
