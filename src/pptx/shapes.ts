// SPDX-License-Identifier: MIT
// The shape tree: shapes and connectors (SVG geometry plus HTML text), pictures, groups and graphic frames
// (tables, charts, SmartArt drawings, OLE fallbacks). Positions are absolute within the parent element.
import { h } from '../dom.js';
import { colorIn, css } from './color.js';
import { type Env, type Source, emuAttr, r2, svg } from './env.js';
import { type Fill, type Line, blipBox, blipEffects, blipParts, fillColor, loadBlip, ownFill, paintFill, paintStroke, parseLine, refFill, refLine } from './fill.js';
import { type Cmd, type Geometry, geometryOf, pathEnd, svgPath } from './geometry.js';
import { type Layer, type BodyProps, defaultBody, mergeBodyPr, renderTextBox, type RunProps } from './text.js';
import { attr, bool, child, elements, int, path } from './xml.js';

/** Maps EMU coordinates of a shape tree into pixels of its container. */
export interface Frame {
  chX: number;
  chY: number;
  /** Pixels per EMU. */
  sx: number;
  sy: number;
  /** Container size in px, for mirroring children of flipped groups. */
  w: number;
  h: number;
  flipH: boolean;
  flipV: boolean;
}

export interface Tree {
  env: Env;
  source: Source;
  /** Shapes a placeholder inherits from, nearest first (layout, then master). Slides only. */
  inherit?: (ph: Element) => Element[];
  /** Inherited text formatting for a shape (null for table cells, chart text …). */
  layers: (sp: Element | null, chain: Element[]) => Layer[];
  /** Placeholders of masters and layouts are prompts for the slide, not content. */
  skipPlaceholders: boolean;
  depth: number;
  groupFill?: Fill;
  /** Hook for graphic frames (tables, charts, diagrams); implemented by the slide renderer. */
  graphicFrame?: (frame: Element, box: Box, tree: Tree) => Promise<HTMLElement | null>;
  /**
   * Text of shapes without a DrawingML text body (Word text boxes, wps:txbx): gets the text rectangle in the
   * shape's coordinates and returns the positioned content.
   */
  text?: (sp: Element, rect: { x: number; y: number; w: number; h: number; rot: number }, box: Box, tree: Tree) => HTMLElement | null;
}

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
  rot: number;
  flipH: boolean;
  flipV: boolean;
  /** Size in EMU, for custom geometry. */
  emuW: number;
  emuH: number;
}

// Limits against hostile files: group nesting, and shape sizes in px (about 40 slide widths).
const MAX_DEPTH = 24;
const MAX_SIZE = 50000;

export function rootFrame(w: number, h: number): Frame {
  return { chX: 0, chY: 0, sx: 1 / 9525, sy: 1 / 9525, w, h, flipH: false, flipV: false };
}

/** Position of an xfrm in the frame of its container. */
export function place(xfrm: Element | null, frame: Frame): Box | null {
  if (!xfrm) return null;
  const off = child(xfrm, 'off'), ext = child(xfrm, 'ext');
  const emuW = Math.max(0, emuAttr(ext, 'cx')), emuH = Math.max(0, emuAttr(ext, 'cy'));
  let x = (emuAttr(off, 'x') - frame.chX) * frame.sx;
  let y = (emuAttr(off, 'y') - frame.chY) * frame.sy;
  const w = Math.min(emuW * frame.sx, MAX_SIZE), h = Math.min(emuH * frame.sy, MAX_SIZE);
  let rot = (int(xfrm, 'rot') ?? 0) / 60000;
  let flipH = bool(xfrm, 'flipH') ?? false, flipV = bool(xfrm, 'flipV') ?? false;
  if (frame.flipH) {
    x = frame.w - x - w;
    flipH = !flipH;
    rot = -rot;
  }
  if (frame.flipV) {
    y = frame.h - y - h;
    flipV = !flipV;
    rot = -rot;
  }
  if (!Number.isFinite(x + y + w + h + rot)) return null;
  return { x, y, w, h, rot, flipH, flipV, emuW, emuH };
}

function positioned(cls: string, box: Box): HTMLElement {
  const el = h('div', { class: cls });
  const s = el.style;
  s.left = `${r2(box.x)}px`;
  s.top = `${r2(box.y)}px`;
  s.width = `${r2(box.w)}px`;
  s.height = `${r2(box.h)}px`;
  if (box.rot) s.transform = `rotate(${r2(box.rot)}deg)`;
  return el;
}

function flipTransform(box: Box): string {
  return box.flipH || box.flipV ? `scale(${box.flipH ? -1 : 1}, ${box.flipV ? -1 : 1})` : '';
}

// ---- Geometry -------------------------------------------------------------------------------------------------

/** Moves the end of a path back so a thick line does not poke through the tip of its arrowhead. */
function shorten(cmds: Cmd[], atStart: boolean, by: number): void {
  const drawn = cmds.filter((c) => c.t !== 'Z') as Exclude<Cmd, { t: 'Z' }>[];
  if (drawn.length < 2 || by <= 0) return;
  const [p, q] = atStart ? [drawn[0], drawn[1]] : [drawn[drawn.length - 1], drawn[drawn.length - 2]];
  const len = Math.hypot(p.x - q.x, p.y - q.y);
  // Only straight end segments; moving the end of a curve would bend it.
  if (len < 1e-6 || (atStart ? q : p).t === 'C') return;
  const k = Math.min(by, len * 0.9) / len;
  p.x += (q.x - p.x) * k;
  p.y += (q.y - p.y) * k;
}

function arrowHead(kind: string, at: { x: number; y: number; dx: number; dy: number }, width: number, length: number, color: string, lineWidth: number): SVGElement {
  const { x, y, dx, dy } = at;
  const nx = -dy, ny = dx;
  const pt = (along: number, across: number) => `${r2(x + dx * along + nx * across)},${r2(y + dy * along + ny * across)}`;
  switch (kind) {
    case 'oval':
      return svg('ellipse', {
        cx: r2(x), cy: r2(y), rx: r2(length / 2), ry: r2(width / 2), fill: color,
        transform: `rotate(${r2((Math.atan2(dy, dx) * 180) / Math.PI)} ${r2(x)} ${r2(y)})`,
      });
    case 'diamond':
      return svg('polygon', { points: [pt(length / 2, 0), pt(0, width / 2), pt(-length / 2, 0), pt(0, -width / 2)].join(' '), fill: color });
    case 'arrow':
      return svg('polyline', {
        points: [pt(-length, width / 2), pt(0, 0), pt(-length, -width / 2)].join(' '),
        fill: 'none', stroke: color, 'stroke-width': r2(lineWidth), 'stroke-linejoin': 'miter', 'stroke-linecap': 'round',
      });
    case 'stealth':
      return svg('polygon', { points: [pt(0, 0), pt(-length, width / 2), pt(-length * 0.6, 0), pt(-length, -width / 2)].join(' '), fill: color });
    default:
      return svg('polygon', { points: [pt(0, 0), pt(-length, width / 2), pt(-length, -width / 2)].join(' '), fill: color });
  }
}

/** SVG for the outline and fill of a shape, arrowheads included. */
export function drawGeometry(geom: Geometry, fill: Fill | undefined, line: Line | undefined, box: Box, env: Env): SVGSVGElement | null {
  const visibleLine = line && line.fill.type !== 'none';
  const visibleFill = fill && fill.type !== 'none' && fill.type !== 'group' && !geom.open;
  if (!visibleLine && !visibleFill) return null;
  const el = svg('svg', { class: 'pptx-geom', width: r2(Math.max(box.w, 1)), height: r2(Math.max(box.h, 1)) });
  const flip = flipTransform(box);
  if (flip) el.style.transform = flip;
  let defs: SVGDefsElement | null = null;
  const getDefs = () => (defs ??= el.insertBefore(svg('defs'), el.firstChild));
  const lineWidth = Math.max(line?.width ?? 1, 0.5);
  const lineColor = line ? css({ ...(fillColor(line.fill) ?? { r: 0, g: 0, b: 0, a: 1 }), a: 1 }) : 'none';
  // PowerPoint keeps arrowheads at a usable size on hairlines.
  const headScale = Math.max(lineWidth, 1.5);
  let arrowsDone = false;
  for (const sub of geom.paths) {
    if (!sub.cmds.length) continue;
    const stroked = sub.stroke && visibleLine;
    const heads: SVGElement[] = [];
    if (stroked && !arrowsDone && line && (line.head || line.tail)) {
      arrowsDone = true;
      for (const [end, atStart] of [[line.head, true], [line.tail, false]] as const) {
        if (!end) continue;
        const at = pathEnd(sub.cmds, atStart);
        if (!at) continue;
        const w = end.w * headScale, len = end.len * headScale;
        const marker = arrowHead(end.type, at, w, len, lineColor, lineWidth);
        if (line.fill.type === 'solid' && line.fill.color.a < 1) marker.setAttribute('opacity', String(r2(line.fill.color.a)));
        heads.push(marker);
        if (end.type === 'triangle' || end.type === 'stealth') shorten(sub.cmds, atStart, len * (end.type === 'stealth' ? 0.5 : 0.8));
        else if (end.type === 'arrow') shorten(sub.cmds, atStart, lineWidth / 2);
      }
    }
    const p = svg('path', { d: svgPath(sub.cmds), 'fill-rule': 'evenodd' });
    if (sub.fill !== 'none' && visibleFill) paintFill(p, fill, getDefs, box.w, box.h, env, sub.fill);
    else p.setAttribute('fill', 'none');
    if (stroked) paintStroke(p, line, getDefs, box.w, box.h);
    else p.setAttribute('stroke', 'none');
    el.append(p, ...heads);
  }
  return el;
}

// ---- Shapes ---------------------------------------------------------------------------------------------------

function placeholderOf(sp: Element): Element | null {
  for (const nv of ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvCxnSpPr', 'nvGrpSpPr']) {
    const ph = path(sp, nv, 'nvPr', 'ph');
    if (ph) return ph;
  }
  return null;
}

function nonVisual(sp: Element): Element | null {
  for (const nv of ['nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvCxnSpPr', 'nvGrpSpPr']) {
    const c = path(sp, nv, 'cNvPr');
    if (c) return c;
  }
  // Word shapes keep it directly in the shape (wps:cNvPr).
  return child(sp, 'cNvPr');
}

function inheritanceChain(sp: Element, tree: Tree): Element[] {
  const ph = placeholderOf(sp);
  return ph && tree.inherit ? tree.inherit(ph) : [];
}

/** First xfrm along the inheritance chain. */
function inheritedXfrm(sp: Element, chain: Element[]): Element | null {
  for (const el of [sp, ...chain]) {
    const xfrm = path(el, 'spPr', 'xfrm') ?? child(el, 'xfrm');
    if (xfrm && child(xfrm, 'ext')) return xfrm;
  }
  return null;
}

function resolveFill(sp: Element, chain: Element[], tree: Tree): Fill | undefined {
  const { env } = tree;
  for (const el of [sp, ...chain]) {
    const own = ownFill(child(el, 'spPr'), env.colors, tree.source);
    if (own) return own.type === 'group' ? tree.groupFill : own;
    const ref = refFill(path(el, 'style', 'fillRef'), env, tree.source);
    if (ref) return ref;
  }
  // "Use slide background": a picture background would have to line up with the slide, so it degrades to a color.
  if (bool(sp, 'useBgFill') && env.background) {
    if (env.background.type === 'solid' || env.background.type === 'grad') return env.background;
    const color = fillColor(env.background);
    return color ? { type: 'solid', color } : undefined;
  }
  return undefined;
}

function resolveLine(sp: Element, chain: Element[], tree: Tree): Line | undefined {
  let line: Line | undefined;
  for (const el of [...chain].reverse().concat(sp)) {
    const ref = refLine(path(el, 'style', 'lnRef'), tree.env, tree.source);
    if (ref) line = ref;
    const ln = path(el, 'spPr', 'ln');
    if (ln) line = parseLine(ln, tree.env.colors, tree.source, line);
  }
  return line;
}

function resolveGeometrySource(sp: Element, chain: Element[]): Element | null {
  for (const el of [sp, ...chain]) {
    const spPr = child(el, 'spPr');
    if (child(spPr, 'prstGeom') || child(spPr, 'custGeom')) return spPr;
  }
  return null;
}

/** Shape style font reference (p:style/a:fontRef) as a text layer: theme font and color. */
function fontRefLayer(sp: Element, chain: Element[], tree: Tree): Layer | null {
  for (const el of [sp, ...chain]) {
    const ref = path(el, 'style', 'fontRef');
    if (!ref) continue;
    const run: RunProps = {};
    const idx = attr(ref, 'idx');
    if (idx === 'major') run.latin = '+mj-lt';
    else if (idx === 'minor') run.latin = '+mn-lt';
    const color = colorIn(ref, tree.env.colors);
    if (color) run.color = color;
    return { run };
  }
  return null;
}

function bodyProps(sp: Element, chain: Element[]): BodyProps {
  const body = defaultBody();
  for (const el of [...chain].reverse().concat(sp)) mergeBodyPr(body, path(el, 'txBody', 'bodyPr'));
  return body;
}

/** Text box rectangle in the shape's own (unrotated) coordinates. */
function textRect(sp: Element, box: Box, geom: Geometry, frame: Frame) {
  const txXfrm = child(sp, 'txXfrm');
  if (txXfrm) {
    // SmartArt drawings give the text its own box in drawing coordinates; bring its center into the shape.
    const tb = place(txXfrm, frame);
    if (tb) {
      const a = (-box.rot * Math.PI) / 180;
      const dx = tb.x + tb.w / 2 - (box.x + box.w / 2), dy = tb.y + tb.h / 2 - (box.y + box.h / 2);
      const lx = dx * Math.cos(a) - dy * Math.sin(a), ly = dx * Math.sin(a) + dy * Math.cos(a);
      return { x: box.w / 2 + lx - tb.w / 2, y: box.h / 2 + ly - tb.h / 2, w: tb.w, h: tb.h, rot: tb.rot };
    }
  }
  const t = geom.text;
  if (!t) return { x: 0, y: 0, w: box.w, h: box.h, rot: 0 };
  let { l, r } = t;
  let { t: top, b } = t;
  if (box.flipH) [l, r] = [box.w - r, box.w - l];
  if (box.flipV) [top, b] = [box.h - b, box.h - top];
  return { x: l, y: top, w: Math.max(0, r - l), h: Math.max(0, b - top), rot: 0 };
}

function renderSp(sp: Element, frame: Frame, tree: Tree, chain: Element[]): HTMLElement | null {
  const box = place(inheritedXfrm(sp, chain), frame);
  if (!box) return null;
  const { env } = tree;
  const el = positioned('pptx-sp', box);
  const geomSource = resolveGeometrySource(sp, chain);
  const geom = geometryOf(geomSource, box.w, box.h, box.emuW, box.emuH);
  const fill = resolveFill(sp, chain, tree);
  const line = resolveLine(sp, chain, tree);
  const shape = drawGeometry(geom, fill, line, box, env);
  if (shape) el.append(shape);

  const txBody = child(sp, 'txBody');
  if (txBody) {
    const layers = tree.layers(sp, chain);
    const fontRef = fontRefLayer(sp, chain, tree);
    if (fontRef) layers.push(fontRef);
    const rect = textRect(sp, box, geom, frame);
    const text = renderTextBox(txBody, layers, bodyProps(sp, chain), rect, { env, source: tree.source });
    if (text) {
      const turns = [rect.rot ? `rotate(${r2(rect.rot)}deg)` : '', box.flipV ? 'rotate(180deg)' : '', text.style.transform].filter(Boolean);
      text.style.transform = turns.join(' ');
      el.append(text);
    }
  } else if (tree.text) {
    const rect = textRect(sp, box, geom, frame);
    const text = tree.text(sp, rect, box, tree);
    if (text) {
      const turns = [rect.rot ? `rotate(${r2(rect.rot)}deg)` : '', box.flipV ? 'rotate(180deg)' : '', text.style.transform].filter(Boolean);
      text.style.transform = turns.join(' ');
      el.append(text);
    }
  }
  return el.childElementCount ? el : null;
}

/** Label of the neutral box shown instead of a picture the browser cannot display. */
function missingLabel(format: string | undefined): string {
  return !format || format === 'unknown' ? 'Image not available' : `${format} image`;
}

function unsupportedImage(box: Box, label: string): HTMLElement {
  const el = positioned('pptx-pic pptx-missing', box);
  el.append(h('span', {}, label));
  return el;
}

function renderPic(pic: Element, frame: Frame, tree: Tree, chain: Element[]): HTMLElement | null {
  const box = place(inheritedXfrm(pic, chain), frame);
  if (!box || box.w <= 0 || box.h <= 0) return null;
  const { env } = tree;
  const blipFill = child(pic, 'blipFill');
  const blip = child(blipFill, 'blip');
  const parts = blipParts(blip, tree.source);
  const el = positioned('pptx-pic', box);
  const crop = h('div', { class: 'pptx-crop' });
  const flip = flipTransform(box);
  if (flip) crop.style.transform = flip;
  const geom = geometryOf(resolveGeometrySource(pic, chain), box.w, box.h, box.emuW, box.emuH);
  const prst = attr(path(pic, 'spPr', 'prstGeom'), 'prst');
  if (prst && prst !== 'rect' && geom.paths[0]) crop.style.clipPath = `path("${svgPath(geom.paths[0].cmds)}")`;
  el.append(crop);

  // Pictures can have a fill behind transparent areas and an outline.
  const fill = ownFill(child(pic, 'spPr'), env.colors, tree.source);
  const line = resolveLine(pic, chain, tree);
  if (fill && fill.type !== 'none') {
    const back = drawGeometry(geom, fill, undefined, box, env);
    if (back) el.insertBefore(back, crop);
  }

  if (parts.length && blipFill) {
    const { crop: c, inset } = blipBox(blipFill);
    const bw = box.w * (1 - inset[0] - inset[2]), bh = box.h * (1 - inset[1] - inset[3]);
    const fullW = bw / Math.max(0.01, 1 - c[0] - c[2]), fullH = bh / Math.max(0.01, 1 - c[1] - c[3]);
    const img = h('img', { alt: attr(nonVisual(pic), 'descr') ?? '', draggable: false, decoding: 'async' });
    const s = img.style;
    s.left = `${r2(box.w * inset[0] - c[0] * fullW)}px`;
    s.top = `${r2(box.h * inset[1] - c[1] * fullH)}px`;
    s.width = `${r2(fullW)}px`;
    s.height = `${r2(fullH)}px`;
    const effects = blipEffects(blip, env);
    if (effects.opacity < 1) s.opacity = String(r2(effects.opacity));
    if (effects.filter) s.filter = effects.filter;
    crop.append(img);
    env.pending.push(
      loadBlip(env, parts).then(async (media) => {
        if (media.url) {
          img.src = media.url;
          await img.decode().catch(() => undefined);
        } else {
          img.remove();
          crop.append(h('div', { class: 'pptx-missing' }, h('span', {}, missingLabel(media.unsupported))));
        }
      }),
    );
  } else if (!fill) {
    return unsupportedImage(box, 'Linked image');
  }
  const outline = drawGeometry(geom, undefined, line, box, env);
  if (outline) el.append(outline);
  return el;
}

async function renderGroup(grp: Element, frame: Frame, tree: Tree): Promise<HTMLElement | null> {
  if (tree.depth >= MAX_DEPTH) return null;
  const xfrm = path(grp, 'grpSpPr', 'xfrm');
  const box = place(xfrm, frame);
  if (!box) return null;
  const el = positioned('pptx-grp', box);
  const chOff = child(xfrm, 'chOff'), chExt = child(xfrm, 'chExt');
  const chW = emuAttr(chExt, 'cx'), chH = emuAttr(chExt, 'cy');
  const inner: Frame = {
    chX: emuAttr(chOff, 'x'),
    chY: emuAttr(chOff, 'y'),
    sx: chW > 0 ? box.w / chW : frame.sx,
    sy: chH > 0 ? box.h / chH : frame.sy,
    w: box.w,
    h: box.h,
    flipH: box.flipH,
    flipV: box.flipV,
  };
  const groupFill = ownFill(child(grp, 'grpSpPr'), tree.env.colors, tree.source);
  await renderTree(grp, el, inner, { ...tree, depth: tree.depth + 1, groupFill: groupFill ?? tree.groupFill });
  return el;
}

async function renderFrame(frameEl: Element, frame: Frame, tree: Tree, chain: Element[]): Promise<HTMLElement | null> {
  const box = place(child(frameEl, 'xfrm') ?? inheritedXfrm(frameEl, chain), frame);
  if (!box) return null;
  const custom = tree.graphicFrame ? await tree.graphicFrame(frameEl, box, tree) : null;
  if (custom) return custom;
  // OLE objects and unknown content usually carry a picture of themselves.
  const data = path(frameEl, 'graphic', 'graphicData');
  for (const el of data ? data.getElementsByTagName('*') : []) {
    if (el.localName !== 'pic') continue;
    const pic = renderPic(el, rootFrame(box.w, box.h), tree, []);
    if (!pic) continue;
    // The picture's own xfrm is in slide coordinates; the frame decides where it goes.
    Object.assign(pic.style, { left: `${r2(box.x)}px`, top: `${r2(box.y)}px`, width: `${r2(box.w)}px`, height: `${r2(box.h)}px` });
    return pic;
  }
  return null;
}

export async function renderShape(el: Element, frame: Frame, tree: Tree): Promise<HTMLElement | null> {
  if (bool(nonVisual(el), 'hidden')) return null;
  if (tree.skipPlaceholders && placeholderOf(el)) return null;
  const chain = inheritanceChain(el, tree);
  switch (el.localName) {
    case 'sp':
    case 'cxnSp':
    case 'wsp':
      return renderSp(el, frame, tree, chain);
    case 'pic':
      return renderPic(el, frame, tree, chain);
    case 'grpSp':
    case 'wgp':
      return renderGroup(el, frame, tree);
    case 'graphicFrame':
      return renderFrame(el, frame, tree, chain);
    default:
      return null;
  }
}

/** Renders the children of an spTree or group into a container element. */
export async function renderTree(treeEl: Element, container: HTMLElement, frame: Frame, tree: Tree): Promise<void> {
  for (const el of elements(treeEl)) {
    if (!['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame', 'wsp', 'wgp'].includes(el.localName)) continue;
    if (tree.env.budget.shapes-- <= 0) return;
    try {
      const out = await renderShape(el, frame, tree);
      if (out) container.append(out);
    } catch (error) {
      // One broken shape should not take the whole slide down.
      console.warn('PowerPoint shape skipped:', error);
    }
  }
}

