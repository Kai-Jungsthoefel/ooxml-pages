// SPDX-License-Identifier: MIT
// SVG content (shapes, charts, VML) to PDF. Each element is drawn in its own user space, placed by its screen
// transform relative to the page; fills may be colors, gradients (PDF shadings) or tiled pictures, and clip paths
// apply to whole groups.
import { parseColor } from './color.js';
import type { PageJob } from './dom.js';
import { paintElement, textStyleOf, toMatrix } from './dom.js';
import type { GradientStop, StrokeStyle } from './painter.js';
import { type Segment, ellipsePath, parsePath, polyPath, rectPath } from './path.js';
import { ascentOf, drawText, measureText } from './text.js';

const SKIP = new Set(['defs', 'linearGradient', 'radialGradient', 'pattern', 'clipPath', 'marker', 'filter', 'mask', 'style', 'title', 'desc', 'symbol', 'metadata', 'script']);

const num = (value: string) => parseFloat(value) || 0;

/** Transform from an element's user space to page px. */
function pageMatrix(el: SVGGraphicsElement, job: PageJob): DOMMatrix | null {
  const m = el.getScreenCTM();
  return m ? new DOMMatrix().translate(-job.left, -job.top).multiply(DOMMatrix.fromMatrix(m)) : null;
}

function geometry(el: SVGElement): Segment[] | null {
  if (el instanceof SVGPathElement) return parsePath(el.getAttribute('d') ?? '');
  if (el instanceof SVGRectElement) {
    const w = el.width.baseVal.value, h = el.height.baseVal.value;
    if (w <= 0 || h <= 0) return null;
    const rx = el.rx.baseVal.value, ry = el.ry.baseVal.value;
    return rectPath(el.x.baseVal.value, el.y.baseVal.value, w, h, rx || ry, ry || rx);
  }
  if (el instanceof SVGCircleElement) return ellipsePath(el.cx.baseVal.value, el.cy.baseVal.value, el.r.baseVal.value, el.r.baseVal.value);
  if (el instanceof SVGEllipseElement) return ellipsePath(el.cx.baseVal.value, el.cy.baseVal.value, el.rx.baseVal.value, el.ry.baseVal.value);
  if (el instanceof SVGLineElement) return [{ k: 'M', x: el.x1.baseVal.value, y: el.y1.baseVal.value }, { k: 'L', x: el.x2.baseVal.value, y: el.y2.baseVal.value }];
  if (el instanceof SVGPolygonElement) return polyPath(el.getAttribute('points') ?? '', true);
  if (el instanceof SVGPolylineElement) return polyPath(el.getAttribute('points') ?? '', false);
  return null;
}

/** The element a paint or clip reference (url("#id")) points to. */
function referenced(value: string, el: Element): Element | null {
  const id = /url\(\s*["']?#([^"')]+)["']?\s*\)/.exec(value)?.[1];
  if (!id) return null;
  return (el.ownerDocument.getElementById(id) as Element | null) ?? null;
}

function length(l: SVGAnimatedLength, bbox: boolean): number {
  const v = l.baseVal;
  if (!bbox) return v.value;
  return v.unitType === SVGLength.SVG_LENGTHTYPE_PERCENTAGE ? v.valueInSpecifiedUnits / 100 : v.valueInSpecifiedUnits;
}

function stopsOf(g: SVGGradientElement, alpha: number): GradientStop[] {
  // Stops may be inherited from a gradient the href points to.
  let source: Element | null = g;
  for (let i = 0; i < 4 && source && !source.querySelector('stop'); i++) source = referenced(`url(${source.getAttribute('href') ?? source.getAttribute('xlink:href') ?? ''})`, g);
  const stops: GradientStop[] = [];
  for (const stop of source?.querySelectorAll('stop') ?? []) {
    const cs = getComputedStyle(stop);
    const color = parseColor(cs.stopColor, num(cs.stopOpacity || '1') * alpha) ?? { r: 0, g: 0, b: 0, a: 0 };
    stops.push({ offset: Math.min(1, Math.max(0, (stop as SVGStopElement).offset.baseVal)), color });
  }
  return stops;
}

/** Fills the current clip with a gradient element, in the shape's user space. */
function paintGradient(g: SVGGradientElement, shape: SVGGraphicsElement, job: PageJob, alpha: number): void {
  const stops = stopsOf(g, alpha);
  if (!stops.length) return;
  const bbox = (g.getAttribute('gradientUnits') ?? 'objectBoundingBox') === 'objectBoundingBox';
  if (bbox) {
    const bb = shape.getBBox();
    if (!bb.width || !bb.height) return;
    job.painter.transform([bb.width, 0, 0, bb.height, bb.x, bb.y]);
  }
  const transform = g.gradientTransform.baseVal.consolidate();
  if (transform) job.painter.transform(toMatrix(transform.matrix));
  if (g instanceof SVGLinearGradientElement) {
    job.painter.gradient('linear', [length(g.x1, bbox), length(g.y1, bbox), length(g.x2, bbox), length(g.y2, bbox)], stops);
  } else if (g instanceof SVGRadialGradientElement) {
    const cx = length(g.cx, bbox), cy = length(g.cy, bbox), r = length(g.r, bbox);
    const fx = g.hasAttribute('fx') ? length(g.fx, bbox) : cx, fy = g.hasAttribute('fy') ? length(g.fy, bbox) : cy;
    job.painter.gradient('radial', [fx, fy, 0, cx, cy, r], stops);
  }
}

/** Fills the current clip with a picture pattern (tiled picture fills), in the shape's user space. */
async function paintPattern(p: SVGPatternElement, shape: SVGGraphicsElement, job: PageJob, alpha: number): Promise<boolean> {
  const img = p.querySelector('image');
  const href = img?.href.baseVal || img?.getAttribute('href');
  if (!img || !href) return false;
  const w = p.width.baseVal.value, h = p.height.baseVal.value;
  const image = await job.images.fromUrl(href);
  if (!image || w <= 0 || h <= 0) return false;
  const bb = shape.getBBox();
  const x0 = p.x.baseVal.value, y0 = p.y.baseVal.value;
  const startX = x0 + Math.floor((bb.x - x0) / w) * w, startY = y0 + Math.floor((bb.y - y0) / h) * h;
  const count = Math.ceil((bb.x + bb.width - startX) / w) * Math.ceil((bb.y + bb.height - startY) / h);
  if (count > 2500) {
    job.painter.image(image, bb.x, bb.y, bb.width, bb.height, alpha);
    return true;
  }
  const iw = img.width.baseVal.value || w, ih = img.height.baseVal.value || h;
  for (let y = startY; y < bb.y + bb.height; y += h) for (let x = startX; x < bb.x + bb.width; x += w) job.painter.image(image, x + img.x.baseVal.value, y + img.y.baseVal.value, iw, ih, alpha);
  return true;
}

function strokeOf(cs: CSSStyleDeclaration): StrokeStyle {
  const dash = cs.strokeDasharray && cs.strokeDasharray !== 'none' ? cs.strokeDasharray.split(/[\s,]+/).map(num).filter((n) => n >= 0) : undefined;
  return {
    width: num(cs.strokeWidth),
    dash: dash && dash.some((n) => n > 0) ? (dash.length % 2 ? [...dash, ...dash] : dash) : undefined,
    dashOffset: num(cs.strokeDashoffset),
    cap: (cs.strokeLinecap as StrokeStyle['cap']) || 'butt',
    join: (cs.strokeLinejoin === 'arcs' || cs.strokeLinejoin === 'miter-clip' ? 'miter' : cs.strokeLinejoin) as StrokeStyle['join'],
  };
}

/** A computed opacity value (fill-opacity, stroke-opacity: a number or a percentage), 1 when unset. */
function opacityOf(value: string): number {
  const n = parseFloat(value);
  if (!Number.isFinite(n)) return 1;
  return Math.min(1, Math.max(0, value.trim().endsWith('%') ? n / 100 : n));
}

async function paintShape(el: SVGGraphicsElement, segments: Segment[], cs: CSSStyleDeclaration, job: PageJob, alpha: number): Promise<void> {
  const m = pageMatrix(el, job);
  if (!m || !segments.length) return;
  const { painter } = job;
  const evenOdd = cs.fillRule === 'evenodd';
  painter.save();
  painter.transform(toMatrix(m));
  const fillAlpha = alpha * opacityOf(cs.fillOpacity);
  if (cs.fill && cs.fill !== 'none') {
    const server = referenced(cs.fill, el);
    if (server instanceof SVGGradientElement || server instanceof SVGPatternElement) {
      painter.save();
      painter.path(segments);
      painter.clip(evenOdd);
      if (server instanceof SVGGradientElement) paintGradient(server, el, job, fillAlpha);
      else if (!(await paintPattern(server, el, job, fillAlpha))) job.notes.add('Some pattern fills are left out.');
      painter.restore();
    } else {
      const color = parseColor(cs.fill.replace(/^url\([^)]*\)\s*/, ''), fillAlpha);
      if (color) {
        painter.path(segments);
        painter.fill(color, evenOdd);
      }
    }
  }
  if (cs.stroke && cs.stroke !== 'none' && num(cs.strokeWidth) > 0) {
    const strokeAlpha = alpha * opacityOf(cs.strokeOpacity);
    const server = referenced(cs.stroke, el);
    // A gradient outline takes the color of its middle.
    const color =
      server instanceof SVGGradientElement
        ? (stopsOf(server, strokeAlpha)[Math.floor(stopsOf(server, strokeAlpha).length / 2)]?.color ?? null)
        : parseColor(cs.stroke.replace(/^url\([^)]*\)\s*/, ''), strokeAlpha);
    if (color) {
      painter.path(segments);
      painter.stroke(color, strokeOf(cs));
    }
  }
  painter.restore();
  if ((cs.markerStart && cs.markerStart !== 'none') || (cs.markerEnd && cs.markerEnd !== 'none')) await paintMarkers(el, segments, cs, job, alpha);
  if (cs.filter && cs.filter !== 'none') job.notes.add('Shadows and glow are left out.');
}

/** Arrow heads (VML lines): the marker's shapes at the path ends, turned along the line. */
async function paintMarkers(el: SVGGraphicsElement, segments: Segment[], cs: CSSStyleDeclaration, job: PageJob, alpha: number): Promise<void> {
  const m = pageMatrix(el, job);
  if (!m) return;
  const points = segments.filter((s) => s.k !== 'Z') as Exclude<Segment, { k: 'Z' }>[];
  if (points.length < 2) return;
  const first = points[0], second = points[1], last = points[points.length - 1], before = points[points.length - 2];
  // Direction of the path at its ends: towards the first control point, from the last one. The start marker
  // points backwards (orient="auto-start-reverse", as the VML renderer writes it).
  const startTo = second.k === 'C' ? { x: second.x1, y: second.y1 } : second;
  const endFrom = last.k === 'C' ? { x: last.x2, y: last.y2 } : before;
  const ends: [string, { x: number; y: number }, { x: number; y: number }][] = [
    [cs.markerStart, startTo, first],
    [cs.markerEnd, endFrom, last],
  ];
  const width = num(cs.strokeWidth) || 1;
  for (const [ref, from, at] of ends) {
    const marker = ref && ref !== 'none' ? referenced(ref, el) : null;
    if (!(marker instanceof SVGMarkerElement)) continue;
    const angle = Math.atan2(at.y - from.y, at.x - from.x);
    const scale = marker.markerUnits.baseVal === SVGMarkerElement.SVG_MARKERUNITS_STROKEWIDTH ? width : 1;
    const vb = marker.viewBox.baseVal;
    const mw = marker.markerWidth.baseVal.value, mh = marker.markerHeight.baseVal.value;
    const sx = vb && vb.width ? mw / vb.width : 1, sy = vb && vb.height ? mh / vb.height : 1;
    const local = new DOMMatrix()
      .translate(at.x, at.y)
      .rotate((angle * 180) / Math.PI)
      .scale(scale)
      .scale(sx, sy)
      .translate(-(marker.refX.baseVal.value), -(marker.refY.baseVal.value));
    for (const child of marker.children) {
      const segs = geometry(child as SVGElement);
      if (!segs) continue;
      const ccs = getComputedStyle(child);
      const color = parseColor(ccs.fill, alpha) ?? parseColor(cs.stroke, alpha);
      if (!color) continue;
      job.painter.save();
      job.painter.transform(toMatrix(m.multiply(local)));
      job.painter.path(segs);
      job.painter.fill(color);
      job.painter.restore();
    }
  }
}

async function paintImageElement(el: SVGImageElement, job: PageJob, alpha: number): Promise<void> {
  const m = pageMatrix(el, job);
  const href = el.href.baseVal || el.getAttribute('href') || el.getAttribute('xlink:href');
  if (!m || !href) return;
  const x = el.x.baseVal.value, y = el.y.baseVal.value, w = el.width.baseVal.value, h = el.height.baseVal.value;
  if (w <= 0 || h <= 0) return;
  const scale = Math.hypot(m.a, m.b);
  const image = await job.images.fromUrl(href, { w: w * scale, h: h * scale });
  if (!image) {
    job.notes.add('A picture could not be carried over.');
    return;
  }
  const par = el.preserveAspectRatio.baseVal;
  let target = { x, y, w, h };
  let clip = false;
  if (par.align !== SVGPreserveAspectRatio.SVG_PRESERVEASPECTRATIO_NONE && image.width && image.height) {
    const slice = par.meetOrSlice === SVGPreserveAspectRatio.SVG_MEETORSLICE_SLICE;
    const s = (slice ? Math.max : Math.min)(w / image.width, h / image.height);
    const iw = image.width * s, ih = image.height * s;
    // Alignment: xMin/xMid/xMax × yMin/yMid/yMax, numbered 2 … 10.
    const a = par.align - 2;
    const fx = [0, 0.5, 1][a % 3] ?? 0.5, fy = [0, 0.5, 1][Math.floor(a / 3)] ?? 0.5;
    target = { x: x + (w - iw) * fx, y: y + (h - ih) * fy, w: iw, h: ih };
    clip = slice;
  }
  job.painter.save();
  job.painter.transform(toMatrix(m));
  if (clip) {
    job.painter.path(rectPath(x, y, w, h));
    job.painter.clip();
  }
  job.painter.image(image, target.x, target.y, target.w, target.h, alpha);
  job.painter.restore();
  if (el.closest('[filter]') || getComputedStyle(el).filter !== 'none') job.notes.add('Picture color effects are left out.');
}

/** Chart labels and WordArt: word by word at the positions the browser gives each character. */
async function paintSvgText(el: SVGTextElement, job: PageJob, alpha: number): Promise<void> {
  const m = pageMatrix(el, job);
  const count = el.getNumberOfChars();
  const text = el.textContent ?? '';
  if (!m || !count || !text.trim()) return;
  const cs = getComputedStyle(el);
  const style = textStyleOf(cs, job, alpha * opacityOf(cs.fillOpacity), cs.fill);
  const strokeColor = cs.stroke && cs.stroke !== 'none' ? parseColor(cs.stroke, alpha * opacityOf(cs.strokeOpacity)) : null;
  if (strokeColor && num(cs.strokeWidth) > 0) style.outline = { color: strokeColor, width: num(cs.strokeWidth) };
  if (!style.families.length || !style.size) return;
  job.painter.save();
  job.painter.transform(toMatrix(m));
  // Text squeezed or stretched to a length (VML WordArt): one run, scaled.
  // The start position lies on the alignment baseline (central for chart labels); the alphabetic baseline is the
  // top of the glyph cell plus the font's ascent, as for HTML text.
  const ascent = ascentOf(`${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`);
  const baseline = (index: number) => el.getExtentOfChar(index).y + ascent;
  const target = el.textLength.baseVal.value;
  if (target > 0 && el.lengthAdjust.baseVal === SVGTextContentElement.LENGTHADJUST_SPACINGANDGLYPHS) {
    const start = el.getStartPositionOfChar(0);
    const natural = measureText(job.fonts, text, style);
    await drawText(job, text, start.x, baseline(0), { ...style, scaleX: natural > 0 ? target / natural : 1 });
  } else {
    for (const word of text.matchAll(/\S+\s*/gu)) {
      const index = word.index!;
      if (index >= count) break;
      const start = el.getStartPositionOfChar(index);
      await drawText(job, word[0].length > word[0].trimEnd().length ? `${word[0].trimEnd()} ` : word[0], start.x, baseline(index), style);
    }
  }
  job.painter.restore();
}

/** A clip path element as segments in the user space of the element it clips. */
function clipSegments(clip: SVGClipPathElement, target: SVGGraphicsElement): { segments: Segment[]; evenOdd: boolean } | null {
  const out: Segment[] = [];
  let evenOdd = false;
  const bbox = clip.clipPathUnits.baseVal === SVGUnitTypes.SVG_UNIT_TYPE_OBJECTBOUNDINGBOX ? target.getBBox() : null;
  for (const child of clip.children) {
    const segs = geometry(child as SVGElement);
    if (!segs) continue;
    let matrix = new DOMMatrix();
    if (bbox) matrix = matrix.translate(bbox.x, bbox.y).scale(bbox.width, bbox.height);
    const t = (child as SVGGraphicsElement).transform?.baseVal.consolidate();
    if (t) matrix = matrix.multiply(DOMMatrix.fromMatrix(t.matrix));
    const map = (x: number, y: number) => matrix.transformPoint(new DOMPoint(x, y));
    for (const s of segs) {
      if (s.k === 'Z') out.push(s);
      else if (s.k === 'C') {
        const a = map(s.x1, s.y1), b = map(s.x2, s.y2), c = map(s.x, s.y);
        out.push({ k: 'C', x1: a.x, y1: a.y, x2: b.x, y2: b.y, x: c.x, y: c.y });
      } else {
        const p = map(s.x, s.y);
        out.push({ k: s.k, x: p.x, y: p.y });
      }
    }
    if (getComputedStyle(child).clipRule === 'evenodd') evenOdd = true;
  }
  return out.length ? { segments: out, evenOdd } : null;
}

async function paintNode(el: Element, job: PageJob, alpha: number): Promise<void> {
  if (el instanceof SVGForeignObjectElement) {
    for (const child of el.children) await paintElement(child, job, new DOMMatrix(), alpha);
    return;
  }
  if (!(el instanceof SVGGraphicsElement) || SKIP.has(el.localName)) return;
  const cs = getComputedStyle(el);
  if (cs.display === 'none') return;
  alpha *= Number.isFinite(parseFloat(cs.opacity)) ? parseFloat(cs.opacity) : 1;
  if (alpha <= 0.002) return;
  const { painter } = job;
  let depth = 0;
  /** Clips what follows to segments in the element's user space; the page space stays current. */
  const clipTo = (m: DOMMatrix, segments: Segment[], evenOdd: boolean) => {
    painter.save();
    painter.transform(toMatrix(m));
    painter.path(segments);
    painter.clip(evenOdd);
    painter.transform(toMatrix(m.inverse()));
    depth++;
  };
  try {
    const clip = cs.clipPath && cs.clipPath !== 'none' ? referenced(cs.clipPath, el) : null;
    if (clip instanceof SVGClipPathElement) {
      const m = pageMatrix(el, job);
      const segs = clipSegments(clip, el);
      if (m && segs) clipTo(m, segs.segments, segs.evenOdd);
    }
    if (el instanceof SVGSVGElement || el instanceof SVGGElement || el instanceof SVGAElement || el instanceof SVGSwitchElement) {
      if (el instanceof SVGSVGElement && cs.overflow !== 'visible') {
        const m = pageMatrix(el, job);
        const vb = el.viewBox.baseVal;
        if (m) clipTo(m, vb && vb.width ? rectPath(vb.x, vb.y, vb.width, vb.height) : rectPath(0, 0, el.width.baseVal.value, el.height.baseVal.value), false);
      }
      for (const child of el.children) await paintNode(child, job, alpha);
      return;
    }
    if (cs.visibility !== 'visible') return;
    if (el instanceof SVGImageElement) await paintImageElement(el, job, alpha);
    else if (el instanceof SVGTextElement) await paintSvgText(el, job, alpha);
    else if (el instanceof SVGUseElement) job.notes.add('Some shapes are left out.');
    else {
      const segments = geometry(el);
      if (segments) await paintShape(el, segments, cs, job, alpha);
    }
  } finally {
    for (let i = 0; i < depth; i++) painter.restore();
  }
}

/** An outer SVG element in the page; its own box (clip, transform) was handled by the DOM walker. */
export async function paintSvg(svg: SVGSVGElement, job: PageJob, alpha: number): Promise<void> {
  for (const child of svg.children) await paintNode(child, job, alpha);
}
