// SPDX-License-Identifier: MIT
// Drawing VML shapes: rectangles, rounded rectangles, ovals, lines, polylines and simple paths with fill and
// stroke, pictures (v:imagedata, with cropping), text boxes (v:textbox with Word content), WordArt text
// (watermarks) and groups.
import { h } from '../dom.js';
import { type Env, type Source, r2, svg, uid } from '../pptx/env.js';
import { loadBlip } from '../pptx/fill.js';
import type { DrawContext } from './drawing.js';
import { layoutStory } from './flow.js';
import { createContext, parseBlocks } from './parse.js';
import { vmlBool, vmlLength, vmlStyle, vmlTextbox } from './vml.js';
import { clampNum, kids, wa } from './xml.js';

const NAMED: Record<string, string> = {
  black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00', silver: '#c0c0c0',
  gray: '#808080', grey: '#808080', maroon: '#800000', navy: '#000080', olive: '#808000', purple: '#800080', teal: '#008080',
  lime: '#00ff00', aqua: '#00ffff', fuchsia: '#ff00ff', windowtext: '#000000', window: '#ffffff', buttonface: '#f0f0f0',
};

/** A VML color ("#rrggbb", "#rgb", a name, optionally followed by a palette index in brackets). */
export function vmlColor(value: string | null | undefined, fallback: string | null): string | null {
  if (!value) return fallback;
  const v = value.trim().split(/\s+/)[0].toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(v)) return v;
  if (/^#[0-9a-f]{3}$/.test(v)) return `#${v[1]}${v[1]}${v[2]}${v[2]}${v[3]}${v[3]}`;
  if (/^[0-9a-f]{6}$/.test(v)) return `#${v}`;
  return NAMED[v] ?? fallback;
}

/** VML fractions: "0.5", "50%", or 16.16 fixed point "32768f". */
function fraction(value: string | null, fallback: number): number {
  if (!value) return fallback;
  const v = value.trim();
  if (v.endsWith('f')) return clampNum(Number(v.slice(0, -1)) / 65536, -10, 10) || 0;
  if (v.endsWith('%')) return clampNum(Number(v.slice(0, -1)) / 100, -10, 10) || 0;
  const n = Number(v);
  return Number.isFinite(n) ? clampNum(n, -10, 10) : fallback;
}

function relId(el: Element | null, name: string): string | null {
  if (!el) return null;
  for (const a of el.attributes) if (a.localName === name && a.namespaceURI && /relationships$/.test(a.namespaceURI)) return a.value;
  return null;
}

const DASHES: Record<string, string> = { dash: '4 3', dot: '1 2', dashdot: '4 2 1 2', shortdash: '3 1', shortdot: '1 1', longdash: '8 3', longdashdot: '8 3 1 3', '1 1': '1 1' };

interface Paint {
  fill: string | null;
  fillOpacity: number;
  stroke: string | null;
  strokeWidth: number;
  dash: string | null;
  startArrow: boolean;
  endArrow: boolean;
  gradient?: { color2: string; angle: number };
}

function paintOf(el: Element): Paint {
  const fillEl = kids(el).find((c) => c.localName === 'fill') ?? null;
  const strokeEl = kids(el).find((c) => c.localName === 'stroke') ?? null;
  const filled = vmlBool(wa(el, 'filled'), true) && vmlBool(wa(fillEl, 'on'), true);
  const stroked = vmlBool(wa(el, 'stroked'), true) && vmlBool(wa(strokeEl, 'on'), true);
  const fill = filled ? vmlColor(wa(fillEl, 'color') ?? wa(el, 'fillcolor'), '#ffffff') : null;
  const type = wa(fillEl, 'type');
  const paint: Paint = {
    fill,
    fillOpacity: fraction(wa(fillEl, 'opacity'), 1),
    stroke: stroked ? vmlColor(wa(strokeEl, 'color') ?? wa(el, 'strokecolor'), '#000000') : null,
    strokeWidth: Math.max(0.5, vmlLength(wa(strokeEl, 'weight') ?? wa(el, 'strokeweight'), 1)),
    dash: DASHES[(wa(strokeEl, 'dashstyle') ?? '').toLowerCase()] ?? null,
    startArrow: !!wa(strokeEl, 'startarrow') && wa(strokeEl, 'startarrow') !== 'none',
    endArrow: !!wa(strokeEl, 'endarrow') && wa(strokeEl, 'endarrow') !== 'none',
  };
  if (fill && (type === 'gradient' || type === 'gradientRadial')) {
    paint.gradient = { color2: vmlColor(wa(fillEl, 'color2'), '#ffffff') ?? '#ffffff', angle: Number(wa(fillEl, 'angle') ?? 0) || 0 };
  }
  return paint;
}

function applyPaint(target: SVGElement, paint: Paint, defs: () => SVGDefsElement, closed: boolean): void {
  if (closed && paint.fill) {
    if (paint.gradient) {
      const g = svg('linearGradient', { id: uid('vml'), gradientTransform: `rotate(${r2(90 - paint.gradient.angle)} 0.5 0.5)` });
      g.append(svg('stop', { offset: 0, 'stop-color': paint.gradient.color2 }), svg('stop', { offset: 1, 'stop-color': paint.fill }));
      defs().append(g);
      target.setAttribute('fill', `url(#${g.id})`);
    } else target.setAttribute('fill', paint.fill);
    if (paint.fillOpacity < 1) target.setAttribute('fill-opacity', String(r2(paint.fillOpacity)));
  } else target.setAttribute('fill', 'none');
  if (paint.stroke) {
    target.setAttribute('stroke', paint.stroke);
    target.setAttribute('stroke-width', String(r2(paint.strokeWidth)));
    if (paint.dash) target.setAttribute('stroke-dasharray', paint.dash.split(' ').map((d) => r2(Number(d) * paint.strokeWidth)).join(' '));
  } else target.setAttribute('stroke', 'none');
}

function arrow(defs: () => SVGDefsElement, color: string, start: boolean): string {
  const marker = svg('marker', { id: uid('arrow'), viewBox: '0 0 10 10', refX: start ? 1 : 9, refY: 5, markerWidth: 4, markerHeight: 4, orient: 'auto-start-reverse' });
  marker.append(svg('path', { d: 'M0 0L10 5L0 10z', fill: color }));
  defs().append(marker);
  return `url(#${marker.id})`;
}

/** VML path commands (m, l, c, x, e, r, v, t, nf, ns) in shape coordinates; formulas are not evaluated. */
function vmlPath(path: string, sx: number, sy: number, ox: number, oy: number): string | null {
  const tokens = path.match(/[a-z]+|-?\d*\.?\d+/gi) ?? [];
  if (tokens.some((t) => t.startsWith('@'))) return null;
  let out = '';
  let cmd = '';
  let x = 0, y = 0;
  const nums: number[] = [];
  const P = (px: number, py: number) => `${r2((px - ox) * sx)} ${r2((py - oy) * sy)}`;
  const flush = () => {
    switch (cmd) {
      case 'm':
        for (let i = 0; i + 1 < nums.length; i += 2) (x = nums[i]), (y = nums[i + 1]), (out += `${i ? 'L' : 'M'}${P(x, y)}`);
        break;
      case 't':
        for (let i = 0; i + 1 < nums.length; i += 2) (x += nums[i]), (y += nums[i + 1]), (out += `M${P(x, y)}`);
        break;
      case 'l':
        for (let i = 0; i + 1 < nums.length; i += 2) (x = nums[i]), (y = nums[i + 1]), (out += `L${P(x, y)}`);
        break;
      case 'r':
        for (let i = 0; i + 1 < nums.length; i += 2) (x += nums[i]), (y += nums[i + 1]), (out += `L${P(x, y)}`);
        break;
      case 'c':
        for (let i = 0; i + 5 < nums.length; i += 6) {
          out += `C${P(nums[i], nums[i + 1])} ${P(nums[i + 2], nums[i + 3])} ${P(nums[i + 4], nums[i + 5])}`;
          x = nums[i + 4];
          y = nums[i + 5];
        }
        break;
      case 'v':
        for (let i = 0; i + 5 < nums.length; i += 6) {
          out += `C${P(x + nums[i], y + nums[i + 1])} ${P(x + nums[i + 2], y + nums[i + 3])} ${P(x + nums[i + 4], y + nums[i + 5])}`;
          x += nums[i + 4];
          y += nums[i + 5];
        }
        break;
      case 'x':
        out += 'Z';
        break;
    }
    nums.length = 0;
  };
  for (const t of tokens.slice(0, 20000)) {
    if (/^[a-z]+$/i.test(t)) {
      flush();
      cmd = t.toLowerCase();
      if (cmd === 'x') flush();
    } else nums.push(Number(t));
  }
  flush();
  return out || null;
}

/** The shape type a v:shape refers to (o:spt or the v:shapetype it names). */
function shapeType(el: Element): { spt: string; path: string | null } {
  let spt = wa(el, 'spt') ?? '';
  let path = wa(el, 'path');
  const ref = (wa(el, 'type') ?? '').replace(/^#/, '');
  if (ref) {
    const m = /_x0000_t(\d+)/.exec(ref);
    if (!spt && m) spt = m[1];
    if (!path) {
      const doc = el.ownerDocument;
      const all = doc.getElementsByTagName('*');
      for (let i = 0; i < all.length; i++) {
        if (all[i].localName === 'shapetype' && all[i].getAttribute('id') === ref) {
          path = wa(all[i], 'path');
          if (!spt) spt = wa(all[i], 'spt') ?? '';
          break;
        }
      }
    }
  }
  return { spt, path };
}

/** A VML shape as an element of the given size. */
export async function renderVml(el: Element, w: number, hgt: number, dc: DrawContext, source: Source, env: Env, depth = 0): Promise<HTMLElement | null> {
  if (depth > 12) return null;
  const style = vmlStyle(el);
  const wrap = h('div', { class: 'docx-object docx-vml' });
  wrap.style.width = `${r2(w)}px`;
  wrap.style.height = `${r2(hgt)}px`;
  const rotation = Number(style.get('rotation') ?? 0) || 0;
  const flip = style.get('flip') ?? '';
  const transforms = [rotation ? `rotate(${r2(rotation)}deg)` : '', flip.includes('x') ? 'scaleX(-1)' : '', flip.includes('y') ? 'scaleY(-1)' : ''].filter(Boolean);
  if (transforms.length) wrap.style.transform = transforms.join(' ');

  if (el.localName === 'group') {
    const [cw, ch] = (wa(el, 'coordsize') ?? `${w},${hgt}`).split(',').map(Number);
    const [ox, oy] = (wa(el, 'coordorigin') ?? '0,0').split(',').map(Number);
    const sx = cw ? w / cw : 1, sy = ch ? hgt / ch : 1;
    for (const c of kids(el)) {
      if (!['shape', 'rect', 'roundrect', 'oval', 'line', 'polyline', 'group', 'image'].includes(c.localName)) continue;
      const cs = vmlStyle(c);
      const num = (v: string | undefined) => (v ? Number(v.replace(/[a-z]+$/i, '')) || 0 : 0);
      const cw2 = num(cs.get('width')) * sx, ch2 = num(cs.get('height')) * sy;
      const child = await renderVml(c, Math.max(cw2, 1), Math.max(ch2, 1), dc, source, env, depth + 1);
      if (!child) continue;
      child.style.position = 'absolute';
      child.style.left = `${r2((num(cs.get('left')) - (ox || 0)) * sx)}px`;
      child.style.top = `${r2((num(cs.get('top')) - (oy || 0)) * sy)}px`;
      wrap.append(child);
    }
    return wrap;
  }

  const paint = paintOf(el);
  let defs: SVGDefsElement | null = null;
  const canvas = svg('svg', { class: 'docx-vml-svg', width: r2(Math.max(w, 1)), height: r2(Math.max(hgt, 1)) });
  const getDefs = () => (defs ??= canvas.insertBefore(svg('defs'), canvas.firstChild));
  const { spt, path } = el.localName === 'shape' ? shapeType(el) : { spt: '', path: null };
  const imagedata = kids(el).find((c) => c.localName === 'imagedata') ?? null;
  const textpath = kids(el).find((c) => c.localName === 'textpath') ?? null;
  const textbox = vmlTextbox(el);
  let drawn = false;

  const draw = (shape: SVGElement, closed = true) => {
    applyPaint(shape, paint, getDefs, closed);
    canvas.append(shape);
    drawn = true;
  };
  switch (el.localName) {
    case 'rect':
      draw(svg('rect', { x: 0, y: 0, width: r2(w), height: r2(hgt) }));
      break;
    case 'roundrect': {
      const arc = fraction(wa(el, 'arcsize'), 0.2);
      const rr = (Math.min(w, hgt) * arc) / 2;
      draw(svg('rect', { x: 0, y: 0, width: r2(w), height: r2(hgt), rx: r2(rr), ry: r2(rr) }));
      break;
    }
    case 'oval':
      draw(svg('ellipse', { cx: r2(w / 2), cy: r2(hgt / 2), rx: r2(w / 2), ry: r2(hgt / 2) }));
      break;
    case 'line': {
      const [x1, y1] = (wa(el, 'from') ?? '0,0').split(',').map((v) => vmlLength(v));
      const [x2, y2] = (wa(el, 'to') ?? '0,0').split(',').map((v) => vmlLength(v));
      const minX = Math.min(x1, x2), minY = Math.min(y1, y2);
      const line = svg('line', { x1: r2(x1 - minX), y1: r2(y1 - minY), x2: r2(x2 - minX), y2: r2(y2 - minY) });
      applyPaint(line, paint, getDefs, false);
      if (paint.stroke && paint.endArrow) line.setAttribute('marker-end', arrow(getDefs, paint.stroke, false));
      if (paint.stroke && paint.startArrow) line.setAttribute('marker-start', arrow(getDefs, paint.stroke, true));
      canvas.append(line);
      drawn = true;
      break;
    }
    case 'polyline': {
      const pts = (wa(el, 'points') ?? '').split(/[\s,]+/).map((v) => vmlLength(v));
      const pairs: string[] = [];
      for (let i = 0; i + 1 < pts.length && pairs.length < 5000; i += 2) pairs.push(`${r2(pts[i])},${r2(pts[i + 1])}`);
      draw(svg('polyline', { points: pairs.join(' ') }), false);
      break;
    }
    case 'image':
    case 'shape': {
      const ellipse = spt === '3';
      const round = spt === '2';
      if (spt === '136' && textpath) break;
      if (path && spt !== '75' && spt !== '202' && spt !== '1' && !ellipse && !round) {
        const [cw, ch] = (wa(el, 'coordsize') ?? '21600,21600').split(',').map(Number);
        const [ox, oy] = (wa(el, 'coordorigin') ?? '0,0').split(',').map(Number);
        const d = vmlPath(path, cw ? w / cw : 1, ch ? hgt / ch : 1, ox || 0, oy || 0);
        if (d) {
          draw(svg('path', { d }), /[zZ]/.test(d));
          break;
        }
      }
      if (imagedata || el.localName === 'image') break;
      // Text boxes and unknown shapes are drawn as their box.
      if (ellipse) draw(svg('ellipse', { cx: r2(w / 2), cy: r2(hgt / 2), rx: r2(w / 2), ry: r2(hgt / 2) }));
      else if (round) draw(svg('rect', { x: 0, y: 0, width: r2(w), height: r2(hgt), rx: r2(Math.min(w, hgt) / 6) }));
      else if (paint.fill || paint.stroke) draw(svg('rect', { x: 0, y: 0, width: r2(w), height: r2(hgt) }));
      break;
    }
  }
  if (drawn) wrap.append(canvas);

  // Picture.
  const rid = relId(imagedata, 'id') ?? relId(el, 'id');
  if ((imagedata || el.localName === 'image') && rid) {
    const rel = source.rels.get(rid);
    if (rel && !rel.external) {
      const crop = h('div', { class: 'docx-vml-crop' });
      const img = h('img', { alt: wa(imagedata, 'title') ?? '', draggable: false });
      const l = fraction(wa(imagedata, 'cropleft'), 0), t = fraction(wa(imagedata, 'croptop'), 0);
      const r = fraction(wa(imagedata, 'cropright'), 0), b = fraction(wa(imagedata, 'cropbottom'), 0);
      const fullW = w / Math.max(0.01, 1 - l - r), fullH = hgt / Math.max(0.01, 1 - t - b);
      Object.assign(img.style, { left: `${r2(-l * fullW)}px`, top: `${r2(-t * fullH)}px`, width: `${r2(fullW)}px`, height: `${r2(fullH)}px` });
      if (vmlBool(wa(imagedata, 'grayscale'), false)) img.style.filter = 'grayscale(1)';
      crop.append(img);
      wrap.append(crop);
      env.pending.push(
        loadBlip(env, [rel.target]).then(async (media) => {
          if (media.url) {
            img.src = media.url;
            await img.decode().catch(() => undefined);
          } else {
            img.remove();
            crop.append(h('div', { class: 'pptx-missing' }, h('span', {}, media.unsupported && media.unsupported !== 'unknown' ? `${media.unsupported} image` : 'Image not available')));
          }
        }),
      );
    }
  }

  // WordArt: the string stretched over the shape.
  if (textpath && spt === '136') {
    const text = wa(textpath, 'string') ?? '';
    const tstyle = vmlStyle(textpath);
    const family = (tstyle.get('font-family') ?? 'Calibri').replace(/["']/g, '').replace(/[;{}<>\\]/g, '');
    const word = svg('svg', { class: 'docx-vml-svg', width: r2(w), height: r2(hgt), viewBox: `0 0 ${r2(w)} ${r2(hgt)}`, preserveAspectRatio: 'none' });
    const t = svg('text', { x: 0, y: r2(hgt * 0.8), 'font-size': r2(hgt * 0.95), textLength: r2(w), lengthAdjust: 'spacingAndGlyphs', 'font-family': `"${family}", sans-serif` });
    t.textContent = text;
    t.setAttribute('fill', paint.fill ?? '#c0c0c0');
    if (paint.fillOpacity < 1) t.setAttribute('fill-opacity', String(r2(paint.fillOpacity)));
    if (paint.stroke && vmlBool(wa(el, 'stroked'), false)) t.setAttribute('stroke', paint.stroke);
    word.append(t);
    wrap.append(word);
  }

  // Text box with Word content.
  if (textbox) {
    const tb = kids(el).find((c) => c.localName === 'textbox') ?? null;
    const inset = (wa(tb, 'inset') ?? '').split(',');
    const pad = (i: number, fallback: number) => (inset[i] !== undefined && inset[i].trim() !== '' ? vmlLength(inset[i].trim()) : fallback);
    const l = pad(0, 9.6), t = pad(1, 4.8), r = pad(2, 9.6), b = pad(3, 4.8);
    let blocks = dc.parsed.get(textbox);
    if (!blocks) {
      blocks = parseBlocks(textbox, createContext(dc.doc, source, dc.flow.notes, false));
      dc.parsed.set(textbox, blocks);
    }
    const width = Math.max(1, w - l - r);
    const layout = layoutStory(dc.flow, blocks, width);
    const inner = h('div', { class: 'docx-story' });
    const anchor = style.get('v-text-anchor') ?? 'top';
    let dy = t;
    const free = hgt - t - b - layout.height;
    if (/middle/.test(anchor)) dy += free / 2;
    else if (/bottom/.test(anchor)) dy += free;
    Object.assign(inner.style, { left: `${r2(l)}px`, top: `${r2(dy)}px`, width: `${r2(width)}px` });
    dc.drawStory(layout, inner, source);
    wrap.append(inner);
  }
  return wrap.childElementCount ? wrap : null;
}
