// SPDX-License-Identifier: MIT
// Drawings in Word documents. DrawingML pictures, shapes, groups and canvases go through the PowerPoint shape
// engine (src/pptx), with Word's text boxes laid out by our own text layout. Charts and SmartArt reuse
// the PowerPoint renderers too; legacy check boxes and ballot boxes are drawn here.
import { h } from '../dom.js';
import { overrideTheme } from '../pptx/color.js';
import { renderChart } from '../pptx/chart.js';
import { type Env, type Source, r2, svg } from '../pptx/env.js';
import { type Box, type Tree, renderShape, renderTree, rootFrame } from '../pptx/shapes.js';
import { renderDiagram } from '../pptx/slide.js';
import { attr, child, descendant, path, relAttr } from '../pptx/xml.js';
import type { DocxDocument } from './document.js';
import { type FlowContext, type StoryLayout, layoutStory } from './flow.js';
import { renderMath } from './math.js';
import type { DrawingObj } from './model.js';
import { type ParseContext, createContext, parseBlocks } from './parse.js';
import { renderVml } from './vml-render.js';
import { clampNum, kid, kids, wa } from './xml.js';

export interface DrawContext {
  doc: DocxDocument;
  flow: FlowContext;
  env: Env;
  /** Draws laid out story content (text box contents) into a container. */
  drawStory(layout: StoryLayout, container: HTMLElement, source: Source): void;
  /** Parse contexts for text box contents, by story part. */
  parsed: WeakMap<Element, ReturnType<typeof parseBlocks>>;
}

const EMU = 9525;

/** Copy of a shape with its xfrm placed at the origin and sized to the drawing's extent. */
function placed(el: Element, w: number, h: number): Element {
  const copy = el.cloneNode(true) as Element;
  const spPr = child(copy, 'spPr') ?? child(copy, 'grpSpPr');
  const xfrm = child(spPr, 'xfrm');
  if (xfrm) {
    const off = child(xfrm, 'off'), ext = child(xfrm, 'ext');
    off?.setAttribute('x', '0');
    off?.setAttribute('y', '0');
    if (copy.localName !== 'wgp' && copy.localName !== 'grpSp') {
      ext?.setAttribute('cx', String(Math.round(w * EMU)));
      ext?.setAttribute('cy', String(Math.round(h * EMU)));
    } else if (ext) {
      ext.setAttribute('cx', String(Math.round(w * EMU)));
      ext.setAttribute('cy', String(Math.round(h * EMU)));
    }
  } else if (spPr) {
    // Shapes without a transform get one, so the engine knows their size.
    const ns = spPr.namespaceURI;
    const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const x = copy.ownerDocument.createElementNS(a, 'a:xfrm');
    const off = copy.ownerDocument.createElementNS(a, 'a:off');
    off.setAttribute('x', '0');
    off.setAttribute('y', '0');
    const ext = copy.ownerDocument.createElementNS(a, 'a:ext');
    ext.setAttribute('cx', String(Math.round(w * EMU)));
    ext.setAttribute('cy', String(Math.round(h * EMU)));
    x.append(off, ext);
    spPr.insertBefore(x, spPr.firstChild);
    void ns;
  }
  return copy;
}

function frameBox(w: number, h: number): Box {
  return { x: 0, y: 0, w, h, rot: 0, flipH: false, flipV: false, emuW: Math.round(w * EMU), emuH: Math.round(h * EMU) };
}

// ---- Text boxes ---------------------------------------------------------------------------------------------

interface LinkedChain {
  /** Content of the first box. */
  content: Element;
  source: Source;
  /** Inner heights of the boxes by sequence number. */
  heights: Map<number, number>;
  width: number;
  layout?: StoryLayout;
}

const chains = new WeakMap<DocxDocument, Map<string, LinkedChain>>();

function bodyInsets(bodyPr: Element | null) {
  const n = (name: string, fallback: number) => {
    const v = Number(attr(bodyPr, name));
    return Number.isFinite(v) && attr(bodyPr, name) !== null ? clampNum(v, 0, 50_000_000) / EMU : fallback / EMU;
  };
  return { l: n('lIns', 91440), t: n('tIns', 45720), r: n('rIns', 91440), b: n('bIns', 45720) };
}

/** Index of linked text boxes (wps:txbx with an id, wps:linkedTxbx id/seq) in the main document. */
function linkedChains(dc: DrawContext): Map<string, LinkedChain> {
  let map = chains.get(dc.doc);
  if (map) return map;
  map = new Map();
  chains.set(dc.doc, map);
  const all = dc.doc.body.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    if (el.localName !== 'txbx' && el.localName !== 'linkedTxbx') continue;
    const id = el.getAttribute('id');
    const sp = el.parentElement;
    if (!id || !sp || sp.localName !== 'wsp') continue;
    const ext = path(sp, 'spPr', 'xfrm', 'ext');
    const insets = bodyInsets(child(sp, 'bodyPr'));
    const height = Math.max(0, Number(attr(ext, 'cy') ?? 0) / EMU - insets.t - insets.b);
    const width = Math.max(1, Number(attr(ext, 'cx') ?? 0) / EMU - insets.l - insets.r);
    let chain = map.get(id);
    if (el.localName === 'txbx') {
      const content = kid(el, 'txbxContent');
      if (!content) continue;
      chain = { content, source: dc.doc.source, heights: chain?.heights ?? new Map(), width };
      chain.heights.set(0, height);
      map.set(id, chain);
    } else {
      if (!chain) {
        chain = { content: null as unknown as Element, source: dc.doc.source, heights: new Map(), width };
        map.set(id, chain);
      }
      chain.heights.set(Number(el.getAttribute('seq') ?? 1), height);
    }
  }
  return map;
}

function blocksOf(content: Element, dc: DrawContext, source: Source) {
  let blocks = dc.parsed.get(content);
  if (!blocks) {
    const pc: ParseContext = createContext(dc.doc, source, dc.flow.notes, false);
    blocks = parseBlocks(content, pc);
    dc.parsed.set(content, blocks);
  }
  return blocks;
}

/** Lines of a story between two heights, moved up to start at 0 (one box of a linked chain). */
function sliceStory(layout: StoryLayout, from: number, to: number): StoryLayout {
  const items = layout.items
    .filter((item) => {
      const top = item.y;
      const bottom = item.t === 'line' ? item.y + item.line.h : item.t === 'group' || item.t === 'fill' ? item.y + item.h : item.y;
      return top >= from - 0.01 && bottom <= to + 0.01;
    })
    .map((item) => ({ ...item, y: item.y - from }));
  return { items, floats: from === 0 ? layout.floats : [], height: Math.min(to, layout.height) - from, width: layout.width };
}

/** Where the boxes of a chain cut the text: line boundaries that fit each box's height. */
function chainCuts(layout: StoryLayout, heights: number[]): number[] {
  const bottoms = layout.items.filter((i) => i.t === 'line').map((i) => (i.t === 'line' ? [i.y, i.y + i.line.h] : [0, 0]));
  const cuts = [0];
  let start = 0;
  for (const h of heights) {
    let end = start;
    for (const [top, bottom] of bottoms) if (top >= start - 0.01 && bottom <= start + h + 0.01) end = Math.max(end, bottom);
    if (end === start) end = start + h;
    cuts.push(end);
    start = end;
  }
  return cuts;
}

/** Word text box content (wps:txbx or a linked box) inside a shape's text rectangle. */
function textBox(sp: Element, rect: { x: number; y: number; w: number; h: number; rot: number }, dc: DrawContext, source: Source): HTMLElement | null {
  const bodyPr = child(sp, 'bodyPr');
  const insets = bodyInsets(bodyPr);
  const txbx = child(sp, 'txbx');
  const linked = child(sp, 'linkedTxbx');
  const vert = attr(bodyPr, 'vert') ?? 'horz';
  const vertical = vert === 'vert' || vert === 'vert270' || vert === 'eaVert';
  const outerW = vertical ? rect.h : rect.w, outerH = vertical ? rect.w : rect.h;
  const width = Math.max(1, outerW - insets.l - insets.r);
  const innerH = Math.max(0, outerH - insets.t - insets.b);
  let layout: StoryLayout | null = null;
  const id = txbx?.getAttribute('id') ?? linked?.getAttribute('id') ?? null;
  const chain = id !== null ? linkedChains(dc).get(id) : undefined;
  if (chain && chain.content && chain.heights.size > 1) {
    chain.layout ??= layoutStory(dc.flow, blocksOf(chain.content, dc, chain.source), chain.width);
    const seqs = [...chain.heights.keys()].sort((a, b) => a - b);
    const cuts = chainCuts(chain.layout, seqs.map((s) => chain.heights.get(s)!));
    const seq = linked ? Number(linked.getAttribute('seq') ?? 1) : 0;
    const index = seqs.indexOf(seq);
    if (index >= 0) layout = sliceStory(chain.layout, cuts[index], cuts[index + 1] ?? chain.layout.height);
  } else {
    const content = kid(txbx, 'txbxContent');
    if (!content) return null;
    layout = layoutStory(dc.flow, blocksOf(content, dc, source), attr(bodyPr, 'wrap') === 'none' ? 4000 : width);
  }
  if (!layout || !layout.items.length) return null;
  const box = h('div', { class: 'docx-textbox' });
  const s = box.style;
  s.left = `${r2(rect.x + (rect.w - outerW) / 2)}px`;
  s.top = `${r2(rect.y + (rect.h - outerH) / 2)}px`;
  s.width = `${r2(outerW)}px`;
  s.height = `${r2(outerH)}px`;
  const anchor = attr(bodyPr, 'anchor') ?? 't';
  let dy = insets.t;
  if (anchor === 'ctr') dy += (innerH - layout.height) / 2;
  else if (anchor === 'b') dy += innerH - layout.height;
  const inner = h('div', { class: 'docx-story' });
  inner.style.left = `${r2(insets.l)}px`;
  inner.style.top = `${r2(dy)}px`;
  inner.style.width = `${r2(width)}px`;
  dc.drawStory(layout, inner, source);
  box.append(inner);
  if (vertical) s.transform = `rotate(${vert === 'vert270' ? -90 : 90}deg)`;
  return box;
}

// ---- Objects ------------------------------------------------------------------------------------------------

function tree(dc: DrawContext, source: Source, env: Env): Tree {
  return {
    env,
    source,
    layers: () => [],
    skipPlaceholders: false,
    depth: 0,
    graphicFrame: async (frameEl, box, t) => {
      const data = path(frameEl, 'graphic', 'graphicData');
      return graphicContent(data, box.w, box.h, dc, t.source, env, box);
    },
    text: (sp, rect, _box, t) => textBox(sp, rect, dc, t.source),
  };
}

/** Chart parts carry their own theme overrides; everything else uses the document's. */
async function chart(data: Element, w: number, hgt: number, dc: DrawContext, source: Source, env: Env, box?: Box): Promise<HTMLElement | null> {
  const rel = source.rels.get(relAttr(child(data, 'chart'), 'id') ?? '');
  if (!rel || rel.external) return null;
  const doc = await dc.doc.pkg.xml(rel.target);
  if (!doc) return null;
  let theme = env.theme;
  for (const r of (await dc.doc.pkg.rels(rel.target)).values()) if (r.type === 'themeOverride' && !r.external) theme = overrideTheme(theme, await dc.doc.pkg.xml(r.target));
  const container = h('div', { class: 'docx-chart' });
  const b = box ?? frameBox(w, hgt);
  Object.assign(container.style, { left: `${r2(b.x)}px`, top: `${r2(b.y)}px`, width: `${r2(w)}px`, height: `${r2(hgt)}px` });
  renderChart(doc.documentElement, container, { ...b, x: 0, y: 0 }, { ...env, theme, colors: { ...env.colors, theme } });
  return container;
}

async function graphicContent(data: Element | null, w: number, hgt: number, dc: DrawContext, source: Source, env: Env, box?: Box): Promise<HTMLElement | null> {
  if (!data) return null;
  const uri = attr(data, 'uri') ?? '';
  const t = tree(dc, source, env);
  if (uri.endsWith('/chart')) return chart(data, w, hgt, dc, source, env, box);
  if (uri.endsWith('/diagram')) return renderDiagram(data, box ?? frameBox(w, hgt), t, dc.doc.pkg);
  for (const el of kids(data)) {
    switch (el.localName) {
      case 'pic':
      case 'wsp':
      case 'wgp':
        return renderShape(placed(el, w, hgt), rootFrame(w, hgt), t);
      case 'wpc':
      case 'lockedCanvas': {
        // Canvas coordinates are EMU from its top left corner.
        const container = h('div', { class: 'docx-canvas' });
        Object.assign(container.style, { width: `${r2(w)}px`, height: `${r2(hgt)}px` });
        const bg = child(el, 'bg');
        const whole = child(el, 'whole');
        if (bg || whole) {
          const fake = el.ownerDocument.createElementNS('http://schemas.microsoft.com/office/word/2010/wordprocessingShape', 'wps:wsp');
          const spPr = el.ownerDocument.createElementNS('http://schemas.microsoft.com/office/word/2010/wordprocessingShape', 'wps:spPr');
          for (const c of [...(bg?.children ?? []), ...(whole?.children ?? [])]) spPr.append(c.cloneNode(true));
          fake.append(spPr);
          const back = await renderShape(placed(fake, w, hgt), rootFrame(w, hgt), t);
          if (back) container.append(back);
        }
        await renderTree(el, container, rootFrame(w, hgt), t);
        return container;
      }
    }
  }
  // OLE objects and other content usually carry a picture of themselves.
  const pic = descendant(data, 'pic');
  return pic ? renderShape(placed(pic, w, hgt), rootFrame(w, hgt), t) : null;
}

/** A legacy form check box: a square with a cross when checked, as Word draws it. */
function checkbox(obj: DrawingObj): HTMLElement {
  const size = obj.w;
  const el = svg('svg', { class: 'docx-checkbox', width: r2(size), height: r2(size), viewBox: '0 0 20 20' });
  el.append(svg('rect', { x: 1, y: 1, width: 18, height: 18, fill: '#fff', stroke: '#000', 'stroke-width': 1.2 }));
  if (obj.graphic.kind === 'checkbox' && obj.graphic.checked) el.append(svg('path', { d: 'M1 1L19 19M19 1L1 19', stroke: '#000', 'stroke-width': 1.2 }));
  const wrap = h('div', { class: 'docx-object' });
  wrap.append(el);
  return wrap;
}

/**
 * Element for a drawing object of the given size; positioned by the caller. Resolves when pictures are in
 * (through env.pending).
 */
export async function renderObject(obj: DrawingObj, dc: DrawContext, env: Env): Promise<HTMLElement | null> {
  const g = obj.graphic;
  const w = obj.w, hgt = obj.h;
  switch (g.kind) {
    case 'checkbox':
      return checkbox(obj);
    case 'math':
      return renderMath(g.el, obj);
    case 'vml':
      return renderVml(g.el, w, hgt, dc, g.source, env);
    case 'dml': {
      const data = kid(g.el, 'graphicData') ?? child(g.el, 'graphicData');
      const el = await graphicContent(data, w, hgt, dc, g.source, env);
      if (!el) return null;
      const wrap = h('div', { class: 'docx-object' });
      if (obj.descr) wrap.title = obj.descr;
      wrap.append(el);
      return wrap;
    }
    default:
      return null;
  }
}

export { textBox, wa };
