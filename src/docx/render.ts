// SPDX-License-Identifier: MIT
// DOM output of laid out pages: every line fragment, fill, border and object is positioned absolutely in page
// pixels, so the browser draws exactly what the layout computed. Headers and footers lie under the body text,
// floating objects behind or in front of it in their z order.
import { h } from '../dom.js';
import type { Env, Source } from '../pptx/env.js';
import { svg } from '../pptx/env.js';
import { loadBlip } from '../pptx/fill.js';
import { type DrawContext, renderObject } from './drawing.js';
import { type FlowContext, type Item, type PageBox, type PageFlow, type PlacedFloat, type StoryLayout, borderWidth } from './flow.js';
import type { Frag, Line } from './line.js';
import { pageNumberText } from './line.js';
import { domBaseline } from './measure.js';
import type { DocxDocument } from './document.js';
import type { Block } from './model.js';
import { textEffects } from './effects.js';
import { type Border, type RunProps, isArtBorder, parseBorder, visible } from './props.js';
import { vmlColor } from './vml-render.js';
import { hexColor, kids, r2, wa } from './xml.js';

const HIGHLIGHT: Record<string, string> = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000', darkBlue: '#000080',
  darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080',
  lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};

const BORDER_STYLES: Record<string, string> = {
  single: 'solid', thick: 'solid', double: 'double', triple: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed',
  dotDash: 'dashed', dotDotDash: 'dashed', dashDotStroked: 'dashed', wave: 'solid', doubleWave: 'double', inset: 'inset', outset: 'outset',
  threeDEmboss: 'ridge', threeDEngrave: 'groove',
};

export function borderCss(b: Border): string {
  // Art borders are pictures we do not have: a chain of dots of the same width stands in for them.
  if (isArtBorder(b)) return `${r2(borderWidth(b))}px dotted ${b.color ?? '#8c8c8c'}`;
  const style = /^(thinThick|thickThin)/.test(b.style) ? 'double' : (BORDER_STYLES[b.style] ?? 'solid');
  return `${r2(borderWidth(b))}px ${style} ${b.color ?? '#000000'}`;
}

type ThemeColors = Map<string, { r: number; g: number; b: number }>;

/** A Word theme color (w:themeColor with w:themeTint / w:themeShade as hex bytes). */
function themeColor(theme: ThemeColors, name: string, tintHex?: string | null, shadeHex?: string | null): string | null {
  const slot = { text1: 'dk1', text2: 'dk2', background1: 'lt1', background2: 'lt2', dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', hyperlink: 'hlink', followedHyperlink: 'folHlink' }[name] ?? name;
  const c = theme.get(slot);
  if (!c) return null;
  let { r, g, b } = c;
  const tint = tintHex ? parseInt(tintHex, 16) / 255 : null;
  const shade = shadeHex ? parseInt(shadeHex, 16) / 255 : null;
  if (tint !== null) [r, g, b] = [r, g, b].map((v) => v * tint + 255 * (1 - tint));
  if (shade !== null) [r, g, b] = [r, g, b].map((v) => v * shade);
  return `rgb(${Math.round(r)}, ${Math.round(g)}, ${Math.round(b)})`;
}

/** Text color of a run: its own, or black (white on dark shading) for automatic color. */
function textColor(run: RunProps, theme: ThemeColors): string {
  if (run.color && run.color !== 'auto') return run.color;
  if (run.themeColor && !run.color) {
    const c = themeColor(theme, run.themeColor, run.themeTint, run.themeShade);
    if (c) return c;
  }
  const fill = run.shd?.fill ?? run.shd?.color;
  if (fill) {
    const n = parseInt(fill.slice(1), 16);
    const lum = 0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255);
    if (lum < 100) return '#ffffff';
  }
  return '#000000';
}

const UNDERLINE: Record<string, [string, string]> = {
  single: ['solid', 'auto'], words: ['solid', 'auto'], double: ['double', 'auto'], thick: ['solid', '0.1em'], dotted: ['dotted', 'auto'],
  dottedHeavy: ['dotted', '0.1em'], dash: ['dashed', 'auto'], dashedHeavy: ['dashed', '0.1em'], dashLong: ['dashed', 'auto'],
  dashLongHeavy: ['dashed', '0.1em'], dotDash: ['dashed', 'auto'], dashDotHeavy: ['dashed', '0.1em'], dotDotDash: ['dotted', 'auto'],
  dashDotDotHeavy: ['dotted', '0.1em'], wave: ['wavy', 'auto'], wavyHeavy: ['wavy', '0.1em'], wavyDouble: ['wavy', 'auto'],
};

export interface RenderShared {
  doc: DocxDocument;
  flow: FlowContext;
  pages: PageBox[];
  pageFlow: PageFlow;
  /** Parsed text box contents. */
  parsed: WeakMap<Element, Block[]>;
}

class PageRenderer {
  readonly dc: DrawContext;
  private readonly theme: Map<string, { r: number; g: number; b: number }>;

  constructor(
    readonly shared: RenderShared,
    readonly env: Env,
    readonly page: PageBox,
  ) {
    this.theme = shared.doc.theme.colors as unknown as Map<string, { r: number; g: number; b: number }>;
    this.dc = {
      doc: shared.doc,
      flow: shared.flow,
      env,
      parsed: shared.parsed,
      drawStory: (layout, container) => this.story(layout, container, 0, 0),
    };
  }

  // ---- Text ----------------------------------------------------------------------------------------------

  private styleText(el: HTMLElement, frag: Frag): void {
    const { style } = frag;
    const run = style.run;
    const s = el.style;
    s.font = style.font.css;
    if (!style.font.kern) s.fontKerning = 'none';
    if (!style.font.ligatures) s.fontVariantLigatures = 'none';
    s.color = textColor(run, this.theme);
    if (style.spacing || frag.charSpacing) s.letterSpacing = `${r2(style.spacing + (frag.charSpacing ?? 0))}px`;
    if (style.scale !== 1) {
      s.transform = `scaleX(${r2(style.scale)})`;
      s.transformOrigin = '0 0';
    }
    if (frag.wordSpacing) s.wordSpacing = `${r2(frag.wordSpacing)}px`;
    const deco: string[] = [];
    if (run.u && run.u !== 'none') {
      deco.push('underline');
      const [kind, thickness] = UNDERLINE[run.u] ?? ['solid', 'auto'];
      s.textDecorationStyle = kind;
      if (thickness !== 'auto') s.textDecorationThickness = thickness;
      if (run.uColor) s.textDecorationColor = run.uColor;
      s.textUnderlineOffset = '0.12em';
    }
    if (run.strike || run.dstrike) {
      deco.push('line-through');
      if (run.dstrike && !run.u) s.textDecorationStyle = 'double';
    }
    if (deco.length) s.textDecorationLine = deco.join(' ');
    const bg = run.highlight && run.highlight !== 'none' ? HIGHLIGHT[run.highlight] : run.shd ? (run.shd.val === 'solid' ? run.shd.color : run.shd.fill) : null;
    if (bg) s.backgroundColor = bg;
    if (run.emboss || run.imprint) s.textShadow = '0 1px 0 rgba(0,0,0,0.25)';
    if (run.outline) {
      s.webkitTextStroke = `0.6px ${s.color}`;
      s.color = 'transparent';
    }
    if (run.textFill || run.textOutline || run.textShadow) {
      const fx = textEffects(run, this.env.colors);
      if (fx.color) s.color = fx.color;
      if (fx.stroke) s.webkitTextStroke = fx.stroke;
      if (fx.gradient) {
        s.backgroundImage = fx.gradient;
        s.backgroundClip = 'text';
        s.webkitBackgroundClip = 'text';
        s.color = 'transparent';
        // A text-shadow would be painted over the gradient; a filter shadows the finished glyphs.
        if (fx.shadow) s.filter = `drop-shadow(${fx.shadow})`;
      } else if (fx.shadow) s.textShadow = fx.shadow;
    }
  }

  /** Small capitals drawn as capitals at 70 %, matching how they were measured. */
  private smallCaps(el: HTMLElement, text: string): void {
    for (const part of text.match(/\p{Ll}+|[^\p{Ll}]+/gu) ?? []) {
      if (/\p{Ll}/u.test(part)) {
        const small = document.createElement('span');
        small.style.fontSize = '70%';
        small.textContent = part.toUpperCase();
        el.append(small);
      } else el.append(part);
    }
  }

  private fieldText(frag: Frag): string {
    const f = frag.field!;
    const pn = this.shared.flow.pageNumber;
    if (f.field === 'page') return pageNumberText(this.page.number, f.format ?? this.page.section.pgNumFmt);
    if (f.field === 'numpages') return pageNumberText(pn.total, f.format ?? 'decimal');
    if (f.field === 'pageref') return (f.arg && this.shared.flow.bookmarks.get(f.arg)) || f.cached;
    return pageNumberText(pn.sectionPages, f.format ?? 'decimal');
  }

  private wrapLink(el: HTMLElement, frag: Frag): HTMLElement {
    const link = frag.style.link;
    if (!link) return el;
    const a = document.createElement('a');
    a.className = 'docx-link';
    if (link.href) {
      a.href = link.href;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.title = link.href;
    } else if (link.anchor) {
      a.dataset.anchor = link.anchor;
      a.tabIndex = 0;
      a.setAttribute('role', 'link');
    }
    // The link takes over the position of what it wraps.
    for (const prop of ['left', 'top', 'width', 'height'] as const) {
      a.style[prop] = el.style[prop];
      el.style[prop] = '';
    }
    el.style.position = 'static';
    el.style.display = 'inline-block';
    a.append(el);
    return a;
  }

  private line(item: { x: number; y: number; line: Line }, container: HTMLElement, dx: number, dy: number): void {
    const { line } = item;
    const baseline = dy + item.y + line.baseline;
    for (const frag of line.frags) {
      const left = dx + item.x + frag.x;
      switch (frag.k) {
        case 'text': {
          const text = frag.field ? this.fieldText(frag) : frag.text;
          if (!text) break;
          const span = document.createElement('span');
          span.className = 'docx-t';
          this.styleText(span, frag);
          if (frag.style.smallCaps) this.smallCaps(span, text);
          else span.textContent = text;
          span.style.left = `${r2(left)}px`;
          span.style.top = `${r2(baseline - frag.style.shift - domBaseline(frag.style.font.css))}px`;
          container.append(this.wrapLink(span, frag));
          break;
        }
        case 'tab': {
          if (!frag.leader && !frag.decorate) break;
          if (frag.w < 1) break;
          const span = document.createElement('span');
          span.className = 'docx-tab';
          this.styleText(span, frag);
          span.style.left = `${r2(left)}px`;
          span.style.top = `${r2(baseline - domBaseline(frag.style.font.css))}px`;
          span.style.width = `${r2(frag.w)}px`;
          span.style.letterSpacing = '';
          if (frag.leader) {
            span.style.textDecorationLine = frag.style.run.u && frag.style.run.u !== 'none' ? 'underline' : 'none';
            span.textContent = frag.leader.repeat(Math.min(2000, Math.ceil(frag.w / Math.max(1, frag.style.font.size * 0.2)) + 2));
          } else span.textContent = ' '.repeat(Math.min(2000, Math.ceil(frag.w / Math.max(1, frag.style.font.size * 0.2)) + 2));
          container.append(span);
          break;
        }
        case 'box': {
          const size = frag.style.font.size;
          const side = size * 0.78;
          const el = svg('svg', { class: 'docx-box', width: r2(side), height: r2(side), viewBox: '0 0 20 20' });
          const color = textColor(frag.style.run, this.theme);
          el.append(svg('rect', { x: 1.5, y: 1.5, width: 17, height: 17, fill: 'none', stroke: color, 'stroke-width': 1.4 }));
          if (frag.text === '☒') el.append(svg('path', { d: 'M5 5L15 15M15 5L5 15', stroke: color, 'stroke-width': 1.8, 'stroke-linecap': 'round' }));
          else if (frag.text === '☑') el.append(svg('path', { d: 'M4.5 10.5L8.5 14.5L15.5 5', fill: 'none', stroke: color, 'stroke-width': 2, 'stroke-linecap': 'round', 'stroke-linejoin': 'round' }));
          el.style.left = `${r2(left + (frag.w - side) / 2)}px`;
          el.style.top = `${r2(baseline - frag.style.shift - side + size * 0.12)}px`;
          container.append(el);
          break;
        }
        case 'sep': {
          const rule = document.createElement('div');
          rule.className = 'docx-rule';
          Object.assign(rule.style, { left: `${r2(left)}px`, top: `${r2(baseline - frag.style.font.size * 0.3)}px`, width: `${r2(frag.w)}px`, borderTop: '0.67px solid #000' });
          container.append(rule);
          break;
        }
        case 'obj': {
          const obj = frag.obj!;
          const slot = h('div', { class: 'docx-slot' });
          const top = baseline - obj.h - obj.effect.b + (obj.descent ?? 0) - (obj.graphic.kind === 'checkbox' ? frag.style.font.size * 0.02 : 0);
          Object.assign(slot.style, { left: `${r2(left + obj.effect.l + (obj.graphic.kind === 'checkbox' ? 1 : 0))}px`, top: `${r2(top)}px`, width: `${r2(obj.w)}px`, height: `${r2(obj.h)}px` });
          container.append(frag.style.link ? this.wrapLink(slot, frag) : slot);
          this.env.pending.push(
            renderObject(obj, this.dc, this.env).then((el) => {
              if (el) slot.append(el);
            }),
          );
          break;
        }
      }
    }
    if (line.bars) {
      for (const x of line.bars) {
        const bar = document.createElement('div');
        bar.className = 'docx-rule';
        Object.assign(bar.style, { left: `${r2(dx + item.x + x)}px`, top: `${r2(dy + item.y)}px`, height: `${r2(line.h)}px`, borderLeft: '0.67px solid #000' });
        container.append(bar);
      }
    }
  }

  // ---- Items ---------------------------------------------------------------------------------------------

  items(items: Item[], container: HTMLElement, dx: number, dy: number): void {
    for (const item of items) {
      switch (item.t) {
        case 'line':
          this.line(item, container, dx, dy);
          break;
        case 'fill': {
          const el = document.createElement('div');
          el.className = 'docx-fill';
          Object.assign(el.style, { left: `${r2(dx + item.x)}px`, top: `${r2(dy + item.y)}px`, width: `${r2(Math.max(0, item.w))}px`, height: `${r2(Math.max(0, item.h))}px`, background: item.color });
          container.append(el);
          break;
        }
        case 'rule': {
          if (!visible(item.b)) break;
          const el = document.createElement('div');
          el.className = 'docx-rule';
          const w = borderWidth(item.b);
          if (item.vertical) Object.assign(el.style, { left: `${r2(dx + item.x - w / 2)}px`, top: `${r2(dy + item.y)}px`, height: `${r2(item.len)}px`, borderLeft: borderCss(item.b) });
          else Object.assign(el.style, { left: `${r2(dx + item.x)}px`, top: `${r2(dy + item.y - w / 2)}px`, width: `${r2(item.len)}px`, borderTop: borderCss(item.b) });
          container.append(el);
          break;
        }
        case 'group': {
          const el = h('div', { class: 'docx-group' });
          Object.assign(el.style, { left: `${r2(dx + item.x)}px`, top: `${r2(dy + item.y)}px`, width: `${r2(item.w)}px`, height: `${r2(item.h)}px` });
          if (item.clip) el.style.overflow = 'hidden';
          let inner: HTMLElement = el;
          if (item.rotate) {
            inner = h('div', { class: 'docx-group' });
            // Local width runs along the box height.
            Object.assign(inner.style, { left: '0', top: '0', width: `${r2(item.h)}px`, height: `${r2(item.w)}px`, transformOrigin: '0 0' });
            inner.style.transform = item.rotate < 0 ? `translate(0, ${r2(item.h)}px) rotate(-90deg)` : `translate(${r2(item.w)}px, 0) rotate(90deg)`;
            el.append(inner);
          }
          this.items(item.items, inner, 0, 0);
          for (const f of item.floats) this.float(f, inner, 0, 0);
          container.append(el);
          break;
        }
      }
    }
  }

  float(f: PlacedFloat, container: HTMLElement, dx: number, dy: number): void {
    if (f.story) {
      const el = h('div', { class: 'docx-story' });
      Object.assign(el.style, { left: `${r2(dx + f.x)}px`, top: `${r2(dy + f.y)}px`, width: `${r2(f.w)}px`, height: `${r2(f.h)}px` });
      el.style.overflow = 'visible';
      this.story(f.story, el, -(f.storyOffset ?? 0), f.storyDy ?? 0);
      container.append(el);
      return;
    }
    const slot = h('div', { class: 'docx-slot' });
    Object.assign(slot.style, { left: `${r2(dx + f.x)}px`, top: `${r2(dy + f.y)}px`, width: `${r2(f.w)}px`, height: `${r2(f.h)}px` });
    container.append(slot);
    const obj = f.w !== f.obj.w || f.h !== f.obj.h ? { ...f.obj, w: f.w, h: f.h } : f.obj;
    this.env.pending.push(
      renderObject(obj, this.dc, this.env).then((el) => {
        if (el) slot.append(el);
      }),
    );
  }

  story(layout: StoryLayout, container: HTMLElement, dx: number, dy: number): void {
    const sorted = [...layout.floats].sort((a, b) => a.z - b.z);
    for (const f of sorted) if (f.behind) this.float(f, container, dx, dy);
    this.items(layout.items, container, dx, dy);
    for (const f of sorted) if (!f.behind) this.float(f, container, dx, dy);
  }

  // ---- Page ----------------------------------------------------------------------------------------------

  /** Page color (w:background) and its VML fill picture; Word shows them only when the settings ask for it. */
  pageBackground(target: HTMLElement): void {
    const { doc } = this.shared;
    const bg = doc.background;
    if (!bg || !doc.settings.displayBackground) return;
    const vml = kids(bg).find((c) => c.localName === 'background');
    const fill = kids(vml).find((c) => c.localName === 'fill');
    const theme = wa(bg, 'themeColor');
    const color = hexColor(wa(bg, 'color'))
      ?? (theme ? themeColor(this.theme, theme, wa(bg, 'themeTint'), wa(bg, 'themeShade')) : null)
      ?? vmlColor(wa(vml, 'fillcolor'), null);
    const el = h('div', { class: 'docx-page-bg' });
    if (color) el.style.backgroundColor = color;
    const rid = fill ? [...fill.attributes].find((a) => a.localName === 'id' && /relationships$/.test(a.namespaceURI ?? ''))?.value : undefined;
    const rel = rid ? doc.source.rels.get(rid) : undefined;
    const type = wa(fill, 'type') ?? 'solid';
    const picture = rel && !rel.external && (type === 'tile' || type === 'frame' || type === 'pattern') ? rel : undefined;
    if (picture) {
      this.env.pending.push(
        loadBlip(this.env, [picture.target]).then((media) => {
          if (!media.url) return;
          el.style.backgroundImage = `url("${media.url}")`;
          if (type === 'frame') el.style.backgroundSize = '100% 100%';
        }),
      );
    }
    if (color || picture) target.append(el);
  }

  pageBorders(target: HTMLElement): void {
    const el = this.page.section.borders;
    if (!el) return;
    const display = el.getAttribute('w:display') ?? [...el.attributes].find((a) => a.localName === 'display')?.value ?? 'allPages';
    const first = this.page.pageInSection === 0;
    if ((display === 'firstPage' && !first) || (display === 'notFirstPage' && first)) return;
    const fromText = ([...el.attributes].find((a) => a.localName === 'offsetFrom')?.value ?? 'text') === 'text';
    const sides: Record<string, Border | undefined> = {};
    for (const c of kids(el)) sides[c.localName] = parseBorder(c);
    const p = this.page;
    const space = (b: Border | undefined) => (b ? (b.space * 4) / 3 : 0);
    const top = fromText ? p.bodyTop - space(sides.top) - borderWidth(sides.top) : space(sides.top);
    const bottom = fromText ? p.bodyBottom + space(sides.bottom) + borderWidth(sides.bottom) : p.height - space(sides.bottom);
    const left = fromText ? p.marginLeft - space(sides.left) - borderWidth(sides.left) : space(sides.left);
    const right = fromText ? p.width - p.marginRight + space(sides.right) + borderWidth(sides.right) : p.width - space(sides.right);
    const box = h('div', { class: 'docx-page-border' });
    Object.assign(box.style, { left: `${r2(left)}px`, top: `${r2(top)}px`, width: `${r2(right - left)}px`, height: `${r2(bottom - top)}px` });
    if (visible(sides.top)) box.style.borderTop = borderCss(sides.top);
    if (visible(sides.bottom)) box.style.borderBottom = borderCss(sides.bottom);
    if (visible(sides.left)) box.style.borderLeft = borderCss(sides.left);
    if (visible(sides.right)) box.style.borderRight = borderCss(sides.right);
    target.append(box);
  }

  render(target: HTMLElement): void {
    const { page, shared } = this;
    const pn = shared.flow.pageNumber;
    pn.value = page.number;
    pn.fmt = page.section.pgNumFmt;
    pn.total = shared.pages.length;
    pn.sectionPages = shared.pages.filter((p) => p.sectionIndex === page.sectionIndex).length;
    const s = page.section;
    this.pageBackground(target);
    // Headers and footers first: they lie under the body text.
    const hf = h('div', { class: 'docx-hf' });
    if (page.header) {
      const layout = shared.pageFlow.headerLayout(page.header, page, false);
      this.story(layout, hf, page.marginLeft, (s.header * 1) / 15);
    }
    if (page.footer) {
      const layout = shared.pageFlow.headerLayout(page.footer, page, true);
      this.story(layout, hf, page.marginLeft, page.height - s.footer / 15 - layout.height);
    }
    target.append(hf);
    const sorted = [...page.floats].sort((a, b) => a.z - b.z);
    const behind = h('div', { class: 'docx-layer' });
    for (const f of sorted) if (f.behind) this.float(f, behind, 0, 0);
    target.append(behind);
    const body = h('div', { class: 'docx-layer' });
    this.items(page.items, body, 0, 0);
    if (page.notes.length) {
      let y = page.bodyBottom - page.notesHeight;
      const sep = shared.pageFlow.separatorLayout();
      this.story(sep, body, page.marginLeft, y);
      y += sep.height;
      for (const note of page.notes) {
        this.story(note.layout, body, page.marginLeft, y);
        y += note.layout.height;
      }
    }
    for (const [x, top, bottom] of page.separators) {
      const rule = h('div', { class: 'docx-rule' });
      Object.assign(rule.style, { left: `${r2(x)}px`, top: `${r2(top)}px`, height: `${r2(bottom - top)}px`, borderLeft: '0.67px solid #000' });
      body.append(rule);
    }
    target.append(body);
    const front = h('div', { class: 'docx-layer' });
    for (const f of sorted) if (!f.behind) this.float(f, front, 0, 0);
    target.append(front);
    this.pageBorders(target);
  }
}

/** Renders one page into its element; returns the work (pictures) to wait for. */
export function renderPage(page: PageBox, target: HTMLElement, shared: RenderShared): Promise<unknown>[] {
  let defs: SVGDefsElement | null = null;
  const env: Env = {
    pkg: shared.doc.pkg,
    theme: shared.doc.theme,
    colors: { theme: shared.doc.theme, clrMap: shared.doc.settings.colorMap },
    media: shared.doc.media,
    deck: { defaultText: null, tableStyles: new Map() },
    slideNumber: page.number,
    pending: [],
    budget: { shapes: 6000, text: 60000 },
    defs: () => {
      if (!defs) {
        defs = svg('defs');
        const holder = svg('svg', { class: 'pptx-defs', width: 0, height: 0, 'aria-hidden': 'true' });
        holder.append(defs);
        target.append(holder);
      }
      return defs;
    },
  };
  const renderer = new PageRenderer(shared, env, page);
  renderer.render(target);
  return env.pending;
}

export type { Source };
