// SPDX-License-Identifier: MIT
// Deck loading and slide assembly: slide → layout → master inheritance (placeholders, text styles, color map,
// background, master and layout graphics), theme, and the graphic frame content (tables, charts, SmartArt).
import { h } from '../dom.js';
import { type Theme, overrideTheme, parseColorMap, parseTheme } from './color.js';
import { renderChart } from './chart.js';
import { type DeckStyles, type Env, MediaCache, type Source, px, r2, svg } from './env.js';
import { type Fill, cssFill, ownFill, refFill } from './fill.js';
import { Package } from './package.js';
import { type Box, type Tree, renderTree, rootFrame } from './shapes.js';
import { renderTable } from './table.js';
import type { Layer } from './text.js';
import { attr, child, clamp, descendant, descendants, elements, int, path, relAttr } from './xml.js';

export interface SlideRef {
  part: string;
  hidden: boolean;
}

export interface Deck {
  pkg: Package;
  media: MediaCache;
  /** Slide size in px. */
  width: number;
  height: number;
  slides: SlideRef[];
  styles: DeckStyles;
  firstNumber: number;
  themes: Map<string, Theme>;
}

// PowerPoint allows 1 to 56 inches per side.
const MIN_SIDE = 914400 / 4;
const MAX_SIDE = 51206400;
// More slides than this would only be a way to exhaust memory; PowerPoint itself stops far earlier.
const MAX_SLIDES = 5000;

/** Reads the package and the list of slides; `progress` gets the share of slides read so far. */
export async function loadDeck(file: Blob, progress?: (fraction: number) => void): Promise<Deck> {
  // Password-protected presentations are stored as an encrypted OLE compound file instead of a ZIP.
  const head = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) {
    throw new Error('This presentation is password-protected (encrypted) and cannot be displayed.');
  }
  let pkg: Package;
  try {
    pkg = await Package.open(file);
  } catch {
    throw new Error('The file is not a valid PowerPoint presentation.');
  }
  const main = (await pkg.related('', 'officeDocument')) ?? 'ppt/presentation.xml';
  const pres = (await pkg.xml(main))?.documentElement;
  if (!pres || pres.localName !== 'presentation') {
    pkg.close();
    throw new Error('The file is not a PowerPoint presentation.');
  }
  const rels = await pkg.rels(main);
  const slides: SlideRef[] = [];
  const ids = elements(child(pres, 'sldIdLst')).slice(0, MAX_SLIDES);
  for (const [i, id] of ids.entries()) {
    progress?.(i / ids.length);
    const rel = rels.get(relAttr(id, 'id') ?? '');
    if (!rel || rel.external) continue;
    const doc = await pkg.xml(rel.target);
    if (!doc) continue;
    slides.push({ part: rel.target, hidden: attr(doc.documentElement, 'show') === '0' });
  }
  const size = child(pres, 'sldSz');
  const tableStyles = new Map<string, Element>();
  for (const rel of rels.values()) {
    if (rel.type !== 'tableStyles' || rel.external) continue;
    const doc = await pkg.xml(rel.target);
    for (const style of elements(doc?.documentElement)) {
      const id = attr(style, 'styleId');
      if (id) tableStyles.set(id, style);
    }
  }
  return {
    pkg,
    media: new MediaCache(pkg),
    width: px(clamp(int(size, 'cx') ?? 12192000, MIN_SIDE, MAX_SIDE)),
    height: px(clamp(int(size, 'cy') ?? 6858000, MIN_SIDE, MAX_SIDE)),
    slides,
    styles: { defaultText: child(pres, 'defaultTextStyle'), tableStyles },
    firstNumber: clamp(int(pres, 'firstSlideNum') ?? 1, 0, 9999),
    themes: new Map(),
  };
}

interface Sheet {
  part: string;
  root: Element;
  source: Source;
  tree: Element | null;
}

async function sheet(pkg: Package, part: string | null): Promise<Sheet | null> {
  if (!part) return null;
  const doc = await pkg.xml(part);
  if (!doc) return null;
  const root = doc.documentElement;
  return { part, root, source: { part, rels: await pkg.rels(part) }, tree: path(root, 'cSld', 'spTree') };
}

function firstRel(s: Sheet | null, type: string): string | null {
  if (!s) return null;
  for (const rel of s.source.rels.values()) if (rel.type === type && !rel.external) return rel.target;
  return null;
}

async function themeOf(deck: Deck, part: string | null): Promise<Theme> {
  const key = part ?? '';
  let theme = deck.themes.get(key);
  if (!theme) {
    theme = parseTheme(part ? await deck.pkg.xml(part) : null);
    deck.themes.set(key, theme);
  }
  return theme;
}

// ---- Placeholders -------------------------------------------------------------------------------------------

interface Placeholder {
  type: string;
  idx: string | null;
  el: Element;
}

function placeholders(tree: Element | null): Placeholder[] {
  const out: Placeholder[] = [];
  const walk = (el: Element | null, depth: number) => {
    for (const c of elements(el)) {
      if (c.localName === 'grpSp' && depth < 8) walk(c, depth + 1);
      const ph = path(c, 'nvSpPr', 'nvPr', 'ph') ?? path(c, 'nvPicPr', 'nvPr', 'ph') ?? path(c, 'nvGraphicFramePr', 'nvPr', 'ph');
      if (ph) out.push({ type: attr(ph, 'type') ?? 'obj', idx: attr(ph, 'idx'), el: c });
    }
  };
  walk(tree, 0);
  return out;
}

const TITLES = new Set(['title', 'ctrTitle']);
const FOOTERS = new Set(['dt', 'ftr', 'sldNum', 'hdr']);

/** Placeholder type as the master knows it: one title, one body, and the footer fields. */
function masterType(type: string): string {
  if (TITLES.has(type)) return 'title';
  if (FOOTERS.has(type)) return type;
  return 'body';
}

function matchLayout(list: Placeholder[], type: string, idx: string | null): Placeholder | undefined {
  if (TITLES.has(type)) return list.find((p) => TITLES.has(p.type));
  if (idx !== null) {
    const byIdx = list.find((p) => p.idx === idx);
    if (byIdx) return byIdx;
  }
  return list.find((p) => p.type === type);
}

// ---- Rendering ---------------------------------------------------------------------------------------------

function background(s: Sheet | null): Element | null {
  return s ? path(s.root, 'cSld', 'bg') : null;
}

function backgroundFill(bg: Element, s: Sheet, env: Env): Fill | undefined {
  const bgPr = child(bg, 'bgPr');
  if (bgPr) return ownFill(bgPr, env.colors, s.source);
  const ref = child(bg, 'bgRef');
  return ref ? refFill(ref, env, s.source) : undefined;
}

function showsMasterShapes(s: Sheet | null): boolean {
  return attr(s?.root, 'showMasterSp') !== '0';
}

export async function renderSlide(deck: Deck, index: number, target: HTMLElement): Promise<Promise<unknown>[]> {
  const { pkg } = deck;
  const slide = await sheet(pkg, deck.slides[index].part);
  if (!slide) throw new Error('Slide part missing');
  const layout = await sheet(pkg, firstRel(slide, 'slideLayout'));
  const master = await sheet(pkg, firstRel(layout, 'slideMaster'));
  let theme = await themeOf(deck, firstRel(master, 'theme'));
  const override = firstRel(slide, 'themeOverride') ?? firstRel(layout, 'themeOverride');
  if (override) theme = overrideTheme(theme, await pkg.xml(override));

  let clrMap = parseColorMap(child(master?.root, 'clrMap'));
  for (const s of [layout, slide]) {
    const ovr = child(child(s?.root, 'clrMapOvr'), 'overrideClrMapping');
    if (ovr) clrMap = parseColorMap(ovr, clrMap);
  }
  let defs: SVGDefsElement | null = null;
  const env: Env = {
    defs: () => {
      if (!defs) {
        defs = svg('defs');
        const holder = svg('svg', { class: 'pptx-defs', width: 0, height: 0, 'aria-hidden': 'true' });
        holder.append(defs);
        target.append(holder);
      }
      return defs;
    },
    pkg,
    theme,
    colors: { theme, clrMap },
    media: deck.media,
    deck: deck.styles,
    slideNumber: deck.firstNumber + index,
    pending: [],
    budget: { shapes: 6000, text: 60000 },
  };

  // Background: the slide's own, else the layout's, else the master's.
  const bgLayer = h('div', { class: 'pptx-bg' });
  target.append(bgLayer);
  for (const s of [slide, layout, master]) {
    const bg = background(s);
    if (!bg || !s) continue;
    const fill = backgroundFill(bg, s, env);
    cssFill(bgLayer, fill, env);
    env.background = fill;
    break;
  }

  const layoutPh = placeholders(layout?.tree ?? null);
  const masterPh = placeholders(master?.tree ?? null);
  const masterStyles = child(master?.root, 'txStyles');
  const defaultText: Layer = { list: deck.styles.defaultText };

  const inherit = (ph: Element): Element[] => {
    const type = attr(ph, 'type') ?? 'obj';
    const fromLayout = matchLayout(layoutPh, type, attr(ph, 'idx'));
    const fromMaster = masterPh.find((p) => p.type === masterType(fromLayout?.type ?? type));
    return [fromLayout?.el, fromMaster?.el].filter((e): e is Element => !!e);
  };
  const phType = (sp: Element, chain: Element[]): string | null => {
    for (const el of [sp, ...chain]) {
      const ph = path(el, 'nvSpPr', 'nvPr', 'ph');
      if (ph) return attr(ph, 'type') ?? 'obj';
    }
    return null;
  };
  const slideLayers = (sp: Element | null, chain: Element[]): Layer[] => {
    const type = sp ? phType(sp, chain) : null;
    if (!type) return [defaultText];
    const kind = TITLES.has(type) ? 'titleStyle' : FOOTERS.has(type) ? 'otherStyle' : 'bodyStyle';
    return [defaultText, { list: child(masterStyles, kind) }, ...[...chain].reverse().map((el) => ({ list: path(el, 'txBody', 'lstStyle') }))];
  };

  const frame = rootFrame(deck.width, deck.height);
  const base = (s: Sheet, skip: boolean): Tree => ({
    env,
    source: s.source,
    layers: () => [defaultText],
    skipPlaceholders: skip,
    depth: 0,
    graphicFrame: (el, box, tree) => graphicFrame(el, box, tree, deck),
  });
  if (master?.tree && showsMasterShapes(slide) && showsMasterShapes(layout)) await renderTree(master.tree, target, frame, base(master, true));
  if (layout?.tree && showsMasterShapes(slide)) await renderTree(layout.tree, target, frame, base(layout, true));
  if (slide.tree) await renderTree(slide.tree, target, frame, { ...base(slide, false), inherit, layers: slideLayers });
  return env.pending;
}

// ---- Graphic frames -----------------------------------------------------------------------------------------

function framed(box: Box, cls: string): HTMLElement {
  const el = h('div', { class: cls });
  Object.assign(el.style, { left: `${r2(box.x)}px`, top: `${r2(box.y)}px`, width: `${r2(box.w)}px`, height: `${r2(box.h)}px` });
  if (box.rot) el.style.transform = `rotate(${r2(box.rot)}deg)`;
  return el;
}

async function graphicFrame(el: Element, box: Box, tree: Tree, deck: Deck): Promise<HTMLElement | null> {
  const data = path(el, 'graphic', 'graphicData');
  const uri = attr(data, 'uri') ?? '';
  const { env } = tree;
  const base: Layer[] = [{ list: env.deck.defaultText }];
  if (uri.endsWith('/table')) {
    const tbl = child(data, 'tbl');
    return tbl ? renderTable(tbl, box, env, tree.source, base) : null;
  }
  if (uri.endsWith('/chart')) {
    const rel = tree.source.rels.get(relAttr(child(data, 'chart'), 'id') ?? '');
    if (!rel || rel.external) return null;
    const doc = await deck.pkg.xml(rel.target);
    const chartRels = await deck.pkg.rels(rel.target);
    let theme = env.theme;
    for (const r of chartRels.values()) if (r.type === 'themeOverride' && !r.external) theme = overrideTheme(theme, await deck.pkg.xml(r.target));
    const container = framed(box, 'pptx-chart');
    if (doc) renderChart(doc.documentElement, container, box, { ...env, theme, colors: { ...env.colors, theme } });
    return container;
  }
  if (uri.endsWith('/diagram')) return renderDiagram(data, box, tree, deck.pkg);
  return null;
}

/**
 * SmartArt without a stored drawing (older files) would need PowerPoint's layout engine. Showing its text in a
 * plain box at least keeps the content readable.
 */
function diagramText(model: Document, box: Box): HTMLElement {
  const el = framed(box, 'pptx-smartart');
  const list = h('ul');
  for (const pt of descendants(model.documentElement, 'pt').slice(0, 2000)) {
    const type = attr(pt, 'type') ?? 'node';
    if (type !== 'node' && type !== 'asst') continue;
    const text = descendants(child(pt, 't'), 'p')
      .map((p) => descendants(p, 't').map((t) => t.textContent ?? '').join(''))
      .join(' ')
      .trim();
    if (text) list.append(h('li', {}, text));
    if (list.childElementCount >= 60) break;
  }
  el.append(h('div', { class: 'pptx-smartart-label' }, 'SmartArt'), list);
  return el;
}

/** SmartArt: PowerPoint (and Word) store a pre-rendered drawing of the diagram, which is plain DrawingML shapes. */
export async function renderDiagram(data: Element | null, box: Box, tree: Tree, pkg: Package): Promise<HTMLElement | null> {
  const ids = descendant(data, 'relIds');
  const rels = tree.source.rels;
  const dataRel = rels.get(relAttr(ids, 'dm') ?? '');
  let drawingPart: string | null = null;
  const model = dataRel && !dataRel.external ? await pkg.xml(dataRel.target) : null;
  if (model) {
    const ext = descendant(model?.documentElement, 'dataModelExt');
    const rel = rels.get(attr(ext, 'relId') ?? '');
    if (rel && !rel.external) drawingPart = rel.target;
  }
  if (!drawingPart) {
    const drawings = [...rels.values()].filter((r) => r.type === 'diagramDrawing' && !r.external);
    if (drawings.length === 1) drawingPart = drawings[0].target;
  }
  const doc = drawingPart ? await pkg.xml(drawingPart) : null;
  const spTree = descendant(doc?.documentElement, 'spTree');
  if (!drawingPart || !spTree) return model ? diagramText(model, box) : null;
  const container = framed(box, 'pptx-grp');
  // Drawing coordinates are EMU from the frame's top left corner.
  const frame = { ...rootFrame(box.w, box.h), sx: box.emuW > 0 ? box.w / box.emuW : 1 / 9525, sy: box.emuH > 0 ? box.h / box.emuH : 1 / 9525 };
  // The drawing leaves out text colors that equal SmartArt's own default, dark 1, whatever the slide maps
  // its text color to.
  const dark = tree.env.theme.colors.get('dk1');
  const base: Layer[] = [{ list: tree.env.deck.defaultText }, ...(dark ? [{ run: { color: dark } }] : [])];
  await renderTree(spTree, container, frame, {
    ...tree,
    source: { part: drawingPart, rels: await pkg.rels(drawingPart) },
    inherit: undefined,
    layers: () => [...base],
    skipPlaceholders: false,
    depth: tree.depth + 1,
  });
  return container;
}
