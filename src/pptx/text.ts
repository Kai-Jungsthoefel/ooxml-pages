// SPDX-License-Identifier: MIT
// Text bodies: property inheritance through the list-style layers (presentation defaults, master text styles,
// placeholder list styles, shape list style, paragraph and run properties) and HTML output.
// Text sizes: 1 pt = 4/3 px at 100 %.
import { mathElement } from '../docx/math.js';
import { h } from '../dom.js';
import { type Color, colorIn, css } from './color.js';
import { type Env, type Source, px, r2 } from './env.js';
import { fillColor, ownFill } from './fill.js';
import { fontFamily, fontMetrics, isSymbolFont, symbolChar, textWidth } from './fonts.js';
import { attr, bool, child, children, clamp, descendant, elements, int, relAttr } from './xml.js';

export interface Spacing {
  pct?: number;
  pts?: number;
}

export type Bullet = { type: 'none' } | { type: 'char'; char: string } | { type: 'auto'; scheme: string; start: number } | { type: 'blip' };

export interface ParaProps {
  algn?: string;
  marL?: number;
  indent?: number;
  lnSpc?: Spacing;
  spcBef?: Spacing;
  spcAft?: Spacing;
  bullet?: Bullet;
  /** null: follow the text. */
  buFont?: string | null;
  buColor?: Color | null;
  buSize?: Spacing | null;
  defTabSz?: number;
  /** Custom tab stops: position in EMU from the left text edge, alignment l/ctr/r/dec. */
  tabs?: { pos: number; algn: string }[];
  rtl?: boolean;
}

export interface RunProps {
  sz?: number;
  b?: boolean;
  i?: boolean;
  u?: string;
  strike?: string;
  baseline?: number;
  cap?: string;
  spc?: number;
  /** null: no fill (invisible text). */
  color?: Color | null;
  highlight?: Color;
  latin?: string;
  ea?: string;
  cs?: string;
  sym?: string;
}

/** One level of inherited text formatting: a list style (lstStyle, txStyles …) and/or direct properties. */
export interface Layer {
  list?: Element | null;
  para?: ParaProps;
  run?: RunProps;
}

export interface BodyProps {
  lIns: number;
  tIns: number;
  rIns: number;
  bIns: number;
  anchor: string;
  wrap: string;
  vert: string;
  rot: number;
  numCol: number;
  spcCol: number;
  autofit: 'none' | 'norm' | 'shape';
  fontScale: number;
  lnSpcReduction: number;
}

export function defaultBody(): BodyProps {
  return {
    lIns: 91440, tIns: 45720, rIns: 91440, bIns: 45720, anchor: 't', wrap: 'square', vert: 'horz', rot: 0, numCol: 1, spcCol: 0,
    autofit: 'none', fontScale: 1, lnSpcReduction: 0,
  };
}

/** Applies a bodyPr on top of inherited body properties, attribute by attribute. */
export function mergeBodyPr(body: BodyProps, el: Element | null): void {
  if (!el) return;
  for (const name of ['lIns', 'tIns', 'rIns', 'bIns'] as const) {
    const v = int(el, name);
    if (v !== undefined) body[name] = clamp(v, 0, 50_000_000);
  }
  body.anchor = attr(el, 'anchor') ?? body.anchor;
  body.wrap = attr(el, 'wrap') ?? body.wrap;
  body.vert = attr(el, 'vert') ?? body.vert;
  body.rot = int(el, 'rot') !== undefined ? int(el, 'rot')! / 60000 : body.rot;
  body.numCol = int(el, 'numCol') !== undefined ? clamp(int(el, 'numCol')!, 1, 16) : body.numCol;
  body.spcCol = clamp(int(el, 'spcCol') ?? body.spcCol, 0, 50_000_000);
  const norm = child(el, 'normAutofit');
  if (norm) {
    body.autofit = 'norm';
    body.fontScale = clamp((int(norm, 'fontScale') ?? 100000) / 100000, 0.05, 1);
    body.lnSpcReduction = clamp((int(norm, 'lnSpcReduction') ?? 0) / 100000, 0, 0.9);
  } else if (child(el, 'spAutoFit')) {
    body.autofit = 'shape';
    body.fontScale = 1;
    body.lnSpcReduction = 0;
  } else if (child(el, 'noAutofit')) {
    body.autofit = 'none';
    body.fontScale = 1;
    body.lnSpcReduction = 0;
  }
}

function spacing(el: Element | null): Spacing | undefined {
  if (!el) return undefined;
  const pct = int(child(el, 'spcPct'), 'val');
  if (pct !== undefined) return { pct: clamp(pct / 100000, 0, 20) };
  const pts = int(child(el, 'spcPts'), 'val');
  if (pts !== undefined) return { pts: clamp(pts / 100, 0, 2000) };
  return undefined;
}

// Run properties have no relationships of their own that a color could need.
const NO_SOURCE: Source = { part: '', rels: new Map() };

export function readRunProps(el: Element | null, into: RunProps, env: Env): void {
  if (!el) return;
  const sz = int(el, 'sz');
  if (sz !== undefined) into.sz = clamp(sz, 100, 400000);
  const b = bool(el, 'b');
  if (b !== undefined) into.b = b;
  const i = bool(el, 'i');
  if (i !== undefined) into.i = i;
  into.u = attr(el, 'u') ?? into.u;
  into.strike = attr(el, 'strike') ?? into.strike;
  const baseline = int(el, 'baseline');
  if (baseline !== undefined) into.baseline = clamp(baseline / 100000, -1, 1);
  into.cap = attr(el, 'cap') ?? into.cap;
  const spc = int(el, 'spc');
  if (spc !== undefined) into.spc = clamp(spc, -40000, 40000);
  const fill = ownFill(el, env.colors, NO_SOURCE);
  if (fill?.type === 'none') into.color = null;
  // Gradient and pattern text gets its main color; picture-filled text keeps the inherited one.
  else if (fill && fillColor(fill)) into.color = fillColor(fill);
  const highlight = colorIn(child(el, 'highlight'), env.colors);
  if (highlight) into.highlight = highlight;
  for (const script of ['latin', 'ea', 'cs', 'sym'] as const) {
    const face = attr(child(el, script), 'typeface');
    if (face) into[script] = face;
  }
}

export function readParaProps(el: Element | null, into: ParaProps, env: Env): void {
  if (!el) return;
  into.algn = attr(el, 'algn') ?? into.algn;
  const marL = int(el, 'marL');
  if (marL !== undefined) into.marL = clamp(marL, -50_000_000, 50_000_000);
  const indent = int(el, 'indent');
  if (indent !== undefined) into.indent = clamp(indent, -50_000_000, 50_000_000);
  const tab = int(el, 'defTabSz');
  if (tab !== undefined) into.defTabSz = clamp(tab, 0, 50_000_000);
  const tabLst = child(el, 'tabLst');
  if (tabLst) {
    into.tabs = children(tabLst, 'tab')
      .slice(0, 64)
      .map((t) => ({ pos: clamp(int(t, 'pos') ?? 0, 0, 50_000_000), algn: attr(t, 'algn') ?? 'l' }))
      .sort((a, b) => a.pos - b.pos);
  }
  const rtl = bool(el, 'rtl');
  if (rtl !== undefined) into.rtl = rtl;
  into.lnSpc = spacing(child(el, 'lnSpc')) ?? into.lnSpc;
  into.spcBef = spacing(child(el, 'spcBef')) ?? into.spcBef;
  into.spcAft = spacing(child(el, 'spcAft')) ?? into.spcAft;
  for (const c of el.children) {
    switch (c.localName) {
      case 'buClrTx':
        into.buColor = null;
        break;
      case 'buClr':
        into.buColor = colorIn(c, env.colors);
        break;
      case 'buSzTx':
        into.buSize = null;
        break;
      case 'buSzPct':
        into.buSize = { pct: clamp((int(c, 'val') ?? 100000) / 100000, 0.25, 4) };
        break;
      case 'buSzPts':
        into.buSize = { pts: clamp((int(c, 'val') ?? 1200) / 100, 1, 4000) };
        break;
      case 'buFontTx':
        into.buFont = null;
        break;
      case 'buFont':
        into.buFont = attr(c, 'typeface');
        break;
      case 'buNone':
        into.bullet = { type: 'none' };
        break;
      case 'buChar':
        into.bullet = { type: 'char', char: (attr(c, 'char') ?? '•').slice(0, 4) };
        break;
      case 'buAutoNum':
        into.bullet = { type: 'auto', scheme: attr(c, 'type') ?? 'arabicPeriod', start: clamp(int(c, 'startAt') ?? 1, 1, 32767) };
        break;
      case 'buBlip':
        into.bullet = { type: 'blip' };
        break;
    }
  }
}

/** Effective paragraph and default run properties of a paragraph at a level (1–9). */
export function resolveLevel(layers: Layer[], level: number, pPr: Element | null, env: Env): { para: ParaProps; run: RunProps } {
  const para: ParaProps = {};
  const run: RunProps = {};
  const apply = (el: Element | null) => {
    if (!el) return;
    readParaProps(el, para, env);
    readRunProps(child(el, 'defRPr'), run, env);
  };
  for (const layer of layers) {
    if (layer.para) Object.assign(para, layer.para);
    if (layer.run) Object.assign(run, layer.run);
    if (layer.list) {
      apply(child(layer.list, 'defPPr'));
      apply(child(layer.list, `lvl${level}pPr`));
    }
  }
  apply(pPr);
  return { para, run };
}

// ---- Output ------------------------------------------------------------------------------------------------

const PT = 4 / 3;

function themeFont(name: string | undefined, env: Env): string | undefined {
  if (!name) return undefined;
  const m = /^\+(mj|mn)-(lt|ea|cs)$/.exec(name);
  if (!m) return name;
  const set = m[1] === 'mj' ? env.theme.major : env.theme.minor;
  return (m[2] === 'lt' ? set.latin : m[2] === 'ea' ? set.ea : set.cs) || undefined;
}

export function runFont(props: RunProps, env: Env): string {
  return fontFamily(themeFont(props.latin, env) ?? env.theme.minor.latin, themeFont(props.ea, env), themeFont(props.cs, env));
}

/** Font size in px with autofit scaling. */
export function runSize(props: RunProps, scale: number): number {
  return ((props.sz ?? 1800) / 100) * PT * scale;
}

function decoration(props: RunProps): string {
  const parts: string[] = [];
  if (props.u && props.u !== 'none') parts.push('underline');
  if (props.strike && props.strike !== 'noStrike') parts.push('line-through');
  return parts.join(' ');
}

export function styleRun(el: HTMLElement, props: RunProps, env: Env, scale: number): void {
  const size = runSize(props, scale);
  const s = el.style;
  s.fontSize = `${r2(props.baseline ? (size * 2) / 3 : size)}px`;
  s.fontFamily = runFont(props, env);
  if (props.b) s.fontWeight = '700';
  if (props.i) s.fontStyle = 'italic';
  const deco = decoration(props);
  if (deco) {
    s.textDecorationLine = deco;
    const u = props.u ?? '';
    if (/^dbl/.test(u) || props.strike === 'dblStrike') s.textDecorationStyle = 'double';
    else if (/^dotted/.test(u)) s.textDecorationStyle = 'dotted';
    else if (/^dash|^dotDash|^dotDotDash/.test(u)) s.textDecorationStyle = 'dashed';
    else if (/^wavy/.test(u)) s.textDecorationStyle = 'wavy';
    if (/Heavy$/.test(u) || u === 'heavy') s.textDecorationThickness = '0.12em';
  }
  if (props.baseline) s.verticalAlign = `${r2(props.baseline * size)}px`;
  if (props.cap === 'all') s.textTransform = 'uppercase';
  else if (props.cap === 'small') s.fontVariant = 'small-caps';
  if (props.spc) s.letterSpacing = `${r2((props.spc / 100) * PT * scale)}px`;
  if (props.color === null) s.color = 'transparent';
  else if (props.color) s.color = css(props.color);
  if (props.highlight) s.backgroundColor = css(props.highlight);
}

const ROMAN: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];

function roman(n: number): string {
  let out = '';
  for (const [value, digits] of ROMAN) while (n >= value) (out += digits), (n -= value);
  return out;
}

function alpha(n: number): string {
  let out = '';
  for (n = Math.max(1, n); n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(97 + ((n - 1) % 26)) + out;
  return out;
}

/** Text of an automatic number, e.g. "3." for arabicPeriod or "(c)" for alphaLcParenBoth. */
export function autoNumber(scheme: string, n: number): string {
  let digits = String(n);
  if (/^romanLc/.test(scheme)) digits = roman(n);
  else if (/^romanUc/.test(scheme)) digits = roman(n).toUpperCase();
  else if (/^alphaLc/.test(scheme)) digits = alpha(n);
  else if (/^alphaUc/.test(scheme)) digits = alpha(n).toUpperCase();
  else if (/^circleNumDb|^circleNumWd/.test(scheme) && n <= 20) digits = String.fromCodePoint(0x2460 + n - 1);
  if (/ParenBoth$/.test(scheme)) return `(${digits})`;
  if (/ParenR$/.test(scheme)) return `${digits})`;
  if (/Period$/.test(scheme)) return `${digits}.`;
  if (/Minus$/.test(scheme)) return `- ${digits} -`;
  return digits;
}

function linkTarget(rPr: Element | null, source: Source): string | null {
  const click = child(rPr, 'hlinkClick');
  const id = relAttr(click, 'id');
  const rel = id ? source.rels.get(id) : undefined;
  if (!rel || !rel.external) return null;
  return /^(https?:|mailto:)/i.test(rel.target.trim()) ? rel.target.trim() : null;
}

/** The plain characters of an equation (Office Math), for measuring the text around it. */
function mathText(el: Element): string {
  return [...el.getElementsByTagName('*')]
    .filter((e) => e.localName === 't')
    .map((e) => e.textContent ?? '')
    .join('');
}

/** Date and time fields whose format the browser can write in the field's language; others keep the saved text. */
const DATE_FIELDS: Record<string, Intl.DateTimeFormatOptions> = {
  datetime: { year: 'numeric', month: 'numeric', day: 'numeric' },
  datetime1: { year: 'numeric', month: 'numeric', day: 'numeric' },
  datetime2: { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' },
  datetime4: { year: 'numeric', month: 'long', day: 'numeric' },
  datetime10: { hour: '2-digit', minute: '2-digit', hourCycle: 'h23' },
  datetime11: { hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' },
  datetime12: { hour: 'numeric', minute: '2-digit', hourCycle: 'h12' },
  datetime13: { hour: 'numeric', minute: '2-digit', second: '2-digit', hourCycle: 'h12' },
};

/** A date field shows the current date, as PowerPoint updates it whenever the slide is shown or printed. */
function dateField(type: string, lang: string | null): string | null {
  const options = DATE_FIELDS[type];
  if (!options) return null;
  try {
    return new Intl.DateTimeFormat(lang || 'en-US', options).format(new Date());
  } catch {
    return null;
  }
}

function symbolText(text: string, font: string | undefined): string {
  if (!isSymbolFont(font)) return text;
  return [...text].map((ch) => symbolChar(font, ch) ?? ch).join('');
}

export interface TextOptions {
  env: Env;
  source: Source;
  scale?: number;
  lnSpcReduction?: number;
}

const MAX_RUN_TEXT = 100_000;

/** Where PowerPoint puts the lines of a paragraph, compared with where CSS puts them. */
interface LineModel {
  /** First baseline below the paragraph top: PowerPoint's, and what CSS produces. */
  first: number;
  firstCss: number;
  /** Space below the last baseline: PowerPoint's, and what CSS keeps. */
  below: number;
  belowCss: number;
}

interface Paragraph {
  el: HTMLElement;
  empty: boolean;
  lines: LineModel;
}

/** Space before/after a paragraph; autofit's line spacing reduction does not apply to it. */
function spacingPx(sp: Spacing | undefined, size: number): number {
  if (!sp) return 0;
  if (sp.pts !== undefined) return sp.pts * PT;
  return (sp.pct ?? 0) * size * 1.2;
}

/**
 * PowerPoint's line layout, measured against its own rendering: the pitch is the spacing times 1.2 em for every
 * font. Above 100 % the first baseline sits at three quarters of the pitch; up to 100 % it is the pitch minus the
 * font's descent, but never higher than three quarters. Below the last baseline only the descent remains.
 * CSS instead centers the font's ascent and descent in each line box.
 */
function lineModel(pitch: number, size: number, family: string): LineModel {
  const { ascent, descent } = fontMetrics(family);
  const a = ascent * size, d = descent * size;
  const firstCss = pitch / 2 + (a - d) / 2;
  const first = pitch > size * 1.2 + 0.01 ? pitch * 0.75 : Math.max(pitch * 0.75, pitch - d);
  return { first, firstCss, below: d, belowCss: pitch - firstCss };
}

interface Piece {
  text: string;
  props: RunProps;
  href: string | null;
  keepColor: boolean;
  /** An equation (m:oMath or m:oMathPara); `text` holds its characters. */
  math?: Element;
}

function canvasFont(props: RunProps, env: Env, scale: number): string {
  const size = runSize(props, scale) * (props.baseline ? 2 / 3 : 1);
  return `${props.i ? 'italic ' : ''}${props.b ? '700 ' : ''}${r2(size)}px ${runFont(props, env)}`;
}

/**
 * Tabs jump to the paragraph's tab stops, or to multiples of the default tab size. Browsers only know the
 * latter, so tabs become spacers whose widths come from measuring the text before them (first line only).
 */
function tabSpacers(pieces: Piece[], para: ParaProps, start: number, env: Env, scale: number): number[][] {
  const stops = (para.tabs ?? []).map((t) => ({ pos: px(t.pos), algn: t.algn }));
  const step = px(para.defTabSz ?? 914400) || 96;
  const parts = pieces.map((piece) => piece.text.split('\t'));
  const fonts = pieces.map((piece) => canvasFont(piece.props, env, scale));
  // Width of the text between a tab and the next one, for right-aligned and centered stops.
  const ahead = (pi: number, ti: number) => {
    let w = 0;
    for (let i = pi; i < pieces.length; i++) {
      for (let k = i === pi ? ti + 1 : 0; k < parts[i].length; k++) {
        w += textWidth(parts[i][k], fonts[i]);
        if (k < parts[i].length - 1) return w;
      }
    }
    return w;
  };
  let x = start;
  return parts.map((list, pi) =>
    list.slice(0, -1).map((part, ti) => {
      x += textWidth(part, fonts[pi]);
      const stop = stops.find((st) => st.pos > x + 0.5);
      let target = stop ? stop.pos : (Math.floor(x / step) + 1) * step;
      if (stop && stop.algn !== 'l') target = Math.max(x, target - (stop.algn === 'ctr' ? ahead(pi, ti) / 2 : ahead(pi, ti)));
      const width = Math.max(0, target - x);
      x = target;
      return width;
    }),
  );
}

type Counters = Map<number, { scheme: string; n: number }>;

function renderParagraph(p: Element, layers: Layer[], opts: TextOptions, counters: Counters, first: boolean): Paragraph {
  const { env, source } = opts;
  const scale = opts.scale ?? 1;
  const reduction = opts.lnSpcReduction ?? 0;
  const pPr = child(p, 'pPr');
  const level = clamp((int(pPr, 'lvl') ?? 0) + 1, 1, 9);
  const { para, run: base } = resolveLevel(layers, level, pPr, env);

  // Collect the runs first: tab widths depend on all of them.
  const pieces: (Piece | 'br')[] = [];
  for (const node of elements(p)) {
    // A slide with more text than this is broken or hostile; what is left is not worth the DOM.
    if (env.budget.text-- <= 0) break;
    const name = node.localName;
    if (name === 'r' || name === 'fld') {
      const props: RunProps = { ...base };
      const rPr = child(node, 'rPr');
      readRunProps(rPr, props, env);
      let text = (child(node, 't')?.textContent ?? '').slice(0, MAX_RUN_TEXT);
      if (name === 'fld' && attr(node, 'type') === 'slidenum') text = String(env.slideNumber);
      else if (name === 'fld') text = dateField(attr(node, 'type') ?? '', attr(rPr, 'lang')) ?? text;
      text = symbolText(text.replace(/\v/g, '\n'), themeFont(props.latin ?? props.sym, env));
      if (!text) continue;
      const keepColor = descendant(child(rPr, 'hlinkClick'), 'hlinkClr')?.getAttribute('val') === 'tx';
      pieces.push({ text, props, href: linkTarget(rPr, source), keepColor });
    } else if (name === 'br') {
      pieces.push('br');
    } else if (name === 'm') {
      // a14:m: the letters come out italic from MathML itself.
      const text = mathText(node);
      const math = child(node, 'oMathPara') ?? child(node, 'oMath');
      if (text) pieces.push({ text, props: { ...base, i: false, latin: 'Cambria Math' }, href: null, keepColor: false, math: math ?? undefined });
    }
  }
  const texts = pieces.filter((x): x is Piece => x !== 'br');
  const hasText = texts.length > 0;

  const endProps: RunProps = { ...base };
  readRunProps(child(p, 'endParaRPr'), endProps, env);
  const sized = hasText ? texts.map((t) => t.props) : [endProps];
  const sizes = sized.map((props) => runSize(props, scale));
  // The paragraph's own font acts as the line "strut"; the smallest size keeps it from inflating lines.
  const strut = Math.min(...sizes);
  const largest = Math.max(...sizes);
  const firstProps = sized[0];
  const largestProps = sized[sizes.indexOf(largest)];

  const el = h('div', { class: 'pptx-p' });
  const s = el.style;
  s.fontSize = `${r2(strut)}px`;
  s.fontFamily = runFont(firstProps, env);
  const lnSpc = para.lnSpc ?? { pct: 1 };
  const factor = Math.max(0.1, (lnSpc.pct ?? 1) - reduction) * 1.2;
  const pitch = lnSpc.pts !== undefined ? lnSpc.pts * PT * scale : factor * largest;
  s.lineHeight = lnSpc.pts !== undefined ? `${r2(pitch)}px` : String(r2(factor));
  const align = para.algn ?? 'l';
  s.textAlign = align === 'ctr' ? 'center' : align === 'r' ? 'right' : align === 'just' || align === 'dist' || align === 'justLow' ? 'justify' : 'left';
  if (align === 'dist') s.textAlignLast = 'justify';
  if (para.rtl) s.direction = 'rtl';
  const marL = px(para.marL ?? 0), indent = px(para.indent ?? 0);
  if (marL) s[para.rtl ? 'paddingRight' : 'paddingLeft'] = `${r2(marL)}px`;
  if (indent) s.textIndent = `${r2(indent)}px`;
  // Padding rather than margins: PowerPoint adds the space after one paragraph to the space before the next.
  if (!first) {
    const before = spacingPx(para.spcBef, largest);
    if (before) s.paddingTop = `${r2(before)}px`;
  }
  const after = spacingPx(para.spcAft, largest);
  if (after) s.paddingBottom = `${r2(after)}px`;
  s.tabSize = `${r2(px(para.defTabSz ?? 914400))}px`;

  // Automatic numbers count paragraphs of the same level; deeper levels do not interrupt them.
  for (const key of [...counters.keys()]) if (key > level) counters.delete(key);
  const bullet = hasText ? para.bullet : undefined;
  if (bullet?.type === 'auto') {
    // A different numbering scheme starts over, like PowerPoint does.
    const prev = counters.get(level);
    counters.set(level, { scheme: bullet.scheme, n: prev?.scheme === bullet.scheme ? prev.n + 1 : bullet.start });
  } else counters.delete(level);

  let textStart = marL + indent;
  if (bullet && bullet.type !== 'none') {
    const firstSize = runSize(firstProps, scale);
    let char = bullet.type === 'char' ? bullet.char : bullet.type === 'auto' ? autoNumber(bullet.scheme, counters.get(level)?.n ?? 1) : '•';
    const buFont = para.buFont === null || para.buFont === undefined ? undefined : themeFont(para.buFont, env);
    const mapped = bullet.type === 'char' ? symbolChar(buFont, char) : null;
    if (mapped) char = mapped;
    const span = h('span', { class: 'pptx-bullet' }, char);
    const bs = span.style;
    const size = para.buSize?.pts !== undefined ? para.buSize.pts * PT * scale : firstSize * (para.buSize?.pct ?? 1);
    const family = buFont && !mapped ? fontFamily(buFont, themeFont(firstProps.latin, env)) : runFont(firstProps, env);
    bs.fontSize = `${r2(size)}px`;
    bs.fontFamily = family;
    if (bullet.type === 'auto') {
      if (firstProps.b) bs.fontWeight = '700';
      if (firstProps.i) bs.fontStyle = 'italic';
    }
    // Bullets follow the text color, including invisible (noFill) text.
    const color = para.buColor ?? firstProps.color;
    if (color === null) bs.color = 'transparent';
    else if (color) bs.color = css(color);
    const width = textWidth(char, `${r2(size)}px ${family}`);
    // A hanging indent reserves the space between the bullet and the text.
    if (indent < 0) {
      bs.minWidth = `${r2(-indent)}px`;
      textStart = marL + Math.max(0, width + indent);
    } else {
      bs.paddingRight = `${r2(size * 0.3)}px`;
      textStart += width + size * 0.3;
    }
    el.append(span);
  }

  if (!hasText) el.append('\u200b');
  const spacers = texts.some((t) => t.text.includes('\t')) ? tabSpacers(texts, para, textStart, env, scale) : null;
  let index = 0;
  for (const piece of pieces) {
    if (piece === 'br') {
      el.append(h('br'));
      continue;
    }
    const span = h('span');
    const tabs = spacers?.[index++];
    if (piece.math) {
      span.append(mathElement(piece.math, piece.math.localName === 'oMathPara'));
    } else if (tabs?.length) {
      piece.text.split('\t').forEach((part, i) => {
        if (part) span.append(part);
        if (i < tabs.length) span.append(h('span', { class: 'pptx-tab', style: `width:${r2(tabs[i])}px` }));
      });
    } else {
      span.textContent = piece.text;
    }
    styleRun(span, piece.props, env, scale);
    if (piece.href) {
      const hlink = env.theme.colors.get(env.colors.clrMap.get('hlink') ?? 'hlink');
      if (!piece.keepColor && hlink) span.style.color = css(hlink);
      span.style.textDecorationLine = 'underline';
      el.append(h('a', { class: 'pptx-link', href: piece.href, target: '_blank', rel: 'noopener noreferrer' }, span));
    } else {
      el.append(span);
    }
  }
  return { el, empty: !hasText, lines: lineModel(pitch, largest, runFont(largestProps, env)) };
}

/** Paragraphs of a text body as block elements. */
export function renderParagraphs(txBody: Element, layers: Layer[], opts: TextOptions): { el: HTMLElement; empty: boolean } {
  const all = [...layers, { list: child(txBody, 'lstStyle') }];
  const container = h('div', { class: 'pptx-paras' });
  const counters: Counters = new Map();
  const paragraphs: Paragraph[] = [];
  for (const p of elements(txBody)) {
    if (p.localName !== 'p') continue;
    if (opts.env.budget.text-- <= 0) break;
    const para = renderParagraph(p, all, opts, counters, paragraphs.length === 0);
    paragraphs.push(para);
    container.append(para.el);
  }
  // Moves the first baseline and the end of the text to where PowerPoint has them; the lines in between keep
  // their pitch. Margins rather than offsets, so anchoring at the middle or bottom sees the corrected height.
  const first = paragraphs[0]?.lines, last = paragraphs[paragraphs.length - 1]?.lines;
  // Space after the last paragraph does not count when PowerPoint anchors the text.
  if (paragraphs.length) paragraphs[paragraphs.length - 1].el.style.paddingBottom = '';
  if (first && last) {
    container.style.marginTop = `${r2(first.first - first.firstCss)}px`;
    container.style.marginBottom = `${r2(last.below - last.belowCss)}px`;
  }
  return { el: container, empty: paragraphs.every((p) => p.empty) };
}

const NOWRAP_SPACE = 4000;

export interface TextRect {
  x: number;
  y: number;
  w: number;
  h: number;
}

function arrangeParagraphs(paras: HTMLElement, body: BodyProps): void {
  if (body.wrap === 'none') {
    // Unwrapped lines may be wider than the box; they grow away from their alignment edge like in PowerPoint.
    for (const p of paras.children as HTMLCollectionOf<HTMLElement>) {
      const align = p.style.textAlign;
      if (align !== 'left' && align !== 'justify') p.style.marginLeft = `-${NOWRAP_SPACE}px`;
      if (align !== 'right') p.style.marginRight = `-${NOWRAP_SPACE}px`;
    }
  }
  if (body.numCol > 1) {
    paras.style.columnCount = String(body.numCol);
    // PowerPoint fills a column before starting the next instead of balancing them.
    paras.style.columnFill = 'auto';
    paras.style.columnGap = `${r2(px(body.spcCol))}px`;
    paras.style.height = '100%';
  }
}

/** A text body positioned inside its shape: insets, anchoring, wrapping, vertical text and columns. */
export function renderTextBox(txBody: Element, layers: Layer[], body: BodyProps, rect: TextRect, opts: TextOptions): HTMLElement | null {
  // PowerPoint draws "shrink text on overflow" with the scale stored in the file; it only recomputes it when
  // the text is edited.
  const { el: paras, empty } = renderParagraphs(txBody, layers, { ...opts, scale: body.fontScale, lnSpcReduction: body.lnSpcReduction });
  if (empty) return null;
  const vertical = body.vert === 'vert' || body.vert === 'vert270' || body.vert === 'eaVert' || body.vert === 'wordArtVert' || body.vert === 'mongolianVert';
  // Vertical text lays out in a box with swapped sides, which is then turned into place.
  const outerW = rect.w, outerH = rect.h;
  const boxW = vertical ? outerH : outerW, boxH = vertical ? outerW : outerH;
  const box = h('div', { class: 'pptx-text' });
  const s = box.style;
  s.left = `${r2(rect.x + (outerW - boxW) / 2)}px`;
  s.top = `${r2(rect.y + (outerH - boxH) / 2)}px`;
  s.width = `${r2(Math.max(0, boxW))}px`;
  s.height = `${r2(Math.max(0, boxH))}px`;
  s.padding = `${r2(px(body.tIns))}px ${r2(px(body.rIns))}px ${r2(px(body.bIns))}px ${r2(px(body.lIns))}px`;
  s.justifyContent = body.anchor === 'ctr' ? 'center' : body.anchor === 'b' ? 'flex-end' : 'flex-start';
  let angle = body.rot;
  if (body.vert === 'vert' || body.vert === 'eaVert') angle += 90;
  else if (body.vert === 'vert270') angle -= 90;
  if (angle) s.transform = `rotate(${r2(angle)}deg)`;
  if (body.wrap === 'none') box.classList.add('pptx-nowrap');
  arrangeParagraphs(paras, body);
  box.append(paras);
  return box;
}
