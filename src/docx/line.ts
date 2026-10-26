// SPDX-License-Identifier: MIT
// Line layout of one paragraph, done by us rather than the browser so lines, tabs and wrapping around floating
// objects come out where Word puts them. A paragraph becomes a list of atoms (words, spaces, tabs, objects)
// with break opportunities; lines are filled greedily into the free intervals of the column at their height.
import type { Theme } from '../pptx/color.js';
import { type FontInfo, SYMBOL_FAMILY, boxWidth, fontInstalled, measure, runFont, scriptOf } from './measure.js';
import type { DrawingObj, FieldInline, Inline, Link, NoteInline, Paragraph } from './model.js';
import { mathSize } from './math.js';
import { formatNumber } from './numbering.js';
import type { ParaProps, RunProps, TabStop } from './props.js';
import { BOX_GLYPHS, isSymbolFont, symbolText } from './symbols.js';
import { halfPt, twip } from './xml.js';

export interface LineEnv {
  theme: Theme;
  /** Default tab stop distance in px. */
  defaultTab: number;
  /** Text of dynamic fields on the page being laid out. */
  field(name: 'page' | 'numpages' | 'sectionpages' | 'pageref', format: string | null, arg?: string, cached?: string): string;
  /** Lines that end with a manual line break are justified too (Word 2013 and later). */
  expandShiftReturn: boolean;
  /** Word 2013+ ignores custom tab stops beyond the right indent less strictly; older modes clamp. */
  compat15: boolean;
  /** Character width for indents given in characters (leftChars …), px. */
  charUnit?: number;
  /** Decimal separator for decimal tab stops (settings w:decimalSymbol). */
  decimal?: string;
}

export interface TextStyle {
  run: RunProps;
  /** Font the text is drawn with (super/subscript already reduced). */
  font: FontInfo;
  /** Baseline shift in px, up positive. */
  shift: number;
  /** Extra space after each character, px. */
  spacing: number;
  /** Horizontal scale (w:w). */
  scale: number;
  link?: Link;
  smallCaps: boolean;
  note?: NoteInline;
  /** Symbol font characters drawn with a Unicode font: the line height follows the original font. */
  symbol?: boolean;
}

export type AtomKind = 'text' | 'space' | 'tab' | 'obj' | 'br' | 'float' | 'mark' | 'box' | 'sep';

export interface Atom {
  k: AtomKind;
  text: string;
  w: number;
  style: TextStyle;
  asc: number;
  desc: number;
  /** A line may break after this atom. */
  breakAfter: boolean;
  inline?: Inline;
  obj?: DrawingObj;
  /** Objects: ascent of the run's font, which the line spacing applies to (the object itself is added on top). */
  textAsc?: number;
}

export interface Frag {
  k: 'text' | 'tab' | 'obj' | 'box' | 'sep';
  x: number;
  w: number;
  text: string;
  style: TextStyle;
  /** Extra width per space for justified text. */
  wordSpacing?: number;
  leader?: string;
  obj?: DrawingObj;
  /** Underlined/struck tab: the decoration spans the tab. */
  decorate?: boolean;
  /** Page number fields are filled in when the page is drawn. */
  field?: FieldInline;
  /** Extra space after each character (distributed alignment). */
  charSpacing?: number;
}

export interface FloatRef {
  obj: DrawingObj;
  /** x of the anchor character in column coordinates. */
  x: number;
}

export interface Line {
  /** Line box height and baseline from its top, px. */
  h: number;
  baseline: number;
  frags: Frag[];
  /** Index of the first atom of the next line. */
  end: number;
  /** The line ends with a page or column break. */
  breakType?: 'page' | 'column';
  clear?: string;
  floats: FloatRef[];
  notes: NoteInline[];
  bookmarks: string[];
  /** Last line of the paragraph. */
  last: boolean;
  /** Text extent (left, right) in column coordinates, for borders and shading. */
  left: number;
  right: number;
  /** Bar tab positions to draw through the line. */
  bars?: number[];
  /** Extent of the text itself from the line top (without line spacing): what text wrapping tests. */
  textTop: number;
  textBottom: number;
}

/** Free horizontal intervals of the column for a band of lines, in column coordinates. */
export type Segments = [number, number][];

const EPS = 0.01;
const CJK = /[⺀-鿿가-힯豈-﫿＀-￯　-〿]/;
/** Characters a line may break after (besides spaces). */
const BREAK_AFTER = new Set(['-', '‐', '–', '—', '­', '​', '‒']);
/** Characters that must not start a line (closing punctuation, small kana). */
const NO_START = /[)\]}»"'.,;:!?、。，．：；？！）〕］｝〉》」』】ーぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ]/;

// ---- Atoms -------------------------------------------------------------------------------------------------

/** Superscript and subscript: two thirds of the size (whole half points), raised 0.35 em or lowered 0.09 em. */
const vertSize = (halfPoints: number) => Math.max(1, Math.floor((halfPoints * 2) / 3));

export function textStyle(run: RunProps, env: LineEnv, script: 'ascii' | 'hAnsi' | 'eastAsia' | 'cs', family?: string, link?: Link): TextStyle {
  // A font name that replaces the run's fonts (symbol fonts, check box states).
  if (family && family !== 'symbol') run = { ...run, fonts: { ascii: family, hAnsi: family, eastAsia: family, cs: family } };
  const full = runFont(run, script, env.theme, family === 'symbol' ? SYMBOL_FAMILY : undefined);
  let font = full;
  let shift = run.position ? halfPt(run.position) : 0;
  if (run.vertAlign === 'superscript' || run.vertAlign === 'subscript') {
    const small = runFont({ ...run, sz: vertSize(run.sz ?? 20), szCs: vertSize(run.szCs ?? run.sz ?? 20) }, script, env.theme, family === 'symbol' ? SYMBOL_FAMILY : undefined);
    font = small;
    shift += run.vertAlign === 'superscript' ? full.size * 0.35 : -full.size * 0.09;
  }
  return {
    run,
    font,
    shift,
    spacing: run.spacing ? twip(run.spacing) : 0,
    scale: run.w ? run.w / 100 : 1,
    link,
    smallCaps: !!run.smallCaps && !run.caps,
    symbol: family === 'symbol',
  };
}

/** Fonts with real ballot box glyphs; elsewhere (or when the font is missing) the renderer draws the boxes. */
const BOX_FONTS = /^(ms gothic|ms mincho|ms pgothic|ms ui gothic|segoe ui symbol|segoe ui emoji|yu gothic|meiryo|simsun|nsimsun|microsoft yahei|microsoft jhenghei|arial unicode ms|malgun gothic|mingliu|pmingliu)$/i;

function boxFont(name: string): boolean {
  return BOX_FONTS.test(name.trim()) && fontInstalled(name);
}

function textWidth(text: string, style: TextStyle): number {
  let w = measure(text, style.font);
  if (style.smallCaps) w = smallCapsWidth(text, style.font);
  if (style.spacing) w += style.spacing * [...text].length;
  return w * style.scale;
}

/** Small capitals: lowercase letters as capitals at 70 % (what the browser synthesizes too). */
function smallCapsWidth(text: string, font: FontInfo): number {
  let w = 0;
  for (const part of text.match(/\p{Ll}+|[^\p{Ll}]+/gu) ?? []) {
    w += /\p{Ll}/u.test(part) ? measure(part.toUpperCase(), font, font.size * 0.7) : measure(part, font);
  }
  return w;
}

/**
 * Height a run takes above and below the baseline. Word puts the font's line gap above the ascent (measured
 * against Word's own PDF output). `full` is the run's font at its own size; raised, lowered and super/subscript
 * text can reach beyond it.
 */
function vertical(style: TextStyle, full: FontInfo): { asc: number; desc: number } {
  const half = (f: FontInfo) => ({ asc: (f.ascent + f.gap) * f.size, desc: f.descent * f.size });
  const big = half(full), small = style.symbol ? big : half(style.font);
  return { asc: Math.max(big.asc, small.asc + style.shift), desc: Math.max(big.desc, small.desc - style.shift) };
}

interface Builder {
  atoms: Atom[];
  env: LineEnv;
}

function pushText(b: Builder, text: string, style: TextStyle, full: FontInfo, inline?: Inline): void {
  const { asc, desc } = vertical(style, full);
  // Split into words, spaces and characters that allow a break.
  let word = '';
  const flushWord = (breakAfter: boolean) => {
    if (!word) return;
    b.atoms.push({ k: 'text', text: word, w: textWidth(word, style), style, asc, desc, breakAfter, inline });
    word = '';
  };
  for (const ch of text) {
    if (ch === ' ' || ch === '　') {
      flushWord(false);
      const last = b.atoms[b.atoms.length - 1];
      if (last && last.k === 'space' && last.style === style) {
        last.text += ch;
        last.w += textWidth(ch, style);
      } else {
        b.atoms.push({ k: 'space', text: ch, w: textWidth(ch, style), style, asc, desc, breakAfter: true, inline });
      }
      continue;
    }
    if (BOX_GLYPHS.has(ch) && !boxFont(style.font.name)) {
      flushWord(false);
      b.atoms.push({ k: 'box', text: ch, w: boxWidth(style.font) * style.scale + style.spacing, style, asc, desc, breakAfter: false, inline });
      continue;
    }
    if (CJK.test(ch)) {
      flushWord(true);
      word = ch;
      flushWord(true);
      continue;
    }
    if (NO_START.test(ch) && !word) {
      // Closing punctuation sticks to what comes before it.
      const prev = b.atoms[b.atoms.length - 1];
      if (prev && prev.k === 'text' && prev.style === style) {
        prev.text += ch;
        prev.w = textWidth(prev.text, style);
        prev.breakAfter = false;
        continue;
      }
    }
    word += ch;
    if (BREAK_AFTER.has(ch)) flushWord(true);
  }
  flushWord(false);
}

/** Splits text into pieces of one script, since each script can have its own font. */
function scriptRuns(text: string, run: RunProps): { text: string; script: 'ascii' | 'hAnsi' | 'eastAsia' | 'cs' }[] {
  const f = run.fonts;
  const same = (f.ascii ?? f.asciiTheme) === (f.hAnsi ?? f.hAnsiTheme) && !run.cs && !run.rtl;
  const out: { text: string; script: 'ascii' | 'hAnsi' | 'eastAsia' | 'cs' }[] = [];
  for (const ch of text) {
    let script = scriptOf(ch.codePointAt(0) ?? 0);
    if (same && (script === 'ascii' || script === 'hAnsi')) script = 'ascii';
    // Spaces and punctuation follow the text around them.
    if (/[\s\p{P}]/u.test(ch) && out.length && script !== 'eastAsia') script = out[out.length - 1].script;
    const last = out[out.length - 1];
    if (last && last.script === script) last.text += ch;
    else out.push({ text: ch, script });
  }
  return out;
}

function applyCaps(text: string, run: RunProps): string {
  return run.caps ? text.toUpperCase() : text;
}

/** The atoms of a paragraph, its number first. */
export function buildAtoms(p: Paragraph, env: LineEnv): Atom[] {
  const b: Builder = { atoms: [], env };
  if (p.num && p.num.text !== undefined) {
    const run = p.num.run;
    const isBullet = p.num.level.numFmt === 'bullet';
    let text = p.num.text;
    let family: string | undefined;
    const font = run.fonts.ascii ?? run.fonts.hAnsi;
    if (isBullet && font && isSymbolFont(font) && fontInstalled(font)) {
      // The bullet's own symbol font draws it exactly like Word.
      text = [...text].map((ch) => (ch.charCodeAt(0) >= 0x20 && ch.charCodeAt(0) < 0x100 ? String.fromCharCode(0xf000 + ch.charCodeAt(0)) : ch)).join('');
      family = font;
    } else if (isBullet && font) {
      const mapped = mapSymbolText(text, font);
      if (mapped !== null) {
        text = mapped;
        family = 'symbol';
      }
    }
    const style = textStyle(run, env, 'ascii', family);
    const full = family === 'symbol' ? runFont(run, 'ascii', env.theme) : style.font;
    if (text) pushText(b, applyCaps(text, run), style, full);
    for (const a of b.atoms) a.breakAfter = false;
    const suff = p.num.level.suff;
    if (suff === 'tab') b.atoms.push({ k: 'tab', text: '', w: 0, style, asc: 0, desc: 0, breakAfter: false, inline: { t: 'tab', run } });
    else if (suff === 'space') pushText(b, ' ', style, full);
    if (b.atoms.length) b.atoms[b.atoms.length - 1].breakAfter = false;
  }
  for (const inline of p.inlines) {
    switch (inline.t) {
      case 'text': {
        // Symbol fonts that are installed draw their own glyphs, exactly like Word.
        const own = inline.original && fontInstalled(inline.original.font);
        const text = own ? inline.original!.text : inline.text;
        const family = own ? inline.original!.font : inline.family;
        for (const piece of scriptRuns(applyCaps(text, inline.run), inline.run)) {
          const style = textStyle(inline.run, env, piece.script, family, inline.link);
          const full = (inline.run.vertAlign && inline.run.vertAlign !== 'baseline') || family === 'symbol' ? runFont(inline.run, piece.script, env.theme) : style.font;
          pushText(b, piece.text, style, full, inline);
        }
        break;
      }
      case 'field': {
        const text = env.field(inline.field, inline.format, inline.arg, inline.cached);
        const style = textStyle(inline.run, env, 'ascii', undefined, inline.link);
        pushText(b, text, style, style.font, inline);
        break;
      }
      case 'note': {
        const style = { ...textStyle(inline.run, env, 'ascii', undefined, inline.link), note: inline.own ? undefined : inline };
        const full = runFont(inline.run, 'ascii', env.theme);
        if (inline.mark) pushText(b, inline.mark, style, full, inline);
        else b.atoms.push({ k: 'mark', text: '', w: 0, style, asc: 0, desc: 0, breakAfter: false, inline });
        break;
      }
      case 'tab': {
        const style = textStyle(inline.run, env, 'ascii', undefined, inline.link);
        const { asc, desc } = vertical(style, style.font);
        b.atoms.push({ k: 'tab', text: '', w: 0, style, asc, desc, breakAfter: true, inline });
        break;
      }
      case 'br': {
        const style = textStyle(inline.run, env, 'ascii');
        const { asc, desc } = vertical(style, style.font);
        b.atoms.push({ k: 'br', text: '', w: 0, style, asc: inline.type === 'line' ? asc : 0, desc: inline.type === 'line' ? desc : 0, breakAfter: true, inline });
        break;
      }
      case 'obj': {
        const style = textStyle(inline.run, env, 'ascii', undefined, inline.link);
        const { obj } = inline;
        const prev = b.atoms[b.atoms.length - 1];
        if (prev) prev.breakAfter = true;
        if (obj.graphic.kind === 'math') {
          // Equations are laid out by the browser; their size is measured once.
          const size = mathSize(obj.graphic.el, obj.graphic.display, style.font.size);
          obj.w = size.w;
          obj.h = size.h;
          obj.descent = size.descent;
          obj.mathSize = style.font.size;
          const { asc: textAsc, desc } = vertical(style, style.font);
          b.atoms.push({ k: 'obj', text: '', w: size.w, style, asc: size.h - size.descent, desc: Math.max(desc, size.descent), breakAfter: true, inline, obj, textAsc });
        } else if (obj.graphic.kind === 'checkbox') {
          // Legacy check boxes sit on the baseline like a character of their size.
          const { asc: textAsc, desc } = vertical(style, style.font);
          b.atoms.push({ k: 'obj', text: '', w: obj.w + 2, style, asc: obj.h, desc, breakAfter: true, inline, obj, textAsc });
        } else {
          const w = obj.w + obj.effect.l + obj.effect.r;
          const h = obj.h + obj.effect.t + obj.effect.b;
          const { asc: textAsc, desc } = vertical(style, style.font);
          b.atoms.push({ k: 'obj', text: '', w, style, asc: h, desc, breakAfter: true, inline, obj, textAsc });
        }
        break;
      }
      case 'float': {
        const style = textStyle(inline.run, env, 'ascii');
        b.atoms.push({ k: 'float', text: '', w: 0, style, asc: 0, desc: 0, breakAfter: false, inline, obj: inline.obj });
        break;
      }
      case 'bookmark': {
        const style = textStyle(inline.run, env, 'ascii');
        b.atoms.push({ k: 'mark', text: '', w: 0, style, asc: 0, desc: 0, breakAfter: false, inline });
        break;
      }
      case 'sep': {
        const style = textStyle(inline.run, env, 'ascii');
        const { asc, desc } = vertical(style, style.font);
        b.atoms.push({ k: 'sep', text: '', w: inline.continuation ? -1 : 192, style, asc, desc, breakAfter: true, inline });
        break;
      }
    }
  }
  return b.atoms;
}

function mapSymbolText(text: string, font: string): string | null {
  if (!/^(wingdings|webdings|symbol)/i.test(font) && !/[-]/.test(text)) return null;
  const r = symbolText(font, text);
  return r.mapped ? r.text : null;
}

// ---- Tabs --------------------------------------------------------------------------------------------------

interface TabInfo {
  pos: number;
  val: string;
  leader: string;
}

function nextTab(x: number, para: ParaProps, env: LineEnv, indLeft: number, lineRight: number, ptab?: { alignment: string; relativeTo: string; leader: string }, colWidth?: number, indRight?: number): TabInfo {
  if (ptab) {
    const left = ptab.relativeTo === 'indent' ? indLeft : 0;
    const right = ptab.relativeTo === 'indent' ? (colWidth ?? lineRight) - (indRight ?? 0) : (colWidth ?? lineRight);
    const pos = ptab.alignment === 'center' ? (left + right) / 2 : ptab.alignment === 'right' ? right : left;
    return { pos, val: ptab.alignment, leader: ptab.leader };
  }
  const stops: TabStop[] = para.tabs ?? [];
  let found: TabInfo | null = null;
  for (const t of stops) {
    if (t.val === 'clear' || t.val === 'bar') continue;
    const pos = twip(t.pos);
    if (pos > x + EPS) {
      found = { pos, val: t.val === 'num' ? 'left' : t.val, leader: t.leader };
      break;
    }
  }
  // A hanging indent acts as a tab stop for the first line.
  if (indLeft > x + EPS && (!found || indLeft < found.pos - EPS)) return { pos: indLeft, val: 'left', leader: 'none' };
  if (found) return found;
  const step = env.defaultTab > 0 ? env.defaultTab : 48;
  const pos = (Math.floor((x + EPS) / step) + 1) * step;
  return { pos, val: 'left', leader: 'none' };
}

const LEADERS: Record<string, string> = { dot: '.', hyphen: '-', underscore: '_', heavy: '_', middleDot: '·' };

// ---- Lines -------------------------------------------------------------------------------------------------

export interface ParaGeometry {
  /** Column width and indents in px. */
  colWidth: number;
  indLeft: number;
  indRight: number;
  indFirst: number;
}

interface Placed {
  atom: Atom;
  x: number;
  w: number;
  seg: number;
  leader?: string;
  /** Extra space per character (distributed alignment). */
  spread?: number;
}

/**
 * Line spacing applied to the natural height: returns height and baseline. With multiple spacing Word keeps the
 * text at the top of the line and adds the extra below it; exact and at-least heights put the text at the bottom.
 */
export function spaceLine(asc: number, desc: number, para: ParaProps, grid?: number): { h: number; baseline: number } {
  const natural = asc + desc;
  let h = natural;
  const rule = para.lineRule ?? 'auto';
  const line = para.line ?? 240;
  let baseline = asc;
  if (rule === 'exact') {
    h = Math.max(1, twip(Math.abs(line)));
    // Extra height goes above the text; missing height is taken from top and bottom alike.
    const excess = h - natural;
    baseline = asc + (excess > 0 ? excess : excess / 2);
  } else if (rule === 'atLeast') {
    h = Math.max(natural, twip(line));
    baseline = h - desc;
  } else {
    h = natural * Math.max(0.06, line / 240);
  }
  if (grid && grid > 0 && rule !== 'exact') {
    const snapped = Math.ceil(h / grid - 0.01) * grid;
    baseline += (snapped - h) / 2;
    h = snapped;
  }
  return { h, baseline };
}

export class ParagraphLines {
  readonly atoms: Atom[];
  /** Font of the paragraph mark: height of empty lines and of the last line. */
  readonly markAsc: number;
  readonly markDesc: number;

  constructor(
    readonly p: Paragraph,
    readonly env: LineEnv,
    readonly geo: ParaGeometry,
    readonly grid?: number,
  ) {
    this.atoms = buildAtoms(p, env);
    // Horizontal rules take their width from the text column they are laid out in.
    for (const a of this.atoms) {
      if (a.k !== 'obj' || a.obj?.hr === undefined) continue;
      a.obj.w = Math.max(1, (geo.colWidth - geo.indLeft - geo.indRight) * a.obj.hr - 1);
      a.w = a.obj.w;
    }
    const markStyle = textStyle(p.mark, env, 'ascii');
    const v = vertical(markStyle, markStyle.font);
    this.markAsc = v.asc;
    this.markDesc = v.desc;
  }

  get empty(): boolean {
    return this.atoms.length === 0;
  }

  /** Left and right edge of a line's text area in column coordinates. */
  bounds(first: boolean): [number, number] {
    const { indLeft, indRight, indFirst, colWidth } = this.geo;
    return [indLeft + (first ? indFirst : 0), colWidth - indRight];
  }

  /** Height of the first line if laid out at full width, for estimating before the real layout. */
  estimateHeight(): number {
    const a = this.atoms[0];
    const asc = Math.max(this.markAsc, a ? a.asc : 0), desc = Math.max(this.markDesc, a ? a.desc : 0);
    return spaceLine(asc, desc, this.p.props, this.grid).h;
  }

  /** Where the text of a line starts its first line, before laying it out: [top, bottom] from the line top. */
  estimateText(): [number, number] {
    const a = this.atoms[0];
    const asc = a && a.k !== 'float' && a.k !== 'mark' ? a.asc : this.markAsc;
    const desc = a && a.k !== 'float' && a.k !== 'mark' ? a.desc : this.markDesc;
    const { baseline } = spaceLine(asc, desc, this.p.props, this.grid);
    return [baseline - asc, baseline + desc];
  }

  /**
   * Lays out the line that starts at atom `start` into the free segments. Returns null when nothing at all fits
   * (the caller moves down past the obstacle).
   */
  line(start: number, segments: Segments, first: boolean, constrained = false): Line | null {
    const { atoms } = this;
    const para = this.p.props;
    const [lineLeft, lineRight] = this.bounds(first);
    const segs: Segments = [];
    for (const [a, b] of segments) {
      const s = Math.max(a, Math.min(lineLeft, this.geo.indLeft)), e = Math.min(b, lineRight);
      if (e - s > EPS) segs.push([s, e]);
    }
    if (!segs.length) return null;
    // The first line starts at the first-line indent when its segment contains it.
    const startX = (seg: number) => (seg === 0 && segs[0][0] <= lineLeft + EPS ? Math.max(segs[0][0], lineLeft) : segs[seg][0]);

    const placed: Placed[] = [];
    let seg = 0;
    let x = startX(0);
    let i = start;
    let lastBreak = -1; // index into placed after which a break is allowed
    let breakType: Line['breakType'];
    let clear: string | undefined;
    let pending: { index: number; tab: TabInfo; x0: number; width: number; decimal: number | null } | null = null;

    /** Width of the text after an alignment tab that goes before the stop. */
    const lead = (p: NonNullable<typeof pending>, width: number) => {
      const t = p.tab;
      return t.val === 'right' || t.val === 'end' ? width : t.val === 'center' ? width / 2 : (p.decimal ?? width);
    };
    const resolvePending = () => {
      if (!pending) return;
      const tabW = Math.max(0, pending.tab.pos - lead(pending, pending.width) - pending.x0);
      const tabAtom = placed[pending.index];
      const delta = tabW - tabAtom.w;
      tabAtom.w = tabW;
      for (let k = pending.index + 1; k < placed.length; k++) placed[k].x += delta;
      x = pending.x0 + tabW + pending.width;
      pending = null;
    };
    /** Right end of the text so far, if `extra` more width followed (alignment tabs grow to the left). */
    const rightWith = (extra: number) => {
      if (!pending) return x + extra;
      const width = pending.width + extra;
      return Math.max(pending.tab.pos - lead(pending, width), pending.x0) + width;
    };
    const currentRight = () => rightWith(0);

    while (i < atoms.length) {
      const a = atoms[i];
      if (a.k === 'br') {
        placed.push({ atom: a, x, w: 0, seg });
        i++;
        const br = a.inline as Extract<Inline, { t: 'br' }>;
        if (br.type !== 'line') breakType = br.type;
        if (br.clear) clear = br.clear;
        break;
      }
      if (a.k === 'float' || a.k === 'mark') {
        placed.push({ atom: a, x: currentRight(), w: 0, seg });
        i++;
        continue;
      }
      if (a.k === 'tab') {
        resolvePending();
        const tabInline = a.inline as Extract<Inline, { t: 'tab' }> | undefined;
        const tab = nextTab(x, para, this.env, this.geo.indLeft, segs[seg][1], tabInline?.ptab, this.geo.colWidth, this.geo.indRight);
        const leader = LEADERS[tab.leader];
        if (tab.val === 'left' || tab.val === 'start' || tab.val === 'num') {
          // A stop beyond the end of the line takes the tab only to the line end; what follows wraps.
          const w = Math.max(0, Math.min(tab.pos, Math.max(segs[seg][1], x)) - x);
          placed.push({ atom: a, x, w, seg, leader });
          x += w;
        } else {
          placed.push({ atom: a, x, w: 0, seg, leader });
          pending = { index: placed.length - 1, tab, x0: x, width: 0, decimal: null };
        }
        lastBreak = placed.length;
        i++;
        continue;
      }
      const w = a.k === 'sep' ? (a.w < 0 ? segs[seg][1] - x : a.w) : a.w;
      const right = rightWith(w);
      const segEnd = segs[seg][1];
      if (right <= segEnd + EPS || a.k === 'space') {
        placed.push({ atom: a, x: pending ? pending.x0 + pending.width : x, w, seg });
        if (pending) {
          if (pending.tab.val === 'decimal' && pending.decimal === null && a.k === 'text') {
            const dot = a.text.indexOf(this.env.decimal ?? '.');
            if (dot >= 0) pending.decimal = pending.width + textWidth(a.text.slice(0, dot), a.style);
          }
          pending.width += w;
        }
        x += w;
        if (a.breakAfter) lastBreak = placed.length;
        i++;
        continue;
      }
      // The atom does not fit.
      if (lastBreak > 0) {
        // Each placed entry stands for one atom, so the break position maps straight back to an atom index.
        const resume = start + lastBreak;
        // Continue after the break in the next free segment of the same line, if there is one.
        if (seg + 1 < segs.length) {
          placed.length = lastBreak;
          resolvePending();
          seg++;
          x = startX(seg);
          i = resume;
          lastBreak = placed.length;
          continue;
        }
        placed.length = lastBreak;
        i = resume;
        break;
      }
      const usedInSeg = placed.some((p) => p.seg === seg && p.atom.k !== 'mark' && p.atom.k !== 'float');
      if (!usedInSeg && seg + 1 < segs.length) {
        // Too narrow for this word: try the next segment.
        seg++;
        x = startX(seg);
        continue;
      }
      const visible = placed.some((p) => p.atom.k !== 'mark' && p.atom.k !== 'float');
      // Next to floating objects nothing fits: the caller moves the line down past them.
      if (!visible && constrained) return null;
      if (!visible) {
        // A word longer than the line: break it between characters.
        if (a.k === 'text' && a.text.length > 1) {
          const avail = segEnd - x;
          const chars = [...a.text];
          let n = 1;
          while (n < chars.length && textWidth(chars.slice(0, n + 1).join(''), a.style) <= avail + EPS) n++;
          const head = chars.slice(0, n).join(''), tail = chars.slice(n).join('');
          const headAtom: Atom = { ...a, text: head, w: textWidth(head, a.style), breakAfter: true };
          const tailAtom: Atom = { ...a, text: tail, w: textWidth(tail, a.style) };
          atoms.splice(i, 1, headAtom, tailAtom);
          placed.push({ atom: headAtom, x, w: headAtom.w, seg });
          x += headAtom.w;
          i++;
          break;
        }
        // An object or single character wider than the line still goes on it.
        placed.push({ atom: a, x, w, seg });
        x += w;
        i++;
        break;
      }
      // Glued atoms (a word across runs) longer than the rest of the line: break before this one.
      break;
    }
    resolvePending();
    const last = i >= atoms.length;
    return this.finish(placed, segs, i, { last, breakType, clear, first });
  }

  private finish(placed: Placed[], segs: Segments, end: number, info: { last: boolean; breakType?: Line['breakType']; clear?: string; first: boolean }): Line {
    const para = this.p.props;
    let asc = 0, desc = 0, objAsc = 0;
    let any = false;
    for (const pl of placed) {
      const a = pl.atom;
      if (a.k === 'float' || a.k === 'mark' || a.k === 'br') continue;
      any = true;
      if (a.k === 'obj') {
        // Objects count with the text font of their run; their height above it comes on top (see below).
        objAsc = Math.max(objAsc, a.asc);
        asc = Math.max(asc, a.textAsc ?? 0);
      } else asc = Math.max(asc, a.asc);
      desc = Math.max(desc, a.desc);
    }
    // A line with nothing visible gets the paragraph mark's height; the last line also when the mark has its own
    // formatting (Word ignores a mark that just follows the paragraph style).
    if (!any || (info.last && this.p.markExplicit)) {
      asc = Math.max(asc, this.markAsc);
      desc = Math.max(desc, this.markDesc);
    }
    for (const pl of placed) {
      if (pl.atom.k === 'br' && pl.atom.inline && (pl.atom.inline as Extract<Inline, { t: 'br' }>).type === 'line') {
        asc = Math.max(asc, pl.atom.asc);
        desc = Math.max(desc, pl.atom.desc);
      }
    }
    // Word applies multiple line spacing to the text only; objects reaching above the text add their height
    // unscaled. With at-least spacing they simply count as text; exact lines clip them.
    const rule = para.lineRule ?? 'auto';
    let { h, baseline } = spaceLine(rule === 'atLeast' ? Math.max(asc, objAsc) : asc, desc, para, this.grid);
    if (rule === 'auto' && objAsc > asc) {
      h += objAsc - asc;
      baseline += objAsc - asc;
    }
    const textAscent = rule === 'atLeast' ? Math.max(asc, objAsc) : asc;

    // Alignment per segment; trailing spaces hang past the end. Display equations are centered.
    let jc = para.jc ?? 'left';
    if (placed.some((p) => p.atom.obj?.graphic.kind === 'math' && p.atom.obj.graphic.display)) jc = 'center';
    if (para.bidi) jc = jc === 'left' || jc === 'start' ? 'right' : jc === 'right' || jc === 'end' ? 'left' : jc;
    const justify = jc === 'both' || jc === 'distribute' || jc === 'thaiDistribute' || jc === 'lowKashida' || jc === 'mediumKashida' || jc === 'highKashida';
    const endsWithBreak = placed.some((p) => p.atom.k === 'br' && (p.atom.inline as Extract<Inline, { t: 'br' }>)?.type === 'line');
    const stretch = justify && !info.last && !info.breakType && (!endsWithBreak || this.env.expandShiftReturn);
    const distributeChars = (jc === 'distribute' || jc === 'thaiDistribute') && info.last;
    let left = Infinity, right = -Infinity;
    for (let s = 0; s < segs.length; s++) {
      const items = placed.filter((p) => p.seg === s);
      if (!items.length) continue;
      let contentEnd = items[0].x;
      let lastTab = -1;
      items.forEach((p, k) => {
        if (p.atom.k === 'tab') lastTab = k;
        if (p.atom.k !== 'space' && p.atom.k !== 'mark' && p.atom.k !== 'float' && p.atom.k !== 'br') contentEnd = Math.max(contentEnd, p.x + p.w);
      });
      const segStart = items[0].x;
      const free = segs[s][1] - contentEnd;
      let offset = 0;
      if (lastTab < 0 || jc === 'center' || jc === 'right' || jc === 'end') {
        if (jc === 'center') offset = free / 2;
        else if (jc === 'right' || jc === 'end') offset = free;
      }
      if (lastTab >= 0 && (jc === 'center' || jc === 'right' || jc === 'end')) offset = Math.max(0, offset);
      if (stretch && free > 0) {
        // Only spaces after the last tab stretch, and not the trailing ones.
        let lastContent = -1;
        items.forEach((p, k) => {
          if (p.atom.k !== 'space' && p.atom.k !== 'mark' && p.atom.k !== 'float' && p.atom.k !== 'br') lastContent = k;
        });
        let spaces = 0;
        for (let k = lastTab + 1; k < lastContent; k++) if (items[k].atom.k === 'space') spaces += [...items[k].atom.text].length;
        if (spaces > 0) {
          const per = free / spaces;
          let shift = 0;
          for (let k = lastTab + 1; k < items.length; k++) {
            items[k].x += shift;
            if (items[k].atom.k === 'space' && k < lastContent) {
              const n = [...items[k].atom.text].length;
              items[k].w += per * n;
              shift += per * n;
            }
          }
        }
      }
      // Distributed alignment spreads the last line over the whole width, character by character.
      if (distributeChars && free > 0) {
        const content = items.filter((p) => p.atom.k !== 'mark' && p.atom.k !== 'float' && p.atom.k !== 'br');
        let lastContent = content.length - 1;
        while (lastContent >= 0 && content[lastContent].atom.k === 'space') lastContent--;
        const chars = content.slice(0, lastContent + 1).reduce((n, p) => n + (p.atom.k === 'text' || p.atom.k === 'space' ? [...p.atom.text].length : 1), 0);
        if (chars > 1) {
          const per = free / (chars - 1);
          let shift = 0;
          for (const p of items) {
            p.x += shift;
            if (p.atom.k === 'text' || p.atom.k === 'space') {
              const n = [...p.atom.text].length;
              p.w += per * n;
              p.spread = per;
              shift += per * n;
            } else if (p.atom.k !== 'mark' && p.atom.k !== 'float' && p.atom.k !== 'br') {
              p.w += per;
              shift += per;
            }
          }
        }
      }
      if (offset) for (const p of items) p.x += offset;
      left = Math.min(left, segStart + offset);
      right = Math.max(right, contentEnd + offset);
    }
    if (!Number.isFinite(left)) {
      const [l, r] = this.bounds(info.first);
      left = l;
      right = r;
    }

    // Fragments: neighbouring text of one style merges into one piece.
    const frags: Frag[] = [];
    const floats: FloatRef[] = [];
    const notes: NoteInline[] = [];
    const bookmarks: string[] = [];
    const bars: number[] = [];
    for (const p of placed) {
      const a = p.atom;
      switch (a.k) {
        case 'float':
          floats.push({ obj: a.obj!, x: p.x });
          break;
        case 'mark':
          if (a.inline?.t === 'bookmark') bookmarks.push(a.inline.name);
          break;
        case 'br':
          break;
        case 'tab': {
          const deco = !!a.style.run.u && a.style.run.u !== 'none';
          frags.push({ k: 'tab', x: p.x, w: p.w, text: '', style: a.style, leader: p.leader, decorate: deco || !!a.style.run.strike });
          break;
        }
        case 'obj':
          frags.push({ k: 'obj', x: p.x, w: p.w, text: '', style: a.style, obj: a.obj });
          break;
        case 'box':
          frags.push({ k: 'box', x: p.x, w: p.w, text: a.text, style: a.style });
          break;
        case 'sep':
          frags.push({ k: 'sep', x: p.x, w: p.w, text: '', style: a.style });
          break;
        default: {
          if (a.style.note && !notes.includes(a.style.note)) notes.push(a.style.note);
          const prev = frags[frags.length - 1];
          const extra = a.k === 'space' ? p.w - a.w - (p.spread ?? 0) * [...a.text].length : 0;
          if (prev && prev.k === 'text' && prev.style === a.style && Math.abs(prev.x + prev.w - p.x) < 0.05 && prev.charSpacing === p.spread) {
            prev.text += a.text;
            prev.w += p.w;
            if (extra) prev.wordSpacing = extra / [...a.text].length;
          } else {
            frags.push({ k: 'text', x: p.x, w: p.w, text: a.text, style: a.style, wordSpacing: extra ? extra / [...a.text].length : undefined, field: a.inline?.t === 'field' ? a.inline : undefined, charSpacing: p.spread });
          }
        }
      }
    }
    // Soft hyphen at the end of a broken line becomes a visible hyphen.
    const lastFrag = frags[frags.length - 1];
    if (!info.last && lastFrag && lastFrag.k === 'text' && lastFrag.text.endsWith('­')) lastFrag.text = `${lastFrag.text.slice(0, -1)}-`;
    for (const t of para.tabs ?? []) if (t.val === 'bar') bars.push(twip(t.pos));
    return { h, baseline, frags, end, breakType: info.breakType, clear: info.clear, floats, notes, bookmarks, last: info.last, left, right, bars: bars.length ? bars : undefined, textTop: baseline - Math.max(textAscent, objAsc), textBottom: baseline + desc };
  }
}

/** Number text for a page field. */
export function pageNumberText(n: number, fmt: string): string {
  if (fmt === 'numberInDash') return `- ${n} -`;
  return formatNumber(n, fmt);
}
