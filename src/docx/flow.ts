// SPDX-License-Identifier: MIT
// Block layout: paragraphs and tables flow into columns and pages. Floating objects take their positions from
// the anchor (page, margin, column, paragraph, line, character) and cut the text area according to their wrap
// type. Pages get headers and footers, footnotes at the bottom, columns per section; keep-with-next, keep
// lines together and widow/orphan control follow Word's rules.
import type { DocxDocument } from './document.js';
import { type Line, type LineEnv, ParagraphLines, type Segments, pageNumberText } from './line.js';
import type { Block, DrawingObj, Paragraph, Section, Table } from './model.js';
import { type NoteNumbers, createContext, parseBlocks } from './parse.js';
import { type Border, type ParaProps, isArtBorder, shadingColor, visible } from './props.js';
import { layoutTable } from './table.js';
import { eighthPt, kid, pt, twip, wa } from './xml.js';

// ---- Output ---------------------------------------------------------------------------------------------------

export interface LineItem {
  t: 'line';
  /** Column left edge and line top, absolute. */
  x: number;
  y: number;
  line: Line;
}

export interface FillItem {
  t: 'fill';
  x: number;
  y: number;
  w: number;
  h: number;
  color: string;
}

/** A border segment: horizontal when `vertical` is false (x..x+len at y), else vertical. */
export interface RuleItem {
  t: 'rule';
  x: number;
  y: number;
  len: number;
  vertical: boolean;
  b: Border;
}

/** Content drawn turned or clipped as a unit (cells with vertical text, clipped rows). */
export interface GroupItem {
  t: 'group';
  x: number;
  y: number;
  w: number;
  h: number;
  rotate: number;
  clip: boolean;
  items: Item[];
  floats: PlacedFloat[];
}

export type Item = LineItem | FillItem | RuleItem | GroupItem;

export interface PlacedFloat {
  obj: DrawingObj;
  /** Object box, absolute (or story coordinates). */
  x: number;
  y: number;
  w: number;
  h: number;
  behind: boolean;
  z: number;
  /** Laid out content of frames and floating tables. */
  story?: StoryLayout;
  /** Shift of the story inside the float box (a table's left edge left of the story origin). */
  storyOffset?: number;
  /** Vertical shift of the story (drop caps sit on a line of their paragraph). */
  storyDy?: number;
}

export interface StoryLayout {
  items: Item[];
  floats: PlacedFloat[];
  height: number;
  /** Extent of the content from x = 0 (or from `left`), for frames with automatic width and floating tables. */
  width: number;
  /** Leftmost content x when it starts left of 0 (tables hanging into the margin). */
  left?: number;
}

export interface NoteBox {
  layout: StoryLayout;
}

export interface PageBox {
  index: number;
  /** Displayed page number. */
  number: number;
  section: Section;
  sectionIndex: number;
  pageInSection: number;
  width: number;
  height: number;
  items: Item[];
  floats: PlacedFloat[];
  excl: Exclusion[];
  /** Header and footer parts (null: none). */
  header: string | null;
  footer: string | null;
  bodyTop: number;
  bodyBottom: number;
  marginLeft: number;
  marginRight: number;
  notes: NoteBox[];
  notesHeight: number;
  /** Column separators: x, top, bottom. */
  separators: [number, number, number][];
  /** Bookmarks that start on this page, with their y. */
  bookmarks: [string, number][];
  blank?: boolean;
}

export interface Exclusion {
  x0: number;
  x1: number;
  y0: number;
  y1: number;
  type: 'square' | 'tight' | 'through' | 'topAndBottom';
  side: string;
  /** Wrap polygon in absolute coordinates, for tight/through wrapping. */
  polygon?: [number, number][];
  distL: number;
  distR: number;
  id: number;
}

// ---- Environment ----------------------------------------------------------------------------------------------

export interface FlowContext {
  doc: DocxDocument;
  notes: NoteNumbers;
  /** Parsed header/footer/note stories by key. */
  stories: Map<string, Block[]>;
  /** Footnote layouts are needed during pagination; their content is parsed once. */
  pageNumber: { value: number; text: string; fmt: string; total: number; sectionPages: number };
  /** From the previous layout pass: page number text of each bookmark, pages per section. */
  bookmarks: Map<string, string>;
  /** Header and footer parts, read before the layout. */
  headers: HeaderDocs;
  sectionCounts: number[];
  budget: { lines: number };
}

const EPS = 0.01;
const MAX_PAGES = 10000;

export function storyBlocks(ctx: FlowContext, key: string, container: Element | null, source: { part: string; rels: Map<string, never> } | import('../pptx/env.js').Source, noteMark?: string): Block[] {
  let blocks = ctx.stories.get(key);
  if (!blocks) {
    const pc = createContext(ctx.doc, source, ctx.notes, false);
    if (noteMark !== undefined) pc.noteMark = noteMark;
    blocks = container ? parseBlocks(container, pc) : [];
    ctx.stories.set(key, blocks);
  }
  return blocks;
}

export function makeLineEnv(ctx: FlowContext): LineEnv {
  const { settings } = ctx.doc;
  return {
    theme: ctx.doc.theme,
    defaultTab: twip(settings.defaultTabStop),
    field: (name, format, arg, cached) => {
      const pn = ctx.pageNumber;
      if (name === 'page') return pageNumberText(pn.value, format ?? pn.fmt);
      if (name === 'numpages') return pageNumberText(pn.total, format ?? 'decimal');
      if (name === 'pageref') return (arg && ctx.bookmarks.get(arg)) || cached || '';
      return pageNumberText(pn.sectionPages, format ?? 'decimal');
    },
    expandShiftReturn: !settings.compat.has('doNotExpandShiftReturn'),
    compat15: settings.compatMode >= 15,
    decimal: settings.decimalSymbol,
  };
}

// ---- Helpers --------------------------------------------------------------------------------------------------

function sameBorders(a: ParaProps, b: ParaProps): boolean {
  return JSON.stringify(a.borders ?? null) === JSON.stringify(b.borders ?? null);
}

function hasBorders(p: ParaProps): boolean {
  const b = p.borders;
  return !!b && (visible(b.top) || visible(b.bottom) || visible(b.left) || visible(b.right) || visible(b.between));
}

/** Border width in px (double styles take three widths). */
export function borderWidth(b: Border | undefined): number {
  if (!visible(b)) return 0;
  if (isArtBorder(b)) return eighthPt(b.size * 8);
  const w = Math.max(eighthPt(b.size), 0.5);
  if (/^(double|triple|thinThick|thickThin)/.test(b.style)) return w * 3;
  return w;
}

function polygonExtent(polygon: [number, number][], y0: number, y1: number): [number, number] | null {
  let min = Infinity, max = -Infinity;
  const n = polygon.length;
  for (let i = 0; i < n; i++) {
    const [ax, ay] = polygon[i];
    const [bx, by] = polygon[(i + 1) % n];
    if (ay >= y0 && ay <= y1) {
      min = Math.min(min, ax);
      max = Math.max(max, ax);
    }
    // Edge crossings with the band's top and bottom.
    for (const yy of [y0, y1]) {
      if ((ay - yy) * (by - yy) < 0) {
        const x = ax + ((yy - ay) / (by - ay)) * (bx - ax);
        min = Math.min(min, x);
        max = Math.max(max, x);
      }
    }
  }
  return Number.isFinite(min) ? [min, max] : null;
}

// ---- Flow -----------------------------------------------------------------------------------------------------

interface Snapshot {
  page: number;
  items: number;
  floats: number;
  excl: number;
  notes: number;
  notesHeight: number;
  y: number;
  col: number;
  regionTop: number;
  group: BorderGroup | null;
}

interface BorderGroup {
  key: string;
  props: ParaProps;
  /** Items index of the group's fill and rules, rewritten as the group grows. */
  top: number;
  bottom: number;
  page: number;
  col: number;
  fill: FillItem | null;
  rules: RuleItem[];
  left: number;
  right: number;
}

export interface Column {
  x: number;
  w: number;
}

/**
 * A table continues a keep-with-next chain when every paragraph of its last row is kept with the next one
 * (Word keeps such a row with what follows).
 */
function tableKeepsWithNext(t: Table): boolean {
  const last = [...t.rows].reverse().find((r) => !r.props.hidden);
  if (!last) return false;
  let any = false;
  for (const cell of last.cells) {
    for (const b of cell.blocks) {
      if (b.t !== 'p' || b.hidden) continue;
      if (!b.props.keepNext) return false;
      any = true;
    }
  }
  return any;
}

/** Paragraph and table placement shared by pages and stories (cells, text boxes, headers, notes). */
export abstract class Flow {
  y = 0;
  col = 0;
  columns: Column[] = [{ x: 0, w: 0 }];
  regionTop = 0;
  items: Item[] = [];
  floats: PlacedFloat[] = [];
  excl: Exclusion[] = [];
  protected group: BorderGroup | null = null;
  /** Space after the previous paragraph and where it ended: Word uses the larger of after and before. */
  protected lastAfter = 0;
  protected lastAfterY = -1;
  /** A manual page or column break came just before: the next paragraph's space before is dropped. */
  protected afterBreak = false;
  /**
   * The page starts a new section: the first paragraph keeps the part of its space before that exceeds the
   * previous section's last space after (as if the two still collapsed across the break).
   */
  protected sectionCarry: number | null = null;
  readonly lineEnv: LineEnv;
  /** Floats already placed (by object id) on the current page/story. */
  protected placedIds = new Set<number>();

  constructor(readonly ctx: FlowContext) {
    this.lineEnv = makeLineEnv(ctx);
  }

  abstract get bottom(): number;
  /** Moves to the next column or page; false when the flow has no more room (stories never break). */
  abstract breakColumn(): boolean;
  abstract breakPage(): boolean;
  abstract snapshot(): Snapshot;
  abstract restore(s: Snapshot): void;
  /** Page geometry for anchored objects. */
  abstract frame(): { pageW: number; pageH: number; left: number; right: number; top: number; bottom: number; odd: boolean; ox: number; oy: number };
  /** Footnotes referenced by a line: returns the extra height they need, or places them. */
  noteSpace(_line: Line, _commit: boolean): number {
    return 0;
  }
  get paged(): boolean {
    return false;
  }

  get colX(): number {
    return this.columns[this.col].x;
  }

  get colW(): number {
    return this.columns[this.col].w;
  }

  atColumnTop(): boolean {
    return this.y <= this.regionTop + EPS;
  }

  // ---- Exclusions --------------------------------------------------------------------------------------------

  /** Free intervals (column coordinates) for a band, and the next y where the obstacles end. */
  segmentsAt(y: number, h: number): { segs: Segments; constrained: boolean; nextY: number } {
    const x0 = this.colX, w = this.colW;
    let segs: Segments = [[0, w]];
    let constrained = false;
    let nextY = Infinity;
    for (const e of this.excl) {
      if (e.y1 <= y + EPS || e.y0 >= y + h - EPS) continue;
      if (e.x1 <= x0 + EPS || e.x0 >= x0 + w - EPS) continue;
      let ex0 = e.x0 - x0, ex1 = e.x1 - x0;
      if (e.polygon && (e.type === 'tight' || e.type === 'through')) {
        const ext = polygonExtent(e.polygon, y, y + h);
        if (!ext) continue;
        ex0 = ext[0] - e.distL - x0;
        ex1 = ext[1] + e.distR - x0;
      }
      constrained = true;
      nextY = Math.min(nextY, e.y1);
      if (e.type === 'topAndBottom') {
        segs = [];
        continue;
      }
      let cut: [number, number];
      const leftSpace = Math.max(0, ex0), rightSpace = Math.max(0, w - ex1);
      if (e.side === 'left') cut = [ex0, w];
      else if (e.side === 'right') cut = [0, ex1];
      else if (e.side === 'largest') cut = leftSpace >= rightSpace ? [ex0, w] : [0, ex1];
      else cut = [ex0, ex1];
      const next: Segments = [];
      for (const [a, b] of segs) {
        if (cut[1] <= a || cut[0] >= b) next.push([a, b]);
        else {
          if (cut[0] > a) next.push([a, cut[0]]);
          if (cut[1] < b) next.push([cut[1], b]);
        }
      }
      segs = next;
    }
    // Word does not put text into slivers next to an object.
    if (constrained) segs = segs.filter(([a, b]) => b - a >= 18);
    return { segs, constrained, nextY };
  }

  // ---- Floats ------------------------------------------------------------------------------------------------

  /** Size of an object, including sizes relative to the page or margins. */
  protected floatSize(obj: DrawingObj): [number, number] {
    let { w, h } = obj;
    const a = obj.anchor;
    if (a?.sizeRelH || a?.sizeRelV) {
      const f = this.frame();
      const ref = (rel: string, horizontal: boolean) => {
        if (horizontal) return rel === 'page' ? f.pageW : rel === 'leftMargin' ? f.left : rel === 'rightMargin' ? f.right : f.pageW - f.left - f.right;
        return rel === 'page' ? f.pageH : rel === 'topMargin' ? f.top : rel === 'bottomMargin' ? f.bottom : f.pageH - f.top - f.bottom;
      };
      if (a.sizeRelH) w = ref(a.sizeRelH[0], true) * a.sizeRelH[1];
      if (a.sizeRelV) h = ref(a.sizeRelV[0], false) * a.sizeRelV[1];
    }
    return [w, h];
  }

  /** Position of an anchored object; `paraTop`/`lineTop` and `charX` are where its anchor is. */
  protected positionFloat(obj: DrawingObj, paraTop: number, lineTop: number, charX: number): { x: number; y: number; w: number; h: number } {
    const a = obj.anchor!;
    const f = this.frame();
    const [w, h] = this.floatSize(obj);
    const col = this.columns[this.col];
    // Horizontal reference area in absolute coordinates.
    let x0 = 0, x1 = f.pageW;
    const rel = a.posH.relativeFrom;
    switch (rel) {
      case 'margin':
        x0 = f.left;
        x1 = f.pageW - f.right;
        break;
      case 'column':
        x0 = col.x + f.ox;
        x1 = col.x + col.w + f.ox;
        break;
      case 'character':
        x0 = x1 = charX + f.ox;
        break;
      case 'leftMargin':
        x1 = f.left;
        break;
      case 'rightMargin':
        x0 = f.pageW - f.right;
        break;
      case 'insideMargin':
        if (f.odd) x1 = f.left;
        else x0 = f.pageW - f.right;
        break;
      case 'outsideMargin':
        if (f.odd) x0 = f.pageW - f.right;
        else x1 = f.left;
        break;
    }
    let x: number;
    let align = a.posH.align;
    if (align === 'inside') align = f.odd ? 'left' : 'right';
    else if (align === 'outside') align = f.odd ? 'right' : 'left';
    // Aligned objects line up with their whole extent (rotation, shadows), offsets refer to the object itself.
    const e = obj.effect;
    if (align === 'left') x = x0 + e.l;
    else if (align === 'right') x = x1 - w - e.r;
    else if (align === 'center') x = (x0 + x1 - w) / 2 + (e.l - e.r) / 2;
    else if (a.posH.pct !== undefined) x = x0 + (x1 - x0) * a.posH.pct;
    else x = x0 + (a.posH.offset ?? 0);

    let y0 = 0, y1 = f.pageH;
    const vrel = a.posV.relativeFrom;
    switch (vrel) {
      case 'margin':
        y0 = f.top;
        y1 = f.pageH - f.bottom;
        break;
      case 'paragraph':
        y0 = y1 = paraTop + f.oy;
        break;
      case 'line':
        y0 = y1 = lineTop + f.oy;
        break;
      case 'topMargin':
        y1 = f.top;
        break;
      case 'bottomMargin':
        y0 = f.pageH - f.bottom;
        break;
      case 'insideMargin':
      case 'outsideMargin':
        y1 = f.top;
        break;
    }
    let y: number;
    const valign = a.posV.align;
    if (valign === 'top' || valign === 'inside') y = y0 + e.t;
    else if (valign === 'bottom' || valign === 'outside') y = y1 - h - e.b;
    else if (valign === 'center') y = (y0 + y1 - h) / 2 + (e.t - e.b) / 2;
    else if (a.posV.pct !== undefined) y = y0 + (y1 - y0) * a.posV.pct;
    else y = y0 + (a.posV.offset ?? 0);
    // Story coordinates.
    return { x: x - f.ox, y: y - f.oy, w, h };
  }

  /** Places an anchored object and makes text wrap around it. */
  registerFloat(obj: DrawingObj, paraTop: number, lineTop: number, charX: number, story?: StoryLayout): PlacedFloat | null {
    if (this.placedIds.has(obj.id)) return null;
    this.placedIds.add(obj.id);
    const a = obj.anchor!;
    const pos = this.positionFloat(obj, paraTop, lineTop, charX);
    const placed: PlacedFloat = { obj, x: pos.x, y: pos.y, w: pos.w, h: pos.h, behind: a.behindDoc, z: a.z, story };
    this.floats.push(placed);
    const wrap = a.wrap;
    if (wrap.type !== 'none') {
      const e = obj.effect;
      const ex: Exclusion = {
        x0: pos.x - e.l - wrap.dist.l,
        x1: pos.x + pos.w + e.r + wrap.dist.r,
        y0: pos.y - e.t - wrap.dist.t,
        y1: pos.y + pos.h + e.b + wrap.dist.b,
        type: wrap.type,
        side: wrap.side,
        distL: wrap.dist.l,
        distR: wrap.dist.r,
        id: obj.id,
      };
      if (wrap.polygon && wrap.polygon.length >= 3 && (wrap.type === 'tight' || wrap.type === 'through')) {
        ex.polygon = wrap.polygon.map(([px, py]) => [pos.x + (px / 21600) * pos.w, pos.y + (py / 21600) * pos.h]);
      }
      if (wrap.type === 'topAndBottom') {
        ex.x0 = Math.min(ex.x0, this.colX);
        ex.x1 = Math.max(ex.x1, this.colX + this.colW);
      }
      this.excl.push(ex);
      this.onLateFloat(ex);
    }
    return placed;
  }

  /** A float that reaches above the current position; pages lay out again with it known from the start. */
  protected onLateFloat(_ex: Exclusion): void {}

  // ---- Paragraphs --------------------------------------------------------------------------------------------

  /** Indents in px; character-based indents use the paragraph's font size. */
  paraGeometry(p: Paragraph): { colWidth: number; indLeft: number; indRight: number; indFirst: number } {
    const props = p.props;
    const charW = ((p.mark.sz ?? 20) / 2) * (4 / 3);
    const left = props.indLeftChars !== undefined && props.indLeftChars !== 0 ? (props.indLeftChars / 100) * charW : twip(props.indLeft ?? 0);
    const right = props.indRightChars !== undefined && props.indRightChars !== 0 ? (props.indRightChars / 100) * charW : twip(props.indRight ?? 0);
    const first = props.indFirstChars !== undefined && props.indFirstChars !== 0 ? (props.indFirstChars / 100) * charW : twip(props.indFirst ?? 0);
    if (props.bidi) return { colWidth: this.colW, indLeft: right, indRight: left, indFirst: first };
    return { colWidth: this.colW, indLeft: left, indRight: right, indFirst: first };
  }

  /** Spacing before/after in px; contextual spacing drops it between paragraphs of one style. */
  spacing(p: Paragraph, prev: Block | null, next: Block | null): { before: number; after: number } {
    const props = p.props;
    const lineUnit = 12 * (4 / 3);
    let before = props.spBeforeLines ? (props.spBeforeLines / 100) * lineUnit : twip(props.spBefore ?? 0);
    let after = props.spAfterLines ? (props.spAfterLines / 100) * lineUnit : twip(props.spAfter ?? 0);
    if (props.spBeforeAuto) before = pt(14);
    if (props.spAfterAuto) after = pt(14);
    if (props.contextualSpacing) {
      if (prev && prev.t === 'p' && prev.styleId === p.styleId) before = 0;
      if (next && next.t === 'p' && next.styleId === p.styleId) after = 0;
    }
    // HTML-style automatic spacing collapses with the neighbour's.
    if (props.spBeforeAuto && prev && prev.t === 'p' && prev.props.spAfterAuto) before = 0;
    return { before, after };
  }

  protected gridPitch(): number | undefined {
    return undefined;
  }

  /** Places a paragraph at the current position. */
  placeParagraph(p: Paragraph, prev: Block | null, next: Block | null, attempt = 0): void {
    if (p.hidden) return;
    const props = p.props;
    if (props.pageBreakBefore && this.paged && !this.atPageTop()) this.breakPage();
    let { before } = this.spacing(p, prev, next);
    const { after } = this.spacing(p, prev, next);
    // Word drops the space before a paragraph at the top of a page or column, except on the first page and at
    // the start of a section.
    if (this.paged && this.atColumnTop() && this.sectionCarry !== null && !this.afterBreak && !props.pageBreakBefore) before = Math.max(0, before - this.sectionCarry);
    else if (this.paged && this.atColumnTop() && (this.afterBreak || props.pageBreakBefore || !this.firstPage())) before = 0;
    else if (Math.abs(this.lastAfterY - this.y) < EPS) before = Math.max(0, before - this.lastAfter);
    this.afterBreak = false;
    this.sectionCarry = null;
    const lines = new ParagraphLines(p, this.lineEnv, this.paraGeometry(p), props.snapToGrid === false ? undefined : this.gridPitch());
    const snap = this.snapshot();
    const startedAtTop = this.atColumnTop();

    // Borders: a group of paragraphs with the same borders shares one box.
    const bordered = hasBorders(props);
    const joins = bordered && this.group && this.group.key === JSON.stringify(props.borders) && sameBorders(this.group.props, props) && this.group.col === this.col && this.group.page === this.pageIndex();
    const b = props.borders ?? {};
    // The box goes on when the next paragraph has the same borders: no bottom edge here, and the line between
    // the two gets its space on both sides.
    const continues = bordered && !!next && next.t === 'p' && !next.hidden && JSON.stringify(next.props.borders ?? null) === JSON.stringify(props.borders ?? null);
    const topPad = bordered && !joins && visible(b.top) ? borderWidth(b.top) + pt(b.top.space) : 0;
    const betweenPad = joins && visible(b.between) ? borderWidth(b.between) + 2 * pt(b.between.space) : 0;
    const bottomPad = bordered && !continues && visible(b.bottom) ? borderWidth(b.bottom) + pt(b.bottom.space) : 0;

    this.y += before;
    const betweenY = this.y;
    this.y += topPad + betweenPad;
    const paraTop = this.y;
    let partTop = this.y;
    let partItems = this.items.length;

    // Frames written before this paragraph belong to it; if they do not fit here, the paragraph moves on.
    if (this.pendingFrames.length) {
      const frames = this.pendingFrames;
      this.pendingFrames = [];
      if (!this.placeFrames(frames, paraTop, lines) && attempt < 2) {
        this.restore(snap);
        this.pendingFrames = frames;
        if (this.breakColumn()) return this.placeParagraph(p, null, next, attempt + 1);
      }
    }
    // Floats anchored in this paragraph; line-relative ones follow their line.
    const anchorFloats = lines.atoms.filter((a) => a.k === 'float' && a.obj?.anchor && a.obj.anchor.posV.relativeFrom !== 'line');
    for (const a of anchorFloats) this.registerFloat(a.obj!, paraTop, paraTop, this.colX + lines.geo.indLeft);

    if (props.keepLines && this.paged && !startedAtTop && attempt === 0) {
      // Measure the paragraph: if it does not fit in what is left of the column, it starts in the next one.
      const height = this.measureParagraph(lines);
      if (this.y + height > this.bottom + EPS) {
        this.restore(snap);
        if (this.breakColumn()) return this.placeParagraph(p, null, next, attempt + 1);
      }
    }

    let state = 0;
    let first = true;
    let index = 0;
    let endedWithBreak = false;
    let markLine = false;
    const placedHere: { item: number; start: number; h: number }[] = [];
    let guard = 0;
    while (state < lines.atoms.length || first || markLine) {
      markLine = false;
      if (++guard > 100000 || this.ctx.budget.lines-- <= 0) break;
      // Word wraps around objects where the text itself is, not where the line spacing reaches.
      const estimate = lines.estimateHeight();
      const [et, eb] = lines.estimateText();
      let band = this.segmentsAt(this.y + et, eb - et);
      let line = band.segs.length ? lines.line(state, band.segs, first, band.constrained) : null;
      if (line && band.constrained && (line.textTop < et - EPS || line.textBottom > eb + EPS)) {
        const again = this.segmentsAt(this.y + line.textTop, line.textBottom - line.textTop);
        if (JSON.stringify(again.segs) !== JSON.stringify(band.segs)) {
          band = again;
          line = again.segs.length ? lines.line(state, again.segs, first, again.constrained) : null;
        }
      }
      if (!line) {
        // Blocked by floating objects: continue below them.
        const nextY = Number.isFinite(band.nextY) ? band.nextY : this.y + estimate;
        if (nextY + estimate > this.bottom + EPS && this.paged) {
          if (!this.breakColumn()) this.y = nextY;
        } else this.y = Math.max(this.y + 0.5, nextY);
        continue;
      }
      const notes = this.noteSpace(line, false);
      const fits = this.y + line.h + notes <= this.bottom + EPS;
      if (!fits && this.paged && !(this.atColumnTop() && placedHere.length === 0)) {
        const linesHere = placedHere.length;
        const widow = props.widowControl !== false;
        if (index === 0 || (widow && index === 1 && !line.last ? false : false)) {
          // Nothing of the paragraph here yet: the whole paragraph goes to the next column.
          if (attempt < 2) {
            this.restore(snap);
            if (this.breakColumn()) return this.placeParagraph(p, null, next, attempt + 1);
          } else if (this.breakColumn()) continue;
        }
        if (widow && linesHere === 1 && index === 1 && attempt < 2) {
          // A single first line would be left behind (orphan): move the paragraph.
          this.restore(snap);
          if (this.breakColumn()) return this.placeParagraph(p, null, next, attempt + 1);
        }
        if (widow && line.last && linesHere >= 2) {
          // The last line alone would start the next page (widow): take one more line along.
          const moved = placedHere.pop()!;
          this.items.splice(moved.item, 1);
          state = moved.start;
          index--;
          first = state === 0;
        }
        this.closeParagraphPart(p, partTop, partItems, false);
        this.breakColumn();
        partTop = this.y;
        partItems = this.items.length;
        placedHere.length = 0;
        continue;
      }
      if (notes > 0) this.noteSpace(line, true);
      this.items.push({ t: 'line', x: this.colX, y: this.y, line });
      placedHere.push({ item: this.items.length - 1, start: state, h: line.h });
      for (const f of line.floats) {
        if (f.obj.anchor?.posV.relativeFrom === 'line') this.registerFloat(f.obj, paraTop, this.y, this.colX + f.x);
      }
      this.onLine(line);
      this.y += line.h;
      state = line.end;
      first = false;
      index++;
      endedWithBreak = line.breakType === 'page' && line.last;
      // After a column break the paragraph mark goes on to the next column as an empty line.
      if (line.breakType === 'column' && line.last && this.paged) markLine = true;
      if (line.breakType === 'page' && this.paged) {
        this.closeParagraphPart(p, partTop, partItems, false);
        this.breakPage();
        partTop = this.y;
        partItems = this.items.length;
        placedHere.length = 0;
      } else if (line.breakType === 'column' && this.paged) {
        this.closeParagraphPart(p, partTop, partItems, false);
        this.breakColumn();
        partTop = this.y;
        partItems = this.items.length;
        placedHere.length = 0;
      }
      if (line.clear) this.clearFloats(line.clear);
      if (line.last && !markLine) break;
    }
    this.y += bottomPad;
    this.closeParagraphPart(p, partTop, partItems, true, { betweenY, joins: !!joins, topPad, bottomPad, after: endedWithBreak ? 0 : after, continues });
    if (endedWithBreak && this.paged) {
      // The paragraph mark after a final break stays on the page before; its spacing does not carry over.
      this.afterBreak = true;
      this.lastAfter = 0;
      this.lastAfterY = -1;
      return;
    }
    this.y += after;
    this.lastAfter = after;
    this.lastAfterY = this.y;
  }

  protected atPageTop(): boolean {
    return this.atColumnTop() && this.col === 0;
  }

  /** On the first page of the document (where Word keeps the space before the first paragraph). */
  protected firstPage(): boolean {
    return true;
  }

  protected pageIndex(): number {
    return 0;
  }

  protected onLine(_line: Line): void {}

  /** Height of a paragraph laid out without obstacles (for keep-together decisions). */
  measureParagraph(lines: ParagraphLines): number {
    let h = 0, state = 0, first = true, guard = 0;
    while ((state < lines.atoms.length || first) && guard++ < 10000) {
      const line = lines.line(state, [[0, this.colW]], first);
      if (!line) break;
      h += line.h;
      state = line.end;
      first = false;
      if (line.last) break;
    }
    return h;
  }

  /** Moves below floats on one side (w:br w:clear). */
  protected clearFloats(side: string): void {
    let y = this.y;
    for (const e of this.excl) {
      if (e.y0 > this.y || e.y1 <= this.y) continue;
      const center = (e.x0 + e.x1) / 2 - this.colX;
      const onLeft = center < this.colW / 2;
      if (side === 'all' || (side === 'left' && onLeft) || (side === 'right' && !onLeft)) y = Math.max(y, e.y1);
    }
    this.y = y;
  }

  /** Shading and borders behind the part of a paragraph in the current column. */
  protected closeParagraphPart(p: Paragraph, top: number, firstItem: number, final: boolean, info?: { betweenY: number; joins: boolean; topPad: number; bottomPad: number; after: number; continues: boolean }): void {
    const props = p.props;
    const shade = shadingColor(props.shd);
    const bordered = hasBorders(props);
    if (!shade && !bordered) {
      if (final) this.group = null;
      return;
    }
    const geo = this.paraGeometry(p);
    const b = props.borders ?? {};
    const space = (x: Border | undefined) => (visible(x) ? pt(x.space) + borderWidth(x) : 0);
    const left = this.colX + geo.indLeft - space(b.left);
    const right = this.colX + this.colW - geo.indRight + space(b.right);
    const bottom = this.y;
    const joins = info?.joins && this.group;
    if (joins && this.group) {
      // Extend the box of the previous paragraph.
      const g = this.group;
      if (g.fill && shade) g.fill.h = bottom - g.fill.y;
      for (const r of g.rules) if (r.vertical) r.len = bottom - r.y;
      if (visible(b.between) && info) {
        const rule: RuleItem = { t: 'rule', x: left, y: info.betweenY + pt(b.between.space) + borderWidth(b.between) / 2, len: right - left, vertical: false, b: b.between };
        this.items.splice(firstItem, 0, rule);
        g.rules.push(rule);
      }
      if (visible(b.bottom) && final && !info?.continues) {
        const rule: RuleItem = { t: 'rule', x: left, y: bottom - borderWidth(b.bottom) / 2, len: right - left, vertical: false, b: b.bottom };
        this.items.push(rule);
        g.rules.push(rule);
      }
      if (!final) this.group = null;
      return;
    }
    const boxTop = top - (info && !info.joins ? info.topPad : 0);
    const fills: Item[] = [];
    let fill: FillItem | null = null;
    if (shade) {
      // Without borders, Word's shading also covers the space after the paragraph.
      const extra = !bordered && info ? info.after : 0;
      fill = { t: 'fill', x: left, y: boxTop, w: right - left, h: bottom - boxTop + extra, color: shade };
      fills.push(fill);
    }
    const rules: RuleItem[] = [];
    if (bordered) {
      if (visible(b.top) && (info ? !info.joins : true)) rules.push({ t: 'rule', x: left, y: boxTop + borderWidth(b.top) / 2, len: right - left, vertical: false, b: b.top });
      if (visible(b.bottom)) rules.push({ t: 'rule', x: left, y: bottom - borderWidth(b.bottom) / 2, len: right - left, vertical: false, b: b.bottom });
      if (visible(b.left)) rules.push({ t: 'rule', x: left + borderWidth(b.left) / 2, y: boxTop, len: bottom - boxTop, vertical: true, b: b.left });
      if (visible(b.right)) rules.push({ t: 'rule', x: right - borderWidth(b.right) / 2, y: boxTop, len: bottom - boxTop, vertical: true, b: b.right });
    }
    // Behind the lines of the paragraph.
    this.items.splice(firstItem, 0, ...fills, ...rules);
    if (final && bordered) {
      this.group = { key: JSON.stringify(props.borders), props, top: boxTop, bottom, page: this.pageIndex(), col: this.col, fill, rules, left, right };
    } else if (final && shade && !bordered) {
      this.group = null;
    } else if (!final) {
      this.group = null;
    }
  }

  // ---- Tables ------------------------------------------------------------------------------------------------

  placeTable(t: Table, prev: Block | null): void {
    this.group = null;
    this.lastAfterY = -1;
    if (t.props.tblpPr) this.placeFloatingTable(t);
    else layoutTable(this, t, prev);
  }

  /** A floating table (w:tblpPr): laid out on its own and placed like an anchored object with square wrapping. */
  private placeFloatingTable(t: Table): void {
    const p = t.props.tblpPr!;
    const n = (name: string) => twip(Number(wa(p, name) ?? 0) || 0);
    const story = layoutStory(this.ctx, [{ ...t, props: { ...t.props, tblpPr: undefined } }], this.colW);
    const hAnchor = wa(p, 'horzAnchor') ?? 'text', vAnchor = wa(p, 'vertAnchor') ?? 'text';
    const xSpec = wa(p, 'tblpXSpec'), ySpec = wa(p, 'tblpYSpec');
    const obj: DrawingObj = {
      id: t.id,
      graphic: { kind: 'table', table: t },
      w: story.width,
      h: story.height,
      effect: { t: 0, b: 0, l: 0, r: 0 },
      anchor: {
        posH: { relativeFrom: hAnchor === 'page' ? 'page' : hAnchor === 'margin' ? 'margin' : 'column', ...(xSpec ? { align: xSpec } : { offset: n('tblpX') }) },
        posV: { relativeFrom: vAnchor === 'page' ? 'page' : vAnchor === 'margin' ? 'margin' : 'paragraph', ...(ySpec && ySpec !== 'inline' ? { align: ySpec } : { offset: n('tblpY') }) },
        behindDoc: false,
        z: 0,
        wrap: { type: 'square', side: 'bothSides', dist: { t: n('topFromText'), b: n('bottomFromText'), l: n('leftFromText'), r: n('rightFromText') } },
        layoutInCell: true,
        allowOverlap: t.props.overlap !== 'never',
      },
    };
    // The table's own left edge sits at its x; the story starts at the column.
    const placed = this.registerFloat(obj, this.y, this.y, this.colX, story);
    if (placed && story.left) placed.storyOffset = story.left;
  }

  /** Places blocks one after another; keep-with-next chains move as a whole when they would be split. */
  placeBlocks(blocks: Block[], onBlock?: (i: number) => void): void {
    let chain: Snapshot | null = null;
    let chainStart = -1;
    let chainRetried = new Set<number>();
    for (let i = 0; i < blocks.length; i++) {
      onBlock?.(i);
      const block = blocks[i];
      const prev = i > 0 ? blocks[i - 1] : null;
      const next = i + 1 < blocks.length ? blocks[i + 1] : null;
      const keep = this.paged && (block.t === 'p' ? !!block.props.keepNext && !block.hidden : tableKeepsWithNext(block));
      if (keep && !chain) {
        chain = this.snapshot();
        chainStart = i;
      }
      const pageBefore = this.pageIndex();
      const colBefore = this.col;
      if (block.t === 'p') {
        if (block.props.framePr && block.props.framePr.dropCap !== 'none') {
          i = this.placeFrame(blocks, i) - 1;
          continue;
        }
        this.placeParagraph(block, prev, next);
      } else this.placeTable(block, prev);
      if (chain && !keep) {
        // The chain ended with this block: did a break come between the chain and this block's first line?
        const moved = this.pageIndex() !== pageBefore || this.col !== colBefore;
        const chainPage = chain.page, chainCol = chain.col;
        const split = moved || chainPage !== pageBefore || chainCol !== colBefore;
        const atTop = chain.y <= chain.regionTop + EPS;
        if (split && !atTop && !chainRetried.has(chainStart) && this.chainBreakable(chain)) {
          chainRetried.add(chainStart);
          this.restore(chain);
          this.breakColumn();
          i = chainStart - 1;
          chain = null;
          continue;
        }
        chain = null;
      }
    }
    void chainRetried;
    // Frames at the very end have no paragraph after them: they stay where the text ends.
    if (this.pendingFrames.length) {
      for (const f of this.pendingFrames.splice(0)) this.registerFloat(f.obj, this.y, this.y, this.colX, f.story);
    }
  }

  /** Whether the start of a keep-with-next chain may be moved to the next column. */
  protected chainBreakable(_s: Snapshot): boolean {
    return this.paged;
  }

  /**
   * Text frames (w:framePr): consecutive paragraphs with the same frame properties form a box positioned like
   * a floating object. Returns the index after the frame.
   */
  placeFrame(blocks: Block[], start: number): number {
    const first = blocks[start] as Paragraph;
    const key = JSON.stringify(first.props.framePr);
    let end = start + 1;
    while (end < blocks.length) {
      const b = blocks[end];
      if (b.t !== 'p' || JSON.stringify(b.props.framePr) !== key) break;
      end++;
    }
    const fp = first.props.framePr!;
    const frameBlocks = blocks.slice(start, end).map((b) => ({ ...(b as Paragraph), props: { ...(b as Paragraph).props, framePr: undefined } }));
    const width = fp.w ? twip(fp.w) : this.colW;
    const story = layoutStory(this.ctx, frameBlocks, width);
    const w = fp.w ? twip(fp.w) : Math.min(this.colW, story.width + 1);
    const h = fp.h ? (fp.hRule === 'exact' ? twip(fp.h) : Math.max(twip(fp.h), story.height)) : story.height;
    const hSpace = twip(fp.hSpace ?? 0), vSpace = twip(fp.vSpace ?? 0);
    const posH = fp.xAlign ? { relativeFrom: fp.hAnchor === 'page' ? 'page' : fp.hAnchor === 'margin' ? 'margin' : 'column', align: fp.xAlign } : { relativeFrom: fp.hAnchor === 'page' ? 'page' : fp.hAnchor === 'margin' ? 'margin' : 'column', offset: twip(fp.x ?? 0) };
    const posV = fp.yAlign ? { relativeFrom: fp.vAnchor === 'page' ? 'page' : fp.vAnchor === 'margin' ? 'margin' : 'paragraph', align: fp.yAlign } : { relativeFrom: fp.vAnchor === 'page' ? 'page' : fp.vAnchor === 'margin' ? 'margin' : 'paragraph', offset: twip(fp.y ?? 0) };
    const obj: DrawingObj = {
      id: first.id,
      graphic: { kind: 'frame', blocks: frameBlocks },
      w,
      h,
      effect: { t: 0, b: 0, l: 0, r: 0 },
      anchor: {
        posH,
        posV,
        behindDoc: false,
        z: 0,
        wrap: { type: fp.wrap === 'none' ? 'none' : fp.wrap === 'notBeside' ? 'topAndBottom' : 'square', side: 'bothSides', dist: { t: vSpace, b: vSpace, l: hSpace, r: hSpace } },
        layoutInCell: true,
        allowOverlap: true,
      },
    };
    // A frame belongs to the paragraph after it and is placed with that paragraph (on its page).
    this.pendingFrames.push({ obj, story, dropCap: fp.dropCap === 'drop' || fp.dropCap === 'margin' ? (fp.lines ?? 3) : 0 });
    return end;
  }

  /** Frames waiting for the paragraph they are anchored to. */
  protected pendingFrames: { obj: DrawingObj; story: StoryLayout; dropCap: number }[] = [];

  /**
   * Places the frames anchored to a paragraph that starts at `paraTop`. A drop cap sits at the left edge with its
   * baseline on the baseline of the paragraph's line `lines`. Returns false when a frame does not fit on the
   * page below the paragraph's start.
   */
  protected placeFrames(frames: { obj: DrawingObj; story: StoryLayout; dropCap: number }[], paraTop: number, lines: ParagraphLines): boolean {
    for (const f of frames) {
      if (f.dropCap) {
        const firstLine = f.story.items.find((i): i is LineItem => i.t === 'line');
        const pitch = lines.estimateHeight();
        const [, textBottom] = lines.estimateText();
        const baseline = textBottom - (lines.markDesc ?? 0) + pitch * (f.dropCap - 1);
        const capBaseline = firstLine ? firstLine.y + firstLine.line.baseline : f.story.height;
        // The frame takes exactly its lines; the letter is moved so its baseline sits on the last of them.
        f.obj.anchor = { ...f.obj.anchor!, posH: { relativeFrom: 'column', offset: lines.geo.indLeft }, posV: { relativeFrom: 'paragraph', offset: 0 } };
        f.obj.h = pitch * f.dropCap;
        const placedCap = this.registerFloat(f.obj, paraTop, paraTop, this.colX, f.story);
        if (placedCap) placedCap.storyDy = baseline - capBaseline;
        continue;
      }
      const placed = this.registerFloat(f.obj, paraTop, paraTop, this.colX, f.story);
      if (placed && this.paged && placed.y + placed.h > this.bottom + EPS && !this.atColumnTop()) return false;
    }
    return true;
  }
}

// ---- Stories --------------------------------------------------------------------------------------------------

/** A single column of unlimited height: table cells, text boxes, headers, footers, notes, frames. */
export class StoryFlow extends Flow {
  constructor(
    ctx: FlowContext,
    width: number,
    private readonly page?: { pageW: number; pageH: number; left: number; right: number; top: number; bottom: number; odd: boolean; ox: number; oy: number },
  ) {
    super(ctx);
    this.columns = [{ x: 0, w: Math.max(1, width) }];
  }

  get bottom(): number {
    return Infinity;
  }

  breakColumn(): boolean {
    return false;
  }

  breakPage(): boolean {
    return false;
  }

  snapshot(): Snapshot {
    return { page: 0, items: this.items.length, floats: this.floats.length, excl: this.excl.length, notes: 0, notesHeight: 0, y: this.y, col: 0, regionTop: this.regionTop, group: this.group };
  }

  restore(s: Snapshot): void {
    this.items.length = s.items;
    for (const f of this.floats.splice(s.floats)) this.placedIds.delete(f.obj.id);
    this.excl.length = s.excl;
    this.y = s.y;
    this.group = s.group;
  }

  frame() {
    return this.page ?? { pageW: this.colW, pageH: Math.max(this.y, 1), left: 0, right: 0, top: 0, bottom: 0, odd: true, ox: 0, oy: 0 };
  }

  result(): StoryLayout {
    let right = 0, left = 0;
    for (const item of this.items) {
      const [l, r] =
        item.t === 'line' ? [item.x + item.line.left, item.x + item.line.right] : item.t === 'rule' ? [item.x, item.x + (item.vertical ? 0 : item.len)] : [item.x, item.x + item.w];
      right = Math.max(right, r);
      left = Math.min(left, l);
    }
    return { items: this.items, floats: this.floats, height: this.y, width: right - left || this.colW, left: left < 0 ? left : undefined };
  }
}

/** Lays out blocks in a box of the given width. */
export function layoutStory(ctx: FlowContext, blocks: Block[], width: number, page?: ConstructorParameters<typeof StoryFlow>[2]): StoryLayout {
  const flow = new StoryFlow(ctx, width, page);
  flow.placeBlocks(blocks);
  return flow.result();
}

// ---- Pages ----------------------------------------------------------------------------------------------------

interface SectionRun {
  section: Section;
  blocks: Block[];
}

/** Splits the body into sections: a paragraph with sectPr ends its section. */
export function sections(blocks: Block[], final: Section): SectionRun[] {
  const out: SectionRun[] = [];
  let current: Block[] = [];
  for (const b of blocks) {
    current.push(b);
    const s = b.t === 'p' ? b.section : undefined;
    if (s) {
      out.push({ section: s, blocks: current });
      current = [];
    }
  }
  out.push({ section: final, blocks: current });
  return out;
}

export class PageFlow extends Flow {
  pages: PageBox[] = [];
  private page!: PageBox;
  private sectionIndex = 0;
  private section!: Section;
  private pageInSection = 0;
  private sectionList: Section[] = [];
  /** Floats found late in the previous pass, by page index. */
  constructor(
    ctx: FlowContext,
    private readonly preFloats: Map<number, Map<number, PlacedFloat & { ex?: Exclusion }>> = new Map(),
    readonly lateFloats: Map<number, Map<number, PlacedFloat & { ex?: Exclusion }>> = new Map(),
  ) {
    super(ctx);
  }

  override get paged(): boolean {
    return true;
  }

  get bottom(): number {
    return this.page.bodyBottom - this.page.notesHeight;
  }

  protected override pageIndex(): number {
    return this.page.index;
  }

  protected override gridPitch(): number | undefined {
    const g = this.section.docGrid;
    return g && (g.type === 'lines' || g.type === 'linesAndChars') && g.linePitch > 0 ? twip(g.linePitch) : undefined;
  }

  frame() {
    const p = this.page;
    return { pageW: p.width, pageH: p.height, left: p.marginLeft, right: p.marginRight, top: twip(Math.abs(this.section.top)), bottom: twip(Math.abs(this.section.bottom)), odd: p.number % 2 === 1, ox: 0, oy: 0 };
  }

  snapshot(): Snapshot {
    const p = this.page;
    return { page: p.index, items: p.items.length, floats: p.floats.length, excl: p.excl.length, notes: p.notes.length, notesHeight: p.notesHeight, y: this.y, col: this.col, regionTop: this.regionTop, group: this.group };
  }

  restore(s: Snapshot): void {
    if (this.page.index !== s.page) {
      for (const dropped of this.pages.splice(s.page + 1)) for (const f of dropped.floats) this.placedIds.delete(f.obj.id);
      this.page = this.pages[s.page];
      this.items = this.page.items;
      this.floats = this.page.floats;
      this.excl = this.page.excl;
      this.applyColumns(this.page);
      this.pageInSection = this.page.pageInSection + 1;
    }
    const p = this.page;
    p.items.length = s.items;
    for (const f of p.floats.splice(s.floats)) this.placedIds.delete(f.obj.id);
    p.excl.length = s.excl;
    p.notes.length = s.notes;
    p.notesHeight = s.notesHeight;
    this.y = s.y;
    this.col = s.col;
    this.regionTop = s.regionTop;
    this.group = s.group;
    this.ctx.pageNumber.value = p.number;
  }

  /** Column geometry of the current section on a page, from `top` down. */
  private sectionColumns(page: PageBox, s: Section): Column[] {
    const left = page.marginLeft, width = page.width - page.marginLeft - page.marginRight;
    const { num, space, equal, list } = s.cols;
    if (num <= 1) return [{ x: left, w: width }];
    if (equal || list.length < num) {
      const sp = twip(space);
      const w = (width - sp * (num - 1)) / num;
      return Array.from({ length: num }, (_, i) => ({ x: left + i * (w + sp), w }));
    }
    const out: Column[] = [];
    let x = left;
    for (const c of list.slice(0, num)) {
      out.push({ x, w: twip(c.w) });
      x += twip(c.w) + twip(c.space);
    }
    return out;
  }

  private applyColumns(page: PageBox): void {
    this.columns = this.sectionColumns(page, page.section);
  }

  private headerFor(kind: 'headers' | 'footers', type: 'default' | 'first' | 'even'): string | null {
    for (let i = this.sectionIndex; i >= 0; i--) {
      const part = this.sectionList[i]?.[kind][type];
      if (part) return part;
    }
    return null;
  }

  /** Header/footer story laid out for the current page. */
  headerLayout(part: string, page: PageBox, footer: boolean): StoryLayout {
    const s = page.section;
    const width = page.width - page.marginLeft - page.marginRight;
    const blocks = this.headerBlocks(part);
    const top = footer ? 0 : twip(s.header);
    return layoutStory(this.ctx, blocks, width, {
      pageW: page.width, pageH: page.height, left: page.marginLeft, right: page.marginRight, top: twip(Math.abs(s.top)), bottom: twip(Math.abs(s.bottom)),
      odd: page.number % 2 === 1, ox: page.marginLeft, oy: top,
    });
  }

  headerBlocks(part: string): Block[] {
    let blocks = this.ctx.stories.get(part);
    if (!blocks) {
      blocks = [];
      this.ctx.stories.set(part, blocks);
      // Parsed on first use; the part was already read by loadHeaders.
      const doc = this.ctx.headers.get(part);
      if (doc) {
        const pc = createContext(this.ctx.doc, doc.source, this.ctx.notes, false);
        blocks.push(...parseBlocks(doc.root, pc));
      }
    }
    return blocks;
  }

  /** Starts a new page for the current section; a blank page (inserted for odd/even page order) stays empty. */
  newPage(blank = false): void {
    if (this.pages.length >= MAX_PAGES) throw new PageLimit();
    if (this.page) this.finishPage(this.page);
    const s = this.section;
    const prev = this.pages[this.pages.length - 1];
    const firstOfSection = this.pageInSection === 0 && !blank;
    const number = firstOfSection && s.pgNumStart !== undefined ? s.pgNumStart : prev ? prev.number + 1 : 1;
    const pageW = twip(s.pageW), pageH = twip(s.pageH);
    const mirror = this.ctx.doc.settings.mirrorMargins && number % 2 === 0;
    let marginLeft = twip(mirror ? s.right : s.left), marginRight = twip(mirror ? s.left : s.right);
    const gutter = this.ctx.doc.settings.gutterAtTop ? 0 : twip(s.gutter);
    if (mirror) marginRight += gutter;
    else marginLeft += gutter;
    const page: PageBox = {
      index: this.pages.length, number, section: s, sectionIndex: this.sectionIndex, pageInSection: this.pageInSection, width: pageW, height: pageH,
      items: [], floats: [], excl: [], header: null, footer: null, bodyTop: 0, bodyBottom: pageH, marginLeft, marginRight, notes: [], notesHeight: 0,
      separators: [], bookmarks: [], blank,
    };
    this.pages.push(page);
    this.page = page;
    this.items = page.items;
    this.floats = page.floats;
    this.excl = page.excl;
    this.ctx.pageNumber.value = number;
    this.ctx.pageNumber.fmt = s.pgNumFmt;
    this.ctx.pageNumber.sectionPages = this.ctx.sectionCounts[this.sectionIndex] ?? this.ctx.pageNumber.sectionPages;

    const even = this.ctx.doc.settings.evenAndOddHeaders && number % 2 === 0;
    const type = s.titlePg && firstOfSection ? 'first' : even ? 'even' : 'default';
    page.header = blank ? null : this.headerFor('headers', type);
    page.footer = blank ? null : this.headerFor('footers', type);
    const top = twip(Math.abs(s.top)), bottom = twip(Math.abs(s.bottom));
    let bodyTop = top, bodyBottom = pageH - bottom;
    if (page.header) {
      const hl = this.headerLayout(page.header, page, false);
      if (s.top >= 0) bodyTop = Math.max(top, twip(s.header) + hl.height);
      this.wrapAround(hl, page.marginLeft, twip(s.header));
    }
    if (page.footer) {
      const fl = this.headerLayout(page.footer, page, true);
      const footerTop = pageH - twip(s.footer) - fl.height;
      if (s.bottom >= 0) bodyBottom = Math.min(bodyBottom, footerTop);
      this.wrapAround(fl, page.marginLeft, footerTop);
    }
    page.bodyTop = bodyTop;
    page.bodyBottom = Math.max(bodyBottom, bodyTop + 20);
    this.applyColumns(page);
    this.col = 0;
    this.y = bodyTop;
    this.regionTop = bodyTop;
    this.regionStart = 0;
    this.regionSection = s;
    if (!blank) this.pageInSection++;
    // Objects found further down this page in the previous pass, so text above them wraps too.
    for (const f of this.preFloats.get(page.index)?.values() ?? []) {
      this.placedIds.add(f.obj.id);
      page.floats.push(f);
      if (f.ex) page.excl.push(f.ex);
    }
  }

  /** Floating objects in headers and footers with text wrapping push the body text aside, too. */
  private wrapAround(layout: StoryLayout, ox: number, oy: number): void {
    for (const f of layout.floats) {
      const wrap = f.obj.anchor?.wrap;
      if (!wrap || wrap.type === 'none') continue;
      const e = f.obj.effect;
      this.page.excl.push({
        x0: f.x + ox - e.l - wrap.dist.l, x1: f.x + ox + f.w + e.r + wrap.dist.r, y0: f.y + oy - e.t - wrap.dist.t, y1: f.y + oy + f.h + e.b + wrap.dist.b,
        type: wrap.type, side: wrap.side, distL: wrap.dist.l, distR: wrap.dist.r, id: -f.obj.id,
      });
    }
  }

  /** First item of the current column region (a section's part of a page) and its section. */
  private regionStart = 0;
  private regionSection: Section | null = null;

  /**
   * Spreads the content of the current column region on this page evenly over its columns. Content moves in
   * units of overlapping items (a line, a table row, a shaded paragraph); explicit column breaks stay put.
   */
  private balanceRegion(page: PageBox): void {
    const items = page.items.slice(this.regionStart);
    if (!items.length) return;
    if (items.some((i) => i.t === 'line' && i.line.breakType === 'column')) return;
    // Text wrapped around a floating object keeps its place: moving it would run it into the object.
    const regionBottom = items.reduce((b, item) => Math.max(b, itemBottom(item)), this.y);
    if (page.excl.some((e) => e.y1 > this.regionTop + EPS && e.y0 < regionBottom)) return;
    const colOf = (item: Item) => {
      let best = 0;
      for (let c = 0; c < this.columns.length; c++) if (item.x >= this.columns[c].x - 1) best = c;
      return best;
    };
    // Units in reading order: by column, then from the top; items that overlap vertically stay together.
    const sorted = items.map((item) => ({ item, col: colOf(item), top: item.y, bottom: itemBottom(item) })).sort((a, b) => a.col - b.col || a.top - b.top);
    const units: { items: Item[]; col: number; top: number; bottom: number }[] = [];
    for (const e of sorted) {
      const last = units[units.length - 1];
      const thin = e.bottom - e.top < EPS && last && e.top <= last.bottom + EPS;
      if (last && last.col === e.col && (e.top < last.bottom - EPS || thin)) {
        last.items.push(e.item);
        last.bottom = Math.max(last.bottom, e.bottom);
      } else units.push({ items: [e.item], col: e.col, top: e.top, bottom: e.bottom });
    }
    // Gap above each unit within its column (paragraph spacing), dropped at the top of a column.
    const gaps = units.map((u, i) => (i > 0 && units[i - 1].col === u.col ? Math.max(0, u.top - units[i - 1].bottom) : 0));
    const total = units.reduce((n, u, i) => n + (u.bottom - u.top) + gaps[i], 0);
    const n = this.columns.length;
    let target = total / n;
    // Grow the target until everything fits into the columns.
    for (let attempt = 0; attempt < 40; attempt++) {
      let col = 0, y = 0;
      let ok = true;
      for (let i = 0; i < units.length; i++) {
        const h = units[i].bottom - units[i].top;
        const gap = y > 0 ? gaps[i] : 0;
        if (y > 0 && y + gap + h > target + EPS) {
          col++;
          y = 0;
          if (col >= n) {
            ok = false;
            break;
          }
          y = h;
        } else y += gap + h;
      }
      if (ok) break;
      target += 2;
    }
    if (target > page.bodyBottom - this.regionTop) return;
    // The space after the last paragraph moves along with it.
    const lastUnit = units[units.length - 1];
    const trail = lastUnit.col === this.col ? Math.max(0, this.y - lastUnit.bottom) : 0;
    const afterAtEnd = Math.abs(this.lastAfterY - this.y) < EPS;
    let col = 0, y = 0;
    for (let i = 0; i < units.length; i++) {
      const u = units[i];
      const h = u.bottom - u.top;
      const gap = y > 0 ? gaps[i] : 0;
      if (y > 0 && y + gap + h > target + EPS && col + 1 < n) {
        col++;
        y = 0;
      } else y += gap;
      const dx = this.columns[col].x - this.columns[u.col].x;
      const dy = this.regionTop + y - u.top;
      for (const item of u.items) {
        item.x += dx;
        item.y += dy;
      }
      y += h;
      u.top += dy;
      u.bottom += dy;
    }
    this.col = col;
    this.y = lastUnit.bottom + trail;
    if (afterAtEnd) this.lastAfterY = this.y;
  }

  /** Ends the current column region: lines between its columns reach down to its lowest content. */
  private closeRegion(page: PageBox): number {
    let bottom = this.regionTop;
    for (const item of page.items.slice(this.regionStart)) bottom = Math.max(bottom, itemBottom(item));
    const s = this.regionSection ?? page.section;
    if (s.cols.sep && this.columns.length > 1) {
      for (let i = 1; i < this.columns.length; i++) {
        const x = (this.columns[i - 1].x + this.columns[i - 1].w + this.columns[i].x) / 2;
        page.separators.push([x, this.regionTop, bottom]);
      }
    }
    return bottom;
  }

  private finishPage(page: PageBox): void {
    this.closeRegion(page);
    // Vertical alignment of the section on the page.
    const v = page.section.vAlign;
    if ((v === 'center' || v === 'bottom') && page.items.length) {
      let bottom = page.bodyTop;
      for (const item of page.items) bottom = Math.max(bottom, itemBottom(item));
      // The space after the last paragraph counts as content when Word centers the page.
      if (page === this.page) bottom = Math.max(bottom, this.y);
      const free = page.bodyBottom - bottom;
      const dy = v === 'center' ? free / 2 : free;
      if (dy > 0) for (const item of page.items) item.y += dy;
    }
  }

  breakColumn(): boolean {
    if (this.col + 1 < this.columns.length) {
      this.col++;
      this.y = this.regionTop;
      return true;
    }
    return this.breakPage();
  }

  breakPage(): boolean {
    this.newPage();
    return true;
  }

  protected override firstPage(): boolean {
    return this.page.index === 0 && this.col === 0;
  }

  protected override atPageTop(): boolean {
    return this.atColumnTop() && this.col === 0 && this.regionTop <= this.page.bodyTop + EPS;
  }

  protected override chainBreakable(s: Snapshot): boolean {
    // Moving a chain that starts at the top of its page would only produce an empty page.
    const page = this.pages[s.page];
    return !(s.y <= page.bodyTop + EPS && s.col === 0);
  }

  protected override onLine(line: Line): void {
    for (const name of line.bookmarks) this.page.bookmarks.push([name, this.y]);
  }

  protected override onLateFloat(ex: Exclusion): void {
    // Text above the object on this page was laid out without it.
    if (ex.y0 < this.y - 1 && ex.id > 0) {
      const f = this.page.floats[this.page.floats.length - 1];
      let map = this.lateFloats.get(this.page.index);
      if (!map) {
        map = new Map();
        this.lateFloats.set(this.page.index, map);
      }
      map.set(f.obj.id, { ...f, ex });
    }
  }

  // ---- Footnotes ---------------------------------------------------------------------------------------------

  private noteLayouts = new Map<string, StoryLayout>();
  private separatorHeight = -1;

  private noteLayout(id: string): StoryLayout {
    let layout = this.noteLayouts.get(id);
    if (!layout) {
      const el = this.ctx.doc.footnotes.get(id) ?? null;
      const source = this.ctx.doc.footnoteSource ?? this.ctx.doc.source;
      const blocks = storyBlocks(this.ctx, `footnote:${id}`, el, source, this.ctx.notes.footnotes.get(id) ?? '');
      layout = layoutStory(this.ctx, blocks, this.page.width - this.page.marginLeft - this.page.marginRight);
      this.noteLayouts.set(id, layout);
    }
    return layout;
  }

  separatorLayout(): StoryLayout {
    const sep = this.ctx.doc.footnoteSeparator;
    const source = this.ctx.doc.footnoteSource ?? this.ctx.doc.source;
    const blocks = storyBlocks(this.ctx, 'footnote:separator', sep, source);
    return layoutStory(this.ctx, blocks, this.page.width - this.page.marginLeft - this.page.marginRight);
  }

  private placedNotes = new Set<string>();

  override noteSpace(line: Line, commit: boolean): number {
    const ids = line.notes.filter((n) => n.kind === 'footnote' && !n.own && !this.placedNotes.has(n.id) && this.ctx.doc.footnotes.has(n.id)).map((n) => n.id);
    if (!ids.length) return 0;
    let extra = 0;
    if (this.page.notes.length === 0) {
      if (this.separatorHeight < 0) this.separatorHeight = this.ctx.doc.footnoteSeparator ? this.separatorLayout().height : 12;
      extra += this.separatorHeight;
    }
    for (const id of ids) extra += this.noteLayout(id).height;
    if (commit) {
      for (const id of ids) {
        this.placedNotes.add(id);
        this.page.notes.push({ layout: this.noteLayout(id) });
      }
      this.page.notesHeight += extra;
    }
    return extra;
  }

  // ---- Document ----------------------------------------------------------------------------------------------

  /** Lays out the whole body; `progress` gets the share of blocks done. */
  async run(runs: SectionRun[], progress?: (fraction: number) => Promise<void> | void): Promise<PageBox[]> {
    this.sectionList = runs.map((r) => r.section);
    const total = runs.reduce((n, r) => n + r.blocks.length, 0) || 1;
    let done = 0;
    let lastYield = performance.now();
    try {
      for (let si = 0; si < runs.length; si++) {
        const run = runs[si];
        const prevSection = this.section;
        this.sectionIndex = si;
        this.section = run.section;
        this.pageInSection = 0;
        this.startSection(prevSection);
        const blocks = run.blocks;
        let pending: Promise<void> | void;
        this.placeBlocks(blocks, (i) => {
          done++;
          void i;
        });
        if (performance.now() - lastYield > 40) {
          pending = progress?.(done / total);
          if (pending) await pending;
          lastYield = performance.now();
        }
      }
      this.placeEndnotes();
    } catch (error) {
      if (!(error instanceof PageLimit)) throw error;
    }
    if (this.page) this.finishPage(this.page);
    return this.pages;
  }

  private startSection(prev: Section | undefined): void {
    const s = this.section;
    const continuous = s.type === 'continuous' && this.page && prev && prev.pageW === s.pageW && prev.pageH === s.pageH;
    if (!this.page) {
      this.newPage();
      return;
    }
    if (continuous) {
      // The new section continues on this page below what is there, with its own columns. Columns before a
      // continuous break are balanced first, like Word does.
      if (this.columns.length > 1) this.balanceRegion(this.page);
      const bottom = Math.max(this.y, this.columns.length > 1 ? this.closeRegion(this.page) : this.y);
      this.regionStart = this.page.items.length;
      this.regionSection = s;
      this.page.section = s;
      this.columns = this.sectionColumns(this.page, s);
      this.col = 0;
      this.y = bottom;
      this.regionTop = bottom;
      this.pageInSection = 1;
      return;
    }
    if (s.type === 'nextColumn' && this.col + 1 < this.columns.length) {
      this.breakColumn();
      return;
    }
    const prevNumber = this.page.number;
    const next = s.pgNumStart ?? prevNumber + 1;
    const carry = Math.abs(this.lastAfterY - this.y) < EPS ? this.lastAfter : 0;
    // Odd/even section starts leave an empty page when needed. With different odd and even headers, Word also
    // inserts one when restarted page numbers would put two odd (or two even) pages in a row.
    const wantOdd = s.type === 'oddPage', wantEven = s.type === 'evenPage';
    const blank =
      (wantOdd && next % 2 === 0) || (wantEven && next % 2 === 1) || (this.ctx.doc.settings.evenAndOddHeaders && s.pgNumStart !== undefined && next % 2 === prevNumber % 2);
    if (blank) this.newPage(true);
    this.newPage();
    this.sectionCarry = carry;
  }

  private placeEndnotes(): void {
    const { doc } = this.ctx;
    const ids = [...this.ctx.notes.endnotes.keys()].filter((id) => doc.endnotes.has(id));
    if (!ids.length) return;
    const source = doc.endnoteSource ?? doc.source;
    const sep = storyBlocks(this.ctx, 'endnote:separator', doc.endnoteSeparator, source);
    this.placeBlocks(sep);
    for (const id of ids) this.placeBlocks(storyBlocks(this.ctx, `endnote:${id}`, doc.endnotes.get(id) ?? null, source, this.ctx.notes.endnotes.get(id) ?? ''));
  }
}

class PageLimit extends Error {}

function itemBottom(item: Item): number {
  if (item.t === 'line') return item.y + item.line.h;
  if (item.t === 'fill' || item.t === 'group') return item.y + item.h;
  return item.y + (item.vertical ? item.len : 0);
}

export type HeaderDocs = Map<string, { root: Element; source: import('../pptx/env.js').Source }>;

/** Header and footer XML read ahead of layout (layout itself is synchronous). */
export async function loadHeaders(doc: DocxDocument, runs: SectionRun[]): Promise<HeaderDocs> {
  const headerDocs: HeaderDocs = new Map();
  for (const r of runs) {
    for (const part of [...Object.values(r.section.headers), ...Object.values(r.section.footers)]) {
      if (!part || headerDocs.has(part)) continue;
      const xml = await doc.pkg.xml(part);
      const root = xml?.documentElement;
      if (!root) continue;
      headerDocs.set(part, { root, source: { part, rels: await doc.pkg.rels(part) } });
    }
  }
  return headerDocs;
}

export { kid };
