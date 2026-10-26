// SPDX-License-Identifier: MIT
// Reads WordprocessingML content into the model: paragraphs with resolved formatting and numbering, runs,
// complex and simple fields (shown with their cached results, page numbers and form fields evaluated), content
// controls, hyperlinks, tracked changes (shown as accepted), drawings, VML, equations and tables with their
// table style formatting.
import type { Source } from '../pptx/env.js';
import type { DocxDocument } from './document.js';
import type { Anchor, AnchorPos, Block, BreakInline, Cell, ColumnDef, DrawingObj, Inline, Link, Paragraph, Row, Section, Table, Wrap } from './model.js';
import { formatNumber } from './numbering.js';
import { type Border, type CellProps, type RowProps, type RunProps, type TableProps, applyCellProps, applyParaProps, applyRowProps, applyTableProps, cloneRun } from './props.js';
import type { StyleParts } from './styles.js';
import { isSymbolFont, symbolChar, symbolText } from './symbols.js';
import { parseVml } from './vml.js';
import { clampNum, emu, kid, kids, wa, wi, wv } from './xml.js';

interface Field {
  instr: string;
  phase: 'code' | 'result';
  ffData: Element | null;
  /** The field's result is replaced by something we compute (page numbers, form fields). */
  replaced: boolean;
  link?: Link;
  /** PAGEREF: the cached result is collected as the fallback text and emitted at the field's end. */
  pageref?: { bookmark: string; format: string | null; props: RunProps };
  cached?: string;
}

export interface NoteNumbers {
  footnotes: Map<string, string>;
  endnotes: Map<string, string>;
  footCount: number;
  endCount: number;
}

export interface ParseContext {
  doc: DocxDocument;
  source: Source;
  fields: Field[];
  tableLayers: StyleParts[];
  notes: NoteNumbers;
  depth: number;
  budget: { blocks: number; inlines: number };
  link?: Link;
  /** Mark of the note being read (for w:footnoteRef). */
  noteMark?: string;
  /** Numbered paragraphs advance the list counters only in the main text. */
  countNumbers: boolean;
  /** Set when the content has fields whose values depend on the finished layout (NUMPAGES, PAGEREF …). */
  totals?: { used: boolean };
}

let nextId = 1;
export const newId = () => nextId++;

const MAX_DEPTH = 16;
const MAX_TEXT = 1_000_000;

export function createContext(doc: DocxDocument, source: Source, notes: NoteNumbers, countNumbers = true): ParseContext {
  return { doc, source, fields: [], tableLayers: [], notes, depth: 0, budget: { blocks: 400_000, inlines: 4_000_000 }, countNumbers };
}

export function emptyNotes(): NoteNumbers {
  return { footnotes: new Map(), endnotes: new Map(), footCount: 0, endCount: 0 };
}

// ---- Blocks ---------------------------------------------------------------------------------------------------

/** Paragraphs and tables of a container (body, cell, header, text box, note). */
export function parseBlocks(container: Element, ctx: ParseContext, out: Block[] = []): Block[] {
  for (const el of kids(container)) {
    if (ctx.budget.blocks-- <= 0) break;
    switch (el.localName) {
      case 'p':
        out.push(parseParagraph(el, ctx));
        break;
      case 'tbl':
        if (ctx.depth < MAX_DEPTH) {
          const table = parseTable(el, ctx);
          if (table) out.push(table);
        }
        break;
      case 'sdt':
        if (ctx.depth < MAX_DEPTH * 2) {
          ctx.depth++;
          parseBlocks(kid(el, 'sdtContent') ?? el, ctx, out);
          ctx.depth--;
        }
        break;
      case 'customXml':
      case 'ins':
      case 'moveTo':
      case 'smartTag':
        if (ctx.depth < MAX_DEPTH * 2) {
          ctx.depth++;
          parseBlocks(el, ctx, out);
          ctx.depth--;
        }
        break;
    }
  }
  return out;
}

// ---- Paragraphs -----------------------------------------------------------------------------------------------

export function parseParagraph(p: Element, ctx: ParseContext): Paragraph {
  const { styles, numbering } = ctx.doc;
  const pPr = kid(p, 'pPr');
  const styleId = styles.paraStyleId(wv(kid(pPr, 'pStyle')));
  const base = styles.paragraph(styleId, ctx.tableLayers);
  let para = base.para;
  const directNum = kid(pPr, 'numPr');
  const numId = (directNum ? wv(kid(directNum, 'numId')) : null) ?? base.para.numId;
  const directLvl = directNum ? wi(kid(directNum, 'ilvl')) : undefined;
  const ilvl = clampNum(directLvl ?? base.para.ilvl ?? numbering.levelForStyle(numId, styleId) ?? 0, 0, 8);
  const level = numId && numId !== '0' ? numbering.level(numId, ilvl) : null;
  if (level?.pPr) {
    const levelLayer: StyleParts = { pPr: level.pPr, rPr: null, tblPr: null, trPr: null, tcPr: null };
    // Numbering set directly on the paragraph overrides the style's indents; numbering that comes with the
    // style is overridden by the style.
    if (directNum && wv(kid(directNum, 'numId')) !== null) applyParaProps(para, level.pPr);
    else para = styles.paragraph(styleId, [...ctx.tableLayers, levelLayer]).para;
  }
  applyParaProps(para, pPr);
  para.numId = numId ?? undefined;
  para.ilvl = ilvl;
  const mark = styles.run(base.run, kid(pPr, 'rPr'));

  const paragraph: Paragraph = { t: 'p', id: newId(), props: para, mark, inlines: [], styleId, markExplicit: !!kid(pPr, 'rPr') };
  parseInlines(p, ctx, base.run, paragraph.inlines);

  if (level) {
    const number = ctx.countNumbers ? numbering.next(numId!, ilvl) : numbering.peek(numId!, ilvl);
    if (number) {
      const run = cloneRun(mark);
      // The number does not take over underlining or highlighting of the paragraph mark's own formatting.
      const numRun = styles.run(run, level.rPr);
      paragraph.num = { text: number.text, run: numRun, level };
    }
  }
  const sectPr = kid(pPr, 'sectPr');
  if (sectPr) paragraph.section = parseSection(sectPr, ctx.source);
  if (mark.vanish && !mark.specVanish && paragraph.inlines.every((i) => i.t === 'bookmark') && !paragraph.section) paragraph.hidden = true;
  return paragraph;
}

// ---- Inlines --------------------------------------------------------------------------------------------------

function visibleNow(ctx: ParseContext): boolean {
  for (const f of ctx.fields) if (f.phase === 'code' || f.replaced) return false;
  return true;
}

function currentLink(ctx: ParseContext): Link | undefined {
  for (let i = ctx.fields.length - 1; i >= 0; i--) if (ctx.fields[i].link) return ctx.fields[i].link;
  return ctx.link;
}

function push(out: Inline[], ctx: ParseContext, inline: Inline): void {
  if (ctx.budget.inlines-- <= 0) return;
  if (inline.t !== 'float' && inline.t !== 'bookmark') {
    const link = currentLink(ctx);
    if (link) inline.link = link;
  }
  out.push(inline);
}

export function parseInlines(container: Element, ctx: ParseContext, base: RunProps, out: Inline[]): void {
  for (const el of kids(container)) {
    switch (el.localName) {
      case 'r':
        parseRun(el, ctx, base, out);
        break;
      case 'hyperlink': {
        const saved = ctx.link;
        ctx.link = hyperlinkTarget(el, ctx) ?? saved;
        parseInlines(el, ctx, base, out);
        ctx.link = saved;
        break;
      }
      case 'fldSimple': {
        const field: Field = { instr: wa(el, 'instr') ?? '', phase: 'code', ffData: null, replaced: false };
        ctx.fields.push(field);
        field.phase = 'result';
        fieldStart(field, ctx, base, out, null);
        parseInlines(el, ctx, base, out);
        ctx.fields.pop();
        break;
      }
      case 'sdt':
        parseInlineSdt(el, ctx, base, out);
        break;
      case 'smartTag':
      case 'customXml':
      case 'ins':
      case 'moveTo':
      case 'dir':
      case 'bdo':
        if (ctx.depth < MAX_DEPTH * 2) {
          ctx.depth++;
          parseInlines(el, ctx, base, out);
          ctx.depth--;
        }
        break;
      case 'bookmarkStart': {
        const name = wa(el, 'name');
        if (name && visibleNow(ctx)) out.push({ t: 'bookmark', name, run: base });
        break;
      }
      case 'oMath':
      case 'oMathPara':
        if (visibleNow(ctx)) {
          push(out, ctx, {
            t: 'obj',
            run: base,
            obj: { id: newId(), graphic: { kind: 'math', el, display: el.localName === 'oMathPara', source: ctx.source }, w: 0, h: 0, effect: { t: 0, b: 0, l: 0, r: 0 } },
          });
        }
        break;
      // Deleted text, comments, proofing marks and permissions take no space.
    }
  }
}

function safeHref(target: string): string | undefined {
  const t = target.trim();
  return /^(https?:|mailto:)/i.test(t) ? t : undefined;
}

function hyperlinkTarget(el: Element, ctx: ParseContext): Link | undefined {
  let id: string | null = null;
  for (const a of el.attributes) if (a.localName === 'id' && a.namespaceURI && /relationships$/.test(a.namespaceURI)) id = a.value;
  const anchor = wa(el, 'anchor') ?? undefined;
  const rel = id ? ctx.source.rels.get(id) : undefined;
  const href = rel && rel.external ? safeHref(rel.target) : undefined;
  if (!href && !anchor) return undefined;
  return { href, anchor };
}

/** Content controls: shown with their content; check boxes with the symbol of their current state. */
function parseInlineSdt(el: Element, ctx: ParseContext, base: RunProps, out: Inline[]): void {
  const sdtPr = kid(el, 'sdtPr');
  const content = kid(el, 'sdtContent');
  if (!content || ctx.depth >= MAX_DEPTH * 2) return;
  ctx.depth++;
  const checkbox = kid(sdtPr, 'checkbox');
  if (!checkbox) {
    parseInlines(content, ctx, base, out);
    ctx.depth--;
    return;
  }
  const inner: Inline[] = [];
  parseInlines(content, ctx, base, inner);
  ctx.depth--;
  const checkedEl = kid(checkbox, 'checked');
  const checked = !!checkedEl && ['1', 'true', 'on'].includes(wv(checkedEl) ?? '1');
  const state = kid(checkbox, checked ? 'checkedState' : 'uncheckedState');
  const code = parseInt(wa(state, 'val') ?? (checked ? '2612' : '2610'), 16);
  const font = wa(state, 'font') ?? 'MS Gothic';
  const firstText = inner.find((i) => i.t === 'text');
  const run = firstText ? firstText.run : base;
  const valid = Number.isFinite(code) && code > 0;
  // Symbol fonts keep the state glyph in the private use area; Unicode fonts (MS Gothic) have ☐ ☒ themselves.
  const glyph: Inline = isSymbolFont(font)
    ? { t: 'text', text: symbolChar(font, code) ?? (checked ? '☒' : '☐'), run, family: 'symbol', original: { font, text: String.fromCharCode(valid && code < 0x100 ? 0xf000 + code : clampNum(code, 0x20, 0xffff)) } }
    : { t: 'text', text: valid ? String.fromCodePoint(clampNum(code, 0x20, 0x10ffff)) : checked ? '☒' : '☐', run, family: font };
  // The content holds the same glyph in the font of the state; anything else (rare) is kept around it.
  let replaced = false;
  for (const item of inner) {
    if (item.t === 'text') {
      if (!replaced) {
        push(out, ctx, glyph);
        replaced = true;
      }
    } else out.push(item);
  }
  if (!replaced) push(out, ctx, glyph);
}

/** The first word of a field instruction and its arguments, with quotes removed. */
function fieldTokens(instr: string): string[] {
  const tokens: string[] = [];
  const re = /"([^"]*)"|(\S+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(instr)) && tokens.length < 64) tokens.push(m[1] ?? m[2]);
  return tokens;
}

const PAGE_FORMATS: Record<string, string> = {
  arabic: 'decimal', roman: 'lowerRoman', alphabetic: 'lowerLetter', cardtext: 'cardinalText', ordtext: 'ordinalText',
  ordinal: 'ordinal', hex: 'hex', arabicdash: 'numberInDash',
};

function numberFormat(tokens: string[]): string | null {
  const i = tokens.indexOf('\\*');
  const name = i >= 0 ? tokens[i + 1] : undefined;
  if (!name) return null;
  if (name === 'ROMAN') return 'upperRoman';
  if (name === 'ALPHABETIC') return 'upperLetter';
  return PAGE_FORMATS[name.toLowerCase()] ?? null;
}

/** Called when the result of a field starts (or at its end when it has none). */
function fieldStart(field: Field, ctx: ParseContext, base: RunProps, out: Inline[], run: RunProps | null): void {
  const tokens = fieldTokens(field.instr);
  const name = (tokens[0] ?? '').toUpperCase();
  const props = run ?? base;
  // Only fields that are themselves visible produce output.
  const outer = ctx.fields.slice(0, -1);
  const visible = outer.every((f) => f.phase === 'result' && !f.replaced);
  switch (name) {
    case 'PAGE':
    case 'NUMPAGES':
    case 'SECTIONPAGES':
      field.replaced = true;
      if (name !== 'PAGE' && ctx.totals) ctx.totals.used = true;
      if (visible) push(out, ctx, { t: 'field', field: name === 'PAGE' ? 'page' : name === 'NUMPAGES' ? 'numpages' : 'sectionpages', format: numberFormat(tokens), cached: '', run: props });
      break;
    case 'PAGEREF': {
      // Word updates page references when it prints; the cached text is only the fallback.
      const bookmark = tokens[1];
      if (!bookmark) break;
      field.replaced = true;
      field.cached = '';
      field.pageref = { bookmark, format: numberFormat(tokens), props };
      if (ctx.totals) ctx.totals.used = true;
      break;
    }
    case 'FORMCHECKBOX': {
      field.replaced = true;
      const box = kid(field.ffData, 'checkBox');
      if (!visible || !box) break;
      const checkedEl = kid(box, 'checked') ?? kid(box, 'default');
      const checked = !!checkedEl && ['1', 'true', 'on'].includes(wv(checkedEl) ?? '1');
      // w:size is in half points; sizeAuto (or nothing) follows the font size.
      const size = wi(kid(box, 'size'));
      const px = size !== undefined ? clampNum(size, 2, 3276) * (2 / 3) : ((props.sz ?? 20) * 2) / 3;
      push(out, ctx, { t: 'obj', run: props, obj: { id: newId(), graphic: { kind: 'checkbox', checked }, w: px, h: px, effect: { t: 0, b: 0, l: 0, r: 0 } } });
      break;
    }
    case 'FORMDROPDOWN': {
      field.replaced = true;
      const list = kid(field.ffData, 'ddList');
      if (!visible || !list) break;
      const entries = kids(list).filter((e) => e.localName === 'listEntry').map((e) => wv(e) ?? '');
      const index = wi(kid(list, 'result')) ?? wi(kid(list, 'default')) ?? 0;
      const text = entries[clampNum(index, 0, Math.max(0, entries.length - 1))] ?? '';
      if (text) push(out, ctx, { t: 'text', text, run: props });
      break;
    }
    case 'HYPERLINK': {
      let href: string | undefined, anchor: string | undefined;
      for (let i = 1; i < tokens.length; i++) {
        if (tokens[i] === '\\l') anchor = tokens[++i];
        else if (tokens[i].startsWith('\\')) i += tokens[i] === '\\o' || tokens[i] === '\\t' ? 1 : 0;
        else if (href === undefined) href = safeHref(tokens[i]) ?? '';
      }
      if (href || anchor) field.link = { href: href || undefined, anchor };
      break;
    }
  }
}

function parseRun(r: Element, ctx: ParseContext, base: RunProps, out: Inline[]): void {
  const rPr = kid(r, 'rPr');
  const run = ctx.doc.styles.run(base, rPr);
  const hidden = !!run.vanish && !run.webHidden;
  for (const c of kids(r)) {
    const name = c.localName;
    if (name === 'fldChar') {
      const type = wa(c, 'fldCharType');
      if (type === 'begin') {
        if (ctx.fields.length < 32) ctx.fields.push({ instr: '', phase: 'code', ffData: kid(c, 'ffData'), replaced: false });
      } else if (type === 'separate') {
        const field = ctx.fields[ctx.fields.length - 1];
        if (field && field.phase === 'code') {
          field.phase = 'result';
          fieldStart(field, ctx, base, out, run);
        }
      } else if (type === 'end') {
        const field = ctx.fields[ctx.fields.length - 1];
        if (field) {
          if (field.phase === 'code') {
            field.phase = 'result';
            fieldStart(field, ctx, base, out, run);
          }
          ctx.fields.pop();
          if (field.pageref && visibleNow(ctx)) {
            const { bookmark, format, props } = field.pageref;
            push(out, ctx, { t: 'field', field: 'pageref', arg: bookmark, format, cached: field.cached ?? '', run: props });
          }
        }
      }
      continue;
    }
    if (name === 'instrText') {
      const field = ctx.fields[ctx.fields.length - 1];
      if (field && field.phase === 'code' && field.instr.length < 4096) field.instr += c.textContent ?? '';
      continue;
    }
    if (name === 't') {
      const top = ctx.fields[ctx.fields.length - 1];
      if (top?.pageref && top.phase === 'result') top.cached = (top.cached ?? '') + (c.textContent ?? '');
    }
    if (!visibleNow(ctx) || hidden) continue;
    switch (name) {
      case 't': {
        let text = (c.textContent ?? '').replace(/[\t\n\r]/g, ' ');
        if (!text) break;
        if (text.length > MAX_TEXT) text = text.slice(0, MAX_TEXT);
        pushText(out, ctx, text, run);
        break;
      }
      case 'tab':
        push(out, ctx, { t: 'tab', run });
        break;
      case 'ptab':
        push(out, ctx, { t: 'tab', run, ptab: { alignment: wa(c, 'alignment') ?? 'left', relativeTo: wa(c, 'relativeTo') ?? 'margin', leader: wa(c, 'leader') ?? 'none' } });
        break;
      case 'br': {
        const type = wa(c, 'type');
        const br: BreakInline = { t: 'br', type: type === 'page' ? 'page' : type === 'column' ? 'column' : 'line', run };
        const clear = wa(c, 'clear');
        if (clear && clear !== 'none') br.clear = clear;
        push(out, ctx, br);
        break;
      }
      case 'cr':
        push(out, ctx, { t: 'br', type: 'line', run });
        break;
      case 'noBreakHyphen':
        pushText(out, ctx, '\u2011', run);
        break;
      case 'softHyphen':
        pushText(out, ctx, '\u00ad', run);
        break;
      case 'sym': {
        const font = wa(c, 'font') ?? '';
        const code = parseInt(wa(c, 'char') ?? '', 16);
        if (!Number.isFinite(code)) break;
        // Symbol fonts address their glyphs in the private use area (F020-F0FF).
        const pua = String.fromCharCode(code < 0x100 ? 0xf000 + code : clampNum(code, 0x20, 0xffff));
        const mapped = symbolChar(font, code);
        if (mapped) push(out, ctx, { t: 'text', text: mapped, run, family: 'symbol', original: { font, text: pua } });
        else push(out, ctx, { t: 'text', text: pua, run, family: font });
        break;
      }
      case 'drawing': {
        const inline = parseDrawing(c, ctx, run);
        if (inline) push(out, ctx, inline);
        break;
      }
      case 'pict':
      case 'object': {
        for (const shape of vmlShapes(c)) {
          const inline = parseVml(shape, ctx, run);
          if (inline) push(out, ctx, inline);
        }
        // Newer embedded objects carry a DrawingML picture instead of VML.
        const drawing = kid(c, 'drawing');
        if (drawing) {
          const inline = parseDrawing(drawing, ctx, run);
          if (inline) push(out, ctx, inline);
        }
        break;
      }
      case 'footnoteReference':
      case 'endnoteReference': {
        const kind = name === 'footnoteReference' ? 'footnote' : 'endnote';
        const id = wa(c, 'id') ?? '';
        const custom = ['1', 'true', 'on'].includes(wa(c, 'customMarkFollows') ?? '');
        let mark = '';
        if (!custom) {
          const notes = ctx.notes;
          const settings = kind === 'footnote' ? ctx.doc.settings.footnotes : ctx.doc.settings.endnotes;
          const count = kind === 'footnote' ? ++notes.footCount : ++notes.endCount;
          mark = formatNumber(settings.numStart + count - 1, settings.numFmt);
          (kind === 'footnote' ? notes.footnotes : notes.endnotes).set(id, mark);
        }
        push(out, ctx, { t: 'note', kind, id, mark, own: false, run });
        break;
      }
      case 'footnoteRef':
      case 'endnoteRef':
        push(out, ctx, { t: 'note', kind: name === 'footnoteRef' ? 'footnote' : 'endnote', id: '', mark: ctx.noteMark ?? '', own: true, run });
        break;
      case 'separator':
      case 'continuationSeparator':
        push(out, ctx, { t: 'sep', continuation: name === 'continuationSeparator', run });
        break;
      case 'pgNum':
        push(out, ctx, { t: 'field', field: 'page', format: null, cached: '', run });
        break;
      case 'ruby': {
        const rubyBase = kid(c, 'rubyBase');
        if (rubyBase) parseInlines(rubyBase, ctx, base, out);
        break;
      }
    }
  }
}

/** Codes of a symbol font text in the private use area, where symbol fonts keep their glyphs. */
function privateUse(text: string): string {
  return [...text]
    .map((ch) => {
      const code = ch.codePointAt(0) ?? 0;
      return code >= 0x20 && code < 0x100 ? String.fromCharCode(0xf000 + code) : ch;
    })
    .join('');
}

/**
 * Text in a symbol font: drawn with the font itself where it is installed, otherwise translated to Unicode.
 * Everything else stays as written.
 */
function pushText(out: Inline[], ctx: ParseContext, text: string, run: RunProps): void {
  const font = run.fonts.ascii ?? run.fonts.hAnsi;
  const symbolFont = !!font && isSymbolFont(font) && !run.fonts.asciiTheme;
  // Characters in the private use area are almost always symbol font codes.
  if (font && (symbolFont || /[\uf020-\uf0ff]/.test(text))) {
    const { text: mapped, mapped: any } = symbolText(font, text);
    if (any || symbolFont) {
      push(out, ctx, { t: 'text', text: mapped, run, family: any ? 'symbol' : undefined, original: { font, text: privateUse(text) } });
      return;
    }
  }
  push(out, ctx, { t: 'text', text, run });
}

function vmlShapes(el: Element): Element[] {
  const out: Element[] = [];
  for (const c of kids(el)) {
    if (['shape', 'rect', 'roundrect', 'oval', 'line', 'polyline', 'arc', 'group', 'image'].includes(c.localName)) out.push(c);
  }
  return out;
}

// ---- Drawings -------------------------------------------------------------------------------------------------

const EMU_LIMIT = 50_000_000;

function emuPx(value: number | undefined): number {
  return emu(clampNum(value ?? 0, -EMU_LIMIT, EMU_LIMIT));
}

function anchorPos(el: Element | null, fallback: string): AnchorPos {
  const pos: AnchorPos = { relativeFrom: wa(el, 'relativeFrom') ?? fallback };
  for (const c of kids(el)) {
    const text = (c.textContent ?? '').trim();
    if (c.localName === 'align') pos.align = text;
    else if (c.localName === 'posOffset') pos.offset = emuPx(Number(text) || 0);
    else if (c.localName === 'pctPosHOffset' || c.localName === 'pctPosVOffset') pos.pct = clampNum((Number(text) || 0) / 100000, -10, 10);
  }
  return pos;
}

function wrapOf(anchor: Element): Wrap {
  const dist = {
    t: emuPx(wi(anchor, 'distT')),
    b: emuPx(wi(anchor, 'distB')),
    l: emuPx(wi(anchor, 'distL')),
    r: emuPx(wi(anchor, 'distR')),
  };
  for (const c of kids(anchor)) {
    switch (c.localName) {
      case 'wrapNone':
        return { type: 'none', side: 'bothSides', dist };
      case 'wrapTopAndBottom':
        return { type: 'topAndBottom', side: 'bothSides', dist: { ...dist, t: emuPx(wi(c, 'distT') ?? wi(anchor, 'distT')), b: emuPx(wi(c, 'distB') ?? wi(anchor, 'distB')) } };
      case 'wrapSquare':
      case 'wrapTight':
      case 'wrapThrough': {
        const type = c.localName === 'wrapSquare' ? 'square' : c.localName === 'wrapTight' ? 'tight' : 'through';
        const own = (name: string, fallback: number) => {
          const v = wi(c, name);
          return v === undefined ? fallback : emuPx(v);
        };
        const wrap: Wrap = { type, side: wa(c, 'wrapText') ?? 'bothSides', dist: { t: own('distT', dist.t), b: own('distB', dist.b), l: own('distL', dist.l), r: own('distR', dist.r) } };
        const polygon = kid(c, 'wrapPolygon');
        if (polygon) {
          wrap.polygon = kids(polygon)
            .filter((pt) => pt.localName === 'start' || pt.localName === 'lineTo')
            .slice(0, 2000)
            .map((pt) => [clampNum(wi(pt, 'x') ?? 0, -21600 * 10, 21600 * 10), clampNum(wi(pt, 'y') ?? 0, -21600 * 10, 21600 * 10)] as [number, number]);
        }
        return wrap;
      }
    }
  }
  return { type: 'none', side: 'bothSides', dist };
}

/** Rotation in degrees of the shape, picture or group inside a graphic. */
function rotationOf(graphic: Element): number {
  const data = kid(graphic, 'graphicData');
  for (const el of kids(data)) {
    const spPr = kid(el, 'spPr') ?? kid(el, 'grpSpPr');
    const xfrm = kid(spPr, 'xfrm');
    const rot = Number(xfrm?.getAttribute('rot') ?? 0);
    if (Number.isFinite(rot) && rot) return (((rot / 60000) % 360) + 360) % 360;
  }
  return 0;
}

function relSize(el: Element | null, name: string): [string, number] | undefined {
  if (!el) return undefined;
  const pct = kid(el, name);
  const v = Number((pct?.textContent ?? '').trim());
  if (!pct || !Number.isFinite(v) || v <= 0) return undefined;
  return [wa(el, 'relativeFrom') ?? 'page', clampNum(v / 100000, 0, 10)];
}

export function parseDrawing(drawing: Element, ctx: ParseContext, run: RunProps): Inline | null {
  const wrapper = kids(drawing).find((e) => e.localName === 'inline' || e.localName === 'anchor');
  if (!wrapper) return null;
  const docPr = kid(wrapper, 'docPr');
  if (['1', 'true'].includes(wa(docPr, 'hidden') ?? '')) return null;
  const extent = kid(wrapper, 'extent');
  const w = Math.max(0, emuPx(wi(extent, 'cx'))), h = Math.max(0, emuPx(wi(extent, 'cy')));
  const ee = kid(wrapper, 'effectExtent');
  const effect = { l: emuPx(wi(ee, 'l')), t: emuPx(wi(ee, 't')), r: emuPx(wi(ee, 'r')), b: emuPx(wi(ee, 'b')) };
  const graphic = kid(wrapper, 'graphic');
  if (!graphic) return null;
  // A turned object takes its turned bounding box in the layout (files may store no or stale extents for it).
  const rot = rotationOf(graphic);
  if (rot) {
    const a = (rot * Math.PI) / 180;
    const bw = Math.abs(w * Math.cos(a)) + Math.abs(h * Math.sin(a));
    const bh = Math.abs(w * Math.sin(a)) + Math.abs(h * Math.cos(a));
    const ex = (bw - w) / 2, ey = (bh - h) / 2;
    effect.l = ex >= 0 ? Math.max(effect.l, ex) : ex;
    effect.r = ex >= 0 ? Math.max(effect.r, ex) : ex;
    effect.t = ey >= 0 ? Math.max(effect.t, ey) : ey;
    effect.b = ey >= 0 ? Math.max(effect.b, ey) : ey;
  }
  const obj: DrawingObj = { id: newId(), graphic: { kind: 'dml', el: graphic, source: ctx.source }, w, h, effect, name: wa(docPr, 'name') ?? undefined, descr: wa(docPr, 'descr') ?? undefined };
  if (wrapper.localName === 'inline') return { t: 'obj', obj, run };

  const anchor: Anchor = {
    posH: anchorPos(kid(wrapper, 'positionH'), 'column'),
    posV: anchorPos(kid(wrapper, 'positionV'), 'paragraph'),
    behindDoc: ['1', 'true'].includes(wa(wrapper, 'behindDoc') ?? ''),
    z: clampNum(wi(wrapper, 'relativeHeight') ?? 0, 0, 1e10),
    wrap: wrapOf(wrapper),
    layoutInCell: !['0', 'false'].includes(wa(wrapper, 'layoutInCell') ?? '1'),
    allowOverlap: !['0', 'false'].includes(wa(wrapper, 'allowOverlap') ?? '1'),
  };
  if (['1', 'true'].includes(wa(wrapper, 'simplePos') ?? '')) {
    const sp = kid(wrapper, 'simplePos');
    anchor.posH = { relativeFrom: 'page', offset: emuPx(wi(sp, 'x')) };
    anchor.posV = { relativeFrom: 'page', offset: emuPx(wi(sp, 'y')) };
  }
  anchor.sizeRelH = relSize(kid(wrapper, 'sizeRelH'), 'pctWidth');
  anchor.sizeRelV = relSize(kid(wrapper, 'sizeRelV'), 'pctHeight');
  obj.anchor = anchor;
  return { t: 'float', obj, run };
}

// ---- Tables ---------------------------------------------------------------------------------------------------

/** Rows of a table, looking through content controls and tracked insertions around them. */
function rowElements(tbl: Element, depth = 0): Element[] {
  const out: Element[] = [];
  for (const el of kids(tbl)) {
    if (el.localName === 'tr') {
      // A deleted row (tracked change) is gone in the final document.
      if (!kid(kid(el, 'trPr'), 'del')) out.push(el);
    } else if (depth < 8 && (el.localName === 'sdt' || el.localName === 'customXml' || el.localName === 'ins' || el.localName === 'moveTo')) {
      out.push(...rowElements(kid(el, 'sdtContent') ?? el, depth + 1));
    }
  }
  return out;
}

function cellElements(tr: Element, depth = 0): Element[] {
  const out: Element[] = [];
  for (const el of kids(tr)) {
    if (el.localName === 'tc') out.push(el);
    else if (depth < 8 && (el.localName === 'sdt' || el.localName === 'customXml')) out.push(...cellElements(kid(el, 'sdtContent') ?? el, depth + 1));
  }
  return out;
}

const DEFAULT_LOOK = { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true };

export function parseTable(tbl: Element, ctx: ParseContext): Table | null {
  const { styles } = ctx.doc;
  const tblPr = kid(tbl, 'tblPr');
  let styleId = wv(kid(tblPr, 'tblStyle'));
  if (!styleId || styles.byId.get(styleId)?.type !== 'table') styleId = styles.defaultTable;
  const chain = styles.chain(styleId).filter((s) => s.type === 'table');
  const props: TableProps = {};
  for (const s of chain) applyTableProps(props, s.tblPr);
  applyTableProps(props, tblPr);
  const look = { ...DEFAULT_LOOK, ...(props.look ?? {}) };
  const grid = kids(kid(tbl, 'tblGrid'))
    .filter((e) => e.localName === 'gridCol')
    .slice(0, 63)
    .map((e) => clampNum(wi(e, 'w') ?? 0, 0, 31680 * 4));

  const rowEls = rowElements(tbl).slice(0, 20000);
  // Grid positions first: conditional formatting needs to know the last row and column.
  const layout = rowEls.map((tr) => {
    const trPr = kid(tr, 'trPr');
    const rowProps: RowProps = {};
    for (const s of chain) applyRowProps(rowProps, s.trPr);
    applyRowProps(rowProps, trPr);
    let col = rowProps.gridBefore ?? 0;
    const cells = cellElements(tr).slice(0, 63).map((tc) => {
      const tcPr = kid(tc, 'tcPr');
      const span = clampNum(wi(kid(tcPr, 'gridSpan')) ?? 1, 1, 63);
      const cell = { tc, col, span };
      col += span;
      return cell;
    });
    return { tr, rowProps, cells, end: col };
  });
  const colCount = Math.max(grid.length, ...layout.map((r) => r.end + (r.rowProps.gridAfter ?? 0)), 1);
  const rowCount = layout.length;
  const rowBand = props.rowBandSize ?? 1, colBand = props.colBandSize ?? 1;

  const conditionsOf = (r: number, c: number, span: number): string[] => {
    const types = ['wholeTable'];
    const lastCol = c + span >= colCount;
    const isFirstRow = look.firstRow && r === 0;
    const isLastRow = look.lastRow && r === rowCount - 1;
    const isFirstCol = look.firstColumn && c === 0;
    const isLastCol = look.lastColumn && lastCol;
    if (!look.noVBand && !isFirstCol && !isLastCol) {
      const index = c - (look.firstColumn ? 1 : 0);
      types.push(Math.floor(index / colBand) % 2 === 0 ? 'band1Vert' : 'band2Vert');
    }
    if (!look.noHBand && !isFirstRow && !isLastRow) {
      const index = r - (look.firstRow ? 1 : 0);
      types.push(Math.floor(index / rowBand) % 2 === 0 ? 'band1Horz' : 'band2Horz');
    }
    if (isFirstCol) types.push('firstCol');
    if (isLastCol) types.push('lastCol');
    if (isFirstRow) types.push('firstRow');
    if (isLastRow) types.push('lastRow');
    if (isFirstRow && isFirstCol) types.push('nwCell');
    if (isFirstRow && isLastCol) types.push('neCell');
    if (isLastRow && isFirstCol) types.push('swCell');
    if (isLastRow && isLastCol) types.push('seCell');
    return types;
  };
  const layersFor = (types: string[]): StyleParts[] => {
    const out: StyleParts[] = [];
    for (const type of types) {
      for (const s of chain) {
        if (type === 'wholeTable') out.push(s);
        const part = s.conditional.get(type);
        if (part) out.push(part);
      }
    }
    return out;
  };
  /** Table borders inside conditional formatting apply to the edges of that region (a row, a column …). */
  const regionEdges = (types: string[], r: number, c: number, span: number): Cell['edges'] => {
    const edges: Cell['edges'] = {};
    const firstCol = c === 0, lastCol = c + span >= colCount, firstRow = r === 0, lastRow = r === rowCount - 1;
    for (const type of types) {
      for (const s of chain) {
        const tblPr = s.conditional.get(type)?.tblPr;
        if (!tblPr) continue;
        const tp: TableProps = {};
        applyTableProps(tp, tblPr);
        const b = tp.borders;
        if (!b) continue;
        const set = (edge: keyof Cell['edges'], value: Border | undefined) => {
          if (value) edges[edge] = value;
        };
        if (/Row$|Horz$/.test(type)) {
          set('top', b.top);
          set('bottom', b.bottom);
          set('left', firstCol ? b.left : b.insideV);
          set('right', lastCol ? b.right : b.insideV);
        } else if (/Col$|Vert$/.test(type)) {
          set('left', b.left);
          set('right', b.right);
          set('top', firstRow ? b.top : b.insideH);
          set('bottom', lastRow ? b.bottom : b.insideH);
        } else if (/Cell$/.test(type)) {
          set('top', b.top);
          set('bottom', b.bottom);
          set('left', b.left);
          set('right', b.right);
        } else {
          set('top', firstRow ? b.top : b.insideH);
          set('bottom', lastRow ? b.bottom : b.insideH);
          set('left', firstCol ? b.left : b.insideV);
          set('right', lastCol ? b.right : b.insideV);
        }
      }
    }
    return edges;
  };
  const styleLayers = layersFor(['wholeTable']);

  const saved = ctx.tableLayers;
  ctx.depth++;
  const rows: Row[] = layout.map((row, r) => {
    const exEl = kid(row.tr, 'tblPrEx');
    let ex: TableProps | undefined;
    if (exEl) {
      ex = {};
      applyTableProps(ex, exEl);
    }
    const cells: Cell[] = row.cells.map(({ tc, col, span }) => {
      const types = conditionsOf(r, col, span);
      const cond = layersFor(types);
      const cellProps: CellProps = {};
      for (const layer of cond) applyCellProps(cellProps, layer.tcPr);
      applyCellProps(cellProps, kid(tc, 'tcPr'));
      ctx.tableLayers = cond;
      const blocks = parseBlocks(tc, ctx);
      return { props: cellProps, blocks, cond, edges: regionEdges(types, r, col, span), col, span };
    });
    return { props: row.rowProps, cells, ex };
  });
  ctx.depth--;
  ctx.tableLayers = saved;
  return { t: 'tbl', id: newId(), props, grid, rows, styleLayers };
}

// ---- Sections -------------------------------------------------------------------------------------------------

const MAX_PAGE = 31680 * 2;

export function parseSection(sectPr: Element, source: Source): Section {
  const pgSz = kid(sectPr, 'pgSz');
  const pgMar = kid(sectPr, 'pgMar');
  const cols = kid(sectPr, 'cols');
  const relTarget = (el: Element) => {
    for (const a of el.attributes) {
      if (a.localName === 'id' && a.namespaceURI && /relationships$/.test(a.namespaceURI)) {
        const rel = source.rels.get(a.value);
        return rel && !rel.external ? rel.target : undefined;
      }
    }
    return undefined;
  };
  const headers: Section['headers'] = {};
  const footers: Section['footers'] = {};
  for (const el of kids(sectPr)) {
    const isHeader = el.localName === 'headerReference';
    if (!isHeader && el.localName !== 'footerReference') continue;
    const type = (wa(el, 'type') ?? 'default') as 'default' | 'first' | 'even';
    const target = relTarget(el);
    if (target) (isHeader ? headers : footers)[type] = target;
  }
  const n = (el: Element | null, name: string, fallback: number, min = -MAX_PAGE, max = MAX_PAGE) => clampNum(wi(el, name) ?? fallback, min, max);
  let pageW = n(pgSz, 'w', 12240, 144, MAX_PAGE);
  let pageH = n(pgSz, 'h', 15840, 144, MAX_PAGE);
  if (wa(pgSz, 'orient') === 'landscape' && pageW < pageH) [pageW, pageH] = [pageH, pageW];
  const colList: ColumnDef[] = kids(cols)
    .filter((c) => c.localName === 'col')
    .slice(0, 45)
    .map((c) => ({ w: n(c, 'w', 0, 0), space: n(c, 'space', 0, 0) }));
  const num = clampNum(wi(cols, 'num') ?? (colList.length || 1), 1, 45);
  const equalAttr = wa(cols, 'equalWidth');
  const pgNum = kid(sectPr, 'pgNumType');
  const grid = kid(sectPr, 'docGrid');
  return {
    pageW,
    pageH,
    top: n(pgMar, 'top', 1440),
    bottom: n(pgMar, 'bottom', 1440),
    left: n(pgMar, 'left', 1440, 0),
    right: n(pgMar, 'right', 1440, 0),
    header: n(pgMar, 'header', 720, 0),
    footer: n(pgMar, 'footer', 720, 0),
    gutter: n(pgMar, 'gutter', 0, 0),
    cols: {
      num,
      space: n(cols, 'space', 720, 0),
      equal: equalAttr === null ? colList.length < 2 : ['1', 'true', 'on'].includes(equalAttr),
      sep: ['1', 'true', 'on'].includes(wa(cols, 'sep') ?? ''),
      list: colList,
    },
    type: wv(kid(sectPr, 'type')) ?? 'nextPage',
    titlePg: !!kid(sectPr, 'titlePg') && !['0', 'false'].includes(wv(kid(sectPr, 'titlePg')) ?? '1'),
    headers,
    footers,
    pgNumStart: wi(pgNum, 'start') !== undefined ? clampNum(wi(pgNum, 'start')!, 0, 1_000_000) : undefined,
    pgNumFmt: wa(pgNum, 'fmt') ?? 'decimal',
    vAlign: wv(kid(sectPr, 'vAlign')) ?? 'top',
    borders: kid(sectPr, 'pgBorders') ?? undefined,
    docGrid: grid ? { type: wa(grid, 'type') ?? 'default', linePitch: clampNum(wi(grid, 'linePitch') ?? 0, 0, 31680), charSpace: wi(grid, 'charSpace') ?? 0 } : undefined,
    bidi: !!kid(sectPr, 'bidi'),
  };
}

export function defaultSection(): Section {
  return {
    pageW: 12240, pageH: 15840, top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720, gutter: 0,
    cols: { num: 1, space: 720, equal: true, sep: false, list: [] }, type: 'nextPage', titlePg: false, headers: {}, footers: {},
    pgNumFmt: 'decimal', vAlign: 'top', bidi: false,
  };
}
