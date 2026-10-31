// SPDX-License-Identifier: MIT
// Word (.docx) content without layout: paragraphs with their style, outline level and list marker, runs with their
// formatting, tables with merged cells, notes, pictures and equations. For converting a document into another
// markup (LaTeX, Markdown, HTML …), where the page layout does not matter. Uses the same reader as the renderer.
import { blipParts } from './pptx/fill.js';
import { loadDocument, type DocxDocument } from './docx/document.js';
import type { Block, Cell, DrawingObj, Inline, Paragraph, Table } from './docx/model.js';
import { createContext, emptyNotes, parseBlocks } from './docx/parse.js';
import type { Source } from './pptx/env.js';

export interface ContentText {
  kind: 'text';
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  script?: 'super' | 'sub';
  smallCaps?: boolean;
  caps?: boolean;
  /** The run's font is a monospaced one (Courier, Consolas …) */
  monospace?: boolean;
  /** Web or mail address (http, https, mailto only) */
  href?: string;
  /** Bookmark the text links to */
  anchor?: string;
}

export type ContentInline =
  | ContentText
  | { kind: 'break'; type: 'line' | 'page' | 'column' }
  | { kind: 'tab' }
  /** Reference to a footnote or endnote; its content is in DocxContent.footnotes / endnotes */
  | { kind: 'note'; note: 'footnote' | 'endnote'; id: string; mark: string }
  /** A picture; its file comes from DocxContent.picture(id). Size in CSS px. */
  | { kind: 'picture'; id: number; width: number; height: number; description?: string; floating: boolean }
  /** An equation as Office MathML (an m:oMath or m:oMathPara element) */
  | { kind: 'math'; omml: Element; display: boolean }
  | { kind: 'bookmark'; name: string }
  /** Content of a text frame or text box, in the place it is anchored */
  | { kind: 'box'; blocks: ContentBlock[] };

export interface ContentParagraph {
  type: 'paragraph';
  /** Style id and name ("Heading1", "heading 1") */
  style: string | null;
  styleName: string | null;
  /** Outline level from the style or the paragraph: 0 for Heading 1 … 8; undefined for body text */
  outlineLevel?: number;
  align?: 'left' | 'center' | 'right' | 'justify';
  /** List membership: the list's id, the level (0–8), numbered or bullets, and the marker as Word shows it */
  list?: { id: string; level: number; ordered: boolean; marker: string };
  inlines: ContentInline[];
}

export interface ContentCell {
  /** Grid column where the cell starts, and how many columns and rows it covers */
  col: number;
  colSpan: number;
  rowSpan: number;
  blocks: ContentBlock[];
}

export interface ContentTable {
  type: 'table';
  /** Grid column widths in CSS px */
  columns: number[];
  rows: { header: boolean; cells: ContentCell[] }[];
}

export type ContentBlock = ContentParagraph | ContentTable;

export interface DocxContent {
  blocks: ContentBlock[];
  footnotes: Map<string, ContentBlock[]>;
  endnotes: Map<string, ContentBlock[]>;
  /** The file of a picture (PNG for Windows metafiles), or null when it cannot be read */
  picture(id: number): Promise<Blob | null>;
  /** Releases the file and the picture URLs */
  close(): void;
}

const MONOSPACE = /courier|consolas|mono|menlo|lucida console|fixedsys|terminal/i;
const TWIP_PX = 96 / 1440;

/** Reads a Word document's content. */
export async function readDocxContent(file: Blob): Promise<DocxContent> {
  const doc = await loadDocument(file);
  try {
    const notes = emptyNotes();
    const pictures = new Map<number, string[]>();
    const convert = (blocks: Block[]) => contentBlocks(blocks, doc, pictures);
    const blocks = convert(parseBlocks(doc.body, createContext(doc, doc.source, notes, true)));
    const read = (map: Map<string, Element>, source: Source | null) => {
      const out = new Map<string, ContentBlock[]>();
      for (const [id, el] of map) out.set(id, convert(parseBlocks(el, createContext(doc, source ?? doc.source, notes, false))));
      return out;
    };
    const footnotes = read(doc.footnotes, doc.footnoteSource);
    const endnotes = read(doc.endnotes, doc.endnoteSource);
    return {
      blocks,
      footnotes,
      endnotes,
      async picture(id) {
        for (const part of pictures.get(id) ?? []) {
          const media = await doc.media.get(part);
          if (media.url) return (await fetch(media.url)).blob();
        }
        return null;
      },
      close() {
        doc.media.revokeAll();
        doc.pkg.close();
      },
    };
  } catch (error) {
    doc.media.revokeAll();
    doc.pkg.close();
    throw error;
  }
}

function contentBlocks(blocks: Block[], doc: DocxDocument, pictures: Map<number, string[]>): ContentBlock[] {
  const out: ContentBlock[] = [];
  for (const b of blocks) {
    if (b.t === 'p') {
      if (!b.hidden) out.push(paragraph(b, doc, pictures));
    } else out.push(table(b, doc, pictures));
  }
  return out;
}

function paragraph(p: Paragraph, doc: DocxDocument, pictures: Map<number, string[]>): ContentParagraph {
  const style = p.styleId ? doc.styles.byId.get(p.styleId) : undefined;
  const jc = p.props.jc;
  const align = jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : jc === 'both' || jc === 'distribute' ? 'justify' : jc ? 'left' : undefined;
  const result: ContentParagraph = {
    type: 'paragraph',
    style: p.styleId,
    styleName: style?.name ?? null,
    inlines: inlines(p.inlines, doc, pictures),
  };
  if (p.props.outlineLvl !== undefined && p.props.outlineLvl < 9) result.outlineLevel = p.props.outlineLvl;
  if (align) result.align = align;
  if (p.num && p.props.numId) {
    const format = p.num.level.numFmt;
    result.list = { id: p.props.numId, level: p.num.level.ilvl, ordered: format !== 'bullet' && format !== 'none', marker: p.num.text };
  }
  return result;
}

function inlines(list: Inline[], doc: DocxDocument, pictures: Map<number, string[]>): ContentInline[] {
  const out: ContentInline[] = [];
  for (const i of list) {
    if (i.t === 'bookmark') {
      out.push({ kind: 'bookmark', name: i.name });
      continue;
    }
    if (i.run.vanish && !i.run.specVanish) continue;
    switch (i.t) {
      case 'text': {
        const r = i.run;
        const t: ContentText = { kind: 'text', text: i.text };
        if (r.b) t.bold = true;
        if (r.i) t.italic = true;
        if (r.u && r.u !== 'none') t.underline = true;
        if (r.strike || r.dstrike) t.strike = true;
        if (r.vertAlign === 'superscript') t.script = 'super';
        else if (r.vertAlign === 'subscript') t.script = 'sub';
        if (r.smallCaps) t.smallCaps = true;
        if (r.caps) t.caps = true;
        if (!i.family && MONOSPACE.test(r.fonts.ascii ?? '')) t.monospace = true;
        if (i.link?.href) t.href = i.link.href;
        else if (i.link?.anchor) t.anchor = i.link.anchor;
        // Runs with the same formatting are one piece of text
        const last = out[out.length - 1];
        if (last?.kind === 'text' && sameLook(last, t)) last.text += t.text;
        else out.push(t);
        break;
      }
      case 'tab':
        out.push({ kind: 'tab' });
        break;
      case 'br':
        out.push({ kind: 'break', type: i.type });
        break;
      case 'field':
        if (i.cached) out.push({ kind: 'text', text: i.cached });
        break;
      case 'note':
        if (!i.own) out.push({ kind: 'note', note: i.kind, id: i.id, mark: i.mark });
        break;
      case 'obj':
      case 'float': {
        const o = object(i.obj, i.t === 'float' || !!i.obj.anchor, doc, pictures);
        if (o) out.push(o);
        break;
      }
    }
  }
  return out;
}

function sameLook(a: ContentText, b: ContentText): boolean {
  const keys: (keyof ContentText)[] = ['bold', 'italic', 'underline', 'strike', 'script', 'smallCaps', 'caps', 'monospace', 'href', 'anchor'];
  return keys.every((k) => a[k] === b[k]);
}

function object(obj: DrawingObj, floating: boolean, doc: DocxDocument, pictures: Map<number, string[]>): ContentInline | null {
  const g = obj.graphic;
  if (g.kind === 'math') return { kind: 'math', omml: g.el, display: g.display };
  if (g.kind === 'frame') return { kind: 'box', blocks: contentBlocks(g.blocks, doc, pictures) };
  if (g.kind === 'table') return { kind: 'box', blocks: [table(g.table, doc, pictures)] };
  if (g.kind === 'checkbox') return { kind: 'text', text: g.checked ? '☒' : '☐' };
  let parts: string[] = [];
  if (g.kind === 'dml') {
    const blip = [...g.el.getElementsByTagName('*')].find((e) => e.localName === 'blip') ?? null;
    parts = blipParts(blip, g.source);
  } else {
    const data = [...g.el.getElementsByTagName('*')].find((e) => e.localName === 'imagedata');
    let id: string | null = null;
    if (data) for (const a of data.attributes) if (a.localName === 'id' && /relationships$/.test(a.namespaceURI ?? '')) id = a.value;
    const rel = id ? g.source.rels.get(id) : undefined;
    if (rel && !rel.external) parts = [rel.target];
  }
  if (!parts.length) return null;
  pictures.set(obj.id, parts);
  const picture: ContentInline = { kind: 'picture', id: obj.id, width: obj.w, height: obj.h, floating };
  if (obj.descr) picture.description = obj.descr;
  return picture;
}

function table(t: Table, doc: DocxDocument, pictures: Map<number, string[]>): ContentTable {
  // Vertically merged cells: a "restart" cell spans the "continue" cells below it in the same grid column
  const grid: (Cell | null)[][] = t.rows.map(() => []);
  t.rows.forEach((row, r) => row.cells.forEach((c) => (grid[r][c.col] = c)));
  const rows: ContentTable['rows'] = t.rows.map((row, r) => ({
    header: !!row.props.tblHeader,
    cells: row.cells.filter((c) => !c.props.vMerge || c.props.vMerge === 'restart').map((c) => {
      let rowSpan = 1;
      if (c.props.vMerge === 'restart') while (grid[r + rowSpan]?.[c.col]?.props.vMerge === 'continue') rowSpan++;
      return { col: c.col, colSpan: Math.max(1, c.span), rowSpan, blocks: contentBlocks(c.blocks, doc, pictures) };
    }),
  }));
  return { type: 'table', columns: t.grid.map((w) => w * TWIP_PX), rows };
}
