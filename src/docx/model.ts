// SPDX-License-Identifier: MIT
// The document content as blocks (paragraphs, tables) and inlines (text, tabs, breaks, objects, fields), with
// all formatting resolved. Layout and rendering only read this model.
import type { Source } from '../pptx/env.js';
import type { Level } from './numbering.js';
import type { Border, CellProps, ParaProps, RowProps, RunProps, TableProps } from './props.js';
import type { StyleParts } from './styles.js';

export interface Link {
  href?: string;
  anchor?: string;
}

interface InlineBase {
  run: RunProps;
  link?: Link;
}

export interface TextInline extends InlineBase {
  t: 'text';
  text: string;
  /**
   * Font that replaces the run's fonts: a Word font name, or 'symbol' for symbol font characters that were
   * translated to Unicode.
   */
  family?: string;
  /** Symbol font text as written (private use area codes), drawn with the font itself where it is installed. */
  original?: { font: string; text: string };
}

export interface TabInline extends InlineBase {
  t: 'tab';
  /** Absolute position tab (w:ptab). */
  ptab?: { alignment: string; relativeTo: string; leader: string };
}

export interface BreakInline extends InlineBase {
  t: 'br';
  type: 'line' | 'page' | 'column';
  clear?: string;
}

export interface ObjectInline extends InlineBase {
  t: 'obj';
  obj: DrawingObj;
}

export interface FloatInline {
  t: 'float';
  obj: DrawingObj;
  run: RunProps;
}

export interface FieldInline extends InlineBase {
  t: 'field';
  field: 'page' | 'numpages' | 'sectionpages' | 'pageref';
  /** Bookmark of a PAGEREF field. */
  arg?: string;
  /** Number format from a \* switch, or null for the section's page number format. */
  format: string | null;
  cached: string;
}

export interface NoteInline extends InlineBase {
  t: 'note';
  kind: 'footnote' | 'endnote';
  id: string;
  /** Text of the reference mark. */
  mark: string;
  /** The w:footnoteRef inside the note itself (shows the number, does not reference anything). */
  own: boolean;
}

export interface BookmarkInline {
  t: 'bookmark';
  name: string;
  run: RunProps;
}

export interface SeparatorInline extends InlineBase {
  t: 'sep';
  continuation: boolean;
}

export type Inline = TextInline | TabInline | BreakInline | ObjectInline | FloatInline | FieldInline | NoteInline | BookmarkInline | SeparatorInline;

export interface Distances {
  t: number;
  b: number;
  l: number;
  r: number;
}

export interface AnchorPos {
  relativeFrom: string;
  align?: string;
  /** px */
  offset?: number;
  /** Fraction of the reference area (wp14:pctPosHOffset …). */
  pct?: number;
}

export interface Wrap {
  type: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  side: string;
  /** Wrap polygon in 21600ths of the object's size. */
  polygon?: [number, number][];
  dist: Distances;
}

export interface Anchor {
  posH: AnchorPos;
  posV: AnchorPos;
  behindDoc: boolean;
  z: number;
  wrap: Wrap;
  layoutInCell: boolean;
  allowOverlap: boolean;
  /** Size relative to the page or margins (wp14:sizeRelH/V): [relativeFrom, fraction]. */
  sizeRelH?: [string, number];
  sizeRelV?: [string, number];
}

export type Graphic =
  | { kind: 'dml'; el: Element; source: Source }
  | { kind: 'vml'; el: Element; source: Source }
  | { kind: 'checkbox'; checked: boolean }
  | { kind: 'math'; el: Element; display: boolean; source: Source }
  | { kind: 'table'; table: Table }
  | { kind: 'frame'; blocks: Block[] };

export interface DrawingObj {
  id: number;
  graphic: Graphic;
  /** Size of the object's box in px. */
  w: number;
  h: number;
  /** Extra space the object takes in the layout (shadows, glow …), px. */
  effect: Distances;
  anchor?: Anchor;
  name?: string;
  descr?: string;
  /** Depth below the baseline for inline objects that are not set on it (equations). */
  descent?: number;
  /** Font size of an equation in px. */
  mathSize?: number;
  /** Horizontal rule (VML o:hr): share of the text width it spans, set when the line is laid out. */
  hr?: number;
}

export interface Paragraph {
  t: 'p';
  id: number;
  props: ParaProps;
  /** Formatting of the paragraph mark: height of empty paragraphs, numbering text. */
  mark: RunProps;
  /** The mark has formatting of its own (pPr/rPr); only then does it add to the height of the last line. */
  markExplicit?: boolean;
  inlines: Inline[];
  styleId: string | null;
  num?: { text: string; run: RunProps; level: Level };
  /** Section that ends with this paragraph. */
  section?: Section;
  /** The paragraph mark is hidden and there is no visible content: the paragraph takes no space. */
  hidden?: boolean;
}

export interface Cell {
  props: CellProps;
  blocks: Block[];
  /** Table style formatting that applies to this cell, lowest priority first. */
  cond: StyleParts[];
  /** Borders of the table style's regions (first row, banded columns …) at this cell's edges. */
  edges: { top?: Border; bottom?: Border; left?: Border; right?: Border };
  /** Grid column where the cell starts. */
  col: number;
  span: number;
}

export interface Row {
  props: RowProps;
  cells: Cell[];
  /** Table property exceptions for this row (tblPrEx). */
  ex?: TableProps;
}

export interface Table {
  t: 'tbl';
  id: number;
  props: TableProps;
  /** Grid column widths in twips. */
  grid: number[];
  rows: Row[];
  /** Whole-table formatting of the table style (borders, cell margins …), lowest priority first. */
  styleLayers: StyleParts[];
  section?: Section;
}

export type Block = Paragraph | Table;

export interface ColumnDef {
  /** Twips. */
  w: number;
  space: number;
}

export interface Section {
  /** Page size and margins in twips. */
  pageW: number;
  pageH: number;
  top: number;
  bottom: number;
  left: number;
  right: number;
  header: number;
  footer: number;
  gutter: number;
  cols: { num: number; space: number; equal: boolean; sep: boolean; list: ColumnDef[] };
  type: string;
  titlePg: boolean;
  headers: { default?: string; first?: string; even?: string };
  footers: { default?: string; first?: string; even?: string };
  pgNumStart?: number;
  pgNumFmt: string;
  vAlign: string;
  borders?: Element;
  docGrid?: { type: string; linePitch: number; charSpace: number };
  bidi: boolean;
}
