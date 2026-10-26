// SPDX-License-Identifier: MIT
// Run, paragraph, table, row and cell properties: plain objects that the style chain fills in layer by layer
// (document defaults, styles, numbering, direct formatting). Lengths stay in Word's units (twips, half points)
// until layout.
import { clampNum, hexColor, kid, kids, onOff, wa, wi, wv } from './xml.js';

export interface Border {
  style: string;
  /** Width in eighths of a point. */
  size: number;
  /** Distance from the text in points (0–31). */
  space: number;
  color: string | null;
  themeColor?: string;
  shadow?: boolean;
  frame?: boolean;
}

export interface Shading {
  val: string;
  color: string | null;
  fill: string | null;
}

export interface Fonts {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
  asciiTheme?: string;
  hAnsiTheme?: string;
  eastAsiaTheme?: string;
  csTheme?: string;
  hint?: string;
}

/** Namespace of the Word 2010 extensions (w14:…). */
const W14 = /\/office\/word\/2010\/wordml$/;

export interface RunProps {
  fonts: Fonts;
  /** Half points. */
  sz?: number;
  szCs?: number;
  b?: boolean;
  bCs?: boolean;
  i?: boolean;
  iCs?: boolean;
  u?: string;
  uColor?: string | null;
  strike?: boolean;
  dstrike?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  vanish?: boolean;
  specVanish?: boolean;
  webHidden?: boolean;
  /** Hex color, 'auto', or undefined. */
  color?: string;
  themeColor?: string;
  themeTint?: string;
  themeShade?: string;
  highlight?: string;
  shd?: Shading;
  vertAlign?: string;
  /** Raised/lowered position in half points. */
  position?: number;
  /** Character spacing in twips. */
  spacing?: number;
  /** Horizontal scale in percent. */
  w?: number;
  /** Kerning threshold in half points (0: off). */
  kern?: number;
  rtl?: boolean;
  cs?: boolean;
  border?: Border;
  ligatures?: string;
  emboss?: boolean;
  imprint?: boolean;
  outline?: boolean;
  shadow?: boolean;
  /** Word 2010 text effects (w14:textFill, w14:textOutline, w14:shadow), drawn with DrawingML colors. */
  textFill?: Element;
  textOutline?: Element;
  textShadow?: Element;
  noProof?: boolean;
  lang?: string;
  rStyle?: string;
}

export interface TabStop {
  /** Twips from the left edge of the text column. */
  pos: number;
  val: string;
  leader: string;
}

export interface FramePr {
  w?: number;
  h?: number;
  hRule?: string;
  x?: number;
  y?: number;
  xAlign?: string;
  yAlign?: string;
  hAnchor?: string;
  vAnchor?: string;
  wrap?: string;
  hSpace?: number;
  vSpace?: number;
  dropCap?: string;
  lines?: number;
}

export interface ParaProps {
  pStyle?: string;
  jc?: string;
  /** Twips. */
  indLeft?: number;
  indRight?: number;
  /** First line offset in twips: positive indent, negative hanging. */
  indFirst?: number;
  indLeftChars?: number;
  indRightChars?: number;
  indFirstChars?: number;
  spBefore?: number;
  spAfter?: number;
  spBeforeLines?: number;
  spAfterLines?: number;
  spBeforeAuto?: boolean;
  spAfterAuto?: boolean;
  /** 240ths of a line for 'auto', twips otherwise. */
  line?: number;
  lineRule?: string;
  contextualSpacing?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  widowControl?: boolean;
  tabs?: TabStop[];
  borders?: { top?: Border; bottom?: Border; left?: Border; right?: Border; between?: Border; bar?: Border };
  shd?: Shading;
  numId?: string;
  ilvl?: number;
  outlineLvl?: number;
  framePr?: FramePr;
  bidi?: boolean;
  textAlignment?: string;
  snapToGrid?: boolean;
  suppressAutoHyphens?: boolean;
  mirrorIndents?: boolean;
}

export interface CellMargins {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
}

/** A width with its type: dxa (twips), pct (fiftieths of a percent, or "50%"), auto, nil. */
export interface Width {
  w: number;
  type: string;
}

export interface TableBorders {
  top?: Border;
  bottom?: Border;
  left?: Border;
  right?: Border;
  insideH?: Border;
  insideV?: Border;
  tl2br?: Border;
  tr2bl?: Border;
}

export interface TableProps {
  tblStyle?: string;
  width?: Width;
  jc?: string;
  ind?: Width;
  borders?: TableBorders;
  shd?: Shading;
  cellMar?: CellMargins;
  layout?: string;
  cellSpacing?: Width;
  look?: Record<string, boolean>;
  rowBandSize?: number;
  colBandSize?: number;
  /** Floating table position (tblpPr). */
  tblpPr?: Element;
  overlap?: string;
  bidiVisual?: boolean;
}

export interface RowProps {
  height?: number;
  hRule?: string;
  cantSplit?: boolean;
  tblHeader?: boolean;
  gridBefore?: number;
  gridAfter?: number;
  wBefore?: Width;
  wAfter?: Width;
  jc?: string;
  cellSpacing?: Width;
  hidden?: boolean;
}

export interface CellProps {
  width?: Width;
  gridSpan?: number;
  vMerge?: string;
  hMerge?: string;
  borders?: TableBorders;
  shd?: Shading;
  mar?: CellMargins;
  vAlign?: string;
  textDirection?: string;
  noWrap?: boolean;
  hideMark?: boolean;
}

// ---- Small parsers -----------------------------------------------------------------------------------------

export function parseBorder(el: Element | null): Border | undefined {
  if (!el) return undefined;
  const style = wv(el) ?? 'nil';
  return {
    style,
    size: clampNum(wi(el, 'sz') ?? 4, 0, 96 * 8),
    space: clampNum(wi(el, 'space') ?? 0, 0, 31),
    color: hexColor(wa(el, 'color')),
    themeColor: wa(el, 'themeColor') ?? undefined,
    shadow: attrOnOff(el, 'shadow') !== null,
    frame: attrOnOff(el, 'frame') !== null ? true : undefined,
  };
}

function attrOnOff(el: Element, name: string): Element | null {
  const v = wa(el, name);
  return v === '1' || v === 'true' || v === 'on' ? el : null;
}

const LINE_BORDERS = new Set([
  'nil', 'none', 'single', 'thick', 'double', 'dotted', 'dashed', 'dotDash', 'dotDotDash', 'triple', 'thinThickSmallGap',
  'thickThinSmallGap', 'thinThickThinSmallGap', 'thinThickMediumGap', 'thickThinMediumGap', 'thinThickThinMediumGap',
  'thinThickLargeGap', 'thickThinLargeGap', 'thinThickThinLargeGap', 'wave', 'doubleWave', 'dashSmallGap', 'dashDotStroked',
  'threeDEmboss', 'threeDEngrave', 'outset', 'inset',
]);

/** Art borders (apples, mapleMuffins … around pages) repeat a picture built into Word; their size is in points. */
export function isArtBorder(b: Border): boolean {
  return !LINE_BORDERS.has(b.style);
}

/** Whether a border draws anything. */
export function visible(b: Border | undefined): b is Border {
  return !!b && b.style !== 'nil' && b.style !== 'none' && b.size > 0;
}

export function parseShading(el: Element | null): Shading | undefined {
  if (!el) return undefined;
  return { val: wv(el) ?? 'clear', color: hexColor(wa(el, 'color')), fill: hexColor(wa(el, 'fill')) };
}

const PERCENTS: Record<string, number> = {
  pct5: 5, pct10: 10, pct12: 12.5, pct15: 15, pct20: 20, pct25: 25, pct30: 30, pct35: 35, pct37: 37.5, pct40: 40,
  pct45: 45, pct50: 50, pct55: 55, pct60: 60, pct62: 62.5, pct65: 65, pct70: 70, pct75: 75, pct80: 80, pct85: 85,
  pct87: 87.5, pct90: 90, pct95: 95,
};

function mixHex(a: string, b: string, t: number): string {
  const pa = parseInt(a.slice(1), 16), pb = parseInt(b.slice(1), 16);
  const ch = (shift: number) => Math.round(((pa >> shift) & 255) * (1 - t) + ((pb >> shift) & 255) * t);
  return `#${((ch(16) << 16) | (ch(8) << 8) | ch(0)).toString(16).padStart(6, '0')}`;
}

/** Background color of a shading: the fill mixed with the pattern color by the pattern's percentage. */
export function shadingColor(shd: Shading | undefined): string | null {
  if (!shd || shd.val === 'nil') return null;
  const fill = shd.fill;
  const color = shd.color ?? '#000000';
  if (shd.val === 'clear' || shd.val === 'none') return fill;
  if (shd.val === 'solid') return shd.color ?? (fill ? '#000000' : '#000000');
  const pct = PERCENTS[shd.val];
  if (pct !== undefined) return mixHex(fill ?? '#ffffff', color, pct / 100);
  // Stripes and crosshatches: a mid tone is closer than either color.
  return mixHex(fill ?? '#ffffff', color, 0.25);
}

export function parseWidth(el: Element | null): Width | undefined {
  if (!el) return undefined;
  const type = wa(el, 'type') ?? 'dxa';
  const raw = wa(el, 'w');
  if (raw === null) return { w: 0, type };
  if (raw.endsWith('%')) return { w: Number(raw.slice(0, -1)) * 50 || 0, type: 'pct' };
  const w = wi(el, 'w') ?? 0;
  return { w: clampNum(w, -31680 * 10, 31680 * 10), type };
}

function parseMargins(el: Element | null, into: CellMargins = {}): CellMargins | undefined {
  if (!el) return undefined;
  for (const c of kids(el)) {
    const w = parseWidth(c);
    if (!w || w.type === 'pct') continue;
    const v = clampNum(w.w, 0, 31680);
    switch (c.localName) {
      case 'top':
        into.top = v;
        break;
      case 'bottom':
        into.bottom = v;
        break;
      case 'left':
      case 'start':
        into.left = v;
        break;
      case 'right':
      case 'end':
        into.right = v;
        break;
    }
  }
  return into;
}

function parseTableBorders(el: Element | null, into: TableBorders = {}): TableBorders | undefined {
  if (!el) return undefined;
  for (const c of kids(el)) {
    const b = parseBorder(c);
    switch (c.localName) {
      case 'top':
        into.top = b;
        break;
      case 'bottom':
        into.bottom = b;
        break;
      case 'left':
      case 'start':
        into.left = b;
        break;
      case 'right':
      case 'end':
        into.right = b;
        break;
      case 'insideH':
        into.insideH = b;
        break;
      case 'insideV':
        into.insideV = b;
        break;
      case 'tl2br':
        into.tl2br = b;
        break;
      case 'tr2bl':
        into.tr2bl = b;
        break;
    }
  }
  return into;
}

// ---- Run properties -------------------------------------------------------------------------------------------

export function emptyRun(): RunProps {
  return { fonts: {} };
}

export function cloneRun(r: RunProps): RunProps {
  return { ...r, fonts: { ...r.fonts } };
}

/** Applies an rPr element on top of run properties. */
export function applyRunProps(into: RunProps, rPr: Element | null): void {
  if (!rPr) return;
  for (const c of kids(rPr)) {
    switch (c.localName) {
      case 'rStyle':
        into.rStyle = wv(c) ?? undefined;
        break;
      case 'rFonts': {
        const f = into.fonts;
        for (const name of ['ascii', 'hAnsi', 'eastAsia', 'cs', 'asciiTheme', 'hAnsiTheme', 'eastAsiaTheme', 'csTheme', 'hint'] as const) {
          const v = wa(c, name);
          if (v !== null) {
            f[name] = v;
            // An explicit font replaces an inherited theme font of the same slot and vice versa.
            if (!name.endsWith('Theme') && name !== 'hint') delete f[`${name}Theme` as keyof Fonts];
          }
        }
        break;
      }
      case 'sz': {
        const v = wi(c);
        if (v !== undefined) into.sz = clampNum(v, 1, 3276);
        break;
      }
      case 'szCs': {
        const v = wi(c);
        if (v !== undefined) into.szCs = clampNum(v, 1, 3276);
        break;
      }
      case 'b':
      case 'bCs':
      case 'i':
      case 'iCs':
      case 'strike':
      case 'dstrike':
      case 'caps':
      case 'smallCaps':
      case 'vanish':
      case 'specVanish':
      case 'webHidden':
      case 'rtl':
      case 'cs':
      case 'emboss':
      case 'imprint':
      case 'outline':
      case 'shadow':
      case 'noProof':
        // w14:shadow shares its name with w:shadow.
        if (c.localName === 'shadow' && W14.test(c.namespaceURI ?? '')) into.textShadow = c;
        else into[c.localName] = onOff(c);
        break;
      case 'textFill':
        into.textFill = c;
        break;
      case 'textOutline':
        into.textOutline = c;
        break;
      case 'u':
        into.u = wv(c) ?? 'single';
        into.uColor = hexColor(wa(c, 'color'));
        break;
      case 'color': {
        const v = wv(c);
        into.color = v === 'auto' ? 'auto' : (hexColor(v) ?? into.color);
        into.themeColor = wa(c, 'themeColor') ?? undefined;
        into.themeTint = wa(c, 'themeTint') ?? undefined;
        into.themeShade = wa(c, 'themeShade') ?? undefined;
        break;
      }
      case 'highlight':
        into.highlight = wv(c) ?? undefined;
        break;
      case 'shd':
        into.shd = parseShading(c);
        break;
      case 'vertAlign':
        into.vertAlign = wv(c) ?? undefined;
        break;
      case 'position': {
        const v = wi(c);
        if (v !== undefined) into.position = clampNum(v, -3168, 3168);
        break;
      }
      case 'spacing': {
        const v = wi(c);
        if (v !== undefined) into.spacing = clampNum(v, -31680, 31680);
        break;
      }
      case 'w': {
        const v = wi(c);
        if (v !== undefined) into.w = clampNum(v, 1, 600);
        break;
      }
      case 'kern': {
        const v = wi(c);
        if (v !== undefined) into.kern = clampNum(v, 0, 3276);
        break;
      }
      case 'bdr':
        into.border = parseBorder(c);
        break;
      case 'ligatures':
        into.ligatures = wv(c) ?? undefined;
        break;
      case 'lang':
        into.lang = wv(c) ?? into.lang;
        break;
    }
  }
}

// ---- Paragraph properties -------------------------------------------------------------------------------------

function parseFrame(el: Element): FramePr {
  const n = (name: string) => wi(el, name);
  return {
    w: n('w'), h: n('h'), hRule: wa(el, 'hRule') ?? undefined, x: n('x'), y: n('y'), xAlign: wa(el, 'xAlign') ?? undefined,
    yAlign: wa(el, 'yAlign') ?? undefined, hAnchor: wa(el, 'hAnchor') ?? undefined, vAnchor: wa(el, 'vAnchor') ?? undefined,
    wrap: wa(el, 'wrap') ?? undefined, hSpace: n('hSpace'), vSpace: n('vSpace'), dropCap: wa(el, 'dropCap') ?? undefined,
    lines: n('lines'),
  };
}

const MAX_TWIPS = 31680 * 4;

/** Applies a pPr element on top of paragraph properties (tabs merge, 'clear' removes inherited stops). */
export function applyParaProps(into: ParaProps, pPr: Element | null): void {
  if (!pPr) return;
  for (const c of kids(pPr)) {
    switch (c.localName) {
      case 'pStyle':
        into.pStyle = wv(c) ?? undefined;
        break;
      case 'jc':
        into.jc = wv(c) ?? undefined;
        break;
      case 'ind': {
        const num = (name: string) => {
          const v = wi(c, name);
          return v === undefined ? undefined : clampNum(v, -MAX_TWIPS, MAX_TWIPS);
        };
        const left = num('left') ?? num('start');
        const right = num('right') ?? num('end');
        if (left !== undefined) into.indLeft = left;
        if (right !== undefined) into.indRight = right;
        const hanging = num('hanging'), first = num('firstLine');
        if (hanging !== undefined) into.indFirst = -hanging;
        else if (first !== undefined) into.indFirst = first;
        const lc = num('leftChars') ?? num('startChars'), rc = num('rightChars') ?? num('endChars');
        if (lc !== undefined) into.indLeftChars = lc;
        if (rc !== undefined) into.indRightChars = rc;
        const hc = num('hangingChars'), fc = num('firstLineChars');
        if (hc !== undefined) into.indFirstChars = -hc;
        else if (fc !== undefined) into.indFirstChars = fc;
        break;
      }
      case 'spacing': {
        const before = wi(c, 'before'), after = wi(c, 'after');
        if (before !== undefined) into.spBefore = clampNum(before, 0, MAX_TWIPS);
        if (after !== undefined) into.spAfter = clampNum(after, 0, MAX_TWIPS);
        const bl = wi(c, 'beforeLines'), al = wi(c, 'afterLines');
        if (bl !== undefined) into.spBeforeLines = clampNum(bl, 0, 10000);
        if (al !== undefined) into.spAfterLines = clampNum(al, 0, 10000);
        const ba = wa(c, 'beforeAutospacing'), aa = wa(c, 'afterAutospacing');
        if (ba !== null) into.spBeforeAuto = ba === '1' || ba === 'true' || ba === 'on';
        if (aa !== null) into.spAfterAuto = aa === '1' || aa === 'true' || aa === 'on';
        const line = wi(c, 'line');
        if (line !== undefined) {
          into.line = clampNum(line, -MAX_TWIPS, MAX_TWIPS);
          into.lineRule = wa(c, 'lineRule') ?? 'auto';
        } else if (wa(c, 'lineRule')) into.lineRule = wa(c, 'lineRule')!;
        break;
      }
      case 'contextualSpacing':
      case 'keepNext':
      case 'keepLines':
      case 'pageBreakBefore':
      case 'widowControl':
      case 'bidi':
      case 'snapToGrid':
      case 'suppressAutoHyphens':
      case 'mirrorIndents':
        into[c.localName] = onOff(c);
        break;
      case 'tabs': {
        const tabs = [...(into.tabs ?? [])];
        for (const t of kids(c)) {
          if (t.localName !== 'tab') continue;
          const pos = clampNum(wi(t, 'pos') ?? 0, -MAX_TWIPS, MAX_TWIPS);
          const val = wv(t) ?? 'left';
          const i = tabs.findIndex((x) => Math.abs(x.pos - pos) < 1);
          if (i >= 0) tabs.splice(i, 1);
          if (val !== 'clear') tabs.push({ pos, val, leader: wa(t, 'leader') ?? 'none' });
        }
        into.tabs = tabs.sort((a, b) => a.pos - b.pos).slice(0, 64);
        break;
      }
      case 'pBdr': {
        const borders = { ...(into.borders ?? {}) };
        for (const b of kids(c)) {
          const name = b.localName === 'start' ? 'left' : b.localName === 'end' ? 'right' : b.localName;
          if (name === 'top' || name === 'bottom' || name === 'left' || name === 'right' || name === 'between' || name === 'bar') borders[name] = parseBorder(b);
        }
        into.borders = borders;
        break;
      }
      case 'shd':
        into.shd = parseShading(c);
        break;
      case 'numPr': {
        const id = wv(kid(c, 'numId'));
        const lvl = wi(kid(c, 'ilvl'));
        if (id !== null) into.numId = id;
        if (lvl !== undefined) into.ilvl = clampNum(lvl, 0, 8);
        break;
      }
      case 'outlineLvl': {
        const v = wi(c);
        if (v !== undefined) into.outlineLvl = v;
        break;
      }
      case 'framePr':
        into.framePr = parseFrame(c);
        break;
      case 'textAlignment':
        into.textAlignment = wv(c) ?? undefined;
        break;
    }
  }
}

// ---- Tables ---------------------------------------------------------------------------------------------------

export function applyTableProps(into: TableProps, tblPr: Element | null): void {
  if (!tblPr) return;
  for (const c of kids(tblPr)) {
    switch (c.localName) {
      case 'tblStyle':
        into.tblStyle = wv(c) ?? undefined;
        break;
      case 'tblW':
        into.width = parseWidth(c);
        break;
      case 'jc':
        into.jc = wv(c) ?? undefined;
        break;
      case 'tblInd':
        into.ind = parseWidth(c);
        break;
      case 'tblBorders':
        into.borders = parseTableBorders(c, { ...(into.borders ?? {}) });
        break;
      case 'shd':
        into.shd = parseShading(c);
        break;
      case 'tblCellMar':
        into.cellMar = parseMargins(c, { ...(into.cellMar ?? {}) });
        break;
      case 'tblLayout':
        into.layout = wa(c, 'type') ?? undefined;
        break;
      case 'tblCellSpacing':
        into.cellSpacing = parseWidth(c);
        break;
      case 'tblLook': {
        const look: Record<string, boolean> = {};
        const hex = parseInt(wv(c) ?? '', 16);
        if (Number.isFinite(hex)) {
          look.firstRow = !!(hex & 0x20);
          look.lastRow = !!(hex & 0x40);
          look.firstColumn = !!(hex & 0x80);
          look.lastColumn = !!(hex & 0x100);
          look.noHBand = !!(hex & 0x200);
          look.noVBand = !!(hex & 0x400);
        }
        for (const name of ['firstRow', 'lastRow', 'firstColumn', 'lastColumn', 'noHBand', 'noVBand']) {
          const v = wa(c, name);
          if (v !== null) look[name] = v === '1' || v === 'true' || v === 'on';
        }
        into.look = look;
        break;
      }
      case 'tblStyleRowBandSize':
        into.rowBandSize = clampNum(wi(c) ?? 1, 1, 100);
        break;
      case 'tblStyleColBandSize':
        into.colBandSize = clampNum(wi(c) ?? 1, 1, 100);
        break;
      case 'tblpPr':
        into.tblpPr = c;
        break;
      case 'tblOverlap':
        into.overlap = wv(c) ?? undefined;
        break;
      case 'bidiVisual':
        into.bidiVisual = onOff(c);
        break;
    }
  }
}

export function applyRowProps(into: RowProps, trPr: Element | null): void {
  if (!trPr) return;
  for (const c of kids(trPr)) {
    switch (c.localName) {
      case 'trHeight':
        into.height = clampNum(wi(c) ?? 0, 0, 31680);
        into.hRule = wa(c, 'hRule') ?? 'atLeast';
        break;
      case 'cantSplit':
        into.cantSplit = onOff(c);
        break;
      case 'tblHeader':
        into.tblHeader = onOff(c);
        break;
      case 'gridBefore':
        into.gridBefore = clampNum(wi(c) ?? 0, 0, 63);
        break;
      case 'gridAfter':
        into.gridAfter = clampNum(wi(c) ?? 0, 0, 63);
        break;
      case 'wBefore':
        into.wBefore = parseWidth(c);
        break;
      case 'wAfter':
        into.wAfter = parseWidth(c);
        break;
      case 'jc':
        into.jc = wv(c) ?? undefined;
        break;
      case 'tblCellSpacing':
        into.cellSpacing = parseWidth(c);
        break;
      case 'hidden':
        into.hidden = onOff(c);
        break;
    }
  }
}

export function applyCellProps(into: CellProps, tcPr: Element | null): void {
  if (!tcPr) return;
  for (const c of kids(tcPr)) {
    switch (c.localName) {
      case 'tcW':
        into.width = parseWidth(c);
        break;
      case 'gridSpan':
        into.gridSpan = clampNum(wi(c) ?? 1, 1, 63);
        break;
      case 'vMerge':
        into.vMerge = wv(c) ?? 'continue';
        break;
      case 'hMerge':
        into.hMerge = wv(c) ?? 'continue';
        break;
      case 'tcBorders':
        into.borders = parseTableBorders(c, { ...(into.borders ?? {}) });
        break;
      case 'shd':
        into.shd = parseShading(c);
        break;
      case 'tcMar':
        into.mar = parseMargins(c, { ...(into.mar ?? {}) });
        break;
      case 'vAlign':
        into.vAlign = wv(c) ?? undefined;
        break;
      case 'textDirection':
        into.textDirection = wv(c) ?? undefined;
        break;
      case 'noWrap':
        into.noWrap = onOff(c);
        break;
      case 'hideMark':
        into.hideMark = onOff(c);
        break;
    }
  }
}
