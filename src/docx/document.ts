// SPDX-License-Identifier: MIT
// The DOCX package: main document part, styles, numbering, theme, settings, notes, and the relationships the
// content needs. Headers and footers are read when a section refers to them.
import { type Theme, parseColorMap, parseTheme } from '../pptx/color.js';
import type { MediaSource, Source } from '../pptx/env.js';
import { Package } from '../pptx/package.js';
import { DocxMedia } from './media.js';
import { Numbering } from './numbering.js';
import { Styles } from './styles.js';
import { clampNum, kid, kids, wa, wi, wv } from './xml.js';

export interface NoteSettings {
  numFmt: string;
  numStart: number;
  numRestart: string;
  pos: string;
}

export interface Settings {
  /** Default tab stop distance in twips. */
  defaultTabStop: number;
  evenAndOddHeaders: boolean;
  /** Word version the layout follows: 11 (2003), 12 (2007), 14 (2010), 15 (2013 and later). */
  compatMode: number;
  footnotes: NoteSettings;
  endnotes: NoteSettings;
  /** w:clrSchemeMapping: bg1 → light1 … */
  colorMap: Map<string, string>;
  displayBackground: boolean;
  /** Compatibility switches that change layout, e.g. doNotExpandShiftReturn. */
  compat: Set<string>;
  mirrorMargins: boolean;
  gutterAtTop: boolean;
  autoHyphenation: boolean;
  decimalSymbol: string;
}

export interface DocxDocument {
  pkg: Package;
  media: MediaSource;
  main: string;
  source: Source;
  body: Element;
  background: Element | null;
  styles: Styles;
  numbering: Numbering;
  theme: Theme;
  settings: Settings;
  footnotes: Map<string, Element>;
  endnotes: Map<string, Element>;
  /** Separator notes (type separator / continuationSeparator) of the footnotes part. */
  footnoteSeparator: Element | null;
  endnoteSeparator: Element | null;
  footnoteSource: Source | null;
  endnoteSource: Source | null;
  /** Font table part, which lists embedded fonts. */
  fontTable: string | null;
}

/** Word's mapping names (light1, dark1 …) to DrawingML scheme slots. */
const MAP_SLOT: Record<string, string> = {
  light1: 'lt1', dark1: 'dk1', light2: 'lt2', dark2: 'dk2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3',
  accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hyperlink: 'hlink', followedHyperlink: 'folHlink',
};

/** Colors of the default Office theme, which Word uses when a document has no theme part. */
const DEFAULT_SCHEME: [string, string][] = [
  ['dk1', '000000'], ['lt1', 'FFFFFF'], ['dk2', '44546A'], ['lt2', 'E7E6E6'], ['accent1', '4472C4'], ['accent2', 'ED7D31'],
  ['accent3', 'A5A5A5'], ['accent4', 'FFC000'], ['accent5', '5B9BD5'], ['accent6', '70AD47'], ['hlink', '0563C1'], ['folHlink', '954F72'],
];

function defaultTheme(): Document {
  const slots = DEFAULT_SCHEME.map(([slot, rgb]) => `<a:${slot}><a:srgbClr val="${rgb}"/></a:${slot}>`).join('');
  const xml = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office">${slots}</a:clrScheme></a:themeElements></a:theme>`;
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function noteSettings(el: Element | null, fallbackFmt: string, base?: NoteSettings): NoteSettings {
  return {
    numFmt: wv(kid(el, 'numFmt')) ?? base?.numFmt ?? fallbackFmt,
    numStart: clampNum(wi(kid(el, 'numStart')) ?? base?.numStart ?? 1, 0, 32767),
    numRestart: wv(kid(el, 'numRestart')) ?? base?.numRestart ?? 'continuous',
    pos: wv(kid(el, 'pos')) ?? base?.pos ?? 'pageBottom',
  };
}

function parseSettings(doc: Document | null): Settings {
  const root = doc?.documentElement ?? null;
  const compat = kid(root, 'compat');
  let compatMode = 12;
  const flags = new Set<string>();
  for (const c of kids(compat)) {
    if (c.localName === 'compatSetting') {
      if (wa(c, 'name') === 'compatibilityMode') compatMode = Number(wa(c, 'val')) || compatMode;
      else if (wa(c, 'val') === '1' || wa(c, 'val') === 'true') flags.add(wa(c, 'name') ?? '');
    } else if (wv(c) !== '0' && wv(c) !== 'false') {
      flags.add(c.localName);
    }
  }
  const mapEl = kid(root, 'clrSchemeMapping');
  const colorMap = parseColorMap(null);
  if (mapEl) {
    for (const a of mapEl.attributes) {
      const slot = MAP_SLOT[a.value];
      // Word's attribute names: bg1, t1, bg2, t2, accent1 …
      const key = a.localName === 't1' ? 'tx1' : a.localName === 't2' ? 'tx2' : a.localName;
      if (slot) colorMap.set(key, slot);
    }
  }
  const on = (name: string) => {
    const el = kid(root, name);
    return !!el && wv(el) !== '0' && wv(el) !== 'false' && wv(el) !== 'off';
  };
  return {
    defaultTabStop: clampNum(wi(kid(root, 'defaultTabStop')) ?? 720, 36, 31680),
    evenAndOddHeaders: on('evenAndOddHeaders'),
    compatMode,
    footnotes: noteSettings(kid(root, 'footnotePr'), 'decimal'),
    endnotes: noteSettings(kid(root, 'endnotePr'), 'lowerRoman'),
    colorMap,
    displayBackground: on('displayBackgroundShape'),
    compat: flags,
    mirrorMargins: on('mirrorMargins'),
    gutterAtTop: on('gutterAtTop'),
    autoHyphenation: on('autoHyphenation'),
    decimalSymbol: (wv(kid(root, 'decimalSymbol')) ?? '.').slice(0, 1) || '.',
  };
}

async function notes(pkg: Package, rels: Map<string, { type: string; target: string; external: boolean }>, type: string, item: string) {
  const out = new Map<string, Element>();
  let separator: Element | null = null;
  for (const rel of rels.values()) {
    if (rel.type !== type || rel.external) continue;
    const doc = await pkg.xml(rel.target);
    for (const el of kids(doc?.documentElement)) {
      if (el.localName !== item) continue;
      const id = wa(el, 'id');
      const kind = wa(el, 'type');
      if (kind === 'separator') separator = el;
      else if (id !== null && !kind) out.set(id, el);
    }
    return { notes: out, separator, source: { part: rel.target, rels: await pkg.rels(rel.target) } };
  }
  return { notes: out, separator, source: null };
}

export async function loadDocument(file: Blob): Promise<DocxDocument> {
  const head = new Uint8Array(await file.slice(0, 4).arrayBuffer());
  if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) {
    throw new Error('This document is password-protected (encrypted) or in the old Word 97–2003 format and cannot be displayed.');
  }
  let pkg: Package;
  try {
    pkg = await Package.open(file);
  } catch {
    throw new Error('The file is not a valid Word document.');
  }
  try {
    const main = (await pkg.related('', 'officeDocument')) ?? 'word/document.xml';
    const doc = await pkg.xml(main);
    const root = doc?.documentElement;
    const body = kid(root, 'body');
    if (!root || root.localName !== 'document' || !body) throw new Error('The file is not a Word document.');
    const rels = await pkg.rels(main);
    const partOf = (type: string) => {
      for (const rel of rels.values()) if (rel.type === type && !rel.external) return rel.target;
      return null;
    };
    const stylesPart = partOf('styles');
    const styles = new Styles(stylesPart ? await pkg.xml(stylesPart) : null);
    const numberingPart = partOf('numbering');
    const numbering = new Numbering(numberingPart ? await pkg.xml(numberingPart) : null, styles);
    const themePart = partOf('theme');
    let theme = parseTheme(themePart ? await pkg.xml(themePart) : null);
    if (!theme.colors.size) theme = { ...theme, colors: parseTheme(defaultTheme()).colors };
    const settingsPart = partOf('settings');
    const settings = parseSettings(settingsPart ? await pkg.xml(settingsPart) : null);
    const foot = await notes(pkg, rels, 'footnotes', 'footnote');
    const end = await notes(pkg, rels, 'endnotes', 'endnote');
    return {
      pkg,
      media: new DocxMedia(pkg),
      main,
      source: { part: main, rels },
      body,
      background: kid(root, 'background'),
      styles,
      numbering,
      theme,
      settings,
      footnotes: foot.notes,
      endnotes: end.notes,
      footnoteSeparator: foot.separator,
      endnoteSeparator: end.separator,
      footnoteSource: foot.source,
      endnoteSource: end.source,
      fontTable: partOf('fontTable'),
    };
  } catch (error) {
    pkg.close();
    throw error;
  }
}
