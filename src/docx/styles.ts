// SPDX-License-Identifier: MIT
// styles.xml: document defaults and paragraph, character and table styles with their basedOn chains.
import { type ParaProps, type RunProps, applyParaProps, applyRunProps, cloneRun, emptyRun } from './props.js';
import { kid, kids, wa, wv } from './xml.js';

/** Conditional formatting of a table style (tblStylePr), by type: firstRow, band1Horz … */
export interface StyleParts {
  pPr: Element | null;
  rPr: Element | null;
  tblPr: Element | null;
  trPr: Element | null;
  tcPr: Element | null;
}

export interface StyleDef extends StyleParts {
  id: string;
  type: string;
  name: string;
  basedOn: string | null;
  isDefault: boolean;
  conditional: Map<string, StyleParts>;
}

const MAX_CHAIN = 32;

function parts(el: Element): StyleParts {
  return { pPr: kid(el, 'pPr'), rPr: kid(el, 'rPr'), tblPr: kid(el, 'tblPr'), trPr: kid(el, 'trPr'), tcPr: kid(el, 'tcPr') };
}

export class Styles {
  readonly byId = new Map<string, StyleDef>();
  readonly defaultRPr: Element | null;
  readonly defaultPPr: Element | null;
  readonly defaultPara: string | null;
  readonly defaultChar: string | null;
  readonly defaultTable: string | null;
  private readonly paraCache = new Map<string, { para: ParaProps; run: RunProps }>();
  private readonly charCache = new Map<string, Element[]>();

  constructor(doc: Document | null) {
    const root = doc?.documentElement ?? null;
    const defaults = kid(root, 'docDefaults');
    this.defaultRPr = kid(kid(defaults, 'rPrDefault'), 'rPr');
    this.defaultPPr = kid(kid(defaults, 'pPrDefault'), 'pPr');
    let para: string | null = null, char: string | null = null, table: string | null = null;
    for (const el of kids(root)) {
      if (el.localName !== 'style') continue;
      const id = wa(el, 'styleId');
      if (!id || this.byId.has(id)) continue;
      const type = wa(el, 'type') ?? 'paragraph';
      const flag = wa(el, 'default');
      const isDefault = flag === '1' || flag === 'true' || flag === 'on';
      const conditional = new Map<string, StyleParts>();
      for (const c of kids(el)) if (c.localName === 'tblStylePr') conditional.set(wa(c, 'type') ?? '', parts(c));
      this.byId.set(id, { id, type, name: wv(kid(el, 'name')) ?? id, basedOn: wv(kid(el, 'basedOn')), isDefault, conditional, ...parts(el) });
      if (isDefault) {
        if (type === 'paragraph' && !para) para = id;
        else if (type === 'character' && !char) char = id;
        else if (type === 'table' && !table) table = id;
      }
    }
    this.defaultPara = para;
    this.defaultChar = char;
    this.defaultTable = table;
  }

  /** The style and its ancestors, base first. Cycles and very long chains are cut off. */
  chain(id: string | null | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let current = id ? this.byId.get(id) : undefined;
    while (current && !seen.has(current.id) && out.length < MAX_CHAIN) {
      seen.add(current.id);
      out.unshift(current);
      current = current.basedOn ? this.byId.get(current.basedOn) : undefined;
    }
    return out;
  }

  /** Paragraph style id to use: the given one if it exists as a paragraph style, else the default. */
  paraStyleId(id: string | null | undefined): string | null {
    const style = id ? this.byId.get(id) : undefined;
    return style && style.type === 'paragraph' ? style.id : this.defaultPara;
  }

  /**
   * Paragraph and run properties a paragraph style gives (document defaults included). `before` adds layers
   * between the document defaults and the style (table style formatting).
   */
  paragraph(id: string | null, before: StyleParts[] = []): { para: ParaProps; run: RunProps } {
    const key = before.length ? null : (id ?? '');
    if (key !== null) {
      const cached = this.paraCache.get(key);
      if (cached) return { para: { ...cached.para, tabs: cached.para.tabs, borders: cached.para.borders }, run: cloneRun(cached.run) };
    }
    const para: ParaProps = {};
    const run = emptyRun();
    applyParaProps(para, this.defaultPPr);
    applyRunProps(run, this.defaultRPr);
    for (const layer of before) {
      applyParaProps(para, layer.pPr);
      applyRunProps(run, layer.rPr);
    }
    for (const style of this.chain(id)) {
      applyParaProps(para, style.pPr);
      applyRunProps(run, style.rPr);
    }
    delete run.rStyle;
    if (key !== null) this.paraCache.set(key, { para: { ...para }, run: cloneRun(run) });
    return { para, run };
  }

  /** rPr elements of a character style chain, base first. */
  character(id: string | null | undefined): Element[] {
    const key = id ?? '';
    let list = this.charCache.get(key);
    if (!list) {
      list = this.chain(id)
        .filter((s) => s.type === 'character' || s.type === 'paragraph')
        .map((s) => s.rPr)
        .filter((e): e is Element => !!e);
      this.charCache.set(key, list);
    }
    return list;
  }

  /** Run properties with a character style and direct formatting on top of a base (the paragraph's). */
  run(base: RunProps, rPr: Element | null): RunProps {
    const run = cloneRun(base);
    const styleId = wv(kid(rPr, 'rStyle'));
    if (styleId) for (const el of this.character(styleId)) applyRunProps(run, el);
    applyRunProps(run, rPr);
    delete run.rStyle;
    return run;
  }

  /** Whether a style id names the given style, e.g. to find built-in styles by name (case-insensitive). */
  idByName(name: string): string | null {
    const lower = name.toLowerCase();
    for (const s of this.byId.values()) if (s.name.toLowerCase() === lower) return s.id;
    return null;
  }
}
