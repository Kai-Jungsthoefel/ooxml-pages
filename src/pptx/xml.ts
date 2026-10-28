// SPDX-License-Identifier: MIT
// Namespace-agnostic helpers for OOXML. Elements are matched by local name only, so transitional and strict
// documents (which use different namespace URIs for the same vocabulary) both work.

export function parseXml(text: string): Document | null {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
}

/**
 * Element children, with markup-compatibility blocks resolved: the Choice where we know what it requires, else
 * the Fallback the producer itself prepared for readers like us (e.g. a picture of a chart type we lack).
 */
export function elements(el: Element | null | undefined): Element[] {
  const out: Element[] = [];
  if (!el) return out;
  for (let node = el.firstElementChild; node; node = node.nextElementSibling) {
    if (node.localName === 'AlternateContent') out.push(...elements(alternate(node)));
    else out.push(node);
  }
  return out;
}

/** Office 2010 drawing extensions: PowerPoint wraps text with equations (a14:m) in a Choice requiring them. */
const SUPPORTED_EXTENSIONS = new Set(['http://schemas.microsoft.com/office/drawing/2010/main']);

function alternate(block: Element): Element | null {
  let choice: Element | null = null;
  let fallback: Element | null = null;
  for (let node = block.firstElementChild; node; node = node.nextElementSibling) {
    if (node.localName === 'Choice') {
      const requires = (node.getAttribute('Requires') ?? '').split(/\s+/).filter(Boolean);
      if (requires.length && requires.every((prefix) => SUPPORTED_EXTENSIONS.has(node.lookupNamespaceURI(prefix) ?? ''))) return node;
      choice ??= node;
    } else if (node.localName === 'Fallback' && node.firstElementChild && !fallback) {
      fallback = node;
    }
  }
  return fallback ?? choice;
}

export function child(el: Element | null | undefined, name: string): Element | null {
  if (!el) return null;
  for (let node = el.firstElementChild; node; node = node.nextElementSibling) {
    if (node.localName === name) return node;
  }
  return null;
}

export function children(el: Element | null | undefined, name: string): Element[] {
  const out: Element[] = [];
  if (!el) return out;
  for (let node = el.firstElementChild; node; node = node.nextElementSibling) {
    if (node.localName === name) out.push(node);
  }
  return out;
}

/** Follows a chain of child names, e.g. path(sp, 'nvSpPr', 'nvPr', 'ph'). */
export function path(el: Element | null | undefined, ...names: string[]): Element | null {
  let current: Element | null = el ?? null;
  for (const name of names) current = child(current, name);
  return current;
}

/** First descendant with the given local name (depth-first). */
export function descendant(el: Element | null | undefined, name: string): Element | null {
  if (!el) return null;
  const all = el.getElementsByTagName('*');
  for (let i = 0; i < all.length; i++) if (all[i].localName === name) return all[i];
  return null;
}

export function descendants(el: Element | null | undefined, name: string): Element[] {
  if (!el) return [];
  return [...el.getElementsByTagName('*')].filter((e) => e.localName === name);
}

export function attr(el: Element | null | undefined, name: string): string | null {
  return el ? el.getAttribute(name) : null;
}

/** Attribute in the relationships namespace (r:embed, r:id …), whatever prefix or URI variant is used. */
export function relAttr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  for (const a of el.attributes) {
    if (a.localName === name && a.namespaceURI && /relationships$/.test(a.namespaceURI)) return a.value;
  }
  return null;
}

/** Integer attribute; undefined when missing or not a finite number. */
export function int(el: Element | null | undefined, name: string): number | undefined {
  const value = attr(el, name);
  if (value === null || value === '') return undefined;
  const n = Number(value);
  return Number.isFinite(n) ? n : undefined;
}

export function bool(el: Element | null | undefined, name: string): boolean | undefined {
  const value = attr(el, name);
  if (value === null) return undefined;
  return value === '1' || value === 'true' || value === 'on';
}

/** Attribute "val" of a child element, a very common OOXML pattern (<c:idx val="3"/>). */
export function val(el: Element | null | undefined, name: string): string | null {
  return attr(child(el, name), 'val');
}

export function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value;
}
