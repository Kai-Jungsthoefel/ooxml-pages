// SPDX-License-Identifier: MIT
// WordprocessingML helpers. Unlike DrawingML, Word puts its attributes into its own namespace (w:val), and the
// namespace URI differs between transitional and strict documents, so attributes are matched by local name.
// Lengths are converted to CSS pixels at 96 dpi.

/** Attribute by local name, whatever prefix or namespace it has. */
export function wa(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) if (attrs[i].localName === name) return attrs[i].value;
  return null;
}

/** w:val of an element. */
export function wv(el: Element | null | undefined): string | null {
  return wa(el, 'val');
}

/** Integer attribute; undefined when missing or not a number. Word also writes "123.0" or "12pt" in some files. */
export function wi(el: Element | null | undefined, name = 'val'): number | undefined {
  const value = wa(el, name);
  if (value === null || value === '') return undefined;
  let n = Number(value);
  if (!Number.isFinite(n)) {
    // Universal measures (ST_UniversalMeasure) like "12pt", "1in", "2.5cm": converted to twips.
    const m = /^(-?\d+(?:\.\d+)?)(mm|cm|in|pt|pc|pi)$/.exec(value.trim());
    if (!m) return undefined;
    const unit: Record<string, number> = { mm: 56.6929, cm: 566.929, in: 1440, pt: 20, pc: 240, pi: 240 };
    n = Number(m[1]) * unit[m[2]];
  }
  return n;
}

/**
 * On/off property (<w:b/>, <w:b w:val="0"/>): undefined when the element is missing, true when present without a
 * value or with a true value.
 */
export function onOff(el: Element | null | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = wv(el);
  return v === null || v === '1' || v === 'true' || v === 'on';
}

const SUPPORTED_EXTENSIONS = new Set([
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
  'http://schemas.microsoft.com/office/word/2010/wordml',
  'http://schemas.microsoft.com/office/word/2012/wordml',
  'http://schemas.microsoft.com/office/drawing/2010/main',
  'http://schemas.openxmlformats.org/officeDocument/2006/math',
]);

/** The branch of an mc:AlternateContent block we can render: a Choice whose requirements we know, else the Fallback. */
function alternate(block: Element): Element | null {
  let fallback: Element | null = null;
  for (let node = block.firstElementChild; node; node = node.nextElementSibling) {
    if (node.localName === 'Choice') {
      const requires = (wa(node, 'Requires') ?? '').split(/\s+/).filter(Boolean);
      if (requires.length && requires.every((prefix) => SUPPORTED_EXTENSIONS.has(node.lookupNamespaceURI(prefix) ?? ''))) return node;
    } else if (node.localName === 'Fallback' && !fallback) {
      fallback = node;
    }
  }
  return fallback;
}

/** Element children with markup-compatibility blocks resolved (DrawingML shapes preferred over VML fallbacks). */
export function kids(el: Element | null | undefined): Element[] {
  const out: Element[] = [];
  if (!el) return out;
  for (let node = el.firstElementChild; node; node = node.nextElementSibling) {
    if (node.localName === 'AlternateContent') {
      const branch = alternate(node);
      if (branch) out.push(...kids(branch));
    } else out.push(node);
  }
  return out;
}

export function kid(el: Element | null | undefined, name: string): Element | null {
  if (!el) return null;
  for (let node = el.firstElementChild; node; node = node.nextElementSibling) {
    if (node.localName === name) return node;
    if (node.localName === 'AlternateContent') {
      const branch = alternate(node);
      const found = branch ? kid(branch, name) : null;
      if (found) return found;
    }
  }
  return null;
}

export function kidsNamed(el: Element | null | undefined, name: string): Element[] {
  return kids(el).filter((e) => e.localName === name);
}

export const clampNum = (v: number, min: number, max: number) => (v < min ? min : v > max ? max : v);

/** Twentieths of a point to px. */
export const twip = (v: number) => v / 15;
/** Half points to px. */
export const halfPt = (v: number) => (v * 2) / 3;
/** Points to px. */
export const pt = (v: number) => (v * 4) / 3;
/** English Metric Units to px. */
export const emu = (v: number) => v / 9525;
/** Eighths of a point (border widths) to px. */
export const eighthPt = (v: number) => v / 6;

/** Rounds to 1/100 px for styles. */
export const r2 = (n: number) => Math.round(n * 100) / 100;

/** Hex color "RRGGBB" from a Word color attribute; null for auto or invalid values. */
export function hexColor(value: string | null | undefined): string | null {
  if (!value || !/^[0-9a-f]{6}$/i.test(value)) return null;
  return `#${value.toLowerCase()}`;
}
