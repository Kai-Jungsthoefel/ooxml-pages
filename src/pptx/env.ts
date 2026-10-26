// SPDX-License-Identifier: MIT
// State shared while rendering one slide: theme and colors, where relationships resolve, images, limits.
import type { ColorContext, Theme } from './color.js';
import type { Fill } from './fill.js';
import type { Package, Rels } from './package.js';
import { clamp, int } from './xml.js';

/** 914400 EMU per inch, 96 px per inch. */
export const EMU_PER_PX = 9525;
/** Coordinates beyond this (in EMU, about 50 slide widths) are treated as broken or hostile input. */
const MAX_EMU = 600_000_000;

export function px(emu: number | undefined, fallback = 0): number {
  if (emu === undefined || !Number.isFinite(emu)) return fallback;
  return clamp(emu, -MAX_EMU, MAX_EMU) / EMU_PER_PX;
}

export function emuAttr(el: Element | null | undefined, name: string, fallback = 0): number {
  const v = int(el, name);
  return v === undefined ? fallback : clamp(v, -MAX_EMU, MAX_EMU);
}

/** Owner of relationships that r:embed, r:id … in an XML part refer to. */
export interface Source {
  part: string;
  rels: Rels;
}

export interface Media {
  url: string | null;
  /** Format name for parts browsers cannot show (EMF, WMF, TIFF …), "unknown" if unrecognized. */
  unsupported?: string;
}

function sniff(bytes: Uint8Array): { mime: string } | { unsupported: string } {
  const at = (i: number, ...sig: number[]) => sig.every((b, k) => bytes[i + k] === b);
  if (at(0, 0x89, 0x50, 0x4e, 0x47)) return { mime: 'image/png' };
  if (at(0, 0xff, 0xd8, 0xff)) return { mime: 'image/jpeg' };
  if (at(0, 0x47, 0x49, 0x46, 0x38)) return { mime: 'image/gif' };
  if (at(0, 0x42, 0x4d)) return { mime: 'image/bmp' };
  if (at(0, 0x52, 0x49, 0x46, 0x46) && at(8, 0x57, 0x45, 0x42, 0x50)) return { mime: 'image/webp' };
  if (at(0, 0x01, 0x00, 0x00, 0x00) && at(40, 0x20, 0x45, 0x4d, 0x46)) return { unsupported: 'EMF' };
  if (at(0, 0xd7, 0xcd, 0xc6, 0x9a) || at(0, 0x01, 0x00, 0x09, 0x00) || at(0, 0x02, 0x00, 0x09, 0x00)) return { unsupported: 'WMF' };
  if (at(0, 0x49, 0x49, 0x2a, 0x00) || at(0, 0x4d, 0x4d, 0x00, 0x2a)) return { unsupported: 'TIFF' };
  if (at(0, 0x1f, 0x8b)) return { unsupported: 'EMZ' };
  const head = new TextDecoder().decode(bytes.subarray(0, 512)).trimStart().toLowerCase();
  // SVG is only ever shown through <img>, where browsers run no scripts and load nothing external.
  if (head.startsWith('<svg') || (head.startsWith('<?xml') && head.includes('<svg'))) return { mime: 'image/svg+xml' };
  return { unsupported: 'unknown' };
}

/** Source of pictures for the renderers (the Word renderer adds metafile support on top of MediaCache). */
export interface MediaSource {
  get(part: string): Promise<Media>;
  revokeAll(): void;
}

/** Images of a package as blob: URLs, shared between slides and revoked together. */
export class MediaCache implements MediaSource {
  private readonly cache = new Map<string, Promise<Media>>();
  private readonly urls: string[] = [];

  constructor(private readonly pkg: Package) {}

  get(part: string): Promise<Media> {
    const key = part.toLowerCase();
    let entry = this.cache.get(key);
    if (!entry) {
      entry = this.pkg.bytes(part).then((bytes) => {
        if (!bytes) return { url: null, unsupported: 'unknown' };
        const kind = sniff(bytes);
        if ('unsupported' in kind) return { url: null, unsupported: kind.unsupported };
        const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: kind.mime }));
        this.urls.push(url);
        return { url };
      });
      this.cache.set(key, entry);
    }
    return entry;
  }

  revokeAll(): void {
    for (const url of this.urls) URL.revokeObjectURL(url);
    this.urls.length = 0;
    this.cache.clear();
  }
}

export interface DeckStyles {
  /** presentation.xml defaultTextStyle. */
  defaultText: Element | null;
  /** ppt/tableStyles.xml styles by id. */
  tableStyles: Map<string, Element>;
}

export interface Env {
  pkg: Package;
  theme: Theme;
  colors: ColorContext;
  media: MediaSource;
  deck: DeckStyles;
  slideNumber: number;
  /** Image loads and other work the slide waits for before it counts as finished. */
  pending: Promise<unknown>[];
  /** Remaining shapes and text pieces (runs, paragraphs) we are willing to create for this slide. */
  budget: { shapes: number; text: number };
  /** Shared SVG definitions of the slide, e.g. filters that HTML pictures refer to. */
  defs: () => SVGDefsElement;
  /** Fill of the slide background, for shapes with useBgFill. */
  background?: Fill;
}

let idCounter = 0;
/** Document-unique ids for SVG gradients, patterns and clip paths. */
export function uid(prefix: string): string {
  return `pptx-${prefix}-${++idCounter}`;
}

export const SVG_NS = 'http://www.w3.org/2000/svg';

export function svg<K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number> = {}): SVGElementTagNameMap[K] {
  const el = document.createElementNS(SVG_NS, tag);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
}

/** Rounds to 1/100 px; enough for layout and keeps the styles short. */
export const r2 = (n: number) => Math.round(n * 100) / 100;
