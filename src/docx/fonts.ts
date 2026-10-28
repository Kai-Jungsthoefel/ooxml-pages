// SPDX-License-Identifier: MIT
// Fonts embedded in a Word document (word/fonts/*.odttf). Word obfuscates them with a key from the font table
// (ECMA-376 Part 1, 17.8.1). Like Word, an embedded font is used only when the font is not installed. The
// faces are added to the page's font set while the document is open and removed again afterwards.
import type { Package } from '../pptx/package.js';
import { fontInstalled, resetFontCaches } from './measure.js';
import { kids, wa } from './xml.js';

const MAX_FACES = 32;
const MAX_FONT_BYTES = 16 * 1024 * 1024;
const MAX_TOTAL_BYTES = 64 * 1024 * 1024;

const VARIANTS: Record<string, { weight: string; style: string }> = {
  embedRegular: { weight: '400', style: 'normal' },
  embedBold: { weight: '700', style: 'normal' },
  embedItalic: { weight: '400', style: 'italic' },
  embedBoldItalic: { weight: '700', style: 'italic' },
};

/** Same cleaning as the font family lists (pptx/fonts.ts), so the face matches the name text asks for. */
const clean = (name: string) => name.replace(/["'\\;{}<>\p{Cc}]/gu, '').trim().slice(0, 64);

/** Undoes the obfuscation: the first 32 bytes are XORed with the GUID key, its bytes in reverse order. */
export function deobfuscate(bytes: Uint8Array, fontKey: string | null): Uint8Array<ArrayBuffer> {
  const out = new Uint8Array(bytes);
  const hex = (fontKey ?? '').replace(/[^0-9a-f]/gi, '');
  if (hex.length !== 32 || out.length < 32) return out;
  for (let i = 0; i < 32; i++) out[i] ^= parseInt(hex.substr(30 - 2 * (i % 16), 2), 16);
  return out;
}

export interface EmbeddedFonts {
  faces: FontFace[];
  release(): void;
}

/** A font file embedded in the document, as PDF output gets it (src/pdf/fonts.ts). */
export interface EmbeddedFontFile {
  name: string;
  bold: boolean;
  italic: boolean;
  bytes: Uint8Array<ArrayBuffer>;
}

let embeddedSink: ((file: EmbeddedFontFile) => void) | null = null;

/** While set, every embedded font file is handed over, also those of fonts that are installed. */
export function setEmbeddedFontSink(sink: ((file: EmbeddedFontFile) => void) | null): void {
  embeddedSink = sink;
}

export async function loadEmbeddedFonts(pkg: Package, fontTablePart: string | null): Promise<EmbeddedFonts> {
  const faces: FontFace[] = [];
  const release = () => {
    if (!faces.length) return;
    for (const face of faces) document.fonts.delete(face);
    faces.length = 0;
    resetFontCaches();
  };
  if (!fontTablePart || typeof FontFace === 'undefined' || !document.fonts) return { faces, release };
  const table = await pkg.xml(fontTablePart);
  const rels = await pkg.rels(fontTablePart);
  let total = 0;
  for (const font of kids(table?.documentElement)) {
    if (font.localName !== 'font') continue;
    const name = clean(wa(font, 'name') ?? '');
    const embeds = kids(font).filter((e) => VARIANTS[e.localName]);
    if (!name || !embeds.length) continue;
    const installed = fontInstalled(name);
    if (installed && !embeddedSink) continue;
    for (const embed of embeds) {
      if (faces.length >= MAX_FACES || total >= MAX_TOTAL_BYTES) break;
      const id = wa(embed, 'id');
      const rel = id ? rels.get(id) : undefined;
      if (!rel || rel.external) continue;
      const bytes = await pkg.bytes(rel.target, MAX_FONT_BYTES);
      if (!bytes) continue;
      total += bytes.length;
      const variant = VARIANTS[embed.localName];
      const clear = deobfuscate(bytes, wa(embed, 'fontKey'));
      embeddedSink?.({ name, bold: variant.weight === '700', italic: variant.style === 'italic', bytes: clear });
      if (installed) continue;
      try {
        const face = new FontFace(name, clear, variant);
        await face.load();
        document.fonts.add(face);
        faces.push(face);
      } catch {
        // A damaged or unsupported font: text falls back to a substitute, as when it is not embedded.
      }
    }
  }
  if (faces.length) resetFontCaches();
  return { faces, release };
}
