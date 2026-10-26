// SPDX-License-Identifier: MIT
// Pictures of a Word document: what browsers show directly comes from the shared MediaCache; Windows
// metafiles (EMF, WMF, EMZ, WMZ) are drawn by our metafile player and handed out as PNG.
import { type Media, MediaCache, type MediaSource } from '../pptx/env.js';
import type { Package } from '../pptx/package.js';
import { metafileToPng } from './metafile.js';

const METAFILES = new Set(['EMF', 'WMF', 'EMZ']);
// Metafiles beyond this size are not worth the drawing time; they keep their placeholder.
const MAX_METAFILE_BYTES = 32 * 1024 * 1024;

export class DocxMedia implements MediaSource {
  private readonly inner: MediaCache;
  private readonly converted = new Map<string, Promise<Media>>();
  private readonly urls: string[] = [];

  constructor(private readonly pkg: Package) {
    this.inner = new MediaCache(pkg);
  }

  get(part: string): Promise<Media> {
    return this.inner.get(part).then((media) => {
      if (media.url || !media.unsupported || !METAFILES.has(media.unsupported)) return media;
      const key = part.toLowerCase();
      let done = this.converted.get(key);
      if (!done) {
        done = this.pkg.bytes(part, MAX_METAFILE_BYTES).then(async (bytes) => {
          const png = bytes ? await metafileToPng(bytes).catch(() => null) : null;
          if (!png) return media;
          const url = URL.createObjectURL(png);
          this.urls.push(url);
          return { url };
        });
        this.converted.set(key, done);
      }
      return done;
    });
  }

  revokeAll(): void {
    this.inner.revokeAll();
    for (const url of this.urls) URL.revokeObjectURL(url);
    this.urls.length = 0;
    this.converted.clear();
  }
}
