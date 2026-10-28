// SPDX-License-Identifier: MIT
// Pictures for the PDF: JPEG goes in as it is, PNG with its transparency; everything else the browser can draw
// (GIF, BMP, WebP, SVG, metafiles already turned into PNG) is redrawn on a canvas and stored as PNG.
import type { PDFDocument, PDFImage } from '@cantoo/pdf-lib';

/** Raster size for vector pictures (SVG) relative to their size on the page. */
const VECTOR_SCALE = 3;
const MAX_SIDE = 4096;

const isJpeg = (b: Uint8Array) => b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff;
const isPng = (b: Uint8Array) => b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47;

async function canvasPng(canvas: HTMLCanvasElement): Promise<Uint8Array | null> {
  const blob = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, 'image/png'));
  return blob ? new Uint8Array(await blob.arrayBuffer()) : null;
}

/** Draws a picture the browser can show onto a canvas; vector pictures at a multiple of their page size. */
async function redraw(blob: Blob, size?: { w: number; h: number }): Promise<Uint8Array | null> {
  const url = URL.createObjectURL(blob);
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    const vector = blob.type.includes('svg') || !img.naturalWidth;
    let w = img.naturalWidth || size?.w || 300;
    let h = img.naturalHeight || size?.h || 150;
    if (vector && size) {
      w = size.w * VECTOR_SCALE;
      h = size.h * VECTOR_SCALE;
    }
    const fit = Math.min(1, MAX_SIDE / Math.max(w, h));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(w * fit));
    canvas.height = Math.max(1, Math.round(h * fit));
    canvas.getContext('2d')?.drawImage(img, 0, 0, canvas.width, canvas.height);
    return await canvasPng(canvas);
  } catch {
    return null;
  } finally {
    URL.revokeObjectURL(url);
  }
}

export class ImageCache {
  private readonly cache = new Map<string, Promise<PDFImage | null>>();
  private readonly canvases = new WeakMap<HTMLCanvasElement, Promise<PDFImage | null>>();

  constructor(private readonly doc: PDFDocument) {}

  /** A picture from a blob: or data: URL; others (remote files) are never fetched. */
  fromUrl(url: string, size?: { w: number; h: number }): Promise<PDFImage | null> {
    if (!/^(blob|data):/.test(url)) return Promise.resolve(null);
    const key = size ? `${url}#${Math.round(size.w)}x${Math.round(size.h)}` : url;
    let image = this.cache.get(key);
    if (!image) {
      image = this.load(url, size).catch(() => null);
      this.cache.set(key, image);
    }
    return image;
  }

  private async load(url: string, size?: { w: number; h: number }): Promise<PDFImage | null> {
    const blob = await (await fetch(url)).blob();
    const bytes = new Uint8Array(await blob.arrayBuffer());
    if (isJpeg(bytes)) {
      try {
        return await this.doc.embedJpg(bytes);
      } catch {
        // Some JPEG variants pdf-lib does not take; the browser decodes them instead.
      }
    }
    if (isPng(bytes)) {
      try {
        return await this.doc.embedPng(bytes);
      } catch {
        // Fall through to redrawing.
      }
    }
    const png = await redraw(blob, size);
    return png ? this.doc.embedPng(png) : null;
  }

  fromCanvas(canvas: HTMLCanvasElement): Promise<PDFImage | null> {
    let image = this.canvases.get(canvas);
    if (!image) {
      image = canvasPng(canvas).then((png) => (png ? this.doc.embedPng(png) : null)).catch(() => null);
      this.canvases.set(canvas, image);
    }
    return image;
  }
}
