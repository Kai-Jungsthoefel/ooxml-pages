// SPDX-License-Identifier: MIT
// The PPTX container: ZIP entries, XML parts and the relationships between them (Open Packaging Conventions).
import { BlobReader, ZipReader } from '@zip.js/zip.js';
import { attr, parseXml } from './xml.js';

// A hostile file can declare parts that inflate to gigabytes; nothing legitimate comes close to these.
const MAX_XML_BYTES = 64 * 1024 * 1024;
const MAX_MEDIA_BYTES = 256 * 1024 * 1024;

export interface Rel {
  id: string;
  /** Last segment of the relationship type URI, e.g. "slideLayout", "image", "hyperlink". */
  type: string;
  /** Resolved part name for internal targets, the raw URI for external ones. */
  target: string;
  external: boolean;
}

export type Rels = Map<string, Rel>;

interface ZipEntry {
  size: number;
  read(limit: number): Promise<Uint8Array>;
}

/** Inflates into memory but gives up as soon as the data grows past the limit: declared sizes can lie. */
async function inflate(getData: (writable: WritableStream<Uint8Array>) => Promise<unknown>, limit: number): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  let total = 0;
  await getData(
    new WritableStream<Uint8Array>({
      write(chunk) {
        total += chunk.length;
        if (total > limit) throw new Error('Part too large');
        chunks.push(chunk);
      },
    }),
  );
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

/** Resolves a relationship target against the folder of the part that owns the relationship. */
export function resolvePart(base: string, target: string): string {
  const segments = target.startsWith('/') ? [] : base.split('/').slice(0, -1);
  for (const segment of target.split('/')) {
    if (segment === '..') segments.pop();
    else if (segment !== '.' && segment !== '') segments.push(segment);
  }
  return segments.join('/');
}

function relsPath(part: string): string {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
}

function decodeText(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  return new TextDecoder().decode(bytes);
}

export class Package {
  private readonly xmlCache = new Map<string, Promise<Document | null>>();
  private readonly relsCache = new Map<string, Promise<Rels>>();

  private constructor(
    private readonly reader: ZipReader<unknown>,
    private readonly entries: Map<string, ZipEntry>,
  ) {}

  static async open(file: Blob): Promise<Package> {
    // Set per reader instead of with configure(), which would change zip.js for the whole host application.
    // Without workers zip.js spawns none from inline code (blocked by strict CSPs); DecompressionStream is fast enough.
    const reader = new ZipReader(new BlobReader(file), { useWebWorkers: false, useCompressionStream: true });
    const entries = new Map<string, ZipEntry>();
    for (const entry of await reader.getEntries()) {
      if (entry.directory) continue;
      // Part names are case-insensitive in OPC.
      entries.set(entry.filename.replace(/\\/g, '/').replace(/^\//, '').toLowerCase(), {
        size: entry.uncompressedSize,
        read: (limit) => inflate((writable) => entry.getData!(writable), limit),
      });
    }
    return new Package(reader, entries);
  }

  private entry(part: string): ZipEntry | undefined {
    const key = part.toLowerCase();
    const found = this.entries.get(key);
    if (found) return found;
    // Targets are URIs, so a part named "my image.png" may be referenced as "my%20image.png".
    try {
      return this.entries.get(decodeURIComponent(key));
    } catch {
      return undefined;
    }
  }

  has(part: string): boolean {
    return this.entry(part) !== undefined;
  }

  async bytes(part: string, limit = MAX_MEDIA_BYTES): Promise<Uint8Array | null> {
    const entry = this.entry(part);
    if (!entry || entry.size > limit) return null;
    try {
      return await entry.read(limit);
    } catch {
      return null;
    }
  }

  xml(part: string): Promise<Document | null> {
    const key = part.toLowerCase();
    let cached = this.xmlCache.get(key);
    if (!cached) {
      cached = this.bytes(part, MAX_XML_BYTES).then((data) => (data ? parseXml(decodeText(data)) : null));
      this.xmlCache.set(key, cached);
    }
    return cached;
  }

  rels(part: string): Promise<Rels> {
    const key = part.toLowerCase();
    let cached = this.relsCache.get(key);
    if (!cached) {
      cached = this.xml(relsPath(part)).then((doc) => {
        const rels: Rels = new Map();
        if (!doc) return rels;
        for (const el of doc.getElementsByTagName('*')) {
          if (el.localName !== 'Relationship') continue;
          const id = attr(el, 'Id');
          const target = attr(el, 'Target');
          if (!id || target === null) continue;
          const external = attr(el, 'TargetMode') === 'External';
          rels.set(id, {
            id,
            type: (attr(el, 'Type') ?? '').split('/').pop() ?? '',
            target: external ? target : resolvePart(part, target),
            external,
          });
        }
        return rels;
      });
      this.relsCache.set(key, cached);
    }
    return cached;
  }

  /** First internal relationship of a type, e.g. the slide layout of a slide. */
  async related(part: string, type: string): Promise<string | null> {
    for (const rel of (await this.rels(part)).values()) if (rel.type === type && !rel.external) return rel.target;
    return null;
  }

  close(): void {
    void this.reader.close();
  }
}
