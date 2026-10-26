// SPDX-License-Identifier: MIT
// numbering.xml: abstract list definitions, list instances with overrides, the counters Word keeps while it
// walks through the document, and the number formats (decimal, roman, letters, bullets …).
import type { Styles } from './styles.js';
import { clampNum, kid, kids, wa, wi, wv } from './xml.js';

export interface Level {
  ilvl: number;
  start: number;
  numFmt: string;
  lvlText: string;
  lvlJc: string;
  suff: string;
  pPr: Element | null;
  rPr: Element | null;
  /** Level (1-based) whose use restarts this one; 0 never restarts; undefined: any higher level. */
  restart?: number;
  isLgl: boolean;
  pStyle: string | null;
  picBullet: Element | null;
}

interface Instance {
  abstractId: string;
  overrides: Map<number, { start?: number; level?: Level }>;
}

const MAX_LISTS = 5000;

function parseLevel(el: Element, pictures: Map<string, Element>): Level {
  const ilvl = clampNum(wi(el, 'ilvl') ?? 0, 0, 8);
  const restart = wi(kid(el, 'lvlRestart'));
  const pic = wv(kid(el, 'lvlPicBulletId'));
  return {
    ilvl,
    start: clampNum(wi(kid(el, 'start')) ?? 1, 0, 2147483647),
    numFmt: wv(kid(el, 'numFmt')) ?? 'decimal',
    lvlText: (wv(kid(el, 'lvlText')) ?? '').slice(0, 256),
    lvlJc: wv(kid(el, 'lvlJc')) ?? 'left',
    suff: wv(kid(el, 'suff')) ?? 'tab',
    pPr: kid(el, 'pPr'),
    rPr: kid(el, 'rPr'),
    restart,
    isLgl: !!kid(el, 'isLgl') && wv(kid(el, 'isLgl')) !== '0',
    pStyle: wv(kid(el, 'pStyle')),
    picBullet: pic !== null ? (pictures.get(pic) ?? null) : null,
  };
}

export interface NumberText {
  text: string;
  level: Level;
}

export class Numbering {
  private readonly abstracts = new Map<string, { levels: Level[]; styleLink: string | null }>();
  private readonly instances = new Map<string, Instance>();
  private readonly counters = new Map<string, (number | undefined)[]>();
  private readonly seen = new Set<string>();

  constructor(doc: Document | null, private readonly styles: Styles) {
    const root = doc?.documentElement ?? null;
    const pictures = new Map<string, Element>();
    for (const el of kids(root)) {
      if (el.localName === 'numPicBullet') {
        const id = wa(el, 'numPicBulletId');
        if (id !== null) pictures.set(id, el);
      }
    }
    for (const el of kids(root)) {
      if (this.abstracts.size + this.instances.size > MAX_LISTS * 2) break;
      if (el.localName === 'abstractNum') {
        const id = wa(el, 'abstractNumId');
        if (id === null) continue;
        const levels: Level[] = [];
        for (const lvl of kids(el)) if (lvl.localName === 'lvl') levels[clampNum(wi(lvl, 'ilvl') ?? 0, 0, 8)] = parseLevel(lvl, pictures);
        this.abstracts.set(id, { levels, styleLink: wv(kid(el, 'numStyleLink')) });
      } else if (el.localName === 'num') {
        const id = wa(el, 'numId');
        const abstractId = wv(kid(el, 'abstractNumId'));
        if (id === null || abstractId === null) continue;
        const overrides = new Map<number, { start?: number; level?: Level }>();
        for (const o of kids(el)) {
          if (o.localName !== 'lvlOverride') continue;
          const ilvl = clampNum(wi(o, 'ilvl') ?? 0, 0, 8);
          const start = wi(kid(o, 'startOverride'));
          const lvl = kid(o, 'lvl');
          overrides.set(ilvl, { start, level: lvl ? parseLevel(lvl, pictures) : undefined });
        }
        this.instances.set(id, { abstractId, overrides });
      }
    }
  }

  /** Levels of an abstract definition, following a numbering style link (numStyleLink) once or twice. */
  private levels(abstractId: string, depth = 0): Level[] {
    const abs = this.abstracts.get(abstractId);
    if (!abs) return [];
    if (abs.styleLink && depth < 4) {
      const style = this.styles.byId.get(abs.styleLink);
      const numId = wv(kid(kid(style?.pPr ?? null, 'numPr'), 'numId'));
      const inst = numId ? this.instances.get(numId) : undefined;
      if (inst && inst.abstractId !== abstractId) return this.levels(inst.abstractId, depth + 1);
    }
    return abs.levels;
  }

  level(numId: string | undefined, ilvl: number): Level | null {
    if (!numId || numId === '0') return null;
    const inst = this.instances.get(numId);
    if (!inst) return null;
    const override = inst.overrides.get(ilvl)?.level;
    if (override) return override;
    return this.levels(inst.abstractId)[ilvl] ?? null;
  }

  /** Level that is linked to a paragraph style (w:pStyle in the level), for style-based numbering. */
  levelForStyle(numId: string | undefined, styleId: string | null): number | null {
    if (!numId || !styleId) return null;
    const inst = this.instances.get(numId);
    if (!inst) return null;
    const found = this.levels(inst.abstractId).find((l) => l && l.pStyle === styleId);
    return found ? found.ilvl : null;
  }

  /**
   * Number text without advancing any counter, for lists in headers, footers and text boxes (which Word numbers
   * on their own): bullets as they are, numbers with their current or start values.
   */
  peek(numId: string, ilvl: number): NumberText | null {
    const inst = this.instances.get(numId);
    const lvl = this.level(numId, ilvl);
    if (!inst || !lvl) return null;
    const startOf = (l: number) => inst.overrides.get(l)?.start ?? this.level(numId, l)?.start ?? 1;
    const text = lvl.lvlText.replace(/%([1-9])/g, (_, d: string) => {
      const l = Number(d) - 1;
      const fmt = lvl.isLgl && l !== ilvl ? 'decimal' : (this.level(numId, l)?.numFmt ?? 'decimal');
      return formatNumber(startOf(l), fmt);
    });
    return { text, level: lvl };
  }

  /** Advances the counters for a numbered paragraph and returns its number text. */
  next(numId: string, ilvl: number): NumberText | null {
    const inst = this.instances.get(numId);
    const lvl = this.level(numId, ilvl);
    if (!inst || !lvl) return null;
    // A list instance that restarts a level counts on its own; plain instances share the abstract list.
    const hasStart = [...inst.overrides.values()].some((o) => o.start !== undefined || o.level);
    const key = hasStart ? `n:${numId}` : `a:${inst.abstractId}`;
    let counters = this.counters.get(key);
    if (!counters) {
      counters = [];
      this.counters.set(key, counters);
    }
    if (!this.seen.has(numId)) {
      this.seen.add(numId);
      for (const [l, o] of inst.overrides) if (o.start !== undefined) counters[l] = o.start - 1;
    }
    const startOf = (l: number) => inst.overrides.get(l)?.start ?? this.level(numId, l)?.start ?? 1;
    counters[ilvl] = counters[ilvl] === undefined ? startOf(ilvl) : counters[ilvl]! + 1;
    for (let deeper = ilvl + 1; deeper < 9; deeper++) {
      const d = this.level(numId, deeper);
      const restart = d?.restart;
      // lvlRestart is 1-based: restart this level when level `restart - 1` (or any higher one) is used.
      if (restart === 0) continue;
      if (restart === undefined || ilvl < restart) counters[deeper] = undefined;
    }
    const text = lvl.lvlText.replace(/%([1-9])/g, (_, d: string) => {
      const l = Number(d) - 1;
      const other = this.level(numId, l);
      const value = counters![l] ?? startOf(l);
      const fmt = lvl.isLgl && l !== ilvl ? 'decimal' : (other?.numFmt ?? 'decimal');
      return formatNumber(value, l === ilvl ? lvl.numFmt : fmt);
    });
    return { text, level: lvl };
  }
}

// ---- Number formats ------------------------------------------------------------------------------------------

const ROMAN: [number, string][] = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];

function roman(n: number): string {
  if (n <= 0 || n >= 4000) return String(n);
  let out = '';
  for (const [value, digits] of ROMAN) while (n >= value) (out += digits), (n -= value);
  return out;
}

/** Word's letter numbering repeats the letter: A … Z, AA, BB … */
function letters(n: number, base: string): string {
  if (n <= 0) return String(n);
  const letter = String.fromCharCode(base.charCodeAt(0) + ((n - 1) % 26));
  return letter.repeat(Math.min(Math.floor((n - 1) / 26) + 1, 10));
}

const ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];

function cardinal(n: number): string {
  if (n < 20) return ONES[n] ?? String(n);
  if (n < 100) return TENS[Math.floor(n / 10)] + (n % 10 ? `-${ONES[n % 10]}` : '');
  if (n < 1000) return `${ONES[Math.floor(n / 100)]} hundred${n % 100 ? ` ${cardinal(n % 100)}` : ''}`;
  return String(n);
}

function ordinalText(n: number): string {
  const word = cardinal(n);
  const irregular: Record<string, string> = { one: 'first', two: 'second', three: 'third', five: 'fifth', eight: 'eighth', nine: 'ninth', twelve: 'twelfth' };
  const m = /(\w+)$/.exec(word);
  if (!m) return word;
  const last = m[1];
  const ord = irregular[last] ?? (last.endsWith('y') ? `${last.slice(0, -1)}ieth` : `${last}th`);
  return word.slice(0, -last.length) + ord;
}

function ordinal(n: number): string {
  const mod100 = n % 100;
  const suffix = mod100 >= 11 && mod100 <= 13 ? 'th' : (['th', 'st', 'nd', 'rd'][n % 10] ?? 'th');
  return `${n}${suffix}`;
}

const capitalize = (s: string) => s.charAt(0).toUpperCase() + s.slice(1);
const CHICAGO = ['*', '†', '‡', '§'];
const IDEOGRAPH_DIGITS = ['〇', '一', '二', '三', '四', '五', '六', '七', '八', '九'];

export function formatNumber(n: number, fmt: string): string {
  switch (fmt) {
    case 'decimal':
      return String(n);
    case 'decimalZero':
      return n < 10 && n >= 0 ? `0${n}` : String(n);
    case 'upperRoman':
      return roman(n);
    case 'lowerRoman':
      return roman(n).toLowerCase();
    case 'upperLetter':
      return letters(n, 'A');
    case 'lowerLetter':
      return letters(n, 'a');
    case 'ordinal':
      return ordinal(n);
    case 'cardinalText':
      return capitalize(cardinal(n));
    case 'ordinalText':
      return capitalize(ordinalText(n));
    case 'hex':
      return n.toString(16).toUpperCase();
    case 'chicago':
      return CHICAGO[(n - 1) % 4].repeat(Math.min(Math.floor((n - 1) / 4) + 1, 10));
    case 'decimalEnclosedCircle':
    case 'decimalEnclosedCircleChinese':
      return n >= 1 && n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : String(n);
    case 'decimalEnclosedParen':
      return n >= 1 && n <= 20 ? String.fromCodePoint(0x2474 + n - 1) : `(${n})`;
    case 'decimalEnclosedFullstop':
      return n >= 1 && n <= 20 ? String.fromCodePoint(0x2488 + n - 1) : `${n}.`;
    case 'decimalFullWidth':
    case 'decimalFullWidth2':
      return String(n).replace(/\d/g, (d) => String.fromCharCode(0xff10 + Number(d)));
    case 'russianLower':
    case 'russianUpper': {
      const alphabet = 'абвгдежзиклмнопрстуфхцчшщэюя';
      const ch = alphabet[(n - 1) % alphabet.length] ?? '';
      const s = ch.repeat(Math.floor((n - 1) / alphabet.length) + 1);
      return fmt === 'russianUpper' ? s.toUpperCase() : s;
    }
    case 'ideographDigital':
    case 'japaneseDigitalTenThousand':
      return String(n).replace(/\d/g, (d) => IDEOGRAPH_DIGITS[Number(d)]);
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'japaneseCounting':
    case 'taiwaneseCounting':
    case 'koreanCounting': {
      if (n <= 0 || n >= 100) return String(n).replace(/\d/g, (d) => IDEOGRAPH_DIGITS[Number(d)]);
      const tens = Math.floor(n / 10), ones = n % 10;
      return `${tens > 1 ? IDEOGRAPH_DIGITS[tens] : ''}${tens ? '十' : ''}${ones || !tens ? IDEOGRAPH_DIGITS[ones] : ''}`;
    }
    case 'none':
      return '';
    case 'bullet':
      return '';
    default:
      return String(n);
  }
}
