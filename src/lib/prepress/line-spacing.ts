import type JSZip from "jszip";
import { DOMParser, XMLSerializer, type Document as XmlDocument, type Element as XmlElement } from "@xmldom/xmldom";

/**
 * Interlineado como en Word. El visor escribe el interlineado relativo al cuerpo de la letra (1,15 × 12 pt), pero Word
 * lo calcula con la altura natural de cada fuente (ascendente + descendente + interlínea de la fuente) por el múltiplo.
 * «Al menos» el visor lo suma (100 % + N pt) y en Word es el mayor de los dos. «Sin espacio entre párrafos del mismo
 * estilo» (contextualSpacing) el visor lo ignora, y entre dos párrafos Word suma el espacio de después y el de antes,
 * mientras que el navegador se queda con el mayor. Todo esto se corrige en la copia que se dibuja, nunca en el Word.
 */

// ---------------------------------------------------------------------------------------------------------------
// Medidas de la fuente
// ---------------------------------------------------------------------------------------------------------------

export type FontMetrics = {
  /** hhea: ascendente (positivo), descendente (negativo o positivo, se toma el valor absoluto) e interlínea. */
  ascent: number;
  descent: number;
  lineGap: number;
  unitsPerEm: number;
  /** OS/2 usWinAscent / usWinDescent: lo que usa Word en Windows. */
  winAscent?: number;
  winDescent?: number;
};

/**
 * Altura de una línea sencilla en Word, en «cuerpos» (multiplicar por el tamaño de la letra). Es la cuenta de Windows:
 * winAscent + winDescent, más la interlínea de hhea que no esté ya incluida en esas dos (Times New Roman da 1,15;
 * Calibri, 1,22).
 */
export function wordLineRatio(m: FontMetrics): number {
  const upm = m.unitsPerEm > 0 ? m.unitsPerEm : 1000;
  const hhea = Math.abs(m.ascent) + Math.abs(m.descent);
  const gap = Math.max(0, m.lineGap);
  const win = (m.winAscent ?? 0) + Math.abs(m.winDescent ?? 0);
  if (win > 0) return (win + Math.max(0, gap - (win - hhea))) / upm;
  return (hhea + gap) / upm;
}

function tag(bytes: Uint8Array, offset: number): string {
  return String.fromCharCode(bytes[offset] ?? 0, bytes[offset + 1] ?? 0, bytes[offset + 2] ?? 0, bytes[offset + 3] ?? 0);
}

/** Lee head, hhea y OS/2 de un TTF, OTF o TTC (la primera fuente). Sin dependencias; falla con WOFF. */
export function readFontMetrics(bytes: Uint8Array): FontMetrics {
  if (bytes.byteLength < 12) throw new Error("Archivo de fuente vacío.");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let base = 0;
  const kind = tag(bytes, 0);
  if (kind === "wOFF" || kind === "wOF2") throw new Error("Fuente WOFF.");
  if (kind === "ttcf") base = view.getUint32(12);
  const numTables = view.getUint16(base + 4);
  const tables = new Map<string, number>();
  for (let i = 0; i < numTables; i++) {
    const rec = base + 12 + i * 16;
    if (rec + 16 > bytes.byteLength) break;
    tables.set(tag(bytes, rec), view.getUint32(rec + 8));
  }
  const head = tables.get("head");
  const hhea = tables.get("hhea");
  if (head === undefined || hhea === undefined || head + 20 > bytes.byteLength || hhea + 10 > bytes.byteLength) {
    throw new Error("La fuente no tiene medidas de renglón.");
  }
  const metrics: FontMetrics = {
    unitsPerEm: view.getUint16(head + 18),
    ascent: view.getInt16(hhea + 4),
    descent: view.getInt16(hhea + 6),
    lineGap: view.getInt16(hhea + 8),
  };
  const os2 = tables.get("OS/2");
  if (os2 !== undefined && os2 + 78 <= bytes.byteLength) {
    metrics.winAscent = view.getUint16(os2 + 74);
    metrics.winDescent = view.getUint16(os2 + 76);
  }
  return metrics;
}

type KitFont = {
  unitsPerEm: number;
  ascent: number;
  descent: number;
  lineGap: number;
  "OS/2"?: { winAscent?: number; winDescent?: number };
};

/** Medidas de un archivo de fuente. TTF, OTF y TTC se leen directo; lo demás (WOFF) pasa por fontkit. */
export async function fontMetricsFromBytes(bytes: Uint8Array): Promise<FontMetrics | null> {
  try {
    return readFontMetrics(bytes);
  } catch {
    /* Probamos con fontkit. */
  }
  try {
    const fontkit = (await import("@pdf-lib/fontkit")).default;
    const created = fontkit.create(bytes) as unknown as KitFont | { fonts?: KitFont[] };
    const font = "fonts" in created && Array.isArray(created.fonts) ? created.fonts[0] : (created as KitFont);
    if (!font || !font.unitsPerEm) return null;
    const os2 = font["OS/2"];
    return {
      unitsPerEm: font.unitsPerEm,
      ascent: font.ascent,
      descent: font.descent,
      lineGap: font.lineGap,
      winAscent: os2?.winAscent,
      winDescent: os2?.winDescent,
    };
  } catch {
    return null;
  }
}

function familyKey(family: string): string {
  return family.trim().replace(/^['"]+|['"]+$/g, "").trim().toLowerCase();
}

/**
 * Medidas de las fuentes de Windows con que se armó el Word. Los sustitutos libres copian el ancho de las letras pero
 * no siempre el alto del renglón (Caladea da 1,30 y Cambria 1,17), así que para estas familias manda esta tabla.
 */
export const WORD_FONT_METRICS: Record<string, FontMetrics> = {
  "Times New Roman": { unitsPerEm: 2048, ascent: 1825, descent: -443, lineGap: 87, winAscent: 1825, winDescent: 443 },
  Times: { unitsPerEm: 2048, ascent: 1825, descent: -443, lineGap: 87, winAscent: 1825, winDescent: 443 },
  Arial: { unitsPerEm: 2048, ascent: 1854, descent: -434, lineGap: 67, winAscent: 1854, winDescent: 434 },
  Helvetica: { unitsPerEm: 2048, ascent: 1854, descent: -434, lineGap: 67, winAscent: 1854, winDescent: 434 },
  "Courier New": { unitsPerEm: 2048, ascent: 1705, descent: -615, lineGap: 0, winAscent: 1705, winDescent: 615 },
  Calibri: { unitsPerEm: 2048, ascent: 1536, descent: -512, lineGap: 452, winAscent: 1950, winDescent: 550 },
  "Calibri Light": { unitsPerEm: 2048, ascent: 1536, descent: -512, lineGap: 452, winAscent: 1950, winDescent: 550 },
  Cambria: { unitsPerEm: 2048, ascent: 1946, descent: -455, lineGap: 0, winAscent: 1946, winDescent: 455 },
  Georgia: { unitsPerEm: 2048, ascent: 1878, descent: -449, lineGap: 0, winAscent: 1878, winDescent: 449 },
  Garamond: { unitsPerEm: 2048, ascent: 1765, descent: -539, lineGap: 0, winAscent: 1765, winDescent: 539 },
  "Book Antiqua": { unitsPerEm: 2048, ascent: 1891, descent: -578, lineGap: 0, winAscent: 1967, winDescent: 578 },
  "Palatino Linotype": { unitsPerEm: 2048, ascent: 1499, descent: -582, lineGap: 682, winAscent: 2150, winDescent: 613 },
  Verdana: { unitsPerEm: 2048, ascent: 2059, descent: -430, lineGap: 0, winAscent: 2059, winDescent: 430 },
  Tahoma: { unitsPerEm: 2048, ascent: 2049, descent: -423, lineGap: 0, winAscent: 2049, winDescent: 423 },
  "Segoe UI": { unitsPerEm: 2048, ascent: 2210, descent: -514, lineGap: 0, winAscent: 2210, winDescent: 514 },
  "Century Gothic": { unitsPerEm: 2048, ascent: 2060, descent: -451, lineGap: 0, winAscent: 1989, winDescent: 451 },
  Candara: { unitsPerEm: 2048, ascent: 1484, descent: -564, lineGap: 452, winAscent: 1950, winDescent: 550 },
  Constantia: { unitsPerEm: 2048, ascent: 1538, descent: -510, lineGap: 452, winAscent: 1950, winDescent: 550 },
  "Trebuchet MS": { unitsPerEm: 2048, ascent: 1923, descent: -455, lineGap: 0, winAscent: 1923, winDescent: 455 },
  Century: { unitsPerEm: 2048, ascent: 2019, descent: -443, lineGap: 0, winAscent: 2019, winDescent: 442 },
  "Bookman Old Style": { unitsPerEm: 2048, ascent: 1929, descent: -475, lineGap: 0, winAscent: 1831, winDescent: 473 },
};

/** Prioridad de cada origen: la fuente que subió la persona, después la tabla de Windows, después los sustitutos. */
const RANK = { substitute: 0, known: 1, uploaded: 2 } as const;

type Source = { rank: number; regular: boolean; load: () => Promise<FontMetrics | null> };

/**
 * Medidas por familia. Se carga con los archivos que de verdad se usan en pantalla (los sustitutos si están activos,
 * las fuentes subidas y sus alias). Los archivos por dirección se leen recién cuando hacen falta.
 */
export class FontMetricsCache {
  private sources = new Map<string, Source>();
  private done = new Map<string, Promise<FontMetrics | null>>();

  constructor(known: Record<string, FontMetrics> = WORD_FONT_METRICS) {
    for (const [family, metrics] of Object.entries(known)) {
      this.put(family, { rank: RANK.known, regular: true, load: async () => metrics });
    }
  }

  /** Gana el origen de más prioridad; dentro del mismo, la primera fuente regular (400, normal). */
  private put(family: string, source: Source) {
    const key = familyKey(family);
    if (!key) return;
    const current = this.sources.get(key);
    if (current) {
      if (current.rank > source.rank) return;
      if (current.rank === source.rank && (current.regular || !source.regular)) return;
    }
    this.sources.set(key, source);
    this.done.delete(key);
  }

  /** Fuente subida por la persona (o su alias): manda sobre todo lo demás. */
  addBytes(family: string, bytes: Uint8Array, face?: { weight?: number; style?: string }) {
    this.put(family, { rank: RANK.uploaded, regular: isRegular(face), load: () => fontMetricsFromBytes(bytes) });
  }

  /** Sustituto libre (public/fonts): solo si no hay nada mejor para esa familia. */
  addUrl(family: string, url: string, face?: { weight?: number; style?: string }) {
    this.put(family, {
      rank: RANK.substitute,
      regular: isRegular(face),
      load: async () => {
        try {
          const response = await fetch(url);
          return response.ok ? fontMetricsFromBytes(new Uint8Array(await response.arrayBuffer())) : null;
        } catch {
          return null;
        }
      },
    });
  }

  metrics = (family: string): Promise<FontMetrics | null> => {
    const key = familyKey(family);
    let found = this.done.get(key);
    if (!found) {
      const source = this.sources.get(key);
      found = source ? source.load().catch(() => null) : Promise.resolve(null);
      this.done.set(key, found);
    }
    return found;
  };
}

function isRegular(face?: { weight?: number; style?: string }): boolean {
  return !face || ((face.weight ?? 400) === 400 && (face.style ?? "normal") === "normal");
}

// ---------------------------------------------------------------------------------------------------------------
// Lo que escribió el visor
// ---------------------------------------------------------------------------------------------------------------

export type LineRule =
  | { kind: "single" }
  | { kind: "auto"; multiple: number }
  | { kind: "atLeast"; pt: number }
  | { kind: "exact"; pt: number }
  | { kind: "other" };

/**
 * Lo que el visor pone en line-height según w:spacing: auto → «1.15», atLeast → «calc(100% + 6pt)», exact → «12pt».
 * Sin valor, interlineado sencillo (lo que hace Word si el párrafo no dice nada).
 */
export function parseDeclaredLineHeight(value: string): LineRule {
  const text = value.trim().toLowerCase();
  if (!text || ["normal", "inherit", "initial", "unset", "revert"].includes(text)) return { kind: "single" };
  if (/^\d*\.?\d+$/.test(text)) {
    const multiple = Number(text);
    return multiple > 0 ? { kind: "auto", multiple } : { kind: "other" };
  }
  const least = /^calc\(\s*100%\s*\+\s*(\d*\.?\d+)(pt|px)\s*\)$/.exec(text);
  if (least) return { kind: "atLeast", pt: Number(least[1]) * (least[2] === "px" ? 0.75 : 1) };
  const exact = /^(\d*\.?\d+)pt$/.exec(text);
  if (exact) return { kind: "exact", pt: Number(exact[1]) };
  return { kind: "other" };
}

/** Alto del renglón en px para una letra de `fontPx` cuya línea sencilla mide `ratio` cuerpos. */
export function wordLinePx(rule: LineRule, ratio: number, fontPx: number): number | null {
  const natural = ratio * fontPx;
  switch (rule.kind) {
    case "single":
      return natural;
    case "auto":
      return natural * rule.multiple;
    case "atLeast":
      return Math.max(natural, (rule.pt * 96) / 72);
    default:
      return null;
  }
}

/** Especificidad aproximada de un selector simple (los que escribe el visor), como número comparable. */
export function specificity(selector: string): number {
  const text = selector
    .replace(/"[^"]*"|'[^']*'/g, "")
    .replace(/::?(before|after|first-line|first-letter|marker|placeholder|selection)\b/g, "")
    .replace(/::[\w-]+/g, "");
  const ids = (text.match(/#[\w-]+/g) ?? []).length;
  const classes =
    (text.match(/\.[\w-]+/g) ?? []).length +
    (text.match(/\[[^\]]*\]/g) ?? []).length +
    (text.match(/:(?!not\(|is\(|where\()[\w-]+/g) ?? []).length;
  const types = (text.replace(/\[[^\]]*\]/g, "").match(/(?:^|[\s>+~(,])([a-zA-Z][\w-]*)/g) ?? []).length;
  return ids * 1_000_000 + classes * 1_000 + types;
}

/** Parte «a, b:not(c, d)» por las comas de afuera. */
export function splitSelectors(text: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === "(" || ch === "[") depth += 1;
    else if (ch === ")" || ch === "]") depth = Math.max(0, depth - 1);
    else if (ch === "," && depth === 0) {
      out.push(text.slice(start, i).trim());
      start = i + 1;
    }
  }
  out.push(text.slice(start).trim());
  return out.filter(Boolean);
}

type LineRuleCss = { selectors: Array<{ text: string; spec: number }>; value: string; important: boolean; order: number };

function collectRules(doc: Document): LineRuleCss[] {
  const found: LineRuleCss[] = [];
  const view = doc.defaultView;
  const walk = (list: CSSRuleList) => {
    for (const rule of Array.from(list)) {
      const media = (rule as CSSMediaRule).media;
      if (media && view && media.mediaText && !view.matchMedia(media.mediaText).matches) continue;
      const style = (rule as CSSStyleRule).style;
      const selectorText = (rule as CSSStyleRule).selectorText;
      if (style && selectorText) {
        const value = style.getPropertyValue("line-height");
        if (value) {
          found.push({
            selectors: splitSelectors(selectorText).map((text) => ({ text, spec: specificity(text) })),
            value,
            important: style.getPropertyPriority("line-height") === "important",
            order: found.length,
          });
        }
      }
      const inner = (rule as CSSGroupingRule).cssRules;
      if (inner && !selectorText) walk(inner);
    }
  };
  for (const sheet of Array.from(doc.styleSheets)) {
    try {
      walk(sheet.cssRules);
    } catch {
      /* Hojas de otro origen: no se pueden leer y no son del Word. */
    }
  }
  return found;
}

/** El line-height que le toca al elemento según el estilo en línea y las reglas (sin herencia). */
function declaredLineHeight(el: HTMLElement, rules: LineRuleCss[]): string {
  let best: { value: string; rank: number; order: number } | null = null;
  for (const rule of rules) {
    let spec = -1;
    for (const selector of rule.selectors) {
      try {
        if (selector.spec > spec && el.matches(selector.text)) spec = selector.spec;
      } catch {
        /* Selector que el navegador no entiende. */
      }
    }
    if (spec < 0) continue;
    const rank = (rule.important ? 1e12 : 0) + spec;
    if (!best || rank > best.rank || (rank === best.rank && rule.order > best.order)) {
      best = { value: rule.value, rank, order: rule.order };
    }
  }
  const inline = el.style.getPropertyValue("line-height");
  if (inline && (el.style.getPropertyPriority("line-height") === "important" || !best || best.rank < 1e12)) return inline;
  return best?.value ?? "";
}


// ---------------------------------------------------------------------------------------------------------------
// «Sin espacio entre párrafos del mismo estilo» (w:contextualSpacing)
// ---------------------------------------------------------------------------------------------------------------

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/** Sufijo de los estilos copiados en la copia de composición para los párrafos que cambian la opción a mano. */
export const CTX_SUFFIX = "PliegoCtx";
const CTX_CLASS = /pliegoctx([01])$/;

/** Las claves con que el visor nombra los estilos en las clases (docx_<clave>). */
export function styleClassKey(id: string): string {
  return id.replace(/[ .]+/g, "-").replace(/[&]+/g, "and").toLowerCase();
}

export type SpacingInfo = {
  /** Estilos de párrafo con «sin espacio entre párrafos del mismo estilo», como clave de clase del visor. */
  contextual: string[];
  /** Estilo de párrafo predeterminado (el de los párrafos sin estilo), como clave de clase. */
  defaultStyle: string | null;
};

export const EMPTY_SPACING: SpacingInfo = { contextual: [], defaultStyle: null };

function localName(node: XmlElement): string {
  const raw = node.localName || node.nodeName || "";
  const i = raw.indexOf(":");
  return i >= 0 ? raw.slice(i + 1) : raw;
}

function childrenOf(el: XmlElement, name?: string): XmlElement[] {
  return Array.from(el.childNodes).filter(
    (node): node is XmlElement => node.nodeType === 1 && (!name || localName(node as XmlElement) === name),
  );
}

function descendants(el: XmlElement | XmlDocument, name: string): XmlElement[] {
  return Array.from(el.getElementsByTagNameNS(W, name)) as XmlElement[];
}

function attr(el: XmlElement | undefined, name: string): string {
  return el ? el.getAttributeNS(W, name) || el.getAttribute(`w:${name}`) || "" : "";
}

/** <w:contextualSpacing/> = sí; w:val="0|false|off" = no. */
function boolValue(el: XmlElement): boolean {
  const value = attr(el, "val").toLowerCase();
  return !["0", "false", "off"].includes(value);
}

function parseXml(xml: string): XmlDocument | null {
  const doc = new DOMParser({
    onError: () => {
      /* Si no se puede leer, esa parte queda igual. */
    },
  }).parseFromString(xml, "text/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) return null;
  return doc;
}

function serialize(doc: XmlDocument): string {
  let out = new XMLSerializer().serializeToString(doc);
  if (!out.startsWith("<?xml")) out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
  return out;
}

type StyleTable = {
  byId: Map<string, XmlElement>;
  defaultId: string | null;
  /** Valor resuelto (con basedOn y docDefaults) de contextualSpacing por estilo de párrafo. */
  contextual: (id: string) => boolean;
};

function readStyles(doc: XmlDocument): StyleTable {
  const byId = new Map<string, XmlElement>();
  let defaultId: string | null = null;
  for (const style of descendants(doc, "style")) {
    if (attr(style, "type") !== "paragraph") continue;
    const id = attr(style, "styleId");
    if (!id) continue;
    byId.set(id, style);
    const isDefault = attr(style, "default").toLowerCase();
    if (!defaultId && (isDefault === "1" || isDefault === "true" || isDefault === "on")) defaultId = id;
  }
  let base = false;
  for (const defaults of descendants(doc, "pPrDefault")) {
    for (const ppr of childrenOf(defaults, "pPr")) {
      const own = childrenOf(ppr, "contextualSpacing")[0];
      if (own) base = boolValue(own);
    }
  }
  const memo = new Map<string, boolean>();
  const contextual = (id: string): boolean => {
    const cached = memo.get(id);
    if (cached !== undefined) return cached;
    const chain = new Set<string>();
    let current: string | null = id;
    let value = base;
    while (current && !chain.has(current)) {
      chain.add(current);
      const style = byId.get(current);
      if (!style) break;
      const ppr = childrenOf(style, "pPr")[0];
      const own = ppr ? childrenOf(ppr, "contextualSpacing")[0] : undefined;
      if (own) {
        value = boolValue(own);
        break;
      }
      current = attr(childrenOf(style, "basedOn")[0], "val") || null;
    }
    memo.set(id, value);
    return value;
  };
  return { byId, defaultId, contextual };
}

/** Qué estilos tienen la opción y cuál es el predeterminado, leído del styles.xml (ya con las copias de abajo). */
export function readSpacingInfo(stylesXml: string): SpacingInfo {
  const doc = parseXml(stylesXml);
  if (!doc) return { ...EMPTY_SPACING };
  const table = readStyles(doc);
  const contextual: string[] = [];
  for (const id of table.byId.keys()) {
    const copy = new RegExp(`${CTX_SUFFIX}([01])$`).exec(id);
    const on = copy ? copy[1] === "1" : table.contextual(id);
    if (on) contextual.push(styleClassKey(id));
  }
  return { contextual, defaultStyle: table.defaultId ? styleClassKey(table.defaultId) : null };
}

/**
 * w:line sin w:lineRule es «múltiple» para Word (auto), pero el visor lo toma como exacto. En la copia de composición se
 * escribe explícito. Cambia solo esas etiquetas; el resto del XML queda byte a byte.
 */
export function normalizeLineRules(xml: string): { xml: string; changed: number } {
  let changed = 0;
  const out = xml.replace(/<w:spacing\b[^>]*?\/?>/g, (tagText) => {
    if (!/\sw:line\s*=/.test(tagText) || /\sw:lineRule\s*=/.test(tagText)) return tagText;
    changed += 1;
    return tagText.replace(/^<w:spacing\b/, '<w:spacing w:lineRule="auto"');
  });
  return { xml: changed ? out : xml, changed };
}

/**
 * Los párrafos que cambian a mano la opción «sin espacio entre párrafos del mismo estilo» pasan, en la copia de
 * composición, a una copia idéntica de su estilo con un sufijo: así el visor deja ver la opción en la clase. Para
 * comparar «mismo estilo» se quita el sufijo. Los estilos con numeración propia no se copian (la numeración del visor
 * se ata al nombre del estilo) y se cuentan en `skipped`.
 */
export function tagDirectContextual(
  stylesXml: string,
  parts: string[],
): { stylesXml: string; parts: string[]; tagged: number; skipped: number } {
  const stylesDoc = parseXml(stylesXml);
  const root = stylesDoc?.documentElement;
  if (!stylesDoc || !root) return { stylesXml, parts, tagged: 0, skipped: 0 };
  const table = readStyles(stylesDoc);
  let tagged = 0;
  let skipped = 0;
  let stylesChanged = false;
  const outParts = parts.map((xml) => {
    if (!/contextualSpacing/.test(xml)) return xml;
    const doc = parseXml(xml);
    if (!doc) return xml;
    let changed = false;
    for (const p of descendants(doc, "p")) {
      const ppr = childrenOf(p, "pPr")[0];
      const own = ppr ? childrenOf(ppr, "contextualSpacing")[0] : undefined;
      if (!ppr || !own) continue;
      const value = boolValue(own);
      const pStyle = childrenOf(ppr, "pStyle")[0];
      const id = attr(pStyle, "val") || table.defaultId;
      const style = id ? table.byId.get(id) : undefined;
      if (!id || !style) {
        skipped += 1;
        continue;
      }
      if (table.contextual(id) === value) continue;
      const styleOwn = childrenOf(style, "pPr")[0];
      if (styleOwn && childrenOf(styleOwn, "numPr").length) {
        skipped += 1;
        continue;
      }
      const copyId = `${id}${CTX_SUFFIX}${value ? 1 : 0}`;
      if (!table.byId.has(copyId)) {
        const copy = style.cloneNode(true) as XmlElement;
        copy.setAttributeNS(W, "w:styleId", copyId);
        if (copy.hasAttributeNS(W, "default")) copy.removeAttributeNS(W, "default");
        if (copy.hasAttribute("w:default")) copy.removeAttribute("w:default");
        const name = childrenOf(copy, "name")[0];
        if (name) name.setAttributeNS(W, "w:val", `${attr(name, "val") || id} ${CTX_SUFFIX}${value ? 1 : 0}`);
        root.appendChild(copy);
        table.byId.set(copyId, copy);
        stylesChanged = true;
      }
      if (pStyle) pStyle.setAttributeNS(W, "w:val", copyId);
      else {
        const created = doc.createElementNS(W, "w:pStyle");
        created.setAttributeNS(W, "w:val", copyId);
        ppr.insertBefore(created, ppr.firstChild);
      }
      tagged += 1;
      changed = true;
    }
    return changed ? serialize(doc) : xml;
  });
  return { stylesXml: stylesChanged ? serialize(stylesDoc) : stylesXml, parts: outParts, tagged, skipped };
}

const SPACING_PARTS = /^word\/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$/;

/**
 * Todo lo de interlineado que va en la copia de composición (prepareRenderCopy): lineRule explícito y la opción
 * «sin espacio entre párrafos del mismo estilo» visible para el visor. Devuelve lo que necesita fixLineSpacing.
 */
export async function prepareLineSpacing(zip: JSZip): Promise<{ spacing: SpacingInfo; warnings: string[] }> {
  const warnings: string[] = [];
  const names = Object.keys(zip.files).filter(
    (name) => SPACING_PARTS.test(name) || name === "word/styles.xml" || name === "word/numbering.xml",
  );
  const originals = new Map<string, string>();
  const texts = new Map<string, string>();
  for (const name of names) {
    const file = zip.file(name);
    if (!file) continue;
    const text = await file.async("string");
    originals.set(name, text);
    texts.set(name, normalizeLineRules(text).xml);
  }
  const styles = texts.get("word/styles.xml");
  if (styles) {
    const partNames = [...texts.keys()].filter((name) => SPACING_PARTS.test(name));
    const result = tagDirectContextual(
      styles,
      partNames.map((name) => texts.get(name)!),
    );
    texts.set("word/styles.xml", result.stylesXml);
    partNames.forEach((name, index) => texts.set(name, result.parts[index]!));
    if (result.skipped) {
      const one = result.skipped === 1;
      warnings.push(
        `${result.skipped} ${one ? "párrafo cambia" : "párrafos cambian"} a mano «No agregar espacio entre párrafos del mismo estilo» y no se ${one ? "pudo" : "pudieron"} reproducir: ahí el espacio entre párrafos puede quedar distinto del Word.`,
      );
    }
  }
  for (const [name, text] of texts) if (text !== originals.get(name)) zip.file(name, text);
  return { spacing: styles ? readSpacingInfo(texts.get("word/styles.xml")!) : { ...EMPTY_SPACING }, warnings };
}

// ---------------------------------------------------------------------------------------------------------------
// Ajuste sobre las páginas dibujadas
// ---------------------------------------------------------------------------------------------------------------

export type LineSpacingOptions = {
  metrics: (family: string) => Promise<FontMetrics | null>;
  /** Lo que devolvió prepareLineSpacing. Sin esto no se aplica «sin espacio entre párrafos del mismo estilo». */
  spacing?: SpacingInfo;
  /** Sumar el espacio de después y el de antes entre dos párrafos, como Word (el navegador toma el mayor). */
  sumSpacing?: boolean;
  /** Prefijo de clases del visor (renderAsync className). */
  className?: string;
};

function firstFamily(list: string): string {
  return splitSelectors(list)[0]?.trim().replace(/^['"]+|['"]+$/g, "") ?? "";
}

function px(value: string): number {
  const n = Number.parseFloat(value);
  return Number.isFinite(n) ? n : 0;
}

const round = (n: number) => `${Math.round(n * 1000) / 1000}px`;

/** Alto de línea normal del navegador para una familia (si no hay archivo de la fuente). */
function measureNormal(doc: Document, fontFamily: string): number | null {
  if (!doc.body) return null;
  const probe = doc.createElement("div");
  probe.style.cssText =
    "position:absolute;left:-10000px;top:0;visibility:hidden;margin:0;padding:0;border:0;line-height:normal;font-size:100px;white-space:nowrap";
  probe.style.fontFamily = fontFamily;
  probe.textContent = "Hg";
  doc.body.appendChild(probe);
  const height = probe.offsetHeight;
  probe.remove();
  return height > 0 ? height / 100 : null;
}

type Face = { family: string; list: string; size: number };

const faceKey = (face: Face) => `${face.list}|${face.size}`;

/**
 * Pone a cada párrafo el interlineado de Word. Va después de dibujar y de cargar las fuentes, y ANTES de repaginate
 * (cambia la altura de todo). Devuelve avisos para la persona.
 */
export async function fixLineSpacing(
  root: HTMLElement,
  opts: LineSpacingOptions,
): Promise<{ paragraphs: number; warnings: string[] }> {
  const doc = root.ownerDocument;
  const view = doc.defaultView;
  if (!view) return { paragraphs: 0, warnings: [] };
  const className = opts.className ?? "docx";
  const prefix = `${className}_`;
  let paragraphs = [...root.querySelectorAll(`section.${className} p`)] as HTMLElement[];
  if (!paragraphs.length) paragraphs = [...root.querySelectorAll("p")] as HTMLElement[];
  if (!paragraphs.length) return { paragraphs: 0, warnings: [] };
  const rules = collectRules(doc);

  // 1. Lectura (todo junto, para no recalcular estilos en cada párrafo).
  const probes = paragraphs.map((p) => {
    const probe = doc.createElement("span");
    probe.dataset.pliegoProbe = "";
    p.appendChild(probe);
    return probe;
  });
  type Read = {
    p: HTMLElement;
    rule: LineRule;
    own: Face;
    base: Face;
    spans: Array<{ el: HTMLElement; face: Face }>;
    top: number;
    bottom: number;
    display: string;
    float: string;
  };
  const faceOf = (style: CSSStyleDeclaration): Face => ({
    family: firstFamily(style.fontFamily),
    list: style.fontFamily,
    size: px(style.fontSize),
  });
  const reads: Read[] = paragraphs.map((p, index) => {
    const style = view.getComputedStyle(p);
    const spans = ([...p.querySelectorAll("span")] as HTMLElement[])
      .filter((span) => span !== probes[index] && span.closest("p") === p)
      .map((el) => ({ el, face: faceOf(view.getComputedStyle(el)) }));
    return {
      p,
      rule: parseDeclaredLineHeight(declaredLineHeight(p, rules)),
      own: faceOf(style),
      base: faceOf(view.getComputedStyle(probes[index]!)),
      spans,
      top: px(style.marginTop),
      bottom: px(style.marginBottom),
      display: style.display,
      float: style.cssFloat,
    };
  });
  for (const probe of probes) probe.remove();

  // 2. Medidas de cada familia.
  const ratios = new Map<string, number | null>();
  const approximate = new Set<string>();
  const families = new Map<string, string>();
  for (const read of reads) {
    families.set(read.base.family, read.base.list);
    for (const span of read.spans) families.set(span.face.family, span.face.list);
  }
  await Promise.all(
    [...families.keys()].map(async (family) => {
      const metrics = family ? await opts.metrics(family).catch(() => null) : null;
      ratios.set(family, metrics ? wordLineRatio(metrics) : null);
    }),
  );
  for (const [family, list] of families) {
    if (ratios.get(family)) continue;
    const measured = measureNormal(doc, list);
    ratios.set(family, measured);
    if (measured && family) approximate.add(family);
  }
  const usedApproximate = new Set<string>();

  // 3. Espacio entre párrafos.
  const contextual = new Set(opts.spacing?.contextual ?? []);
  const styleOf = (p: HTMLElement): string | null => {
    for (const name of p.classList) if (name.startsWith(prefix)) return name.slice(prefix.length);
    return opts.spacing?.defaultStyle ?? null;
  };
  const sameStyle = (a: string | null, b: string | null) =>
    a !== null && b !== null && a.replace(CTX_CLASS, "") === b.replace(CTX_CLASS, "");
  const byElement = new Map(reads.map((read) => [read.p, read]));
  const margins = new Map(reads.map((read) => [read.p, { top: read.top, bottom: read.bottom }]));
  const flowing = (read: Read | undefined) => !!read && read.float === "none" && /^(block|list-item|flow-root)$/.test(read.display);
  if (opts.spacing) {
    for (const read of reads) {
      const key = styleOf(read.p);
      if (!key || !contextual.has(key)) continue;
      const prev = byElement.get(read.p.previousElementSibling as HTMLElement);
      const next = byElement.get(read.p.nextElementSibling as HTMLElement);
      if (prev && sameStyle(key, styleOf(prev.p))) margins.get(read.p)!.top = 0;
      if (next && sameStyle(key, styleOf(next.p))) margins.get(read.p)!.bottom = 0;
    }
  }
  if (opts.sumSpacing ?? true) {
    for (const read of reads) {
      const prev = byElement.get(read.p.previousElementSibling as HTMLElement);
      if (!flowing(prev) || !flowing(read)) continue;
      const mine = margins.get(read.p)!;
      const theirs = margins.get(prev!.p)!;
      if (mine.top <= 0 || theirs.bottom <= 0) continue;
      theirs.bottom += mine.top;
      mine.top = 0;
    }
  }

  // 4. Escritura.
  let changed = 0;
  for (const read of reads) {
    const { p, rule } = read;
    const margin = margins.get(p)!;
    if (Math.abs(margin.top - read.top) > 0.01) p.style.marginTop = round(margin.top);
    if (Math.abs(margin.bottom - read.bottom) > 0.01) p.style.marginBottom = round(margin.bottom);
    if (rule.kind === "other") continue;
    // La letra del párrafo pasa a ser la de su texto: si no, el «puntal» de cada renglón usa la letra de la app y
    // agranda los renglones.
    if (read.own.list !== read.base.list) p.style.fontFamily = read.base.list;
    if (Math.abs(read.own.size - read.base.size) > 0.01) p.style.fontSize = round(read.base.size);
    if (rule.kind === "exact") {
      changed += 1;
      continue;
    }
    const lineFor = (face: Face): number | null => {
      const ratio = ratios.get(face.family);
      if (!ratio || !(face.size > 0)) return null;
      if (approximate.has(face.family)) usedApproximate.add(face.family);
      return wordLinePx(rule, ratio, face.size);
    };
    const line = lineFor(read.base);
    if (line === null) continue;
    p.style.lineHeight = round(line);
    p.style.minHeight = round(line);
    changed += 1;
    // Cada tramo con otra letra u otro tamaño lleva su propio renglón (Word usa el mayor de la línea).
    const assigned = new Map<HTMLElement, string>();
    for (const span of read.spans) {
      let parentKey = faceKey(read.base);
      for (let up = span.el.parentElement; up && up !== p; up = up.parentElement) {
        const key = assigned.get(up);
        if (key) {
          parentKey = key;
          break;
        }
      }
      const key = faceKey(span.face);
      assigned.set(span.el, key);
      if (key === parentKey) continue;
      const own = lineFor(span.face);
      if (own !== null) span.el.style.lineHeight = round(own);
    }
  }

  const warnings: string[] = [];
  if (usedApproximate.size) {
    const names = [...usedApproximate].map((family) => `«${family}»`).join(", ");
    warnings.push(
      `No tengo el archivo de ${usedApproximate.size === 1 ? "la letra" : "las letras"} ${names}: el interlineado se calculó con las medidas del navegador y puede quedar un poco distinto del Word. Subí la fuente para que coincida.`,
    );
  }
  return { paragraphs: changed, warnings };
}

// ---------------------------------------------------------------------------------------------------------------
// Columnas
// ---------------------------------------------------------------------------------------------------------------

/**
 * Word llena las columnas de izquierda a derecha y solo las equilibra antes de un salto de sección continuo; el
 * navegador siempre las equilibra. Va DESPUÉS de repaginate: en el último bloque de columnas de cada página se fija
 * el alto disponible y se llena en orden. Si así algo no entra, se vuelve a dejar equilibrado.
 * `balanced` dice qué bloques van equilibrados igual (por ejemplo, el final de una sección seguida de una continua).
 */
export function fillLastColumns(
  root: HTMLElement,
  balanced?: (article: HTMLElement, page: HTMLElement) => boolean,
): number {
  const view = root.ownerDocument.defaultView;
  if (!view) return 0;
  let filled = 0;
  for (const page of [...root.querySelectorAll("section.docx")] as HTMLElement[]) {
    const articles = ([...page.children] as HTMLElement[]).filter((child) => child.tagName === "ARTICLE");
    const last = articles[articles.length - 1];
    if (!last || !last.children.length) continue;
    if (!(Number.parseInt(view.getComputedStyle(last).columnCount, 10) > 1)) continue;
    if (balanced?.(last, page)) continue;
    const sheet = px(view.getComputedStyle(page).minHeight);
    if (!(sheet > 0) || page.offsetHeight > sheet + 1) continue;
    const box = page.getBoundingClientRect();
    if (!(box.height > 0) || !(page.offsetHeight > 0)) continue;
    const k = box.height / page.offsetHeight;
    const next = last.nextElementSibling as HTMLElement | null;
    const limit = next
      ? next.getBoundingClientRect().top - px(view.getComputedStyle(next).marginTop) * k
      : box.bottom - px(view.getComputedStyle(page).paddingBottom) * k;
    const spare = (limit - last.getBoundingClientRect().bottom) / k;
    if (!(spare > 2)) continue;
    const before = { height: last.style.height, fill: last.style.columnFill };
    last.style.height = round(last.offsetHeight + spare - 1);
    last.style.columnFill = "auto";
    const fits =
      page.offsetHeight <= sheet + 1 && last.scrollWidth <= last.clientWidth + 1 && last.scrollHeight <= last.clientHeight + 1;
    if (fits) filled += 1;
    else {
      last.style.height = before.height;
      last.style.columnFill = before.fill;
    }
  }
  return filled;
}

/** Aviso cuando el Word tiene división automática de palabras (Pliego la deja apagada). */
export const HYPHENATION_WARNING =
  "El Word tiene activada la división automática de palabras. Pliego no puede cortar las palabras igual que Word, así que acá no se cortan solas: algunos renglones y páginas pueden quedar distintos. Si necesitás cortes, ponelos en el Word con guiones opcionales (Ctrl + guion).";
