import JSZip from "jszip";
import {
  DOMParser,
  XMLSerializer,
  type Document as XmlDocument,
  type Element as XmlElement,
  type Node as XmlNode,
} from "@xmldom/xmldom";
import { FOOTNOTE_MARK, MARK_CLOSE, MARK_OPEN } from "./fields.ts";
import { EMPTY_SPACING, prepareLineSpacing, type SpacingInfo } from "./line-spacing.ts";

/**
 * Ajustes del .docx que solo sirven para componer (la copia que ve el visor, nunca el Word que se descarga) y datos del
 * diseño de página que el visor no aplica y que se completan después de dibujar.
 */
export type SectionLayout = {
  /** Cómo empieza la sección (w:type del sectPr). Ausente = página siguiente. */
  type?: "nextPage" | "oddPage" | "evenPage" | "continuous" | "nextColumn";
  /** Margen de encuadernación (w:gutter), en puntos. En la copia para el visor ya va sumado al margen. */
  gutterPt?: number;
  /** Márgenes izquierdo y derecho del Word, en puntos. */
  leftPt?: number;
  rightPt?: number;
  /** Primera página distinta (w:titlePg): la primera página de la sección lleva su propio encabezado y pie. */
  titlePage?: boolean;
};

export type RenderInfo = {
  /** Márgenes simétricos (settings.xml w:mirrorMargins): en las páginas pares se invierten izquierda y derecha. */
  mirrorMargins: boolean;
  /** Encuadernación arriba (w:gutterAtTop). */
  gutterAtTop: boolean;
  /** División automática de palabras activada en el Word (w:autoHyphenation). */
  autoHyphenation: boolean;
  /** Numeración de notas al pie (w:footnotePr en settings o en el sectPr final). */
  footnoteRestart: "continuous" | "eachSect" | "eachPage";
  footnoteStart: number;
  /** Una entrada por sección, en el mismo orden que SectionNumbering de fields.ts. */
  sections: SectionLayout[];
  /** Interlineado y «sin espacio entre párrafos del mismo estilo», para la composición (line-spacing.ts). */
  spacing: SpacingInfo;
  /** Modo de compatibilidad del Word (w:compatSetting compatibilityMode): 15 = Word 2013 o posterior; 0 = no lo dice. */
  compatibilityMode: number;
  /** «Suprimir espacio anterior después de un salto de página manual» (w:suppressSpBfAfterPgBrk). */
  suppressSpBfAfterPgBrk: boolean;
};

export const DEFAULT_RENDER_INFO: RenderInfo = {
  mirrorMargins: false,
  gutterAtTop: false,
  autoHyphenation: false,
  footnoteRestart: "continuous",
  footnoteStart: 1,
  sections: [],
  spacing: EMPTY_SPACING,
  compatibilityMode: 0,
  suppressSpBfAfterPgBrk: false,
};

/**
 * Como Word: después de un salto de página manual el primer párrafo de la página pierde el espacio anterior desde
 * Word 2013 (modo de compatibilidad 15), o antes si el Word tiene activada esa opción.
 */
export function suppressesSpaceAfterManualBreak(info: RenderInfo): boolean {
  return info.suppressSpBfAfterPgBrk || info.compatibilityMode >= 15;
}

/**
 * Va justo antes de la llamada de una nota con símbolo propio (w:customMarkFollows): el visor igual le pone número, y
 * el símbolo de verdad es el texto que sigue. Queda al final del mismo texto que la marca de nota de fields.ts.
 */
export const FOOTNOTE_CUSTOM_MARK = `${MARK_OPEN}C${MARK_CLOSE}`;
/** Justo después de un salto de página manual que corta un párrafo: lo que sigue es la continuación del párrafo. */
export const BREAK_MARK = `${MARK_OPEN}B${MARK_CLOSE}`;
/** En la primera celda de cada fila de título de una tabla (w:tblHeader): se repite si la tabla sigue en otra página. */
export const HEADER_ROW_MARK = `${MARK_OPEN}H${MARK_CLOSE}`;
/**
 * Párrafo de muestra: cada sección termina con una o dos páginas de muestra (una impar y una par) que solo sirven para
 * que el visor dibuje el encabezado y el pie comunes y de página par de la sección. Se sacan antes de paginar.
 */
export const PROBE_MARK = `${MARK_OPEN}Q${MARK_CLOSE}`;
/**
 * Al final de un párrafo vacío, en un tramo con la letra de su marca de párrafo (w:pPr/w:rPr): en Word un párrafo sin
 * texto tiene el alto de esa letra, no la del estilo. Se lee y se saca apenas se dibuja (tab-stops.ts).
 */
export const EMPTY_PARAGRAPH_MARK = `${MARK_OPEN}E${MARK_CLOSE}`;
/**
 * Al principio de un párrafo con tabulaciones alineadas a la derecha, al centro o decimales: sus tabulaciones, en
 * puntos desde el borde de la columna («T» + «pos:alineación:relleno;…»). Se lee y se saca apenas se dibuja.
 */
export const TABS_MARK_OPEN = `${MARK_OPEN}T`;
/**
 * Un párrafo de muestra al final del documento, copia del separador de notas al pie del Word (footnotes.xml): Word le
 * reserva su alto entre el texto y las notas de cada página, y dibuja ahí la raya. Se mide después del interlineado y se
 * saca antes de paginar.
 */
export const NOTE_SEPARATOR_MARK = `${MARK_OPEN}Z${MARK_CLOSE}`;
/**
 * En lugar del número de cada nota al pie dentro de su propio texto (w:footnoteRef): el visor lo deja vacío y la nota
 * salía sin su número. Va seguido de un «0» que ocupa el lugar del número hasta que se numeran las notas.
 */
export const NOTE_NUMBER_MARK = `${MARK_OPEN}R${MARK_CLOSE}`;

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const O = "urn:schemas-microsoft-com:office:office";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const PKG_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types";

function localName(node: XmlElement): string {
  const raw = node.localName || node.nodeName || "";
  const i = raw.indexOf(":");
  return i >= 0 ? raw.slice(i + 1) : raw;
}

function children(el: XmlElement, name?: string): XmlElement[] {
  return Array.from(el.childNodes).filter(
    (node): node is XmlElement => node.nodeType === 1 && (!name || localName(node as XmlElement) === name),
  );
}

function descendants(el: XmlElement | XmlDocument, name: string): XmlElement[] {
  return Array.from(el.getElementsByTagNameNS(W, name)) as XmlElement[];
}

/** Elementos de cualquier espacio de nombres con ese nombre local (VML, DrawingML, mc:…). */
function anyDescendants(el: XmlElement | XmlDocument, name: string): XmlElement[] {
  return (Array.from(el.getElementsByTagName("*")) as XmlElement[]).filter((node) => localName(node) === name);
}

function attr(el: XmlElement | undefined, name: string): string {
  return el ? el.getAttributeNS(W, name) || el.getAttribute(`w:${name}`) || "" : "";
}

/** Interruptores de Word: el elemento solo, o con w:val="1/true/on", vale sí. */
function isOn(el: XmlElement | undefined): boolean {
  if (!el) return false;
  return !["0", "false", "off"].includes(attr(el, "val").toLowerCase());
}

function isOnValue(value: string): boolean {
  return ["1", "true", "on"].includes(value.toLowerCase());
}

function ancestor(el: XmlNode, names: string[]): XmlElement | null {
  for (let node: XmlNode | null = el.parentNode; node; node = node.parentNode) {
    if (node.nodeType === 1 && names.includes(localName(node as XmlElement))) return node as XmlElement;
  }
  return null;
}

function contains(outer: XmlNode, inner: XmlNode): boolean {
  for (let node: XmlNode | null = inner; node; node = node.parentNode) if (node === outer) return true;
  return false;
}

/** Lo que va en mc:Choice el visor nunca lo dibuja (usa siempre mc:Fallback). */
function inChoice(el: XmlElement): boolean {
  return ancestor(el, ["Choice"]) !== null;
}

function newText(doc: XmlDocument, text: string): XmlElement {
  const t = doc.createElementNS(W, "w:t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  return t;
}

function make(doc: XmlDocument, ns: string, name: string, attrs: Record<string, string> = {}, kids: XmlElement[] = []) {
  const el = doc.createElementNS(ns, name);
  for (const [key, value] of Object.entries(attrs)) el.setAttribute(key, value);
  for (const kid of kids) el.appendChild(kid);
  return el;
}

function parse(xml: string): XmlDocument | null {
  try {
    const doc = new DOMParser({ onError: () => undefined }).parseFromString(xml.replace(/^\uFEFF/, ""), "text/xml");
    return doc.getElementsByTagName("parsererror").length ? null : doc;
  } catch {
    return null;
  }
}

function serialize(doc: XmlDocument): string {
  const out = new XMLSerializer().serializeToString(doc);
  return out.startsWith("<?xml") ? out : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
}

async function readXml(zip: JSZip, name: string): Promise<XmlDocument | null> {
  const xml = await zip.file(name)?.async("string");
  return xml == null ? null : parse(xml);
}

// ——— Secciones (mismo orden que markSections de fields.ts) ———

function bodyBlocks(body: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  for (const child of children(body)) {
    const name = localName(child);
    if (name === "sdt") out.push(...bodyBlocks(children(child, "sdtContent")[0] ?? child));
    else if (name === "customXml") out.push(...bodyBlocks(child));
    else if (name === "p" || name === "tbl") out.push(child);
  }
  return out;
}

/** El sectPr de cada sección del cuerpo; undefined si el documento termina sin sectPr final. */
function bodySections(doc: XmlDocument): Array<XmlElement | undefined> {
  const body = descendants(doc, "body")[0];
  if (!body) return [];
  const out: Array<XmlElement | undefined> = [];
  let opened = false;
  for (const block of bodyBlocks(body)) {
    if (!opened && (localName(block) === "p" || descendants(block, "p").length)) opened = true;
    const props = localName(block) === "p" ? children(block, "pPr")[0] : undefined;
    const sectPr = props ? children(props, "sectPr")[0] : undefined;
    if (sectPr) {
      out.push(sectPr);
      opened = false;
    }
  }
  const last = children(body, "sectPr")[0];
  if (opened || last) out.push(last);
  return out;
}

const SECTION_TYPES = ["nextPage", "oddPage", "evenPage", "continuous", "nextColumn"] as const;

function twipsToPt(value: string): number | undefined {
  if (!/^-?\d+(\.\d+)?$/.test(value.trim())) return undefined;
  return Number.parseFloat(value) / 20;
}

function readLayout(sectPr: XmlElement | undefined): SectionLayout {
  if (!sectPr) return {};
  const layout: SectionLayout = {};
  const type = attr(children(sectPr, "type")[0], "val");
  const known = SECTION_TYPES.find((item) => item === type);
  if (known) layout.type = known;
  const margins = children(sectPr, "pgMar")[0];
  if (margins) {
    const gutter = twipsToPt(attr(margins, "gutter"));
    const left = twipsToPt(attr(margins, "left") || attr(margins, "start"));
    const right = twipsToPt(attr(margins, "right") || attr(margins, "end"));
    if (gutter !== undefined) layout.gutterPt = gutter;
    if (left !== undefined) layout.leftPt = left;
    if (right !== undefined) layout.rightPt = right;
  }
  if (isOn(children(sectPr, "titlePg")[0])) layout.titlePage = true;
  return layout;
}

/**
 * El visor lee el margen de encuadernación pero no lo aplica. Se suma en la copia al margen interior (el izquierdo; con
 * márgenes simétricos el izquierdo es el interior) o al de arriba si el Word lo pide, así las páginas se llenan desde
 * el principio con el ancho final. Después solo se invierten izquierda y derecha en las páginas pares, sin reacomodar.
 */
function bakeGutter(sectPr: XmlElement, info: RenderInfo): boolean {
  const margins = children(sectPr, "pgMar")[0];
  if (!margins) return false;
  const gutter = twipsToPt(attr(margins, "gutter"));
  if (gutter === undefined || gutter <= 0) return false;
  const side = !info.mirrorMargins && info.gutterAtTop ? "top" : !attr(margins, "left") && attr(margins, "start") ? "start" : "left";
  const raw = attr(margins, side);
  const base = raw === "" ? 0 : twipsToPt(raw);
  if (base === undefined) return false;
  // Un margen de arriba negativo es fijo (el texto no se corre por el encabezado): se agranda igual, con su signo.
  const value = (base < 0 ? base - gutter : base + gutter) * 20;
  margins.setAttributeNS(W, `w:${side}`, String(Math.round(value)));
  margins.setAttributeNS(W, "w:gutter", "0");
  return true;
}

function readFootnotePr(props: XmlElement | undefined, info: RenderInfo) {
  if (!props) return;
  const restart = attr(children(props, "numRestart")[0], "val");
  if (restart === "continuous" || restart === "eachSect" || restart === "eachPage") info.footnoteRestart = restart;
  const start = Number.parseInt(attr(children(props, "numStart")[0], "val"), 10);
  if (Number.isFinite(start)) info.footnoteStart = start;
}

// ——— Encabezados y pies «vinculados al anterior» ———

const REF_KINDS = ["headerReference", "footerReference"] as const;
type RefKind = (typeof REF_KINDS)[number];
type PartKind = "header" | "footer";
const REF_TYPES = ["default", "first", "even"] as const;

/** Las referencias van primero en el sectPr: la nueva entra después de las que ya hay. */
function insertReference(sectPr: XmlElement, ref: XmlElement) {
  const after = children(sectPr).find((child) => !(REF_KINDS as readonly string[]).includes(localName(child)));
  if (after) sectPr.insertBefore(ref, after);
  else sectPr.appendChild(ref);
}

type BlankIds = Partial<Record<PartKind, string>>;

/**
 * En Word, una sección sin encabezado (o pie) de algún tipo usa el de la sección anterior; el visor no. Se copia la
 * referencia de la sección anterior. Si ninguna anterior lo tiene, Word deja ese encabezado en blanco (primera página
 * distinta, o pares e impares distintos) y el visor pondría el común: se apunta a uno vacío.
 */
function linkHeaders(
  doc: XmlDocument,
  sections: Array<XmlElement | undefined>,
  evenAndOdd: boolean,
  blank: BlankIds,
): { inherited: number; blanks: Set<PartKind> } {
  const previous: Record<RefKind, Map<string, XmlElement>> = { headerReference: new Map(), footerReference: new Map() };
  let inherited = 0;
  const blanks = new Set<PartKind>();
  for (const sectPr of sections) {
    if (!sectPr) continue;
    const titlePage = isOn(children(sectPr, "titlePg")[0]);
    for (const kind of REF_KINDS) {
      const own = new Map<string, XmlElement>();
      for (const ref of children(sectPr, kind)) own.set(attr(ref, "type") || "default", ref);
      for (const type of REF_TYPES) {
        const before = previous[kind].get(type);
        if (own.has(type) || !before) continue;
        const copy = before.cloneNode(true) as XmlElement;
        insertReference(sectPr, copy);
        own.set(type, copy);
        inherited += 1;
      }
      const part: PartKind = kind === "headerReference" ? "header" : "footer";
      const blankId = blank[part];
      if (blankId && own.size) {
        const wanted = [titlePage ? "first" : "", evenAndOdd ? "even" : ""].filter(Boolean);
        for (const type of wanted) {
          if (own.has(type)) continue;
          const ref = doc.createElementNS(W, `w:${kind}`);
          ref.setAttributeNS(W, "w:type", type);
          ref.setAttributeNS(R, "r:id", blankId);
          insertReference(sectPr, ref);
          own.set(type, ref);
          blanks.add(part);
        }
      }
      previous[kind] = own;
    }
  }
  return { inherited, blanks };
}

const BLANK_TARGETS: Record<PartKind, string> = { header: "pliego-blank-header.xml", footer: "pliego-blank-footer.xml" };

/** Ids libres en las relaciones del documento para el encabezado y el pie vacíos (sin tocar nada todavía). */
function blankIds(rels: XmlDocument | null): BlankIds {
  if (!rels?.documentElement) return {};
  const used = new Set(
    (Array.from(rels.getElementsByTagName("Relationship")) as XmlElement[]).map((rel) => rel.getAttribute("Id") || ""),
  );
  const free = (base: string) => {
    let id = base;
    for (let n = 2; used.has(id); n += 1) id = `${base}${n}`;
    used.add(id);
    return id;
  };
  return { header: free("rIdPliegoBlankHeader"), footer: free("rIdPliegoBlankFooter") };
}

function addBlankParts(zip: JSZip, rels: XmlDocument, types: XmlDocument | null, ids: BlankIds, which: Set<PartKind>) {
  const root = rels.documentElement;
  if (!root) return;
  for (const part of which) {
    const id = ids[part];
    if (!id) continue;
    const tag = part === "header" ? "w:hdr" : "w:ftr";
    zip.file(
      `word/${BLANK_TARGETS[part]}`,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${tag} xmlns:w="${W}"><w:p/></${tag}>`,
    );
    const rel = rels.createElementNS(PKG_RELS, "Relationship");
    rel.setAttribute("Id", id);
    rel.setAttribute("Type", `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${part}`);
    rel.setAttribute("Target", BLANK_TARGETS[part]);
    root.appendChild(rel);
    const typesRoot = types?.documentElement;
    if (types && typesRoot) {
      const override = types.createElementNS(CONTENT_TYPES, "Override");
      override.setAttribute("PartName", `/word/${BLANK_TARGETS[part]}`);
      override.setAttribute("ContentType", `application/vnd.openxmlformats-officedocument.wordprocessingml.${part}+xml`);
      typesRoot.appendChild(override);
    }
  }
}

// ——— Guiones y notas con símbolo propio ———

/**
 * El visor dibuja el guion de no separación como un punto de corte (no se ve) y el guion opcional no lo conoce. El de
 * no separación pasa a un guion común entre dos «unión de palabras» (U+2060): se dibuja con el guion de la misma letra
 * (U+2011 falta en muchas fuentes y en los subconjuntos que incrusta Word) y el renglón igual no se corta ahí.
 */
function fixHyphens(doc: XmlDocument): number {
  let count = 0;
  for (const [name, text] of [["noBreakHyphen", "\u2060-\u2060"], ["softHyphen", "\u00AD"]] as const) {
    for (const node of descendants(doc, name)) {
      node.parentNode?.replaceChild(newText(doc, text), node);
      count += 1;
    }
  }
  return count;
}

/**
 * Los estilos de carácter «Hyperlink» del Word (el id cambia según el idioma del Word: se buscan por nombre, que no
 * cambia). Siempre incluye «Hyperlink», el id que usa Word en inglés y en la mayoría de los documentos.
 */
function hyperlinkStyleIds(styles: XmlDocument | null): Set<string> {
  const ids = new Set(["Hyperlink"]);
  for (const style of styles ? descendants(styles, "style") : []) {
    if (attr(style, "type") !== "character") continue;
    if (attr(children(style, "name")[0], "val").toLowerCase() === "hyperlink") ids.add(attr(style, "styleId"));
  }
  return ids;
}

/** Un campo del índice (TOC) o un enlace de campo a una entrada del índice (HYPERLINK \l "_Toc…"). */
const TOC_FIELD = /^\s*(?:TOC\b|HYPERLINK\b[\s\S]*\\l\s+"?_Toc)/i;

/**
 * Las entradas de un índice (o de una tabla de ilustraciones) que arma Word llevan el estilo «Hyperlink», que es azul
 * y subrayado. Word igual las muestra con la letra del índice, sin azul ni subrayado (así salen en el PDF que hace
 * Word); el visor les pondría el azul. Word las escribe de dos maneras: un enlace a «_Toc…» (w:hyperlink) o un campo
 * HYPERLINK \l "_Toc…", y las dos van dentro del campo TOC. Se les quita ese estilo solo ahí: un enlace a una página
 * web en el texto sigue como en el Word.
 */
function plainTocLinks(doc: XmlDocument, hyperlinkIds: Set<string>): number {
  const plain = new Set<XmlElement>();
  for (const link of descendants(doc, "hyperlink")) {
    if (attr(link, "anchor").startsWith("_Toc")) for (const run of descendants(link, "r")) plain.add(run);
  }
  for (const field of descendants(doc, "fldSimple")) {
    if (TOC_FIELD.test(attr(field, "instr"))) for (const run of descendants(field, "r")) plain.add(run);
  }
  // Campos complejos (begin … instrText … separate … end), que pueden ocupar varios párrafos y anidarse: cada tramo
  // se anota con los campos abiertos, y al final (cuando ya se conoce el texto de cada campo) se ve cuáles son del índice.
  type Field = { instr: string };
  const open: Field[] = [];
  const runFields: Array<[XmlElement, Field[]]> = [];
  for (const run of descendants(doc, "r")) {
    let inside = [...open];
    for (const child of children(run)) {
      const name = localName(child);
      if (name === "fldChar") {
        const type = attr(child, "fldCharType");
        if (type === "begin") {
          open.push({ instr: "" });
          inside = [...open];
        } else if (type === "end") open.pop();
      } else if (name === "instrText" && open.length) {
        open[open.length - 1]!.instr += child.textContent ?? "";
      }
    }
    if (inside.length) runFields.push([run, inside]);
  }
  for (const [run, fields] of runFields) if (fields.some((field) => TOC_FIELD.test(field.instr))) plain.add(run);

  let count = 0;
  for (const run of plain) {
    for (const props of children(run, "rPr")) {
      for (const style of children(props, "rStyle")) {
        if (!hyperlinkIds.has(attr(style, "val"))) continue;
        props.removeChild(style);
        count += 1;
      }
    }
  }
  return count;
}

// ——— Nombres de letra que el navegador leería como palabras de CSS ———

/**
 * Nombres de letra que en CSS son palabras reservadas o familias genéricas: un estilo pegado de una página web puede
 * decir que su letra es «inherit». Word no tiene esa letra y usa la que su tabla de fuentes da como alternativa
 * (w:altName); el navegador, en cambio, lo lee como «heredar» y termina en otra letra (la configuración general).
 */
const CSS_FONT_WORDS = new Set([
  "inherit",
  "initial",
  "unset",
  "revert",
  "revert-layer",
  "default",
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
  "math",
  "emoji",
  "fangsong",
]);

/** Para cada nombre de esos, la letra que usaría Word: la alternativa de la tabla de fuentes, o una según su familia. */
function keywordFontReplacements(fontTable: XmlDocument | null): Map<string, string> {
  const out = new Map<string, string>();
  for (const font of fontTable ? descendants(fontTable, "font") : []) {
    const name = attr(font, "name");
    if (!CSS_FONT_WORDS.has(name.trim().toLowerCase())) continue;
    const alt = attr(children(font, "altName")[0], "val").trim();
    const family = attr(children(font, "family")[0], "val");
    const fallback = family === "swiss" ? "Arial" : family === "modern" ? "Courier New" : "Times New Roman";
    out.set(name.trim().toLowerCase(), alt && !CSS_FONT_WORDS.has(alt.toLowerCase()) ? alt : fallback);
  }
  return out;
}

const FONT_ATTRIBUTES = ["ascii", "hAnsi", "eastAsia", "cs"];

/** Cambia esos nombres en los w:rFonts de una parte ya leída. Devuelve cuántos cambió. */
function replaceKeywordFonts(doc: XmlDocument, replacements: Map<string, string>): number {
  let count = 0;
  for (const fonts of descendants(doc, "rFonts")) {
    for (const name of FONT_ATTRIBUTES) {
      const value = attr(fonts, name).trim().toLowerCase();
      if (!value || !CSS_FONT_WORDS.has(value)) continue;
      fonts.setAttributeNS(W, `w:${name}`, replacements.get(value) ?? "Times New Roman");
      count += 1;
    }
  }
  return count;
}

// ——— Párrafos vacíos y tabulaciones: datos para después de dibujar ———

type ParagraphStyle = { basedOn: string; tabs: XmlElement[] };
type ParagraphStyles = { byId: Map<string, ParagraphStyle>; defaultId: string };

function readParagraphStyles(styles: XmlDocument | null): ParagraphStyles {
  const byId = new Map<string, ParagraphStyle>();
  let defaultId = "";
  for (const style of styles ? descendants(styles, "style") : []) {
    if (attr(style, "type") !== "paragraph") continue;
    const id = attr(style, "styleId");
    if (!id) continue;
    if (isOnValue(attr(style, "default"))) defaultId = id;
    const pPr = children(style, "pPr")[0];
    const tabs = pPr ? children(pPr, "tabs")[0] : undefined;
    byId.set(id, { basedOn: attr(children(style, "basedOn")[0], "val"), tabs: tabs ? children(tabs, "tab") : [] });
  }
  return { byId, defaultId };
}

export type TabStopMark = { pos: number; align: string; leader: string };

/** Las tabulaciones de un párrafo: las de su estilo (y los estilos en que se basa) y después las propias. */
function paragraphTabStops(paragraph: XmlElement, styles: ParagraphStyles): TabStopMark[] {
  const pPr = children(paragraph, "pPr")[0];
  const chain: XmlElement[][] = [];
  let id = attr(children(pPr ?? paragraph, "pStyle")[0], "val") || styles.defaultId;
  const seen = new Set<string>();
  while (id && !seen.has(id)) {
    seen.add(id);
    const style = styles.byId.get(id);
    if (!style) break;
    chain.unshift(style.tabs);
    id = style.basedOn;
  }
  const own = pPr ? children(pPr, "tabs")[0] : undefined;
  if (own) chain.push(children(own, "tab"));
  const stops = new Map<number, TabStopMark>();
  for (const tabs of chain) {
    for (const tab of tabs) {
      const twips = Number.parseInt(attr(tab, "pos"), 10);
      if (!Number.isFinite(twips)) continue;
      const align = attr(tab, "val") || "left";
      if (align === "clear") stops.delete(twips);
      else stops.set(twips, { pos: twips / 20, align, leader: attr(tab, "leader") || "none" });
    }
  }
  return [...stops.entries()].sort((a, b) => a[0] - b[0]).map(([, stop]) => stop);
}

/** Lo que ocupa lugar en un renglón: si un párrafo no tiene nada de esto, es un párrafo vacío. */
const LINE_CONTENT = new Set([
  "tab",
  "ptab",
  "br",
  "cr",
  "drawing",
  "pict",
  "object",
  "sym",
  "noBreakHyphen",
  "softHyphen",
  "footnoteReference",
  "endnoteReference",
  "separator",
  "continuationSeparator",
]);

function hasLineContent(paragraph: XmlElement): boolean {
  for (const run of descendants(paragraph, "r")) {
    for (const child of children(run)) {
      const name = localName(child);
      if (name === "t" ? Boolean(child.textContent) : LINE_CONTENT.has(name)) return true;
    }
  }
  return anyDescendants(paragraph, "drawing").length > 0 || anyDescendants(paragraph, "pict").length > 0;
}

/** Un tramo nuevo al final o al principio del párrafo, con ese texto y una copia de esas propiedades de letra. */
function markRun(doc: XmlDocument, text: string, props?: XmlElement): XmlElement {
  const run = make(doc, W, "w:r");
  if (props) {
    const copy = props.cloneNode(true) as XmlElement;
    // Las marcas de revisión de la marca de párrafo no son letra.
    for (const child of children(copy)) {
      if (["ins", "del", "moveFrom", "moveTo", "rPrChange"].includes(localName(child))) copy.removeChild(child);
    }
    run.appendChild(copy);
  }
  run.appendChild(newText(doc, text));
  return run;
}

/**
 * Deja dos datos para después de dibujar: en cada párrafo vacío, la letra de su marca de párrafo; en cada párrafo con
 * tabulaciones a la derecha, al centro o decimales, sus tabulaciones (el visor las calcula mal cuando lo que sigue no
 * entra en el renglón). Devuelve cuántos párrafos marcó.
 */
function markParagraphs(doc: XmlDocument, styles: ParagraphStyles): number {
  let count = 0;
  for (const paragraph of descendants(doc, "p")) {
    if (inChoice(paragraph)) continue;
    const pPr = children(paragraph, "pPr")[0];
    if (!hasLineContent(paragraph)) {
      const props = pPr ? children(pPr, "rPr")[0] : undefined;
      if (props) {
        paragraph.appendChild(markRun(doc, EMPTY_PARAGRAPH_MARK, props));
        count += 1;
      }
      continue;
    }
    const usesTab = descendants(paragraph, "tab").some((tab) => tab.parentNode && localName(tab.parentNode as XmlElement) === "r");
    if (!usesTab) continue;
    const stops = paragraphTabStops(paragraph, styles);
    if (!stops.some((stop) => stop.align === "right" || stop.align === "end" || stop.align === "center" || stop.align === "decimal")) continue;
    const payload = stops.map((stop) => `${Math.round(stop.pos * 100) / 100}:${stop.align}:${stop.leader}`).join(";");
    const run = markRun(doc, `${TABS_MARK_OPEN}${payload}${MARK_CLOSE}`);
    const first = children(paragraph).find((child) => localName(child) !== "pPr");
    if (first) paragraph.insertBefore(run, first);
    else paragraph.appendChild(run);
    count += 1;
  }
  return count;
}

/**
 * Agrega al final del documento la muestra del separador de notas (NOTE_SEPARATOR_MARK) si el libro tiene notas al pie:
 * las propiedades del párrafo del separador del Word y un tramo con la marca (con la letra de ese párrafo, como el
 * separador). Devuelve si la agregó.
 */
function addNoteSeparatorProbe(doc: XmlDocument, footnotes: XmlDocument | null): boolean {
  if (!footnotes || !descendants(doc, "footnoteReference").length) return false;
  const separator = descendants(footnotes, "footnote").find((note) => attr(note, "type") === "separator");
  const source = separator ? descendants(separator, "p")[0] : undefined;
  const body = descendants(doc, "body")[0];
  if (!source || !body) return false;
  const paragraph = make(doc, W, "w:p");
  const pPr = children(source, "pPr")[0];
  if (pPr) {
    const copy = pPr.cloneNode(true) as XmlElement;
    for (const child of children(copy)) if (["sectPr", "pPrChange"].includes(localName(child))) copy.removeChild(child);
    paragraph.appendChild(copy);
  }
  // Con la letra del tramo del separador (o la de la marca de párrafo): es la que mide el renglón en Word.
  const separatorRun = descendants(source, "r").find((run) => children(run, "separator").length > 0);
  const props = (separatorRun && children(separatorRun, "rPr")[0]) ?? (pPr ? children(pPr, "rPr")[0] : undefined);
  paragraph.appendChild(markRun(doc, NOTE_SEPARATOR_MARK, props));
  const last = children(body, "sectPr")[0];
  body.insertBefore(paragraph, last ?? null);
  return true;
}

/**
 * Los estilos de carácter que suben o bajan el texto (w:vertAlign: superíndice o subíndice), con lo heredado de los
 * estilos en que se basan. El visor no lee esa propiedad en los estilos, solo escrita en el texto.
 */
function vertAlignStyles(styles: XmlDocument | null): Map<string, string> {
  const own = new Map<string, { basedOn: string; value: string }>();
  for (const style of styles ? descendants(styles, "style") : []) {
    if (attr(style, "type") !== "character") continue;
    const rPr = children(style, "rPr")[0];
    own.set(attr(style, "styleId"), {
      basedOn: attr(children(style, "basedOn")[0], "val"),
      value: rPr ? attr(children(rPr, "vertAlign")[0], "val") : "",
    });
  }
  const out = new Map<string, string>();
  for (const id of own.keys()) {
    const seen = new Set<string>();
    for (let current: string | undefined = id; current && !seen.has(current); current = own.get(current)?.basedOn) {
      seen.add(current);
      const value = own.get(current)?.value;
      if (value) {
        if (value !== "baseline") out.set(id, value);
        break;
      }
    }
  }
  return out;
}

/**
 * A los tramos con un estilo de superíndice o subíndice les escribe esa propiedad en el texto, para que el visor la
 * dibuje (el número de cada nota en su propio texto, «m²» con un estilo…). Las llamadas a notas no: el visor ya las sube.
 */
function applyStyleVertAlign(doc: XmlDocument, byStyle: Map<string, string>): number {
  if (!byStyle.size) return 0;
  let count = 0;
  for (const run of descendants(doc, "r")) {
    const rPr = children(run, "rPr")[0];
    const value = rPr ? byStyle.get(attr(children(rPr, "rStyle")[0], "val")) : undefined;
    if (!rPr || !value || children(rPr, "vertAlign").length) continue;
    if (children(run).some((child) => ["footnoteReference", "endnoteReference"].includes(localName(child)))) continue;
    rPr.appendChild(make(doc, W, "w:vertAlign", { "w:val": value }));
    count += 1;
  }
  return count;
}

/** Cambia cada w:footnoteRef (el número de la nota en su texto) por la marca y un «0» provisorio, en su mismo tramo. */
function markNoteNumbers(doc: XmlDocument): number {
  // El provisorio tiene tantas cifras como el número más alto posible: uno más angosto podía dejar un renglón de nota
  // que, con el número de verdad, ya no entra.
  const notes = descendants(doc, "footnote").filter((note) => !["separator", "continuationSeparator", "continuationNotice"].includes(attr(note, "type")));
  const placeholder = "0".repeat(String(Math.max(1, notes.length + 9)).length);
  let count = 0;
  for (const ref of descendants(doc, "footnoteRef")) {
    ref.parentNode?.replaceChild(newText(doc, `${NOTE_NUMBER_MARK}${placeholder}`), ref);
    count += 1;
  }
  return count;
}

// ——— Marcas para la paginación: saltos manuales, filas de título y páginas de muestra ———

/** Lo que se ve en un renglón (para saber si un salto de página manual deja algo después en el mismo párrafo). */
const VISIBLE = new Set(["tab", "ptab", "drawing", "pict", "object", "sym", "noBreakHyphen", "footnoteReference", "endnoteReference"]);

function isVisible(node: XmlElement): boolean {
  const name = localName(node);
  if (name === "t") return (node.textContent ?? "") !== "";
  return VISIBLE.has(name);
}

/**
 * Un salto de página manual a mitad de párrafo: el visor parte el párrafo y la segunda parte empieza la página con el
 * espacio anterior del párrafo, que en Word no tiene (no es el comienzo del párrafo). Se marca lo que sigue al salto.
 */
function markManualBreaks(body: XmlElement): number {
  let count = 0;
  for (const paragraph of bodyBlocks(body)) {
    if (localName(paragraph) !== "p") continue;
    const all = Array.from(paragraph.getElementsByTagName("*")) as XmlElement[];
    all.forEach((node, index) => {
      if (localName(node) !== "br" || attr(node, "type") !== "page" || inChoice(node)) return;
      if (!all.slice(index + 1).some(isVisible)) return;
      const doc = node.ownerDocument;
      if (!doc) return;
      node.parentNode?.insertBefore(newText(doc, BREAK_MARK), node.nextSibling);
      count += 1;
    });
  }
  return count;
}

/** Las filas de título del principio de cada tabla (w:tblHeader) llevan una marca en su primera celda. */
function markHeaderRows(doc: XmlDocument): number {
  let count = 0;
  for (const table of descendants(doc, "tbl")) {
    if (inChoice(table)) continue;
    for (const row of children(table, "tr")) {
      const props = children(row, "trPr")[0];
      if (!props || !isOn(children(props, "tblHeader")[0])) break;
      const cell = children(row, "tc")[0];
      const paragraph = cell ? children(cell, "p")[0] : undefined;
      if (!paragraph) continue;
      const run = doc.createElementNS(W, "w:r");
      run.appendChild(newText(doc, HEADER_ROW_MARK));
      const pPr = children(paragraph, "pPr")[0];
      paragraph.insertBefore(run, pPr ? pPr.nextSibling : paragraph.firstChild);
      count += 1;
    }
  }
  return count;
}

function probeParagraph(doc: XmlDocument): XmlElement {
  const style = make(doc, W, "w:pStyle", { "w:val": "PliegoMuestra" });
  const pPr = make(doc, W, "w:pPr", {}, [style, make(doc, W, "w:pageBreakBefore")]);
  const run = make(doc, W, "w:r", {}, [newText(doc, PROBE_MARK)]);
  return make(doc, W, "w:p", {}, [pPr, run]);
}

/**
 * Al final de cada sección, una página de muestra (dos si hay encabezados distintos en páginas pares e impares, para
 * que salga uno de cada uno). El visor arma una sola página por sección, con el encabezado de primera página si la
 * sección lo tiene; de las muestras salen el encabezado y el pie de las páginas siguientes. La última de todas recibe
 * además las notas finales. El sectPr pasa a la última muestra para que sigan dentro de su sección.
 */
function addProbes(doc: XmlDocument, sections: Array<XmlElement | undefined>, evenAndOdd: boolean): number {
  const body = descendants(doc, "body")[0];
  if (!body) return 0;
  const count = evenAndOdd ? 2 : 1;
  let added = 0;
  for (const sectPr of sections) {
    const probes = Array.from({ length: count }, () => probeParagraph(doc));
    const props = sectPr?.parentNode as XmlElement | null | undefined;
    const paragraph = props?.parentNode as XmlElement | null | undefined;
    const parent = paragraph?.parentNode;
    if (sectPr && props && paragraph && parent && localName(props) === "pPr") {
      props.removeChild(sectPr);
      children(probes[probes.length - 1]!, "pPr")[0]!.appendChild(sectPr);
      const anchor = paragraph.nextSibling;
      for (const probe of probes) parent.insertBefore(probe, anchor);
    } else {
      // La última sección: antes del sectPr final del cuerpo.
      const last = children(body, "sectPr")[0] ?? null;
      for (const probe of probes) body.insertBefore(probe, last);
    }
    added += probes.length;
  }
  return added;
}

function markCustomFootnotes(doc: XmlDocument): number {
  let count = 0;
  for (const reference of descendants(doc, "footnoteReference")) {
    if (!isOnValue(attr(reference, "customMarkFollows"))) continue;
    const before = reference.previousSibling;
    const text = before && before.nodeType === 1 && localName(before as XmlElement) === "t" ? (before as XmlElement) : null;
    // En el mismo texto que la marca de nota: así las dos quedan justo antes del número que pone el visor.
    if (text && (text.textContent ?? "").endsWith(FOOTNOTE_MARK)) text.appendChild(doc.createTextNode(FOOTNOTE_CUSTOM_MARK));
    else reference.parentNode?.insertBefore(newText(doc, FOOTNOTE_CUSTOM_MARK), reference);
    count += 1;
  }
  return count;
}

// ——— Imágenes VML y avisos de dibujos ———

const EMU_PER: Record<string, number> = { pt: 12700, in: 914400, cm: 360000, mm: 36000, pc: 152400, px: 9525, "": 9525 };

function cssStyle(el: XmlElement): Map<string, string> {
  const out = new Map<string, string>();
  for (const item of (el.getAttribute("style") || "").split(";")) {
    const i = item.indexOf(":");
    if (i > 0) out.set(item.slice(0, i).trim().toLowerCase(), item.slice(i + 1).trim());
  }
  return out;
}

/** «120.5pt», «2in», «3cm» → EMU. Sin unidad, VML cuenta píxeles. */
export function lengthToEmu(value: string | undefined): number | null {
  const match = /^(\d*\.?\d+)\s*(pt|in|cm|mm|pc|px)?$/i.exec((value ?? "").trim());
  if (!match) return null;
  const emu = Math.round(Number.parseFloat(match[1]!) * EMU_PER[(match[2] ?? "").toLowerCase()]!);
  return emu > 0 ? emu : null;
}

/** Recorte de VML («6554f» = 6554/65536, o una fracción común) → milésimas de por ciento, como DrawingML. */
export function cropToPercent(value: string): number | null {
  const match = /^(-?\d*\.?\d+)(f?)$/.exec(value.trim());
  if (!match) return null;
  const fraction = Number.parseFloat(match[1]!) / (match[2] ? 65536 : 1);
  return Number.isFinite(fraction) ? Math.round(fraction * 100000) : null;
}

const VML_SHAPES = ["shape", "rect", "roundrect", "oval", "line", "polyline", "arc", "curve", "group", "image"];

type Found = {
  texts: string[];
  shapes: string[];
  charts: number;
  smartArt: number;
  objects: number;
  floating: number;
  behind: number;
  noSize: number;
  unknown: number;
};

const nothingFound = (): Found => ({ texts: [], shapes: [], charts: 0, smartArt: 0, objects: 0, floating: 0, behind: 0, noSize: 0, unknown: 0 });

function cleanText(text: string): string {
  const plain = text.replace(/[^]*/g, "").replace(/\s+/g, " ").trim();
  return plain.length > 40 ? `${plain.slice(0, 40).trimEnd()}…` : plain;
}

/** Texto de un párrafo sin el de los cuadros de texto que tenga adentro. */
function ownText(paragraph: XmlElement): string {
  const texts = descendants(paragraph, "t").filter((t) => {
    const box = ancestor(t, ["txbxContent"]);
    return !box || !contains(paragraph, box);
  });
  return cleanText(texts.map((t) => t.textContent ?? "").join(""));
}

/** El texto del párrafo donde está el dibujo, o el de alguno de los anteriores. */
function nearText(el: XmlElement): string {
  let node: XmlNode | null = ancestor(el, ["p"]);
  for (let tries = 0; node && tries < 4; node = node.previousSibling) {
    if (node.nodeType !== 1 || localName(node as XmlElement) !== "p") continue;
    tries += 1;
    const text = ownText(node as XmlElement);
    if (text) return text;
  }
  return "";
}

let pictureId = 0;

function vmlToDrawing(doc: XmlDocument, shape: XmlElement, image: XmlElement, relId: string, cx: number, cy: number): XmlElement {
  pictureId += 1;
  const id = String(900000 + pictureId);
  const blip = doc.createElementNS(A, "a:blip");
  blip.setAttributeNS(R, "r:embed", relId);
  const blipFill = make(doc, PIC, "pic:blipFill", {}, [blip]);
  const crop: Record<string, string> = {};
  for (const [side, key] of [["l", "cropleft"], ["t", "croptop"], ["r", "cropright"], ["b", "cropbottom"]] as const) {
    const value = cropToPercent(image.getAttribute(key) || "");
    if (value) crop[side] = String(value);
  }
  if (Object.keys(crop).length) blipFill.appendChild(make(doc, A, "a:srcRect", crop));
  blipFill.appendChild(make(doc, A, "a:stretch", {}, [make(doc, A, "a:fillRect")]));
  const extent = { cx: String(cx), cy: String(cy) };
  const pic = make(doc, PIC, "pic:pic", {}, [
    make(doc, PIC, "pic:nvPicPr", {}, [make(doc, PIC, "pic:cNvPr", { id, name: "" }), make(doc, PIC, "pic:cNvPicPr")]),
    blipFill,
    make(doc, PIC, "pic:spPr", {}, [
      make(doc, A, "a:xfrm", {}, [make(doc, A, "a:off", { x: "0", y: "0" }), make(doc, A, "a:ext", extent)]),
      make(doc, A, "a:prstGeom", { prst: "rect" }, [make(doc, A, "a:avLst")]),
    ]),
  ]);
  const graphic = make(doc, A, "a:graphic", {}, [make(doc, A, "a:graphicData", { uri: PIC }, [pic])]);
  const inline = make(doc, WP, "wp:inline", { distT: "0", distB: "0", distL: "0", distR: "0" }, [
    make(doc, WP, "wp:extent", extent),
    make(doc, WP, "wp:docPr", { id, name: `Imagen ${id}`, descr: shape.getAttribute("alt") || "" }),
    graphic,
  ]);
  return make(doc, W, "w:drawing", {}, [inline]);
}

/**
 * Las imágenes viejas de Word (VML: w:pict, o un objeto incrustado con su vista previa) el visor las dibuja como un
 * SVG de tamaño calculado a ojo, o no las dibuja. Se cambian por la misma imagen en el formato nuevo, en su renglón y
 * con la misma relación: el archivo de la imagen pasa sin tocar.
 */
function convertVml(doc: XmlDocument, found: Found): number {
  let converted = 0;
  for (const holder of [...descendants(doc, "pict"), ...descendants(doc, "object")]) {
    const parent = holder.parentNode;
    if (!parent || inChoice(holder)) continue;
    if (anyDescendants(holder, "txbxContent").length) continue; // el cuadro de texto tiene su propio aviso
    const shapes = children(holder).filter((child) => VML_SHAPES.includes(localName(child)));
    const shape = shapes.length === 1 ? shapes[0]! : undefined;
    const image = shape && ["shape", "rect", "image"].includes(localName(shape)) ? children(shape, "imagedata")[0] : undefined;
    const relId = image
      ? image.getAttributeNS(R, "id") || image.getAttribute("r:id") || image.getAttributeNS(O, "relid") || image.getAttribute("o:relid") || ""
      : "";
    if (!shape || !image || !relId) {
      if (localName(holder) === "object") found.objects += 1;
      else if (shapes.length) found.shapes.push(nearText(holder));
      continue;
    }
    const style = cssStyle(shape);
    const cx = lengthToEmu(style.get("width"));
    const cy = lengthToEmu(style.get("height"));
    if (!cx || !cy) {
      found.noSize += 1;
      continue;
    }
    // Detrás del texto (una marca de agua, un fondo): en el renglón empujaría todo el texto. Queda como estaba.
    if (style.get("position") === "absolute" && Number.parseFloat(style.get("z-index") ?? "0") < 0) {
      found.behind += 1;
      continue;
    }
    if (style.get("position") === "absolute") found.floating += 1;
    parent.replaceChild(vmlToDrawing(doc, shape, image, relId, cx, cy), holder);
    converted += 1;
  }
  return converted;
}

/** Cuadros de texto, formas, gráficos y SmartArt: el visor no los dibuja, o no siempre. */
function findDrawings(doc: XmlDocument, found: Found) {
  for (const box of anyDescendants(doc, "txbxContent")) {
    if (inChoice(box) || ancestor(box, ["txbxContent"])) continue;
    found.texts.push(cleanText(descendants(box, "t").map((t) => t.textContent ?? "").join("")));
  }
  for (const drawing of descendants(doc, "drawing")) {
    if (inChoice(drawing)) continue;
    for (const data of anyDescendants(drawing, "graphicData")) {
      for (const kind of children(data)) {
        const name = localName(kind);
        if (name === "pic") continue;
        if (name === "chart") found.charts += 1;
        else if (name === "relIds") found.smartArt += 1;
        else if (!(name === "wsp" && anyDescendants(kind, "txbxContent").length)) found.shapes.push(nearText(drawing));
      }
    }
  }
  for (const choice of anyDescendants(doc, "AlternateContent")) {
    if (inChoice(choice)) continue;
    const fallback = children(choice, "Fallback")[0];
    if (!fallback || !children(fallback).length) found.unknown += 1;
  }
}

const WHERE: Array<[RegExp, string]> = [
  [/^word\/document\.xml$/, "el texto del libro"],
  [/^word\/header\d*\.xml$/, "un encabezado"],
  [/^word\/footer\d*\.xml$/, "un pie de página"],
  [/^word\/footnotes\.xml$/, "las notas al pie"],
  [/^word\/endnotes\.xml$/, "las notas finales"],
];

function warningsFor(where: string, found: Found): string[] {
  const out: string[] = [];
  const texts = found.texts.length;
  const firstText = found.texts.find(Boolean) ?? "";
  if (texts === 1) out.push(`Hay un cuadro de texto en ${where}${firstText ? ` («${firstText}»)` : ""}: su texto puede no salir.`);
  else if (texts > 1) {
    out.push(`Hay ${texts} cuadros de texto en ${where}${firstText ? ` (el primero: «${firstText}»)` : ""}: su texto puede no salir.`);
  }
  const shapes = found.shapes.length;
  const near = found.shapes.find(Boolean) ?? "";
  if (shapes === 1) {
    out.push(`Hay una forma o dibujo de Word en ${where}${near ? ` (cerca de «${near}»)` : ""}: puede no salir o salir distinto.`);
  } else if (shapes > 1) {
    out.push(`Hay ${shapes} formas o dibujos de Word en ${where}${near ? ` (uno cerca de «${near}»)` : ""}: pueden no salir o salir distintos.`);
  }
  if (found.charts) {
    const what = found.charts === 1 ? "Hay un gráfico" : `Hay ${found.charts} gráficos`;
    out.push(`${what} de Word en ${where}: no sale en la vista previa ni en el PDF. Conviene pegarlo como imagen.`);
  }
  if (found.smartArt) {
    const what = found.smartArt === 1 ? "Hay un SmartArt" : `Hay ${found.smartArt} SmartArt`;
    out.push(`${what} en ${where}: no sale en la vista previa ni en el PDF. Conviene pegarlo como imagen.`);
  }
  if (found.objects) {
    const what = found.objects === 1 ? "Hay un objeto incrustado" : `Hay ${found.objects} objetos incrustados`;
    out.push(`${what} (de otro programa) en ${where} que no trae su imagen: no sale.`);
  }
  if (found.noSize) {
    const what = found.noSize === 1 ? "Una imagen de formato viejo" : `${found.noSize} imágenes de formato viejo`;
    out.push(`${what} en ${where} no dice su tamaño: puede no salir.`);
  }
  if (found.floating) {
    const what = found.floating === 1 ? "Una imagen de formato viejo" : `${found.floating} imágenes de formato viejo`;
    out.push(`${what} en ${where} estaba suelta sobre la página: en el PDF queda en su renglón. Revisá que se vea bien.`);
  }
  if (found.behind) {
    const what = found.behind === 1 ? "Una imagen detrás del texto" : `${found.behind} imágenes detrás del texto`;
    out.push(`${what} (como una marca de agua) en ${where} puede no salir o salir corrida.`);
  }
  if (found.unknown) {
    const what = found.unknown === 1 ? "un elemento" : `${found.unknown} elementos`;
    out.push(`Hay ${what} de Word en ${where} que la vista previa no sabe mostrar: puede no salir.`);
  }
  return out;
}

/** Recibe la copia ya marcada por markPageFields y la deja lista para el visor. */
export async function prepareRenderCopy(
  bytes: Uint8Array,
): Promise<{ bytes: Uint8Array; info: RenderInfo; warnings: string[] }> {
  const info: RenderInfo = { ...DEFAULT_RENDER_INFO, sections: [] };
  const warnings: string[] = [];
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return { bytes, info, warnings };
  }
  pictureId = 0;

  const settingsRoot = (await readXml(zip, "word/settings.xml"))?.documentElement ?? undefined;
  let evenAndOdd = false;
  if (settingsRoot) {
    info.mirrorMargins = isOn(children(settingsRoot, "mirrorMargins")[0]);
    info.gutterAtTop = isOn(children(settingsRoot, "gutterAtTop")[0]);
    info.autoHyphenation = isOn(children(settingsRoot, "autoHyphenation")[0]);
    evenAndOdd = isOn(children(settingsRoot, "evenAndOddHeaders")[0]);
    readFootnotePr(children(settingsRoot, "footnotePr")[0], info);
    const compat = children(settingsRoot, "compat")[0];
    if (compat) {
      info.suppressSpBfAfterPgBrk = isOn(children(compat, "suppressSpBfAfterPgBrk")[0]);
      for (const setting of children(compat, "compatSetting")) {
        if (attr(setting, "name") !== "compatibilityMode") continue;
        const mode = Number.parseInt(attr(setting, "val"), 10);
        if (Number.isFinite(mode)) info.compatibilityMode = mode;
      }
    }
  }

  // Letras con nombre de palabra de CSS («inherit»): la que usaría Word, en los estilos, la numeración y el texto.
  const fontWords = keywordFontReplacements(await readXml(zip, "word/fontTable.xml"));
  for (const name of ["word/styles.xml", "word/numbering.xml"]) {
    const part = await readXml(zip, name);
    if (part && replaceKeywordFonts(part, fontWords)) zip.file(name, serialize(part));
  }
  const stylesDoc = await readXml(zip, "word/styles.xml");
  const hyperlinkIds = hyperlinkStyleIds(stylesDoc);
  const paragraphStyles = readParagraphStyles(stylesDoc);
  const raised = vertAlignStyles(stylesDoc);
  const parts = Object.keys(zip.files).filter((name) => WHERE.some(([pattern]) => pattern.test(name)));
  for (const name of parts) {
    const doc = await readXml(zip, name);
    if (!doc) continue; // si una parte no se puede leer, queda igual que en el Word
    if (name === "word/document.xml") {
      const sections = bodySections(doc);
      info.sections = sections.map(readLayout);
      const body = descendants(doc, "body")[0];
      const last = body ? children(body, "sectPr")[0] : undefined;
      if (last) readFootnotePr(children(last, "footnotePr")[0], info);
      const relsName = "word/_rels/document.xml.rels";
      const rels = await readXml(zip, relsName);
      const ids = blankIds(rels);
      const { blanks } = linkHeaders(doc, sections, evenAndOdd, ids);
      if (rels && blanks.size) {
        const types = await readXml(zip, "[Content_Types].xml");
        addBlankParts(zip, rels, types, ids, blanks);
        zip.file(relsName, serialize(rels));
        if (types) zip.file("[Content_Types].xml", serialize(types));
      }
      for (const sectPr of sections) if (sectPr) bakeGutter(sectPr, info);
      markCustomFootnotes(doc);
      if (body) markManualBreaks(body);
      markHeaderRows(doc);
      addProbes(doc, sections, evenAndOdd);
      addNoteSeparatorProbe(doc, await readXml(zip, "word/footnotes.xml"));
      if (descendants(doc, "endnoteReference").some((ref) => isOnValue(attr(ref, "customMarkFollows")))) {
        warnings.push("Una nota final usa un símbolo propio en vez de número: en la vista previa puede salir también con número.");
      }
    }
    fixHyphens(doc);
    if (name === "word/footnotes.xml") markNoteNumbers(doc);
    applyStyleVertAlign(doc, raised);
    replaceKeywordFonts(doc, fontWords);
    plainTocLinks(doc, hyperlinkIds);
    markParagraphs(doc, paragraphStyles);
    const found = nothingFound();
    convertVml(doc, found);
    findDrawings(doc, found);
    warnings.push(...warningsFor(WHERE.find(([pattern]) => pattern.test(name))![1], found));
    zip.file(name, serialize(doc));
  }

  const spacing = await prepareLineSpacing(zip);
  info.spacing = spacing.spacing;
  warnings.push(...spacing.warnings);

  // STORE: es una copia para el visor; las imágenes pasan igual.
  const out = await zip.generateAsync({ type: "uint8array", compression: "STORE" });
  // El mismo aviso de varios encabezados o pies se dice una vez.
  return { bytes: out, info, warnings: [...new Set(warnings)] };
}
