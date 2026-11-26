import JSZip from "jszip";
import { DOMParser, XMLSerializer, type Document as XmlDocument, type Element as XmlElement, type Node as XmlNode } from "@xmldom/xmldom";

export type CleanOptions = {
  /** Acepta inserciones y descarta texto borrado. Quita las marcas de revisión. */
  acceptRevisions: boolean;
  /** Quita comentarios y sus globos. No toca las notas al pie. */
  removeComments: boolean;
  /** Conserva las notas al final. Las notas al pie siempre se conservan. */
  keepEndnotes: boolean;
  /** Quita el resaltado de color (no es una marca de revisión). */
  removeHighlights: boolean;
};

export const DEFAULT_CLEAN_OPTIONS: CleanOptions = {
  acceptRevisions: true,
  removeComments: true,
  keepEndnotes: true,
  removeHighlights: false,
};

const NOT_A_FONT = new Set([
  "inherit",
  "initial",
  "unset",
  "revert",
  "serif",
  "sans-serif",
  "monospace",
  "cursive",
  "fantasy",
  "system-ui",
  "emoji",
  "math",
]);

/** Descarta códigos de idioma y palabras de CSS que a veces aparecen como si fueran una fuente. */
export function isFontName(name: string): boolean {
  const trimmed = name.trim();
  if (!trimmed || trimmed.startsWith("+")) return false;
  const key = trimmed.toLowerCase();
  if (NOT_A_FONT.has(key)) return false;
  if (/^[a-z]{2,3}-[a-z]{2,4}$/i.test(trimmed)) return false;
  return true;
}
export const FONT_SUBSTITUTES: Record<string, string> = {
  calibri: "Carlito",
  "calibri light": "Carlito",
  cambria: "Caladea",
  arial: "Liberation Sans",
  helvetica: "Liberation Sans",
  "times new roman": "Liberation Serif",
  times: "Liberation Serif",
  "courier new": "Liberation Mono",
  courier: "Liberation Mono",
  tahoma: "Tahoma",
  candara: "Candara",
  "candara light": "Candara",
  "ms gothic": "MS Gothic",
  symbol: "Symbol",
  wingdings: "Wingdings",
};

export type FontReport = {
  name: string;
  embedded: boolean;
  /** Familia que se incrusta en el PDF si el Word no traía el archivo. */
  substitute: string | null;
};

export type ImageReport = {
  name: string;
  bytes: number;
};

export type CleanReport = {
  commentsRemoved: number;
  insertionsAccepted: number;
  deletionsRemoved: number;
  footnotesKept: number;
  endnotesKept: number;
  highlightsRemoved: number;
  images: ImageReport[];
  fonts: FontReport[];
  warnings: string[];
};

export type CleanResult = {
  bytes: Uint8Array;
  report: CleanReport;
};

const UNWRAP = new Set(["ins", "moveTo"]);

const DROP = new Set([
  "del",
  "moveFrom",
  "cellDel",
  "cellIns",
  "cellMerge",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
  "customXmlDelRangeStart",
  "customXmlDelRangeEnd",
  "customXmlInsRangeStart",
  "customXmlInsRangeEnd",
  "customXmlMoveFromRangeStart",
  "customXmlMoveFromRangeEnd",
  "customXmlMoveToRangeStart",
  "customXmlMoveToRangeEnd",
]);

const COMMENT_MARKERS = new Set(["commentRangeStart", "commentRangeEnd", "commentReference"]);

const CHANGE_PROPS = new Set([
  "rPrChange",
  "pPrChange",
  "sectPrChange",
  "tblPrChange",
  "tblGridChange",
  "trPrChange",
  "tcPrChange",
  "numPrChange",
]);

type Stats = {
  insertionsAccepted: number;
  deletionsRemoved: number;
  commentMarkers: number;
  highlightsRemoved: number;
  cellMarks: number;
  rowsRemoved: number;
  cellsRemoved: number;
  paragraphsJoined: number;
  paragraphMarksKept: number;
  sectionBreaksRemoved: number;
};

function localName(node: XmlElement): string {
  const raw = node.localName || node.nodeName || "";
  const i = raw.indexOf(":");
  return i >= 0 ? raw.slice(i + 1) : raw;
}

function cleanTree(root: XmlNode, options: CleanOptions, stats: Stats) {
  const children = Array.from(root.childNodes);
  for (const child of children) {
    if (child.nodeType !== 1) continue;
    const el = child as XmlElement;
    const name = localName(el);

    if (options.removeComments && COMMENT_MARKERS.has(name)) {
      stats.commentMarkers += 1;
      el.parentNode?.removeChild(el);
      continue;
    }

    if (options.acceptRevisions && (DROP.has(name) || CHANGE_PROPS.has(name))) {
      if (name === "del" || name === "moveFrom") stats.deletionsRemoved += 1;
      if (name === "cellMerge") stats.cellMarks += 1;
      el.parentNode?.removeChild(el);
      continue;
    }

    if (options.removeHighlights && name === "highlight") {
      stats.highlightsRemoved += 1;
      el.parentNode?.removeChild(el);
      continue;
    }

    if (!options.keepEndnotes && name === "endnoteReference") {
      el.parentNode?.removeChild(el);
      continue;
    }

    cleanTree(el, options, stats);

    if (options.acceptRevisions && UNWRAP.has(name) && el.parentNode) {
      stats.insertionsAccepted += 1;
      const parent = el.parentNode;
      while (el.firstChild) parent.insertBefore(el.firstChild, el);
      parent.removeChild(el);
    }
  }
}

function elementChildren(node: XmlNode): XmlElement[] {
  return Array.from(node.childNodes).filter((n) => n.nodeType === 1) as XmlElement[];
}

function childNamed(node: XmlNode, name: string): XmlElement | undefined {
  return elementChildren(node).find((el) => localName(el) === name);
}

const WORD_NS = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);

/** Solo elementos de Word (w:p, no a:p de un gráfico), en orden de documento. */
function elementsNamed(root: XmlNode, name: string): XmlElement[] {
  const all = (root as XmlDocument | XmlElement).getElementsByTagName("*");
  return (Array.from(all) as XmlElement[]).filter(
    (el) => localName(el) === name && WORD_NS.has(el.namespaceURI ?? ""),
  );
}

/** Si un antepasado ya se quitó, el nodo cuelga de un árbol suelto y no hay que contarlo. */
function isAttached(node: XmlNode, doc: XmlDocument): boolean {
  let cur: XmlNode | null = node;
  while (cur && cur !== doc) cur = cur.parentNode;
  return cur === doc;
}

/** Marcas que pueden quedar entre dos párrafos sin ser contenido. */
const BETWEEN_PARAGRAPHS = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "permStart",
  "permEnd",
  "proofErr",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);

function nextParagraph(p: XmlElement): XmlElement | null {
  let cur = p.nextSibling;
  while (cur) {
    if (cur.nodeType === 1) {
      const name = localName(cur as XmlElement);
      if (name === "p") return cur as XmlElement;
      if (!BETWEEN_PARAGRAPHS.has(name)) return null;
    }
    cur = cur.nextSibling;
  }
  return null;
}

/** ¿Queda algo vivo en el párrafo además de sus propiedades y lo borrado? */
function hasLiveContent(p: XmlElement): boolean {
  return elementChildren(p).some((el) => {
    const name = localName(el);
    return name !== "pPr" && name !== "del" && name !== "moveFrom";
  });
}

/** Una celda de Word siempre tiene que terminar en un párrafo. */
function ensureCellParagraph(cell: XmlNode | null, doc: XmlDocument) {
  if (!cell || cell.nodeType !== 1 || localName(cell as XmlElement) !== "tc") return;
  const el = cell as XmlElement;
  const last = elementChildren(el).pop();
  if (last && localName(last) === "p") return;
  const qname = el.prefix ? `${el.prefix}:p` : "p";
  el.appendChild(doc.createElementNS(el.namespaceURI, qname));
}

function removeNode(el: XmlElement) {
  el.parentNode?.removeChild(el);
}

/**
 * Acepta las revisiones que cambian la estructura, antes de quitar las marcas sueltas:
 * filas y celdas borradas se van enteras, y un salto de párrafo borrado une ese
 * párrafo con el siguiente, que conserva sus propiedades (así lo hace Word).
 * Las marcas de fila o celda insertada las quita después cleanTree.
 */
function acceptStructure(doc: XmlDocument, stats: Stats) {
  const before = stats.rowsRemoved + stats.cellsRemoved;
  for (const tr of elementsNamed(doc, "tr")) {
    const trPr = childNamed(tr, "trPr");
    if (!trPr || !childNamed(trPr, "del") || !isAttached(tr, doc)) continue;
    removeNode(tr);
    stats.rowsRemoved += 1;
  }

  for (const tc of elementsNamed(doc, "tc")) {
    const tcPr = childNamed(tc, "tcPr");
    if (!tcPr || !childNamed(tcPr, "cellDel") || !isAttached(tc, doc)) continue;
    removeNode(tc);
    stats.cellsRemoved += 1;
  }

  // Una fila sin celdas o una tabla sin filas no valen en Word: se van también.
  if (stats.rowsRemoved + stats.cellsRemoved > before) {
    for (const tr of elementsNamed(doc, "tr").reverse()) {
      if (isAttached(tr, doc) && !hasDescendant(tr, "tc")) removeNode(tr);
    }
    for (const tbl of elementsNamed(doc, "tbl").reverse()) {
      if (!isAttached(tbl, doc) || hasDescendant(tbl, "tr")) continue;
      const parent = tbl.parentNode;
      removeNode(tbl);
      ensureCellParagraph(parent, doc);
    }
  }

  for (const p of elementsNamed(doc, "p")) {
    if (!isAttached(p, doc)) continue;
    const pPr = childNamed(p, "pPr");
    const markProps = pPr ? childNamed(pPr, "rPr") : undefined;
    const mark = markProps ? (childNamed(markProps, "del") ?? childNamed(markProps, "moveFrom")) : undefined;
    if (!pPr || !markProps || !mark) continue;
    const breaksSection = Boolean(childNamed(pPr, "sectPr"));
    const next = nextParagraph(p);

    if (next) {
      const anchor = elementChildren(next).find((el) => localName(el) !== "pPr") ?? null;
      for (const node of Array.from(p.childNodes)) {
        if (node === pPr) continue;
        next.insertBefore(node, anchor);
      }
      removeNode(p);
      stats.paragraphsJoined += 1;
      if (breaksSection) stats.sectionBreaksRemoved += 1;
      continue;
    }

    if (!breaksSection && !hasLiveContent(p)) {
      const parent = p.parentNode;
      removeNode(p);
      ensureCellParagraph(parent, doc);
      stats.paragraphsJoined += 1;
      continue;
    }

    // No hay párrafo siguiente con el que unirlo (fin de celda, antes de una tabla…):
    // se deja el párrafo y solo se quita la marca.
    markProps.removeChild(mark);
    stats.paragraphMarksKept += 1;
  }
}

function hasDescendant(el: XmlElement, name: string): boolean {
  return elementsNamed(el, name).length > 0;
}

function serialize(doc: XmlDocument): string {
  let out = new XMLSerializer().serializeToString(doc);
  if (!out.startsWith("<?xml")) {
    out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
  }
  return out;
}

function parseXml(xml: string): XmlDocument | null {
  const doc = new DOMParser({
    onError: () => {
      /* El aviso lo pone el llamador si el resultado trae parsererror. */
    },
  }).parseFromString(xml, "text/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) return null;
  return doc;
}

function cleanXml(xml: string, options: CleanOptions, stats: Stats, warnings: string[], part: string): string {
  const doc = parseXml(xml);
  if (!doc) {
    warnings.push(`No se pudo interpretar ${part}. Esa parte quedó igual.`);
    return xml;
  }
  if (options.acceptRevisions) acceptStructure(doc, stats);
  cleanTree(doc, options, stats);
  return serialize(doc);
}

function shouldCleanPart(path: string, commentParts: Map<string, string>): boolean {
  if (!path.startsWith("word/") || !path.endsWith(".xml")) return false;
  if (path.endsWith("/fontTable.xml") || path.endsWith("/settings.xml")) return false;
  if (path.endsWith("/styles.xml") || path.endsWith("/numbering.xml") || path.endsWith("/webSettings.xml")) {
    return false;
  }
  if (path.includes("/theme/")) return false;
  if (commentParts.has(path)) return false;
  return true;
}

const COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

/** Solo estos tipos de relación son comentarios. Se mira el tipo, nunca el nombre del archivo. */
const COMMENT_REL_TYPES = new Set([
  COMMENTS_REL,
  "http://purl.oclc.org/ooxml/officeDocument/relationships/comments",
  "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
  "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
  "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
  "http://schemas.microsoft.com/office/2011/relationships/people",
]);

function isRelsFile(name: string): boolean {
  return /(^|\/)_rels\/[^/]*\.rels$/.test(name);
}

/** "word/_rels/document.xml.rels" → "word/": la carpeta contra la que se resuelven sus destinos. */
function relsBase(relsPath: string): string {
  const i = relsPath.lastIndexOf("_rels/");
  return i >= 0 ? relsPath.slice(0, i) : "";
}

function resolveTarget(base: string, target: string): string {
  let raw = target;
  try {
    raw = decodeURIComponent(target);
  } catch {
    /* Se usa tal cual. */
  }
  const joined = raw.startsWith("/") ? raw.slice(1) : base + raw;
  const out: string[] = [];
  for (const segment of joined.split("/")) {
    if (segment === "" || segment === ".") continue;
    if (segment === "..") out.pop();
    else out.push(segment);
  }
  return out.join("/");
}

/** Relaciones propias de una parte: "word/comments.xml" → "word/_rels/comments.xml.rels". */
function relsOf(part: string): string {
  const i = part.lastIndexOf("/");
  return `${part.slice(0, i + 1)}_rels/${part.slice(i + 1)}.rels`;
}

function relationships(doc: XmlDocument): XmlElement[] {
  return (Array.from(doc.getElementsByTagName("*")) as XmlElement[]).filter(
    (el) => localName(el) === "Relationship",
  );
}

/** Partes de comentarios (ruta en el zip → tipo de relación), halladas por el tipo de relación. */
async function findCommentParts(zip: JSZip): Promise<Map<string, string>> {
  const byLower = new Map<string, string>();
  for (const name of Object.keys(zip.files)) byLower.set(name.toLowerCase(), name);

  const parts = new Map<string, string>();
  for (const name of Object.keys(zip.files)) {
    if (!isRelsFile(name)) continue;
    const xml = await zip.file(name)?.async("string");
    const doc = xml ? parseXml(xml) : null;
    if (!doc) continue;
    for (const rel of relationships(doc)) {
      const type = rel.getAttribute("Type") || "";
      const target = rel.getAttribute("Target") || "";
      if (!COMMENT_REL_TYPES.has(type) || !target || rel.getAttribute("TargetMode") === "External") continue;
      const found = byLower.get(resolveTarget(relsBase(name), target).toLowerCase());
      if (found && !zip.files[found]?.dir) parts.set(found, type);
    }
  }
  return parts;
}

/**
 * Quita los comentarios del paquete: solo las partes que apuntan las relaciones de
 * tipo comentario, esas relaciones y sus entradas en [Content_Types].xml.
 * Un archivo que solo se llame parecido (word/media/people.jpg) no se toca.
 */
async function stripCommentPackage(zip: JSZip, commentParts: Map<string, string>) {
  const removed = new Set<string>();
  for (const part of commentParts.keys()) {
    zip.remove(part);
    removed.add(part.toLowerCase());
    if (zip.file(relsOf(part))) zip.remove(relsOf(part));
  }

  for (const name of Object.keys(zip.files)) {
    if (!isRelsFile(name)) continue;
    const xml = await zip.file(name)?.async("string");
    const doc = xml ? parseXml(xml) : null;
    if (!doc) continue;
    let changed = false;
    for (const rel of relationships(doc)) {
      if (!COMMENT_REL_TYPES.has(rel.getAttribute("Type") || "")) continue;
      rel.parentNode?.removeChild(rel);
      changed = true;
    }
    if (changed) zip.file(name, serialize(doc));
  }

  if (removed.size === 0) return;
  const typesXml = await zip.file("[Content_Types].xml")?.async("string");
  const types = typesXml ? parseXml(typesXml) : null;
  if (!types) return;
  let changed = false;
  for (const el of Array.from(types.getElementsByTagName("*")) as XmlElement[]) {
    if (localName(el) !== "Override") continue;
    if (!removed.has(resolveTarget("", el.getAttribute("PartName") || "").toLowerCase())) continue;
    el.parentNode?.removeChild(el);
    changed = true;
  }
  if (changed) zip.file("[Content_Types].xml", serialize(types));
}

async function clearTrackRevisions(zip: JSZip) {
  const file = zip.file("word/settings.xml");
  if (!file) return;
  const xml = await file.async("string");
  const next = xml
    .replace(/<w:trackRevisions\b[^>]*\/>/g, "")
    .replace(/<w:trackRevisions\b[^>]*>[\s\S]*?<\/w:trackRevisions>/g, "");
  if (next !== xml) zip.file("word/settings.xml", next);
}

async function readFonts(zip: JSZip): Promise<FontReport[]> {
  const embedded = new Set<string>();
  const names = new Set<string>();

  const fontTable = await zip.file("word/fontTable.xml")?.async("string");
  if (fontTable) {
    for (const chunk of fontTable.split(/<w:font\b/).slice(1)) {
      const name = /w:name="([^"]+)"/.exec(chunk)?.[1];
      if (!name || !isFontName(name)) continue;
      names.add(name);
      if (/<w:embed(?:Regular|Bold|Italic|BoldItalic)\b/.test(chunk)) embedded.add(name);
    }
  }

  const usedPart = (path: string) =>
    /^word\/(document|footnotes|endnotes|header\d+|footer\d+)\.xml$/.test(path);

  const attr = /w:(?:ascii|hAnsi|cs|eastAsia)="([^"]+)"/g;
  for (const path of Object.keys(zip.files)) {
    if (!usedPart(path)) continue;
    const xml = await zip.file(path)?.async("string");
    if (!xml) continue;
    for (const match of xml.matchAll(attr)) {
      if (isFontName(match[1] ?? "")) names.add(match[1]!);
    }
  }

  const themeXml = await zip.file("word/theme/theme1.xml")?.async("string");
  if (themeXml) {
    for (const block of themeXml.match(/<a:(?:major|minor)Font>[\s\S]*?<\/a:(?:major|minor)Font>/g) ?? []) {
      const face = /<a:latin[^>]*typeface="([^"]+)"/.exec(block)?.[1];
      if (face && isFontName(face)) names.add(face);
    }
  }

  return [...names]
    .sort((a, b) => a.localeCompare(b))
    .map((name) => {
      const isEmbedded = embedded.has(name);
      return {
        name,
        embedded: isEmbedded,
        substitute: isEmbedded ? null : (FONT_SUBSTITUTES[name.trim().toLowerCase()] ?? null),
      };
    });
}

async function readImages(zip: JSZip): Promise<ImageReport[]> {
  const images: ImageReport[] = [];
  for (const name of Object.keys(zip.files)) {
    if (!name.startsWith("word/media/") || zip.files[name]?.dir) continue;
    const data = await zip.file(name)?.async("uint8array");
    if (!data) continue;
    images.push({ name: name.split("/").pop() || name, bytes: data.byteLength });
  }
  images.sort((a, b) => a.name.localeCompare(b.name));
  return images;
}

function countTag(xml: string, tag: string): number {
  const re = new RegExp(`<w:${tag}\\b`, "g");
  return xml.match(re)?.length ?? 0;
}

/**
 * Limpia un .docx para premedia sin volver a codificar imágenes ni fuentes.
 * Las notas al pie (w:footnoteReference y footnotes.xml) no se tocan,
 * salvo las marcas de revisión o comentarios que haya dentro de ellas.
 */
export async function cleanDocx(input: Uint8Array, options: CleanOptions): Promise<CleanResult> {
  const zip = await JSZip.loadAsync(input);
  const stats: Stats = {
    insertionsAccepted: 0,
    deletionsRemoved: 0,
    commentMarkers: 0,
    highlightsRemoved: 0,
    cellMarks: 0,
    rowsRemoved: 0,
    cellsRemoved: 0,
    paragraphsJoined: 0,
    paragraphMarksKept: 0,
    sectionBreaksRemoved: 0,
  };
  const warnings: string[] = [];

  const commentParts = await findCommentParts(zip);
  let commentsRemoved = 0;
  if (options.removeComments) {
    for (const [part, type] of commentParts) {
      if (!type.endsWith("/comments")) continue;
      const commentsXml = await zip.file(part)?.async("string");
      if (commentsXml) commentsRemoved += commentsXml.match(/<w:comment[\s>]/g)?.length ?? 0;
    }
  }

  const parts = Object.keys(zip.files).filter(
    (name) => !zip.files[name]?.dir && shouldCleanPart(name, commentParts),
  );
  for (const name of parts) {
    const xml = await zip.file(name)?.async("string");
    if (xml == null) continue;
    zip.file(name, cleanXml(xml, options, stats, warnings, name));
  }

  if (options.removeComments) await stripCommentPackage(zip, commentParts);
  if (options.acceptRevisions) await clearTrackRevisions(zip);

  if (stats.cellMarks > 0) {
    warnings.push(
      "Había celdas combinadas con marca de revisión. Se quitó la marca y la tabla quedó como estaba: revisala antes de mandar a imprenta.",
    );
  }
  if (stats.cellsRemoved > 0) {
    const n = stats.cellsRemoved;
    warnings.push(
      n === 1
        ? "Se borró 1 celda que estaba marcada como eliminada. Revisá que esa tabla haya quedado pareja."
        : `Se borraron ${n} celdas que estaban marcadas como eliminadas. Revisá que esas tablas hayan quedado parejas.`,
    );
  }
  if (stats.sectionBreaksRemoved > 0) {
    const n = stats.sectionBreaksRemoved;
    warnings.push(
      n === 1
        ? "Se aceptó el borrado de un salto de sección: esa parte tomó el formato de página de la sección siguiente."
        : `Se aceptó el borrado de ${n} saltos de sección: esas partes tomaron el formato de página de la sección siguiente.`,
    );
  }
  if (stats.paragraphMarksKept > 0) {
    const n = stats.paragraphMarksKept;
    warnings.push(
      n === 1
        ? "Un párrafo tenía borrado el fin de párrafo, pero no había un párrafo siguiente con el cual unirlo (por ejemplo, antes de una tabla). Quedó separado: revisalo."
        : `${n} párrafos tenían borrado el fin de párrafo, pero no había un párrafo siguiente con el cual unirlos (por ejemplo, antes de una tabla). Quedaron separados: revisalos.`,
    );
  }

  const documentXml = (await zip.file("word/document.xml")?.async("string")) ?? "";
  const footnotesKept = countTag(documentXml, "footnoteReference");
  const endnotesKept = countTag(documentXml, "endnoteReference");

  if (options.removeComments && /<w:(commentReference|commentRangeStart)\b/.test(documentXml)) {
    warnings.push("Quedó alguna marca de comentario en el cuerpo.");
  }
  if (options.acceptRevisions && /<w:(ins|del)\b/.test(documentXml)) {
    warnings.push("Quedó alguna marca de revisión en el cuerpo.");
  }

  const fonts = await readFonts(zip);
  for (const font of fonts) {
    if (!font.embedded && !font.substitute) {
      warnings.push(
        `«${font.name}» no viene incrustada en el Word y no hay sustituta métrica. La paginación puede moverse.`,
      );
    }
  }

  const images = await readImages(zip);
  const bytes = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return {
    bytes,
    report: {
      commentsRemoved,
      insertionsAccepted: stats.insertionsAccepted,
      deletionsRemoved: stats.deletionsRemoved + stats.rowsRemoved + stats.cellsRemoved + stats.paragraphsJoined,
      footnotesKept,
      endnotesKept,
      highlightsRemoved: stats.highlightsRemoved,
      images,
      fonts,
      warnings,
    },
  };
}
