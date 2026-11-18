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
      if (name === "cellDel" || name === "cellIns") stats.cellMarks += 1;
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
  cleanTree(doc, options, stats);
  return serialize(doc);
}

function shouldCleanPart(path: string): boolean {
  if (!path.startsWith("word/") || !path.endsWith(".xml")) return false;
  if (path.endsWith("/fontTable.xml") || path.endsWith("/settings.xml")) return false;
  if (path.endsWith("/styles.xml") || path.endsWith("/numbering.xml") || path.endsWith("/webSettings.xml")) {
    return false;
  }
  if (path.includes("/theme/")) return false;
  if (/comments/i.test(path)) return false;
  return true;
}

async function stripCommentPackage(zip: JSZip) {
  const names = Object.keys(zip.files);
  for (const name of names) {
    if (/comments/i.test(name) || /(^|\/)people\.xml$/.test(name)) zip.remove(name);
  }

  const rels = Object.keys(zip.files).filter(
    (name) => name.endsWith(".rels") || name === "[Content_Types].xml",
  );
  for (const name of rels) {
    const file = zip.file(name);
    if (!file) continue;
    const xml = await file.async("string");
    const doc = parseXml(xml);
    if (!doc) continue;
    const nodes = Array.from(doc.getElementsByTagName("*"));
    for (const el of nodes) {
      const blob = `${el.getAttribute("Type") || ""} ${el.getAttribute("Target") || ""} ${el.getAttribute("PartName") || ""}`.toLowerCase();
      if (blob.includes("comment") || blob.includes("people")) el.parentNode?.removeChild(el);
    }
    zip.file(name, serialize(doc));
  }
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
  };
  const warnings: string[] = [];

  let commentsRemoved = 0;
  if (options.removeComments) {
    const commentsXml = await zip.file("word/comments.xml")?.async("string");
    if (commentsXml) commentsRemoved = commentsXml.match(/<w:comment[\s>]/g)?.length ?? 0;
  }

  const parts = Object.keys(zip.files).filter((name) => !zip.files[name]?.dir && shouldCleanPart(name));
  for (const name of parts) {
    const xml = await zip.file(name)?.async("string");
    if (xml == null) continue;
    zip.file(name, cleanXml(xml, options, stats, warnings, name));
  }

  if (options.removeComments) await stripCommentPackage(zip);
  if (options.acceptRevisions) await clearTrackRevisions(zip);

  if (stats.cellMarks > 0) {
    warnings.push(
      "Había celdas con marca de revisión. Se quitó la marca y se conservó la celda: revísala antes de filmar.",
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
      deletionsRemoved: stats.deletionsRemoved,
      footnotesKept,
      endnotesKept,
      highlightsRemoved: stats.highlightsRemoved,
      images,
      fonts,
      warnings,
    },
  };
}
