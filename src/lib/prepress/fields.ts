import JSZip from "jszip";
import { DOMParser, XMLSerializer, type Document as XmlDocument, type Element as XmlElement } from "@xmldom/xmldom";

/**
 * Campos de número de página. El visor muestra el número que Word dejó guardado (el mismo en todas las páginas),
 * así que en una copia que solo sirve para componer se cambian por marcas que después se numeran página por página.
 * También se marcan el comienzo de cada sección y las llamadas a notas al pie (para que viajen con su nota si un
 * párrafo pasa a otra página). El Word limpio que se descarga no pasa por acá.
 */
export const MARK_OPEN = "";
export const MARK_CLOSE = "";
/** Número de página; con formato propio del campo (\* roman…) va como «P:lowerRoman». */
export const pageMark = (format?: string) => `${MARK_OPEN}P${format ? `:${format}` : ""}${MARK_CLOSE}`;
export const PAGE_MARK = pageMark();
export const PAGES_MARK = `${MARK_OPEN}N${MARK_CLOSE}`;
export const SECTION_PAGES_MARK = `${MARK_OPEN}M${MARK_CLOSE}`;
export const FOOTNOTE_MARK = `${MARK_OPEN}F${MARK_CLOSE}`;
export const sectionMark = (index: number) => `${MARK_OPEN}S${index}${MARK_CLOSE}`;

export type SectionNumbering = { start?: number; format?: string };

export type MarkedDocx = { bytes: Uint8Array; sections: SectionNumbering[]; fields: number };

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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

function attr(el: XmlElement | undefined, name: string): string {
  return el ? el.getAttributeNS(W, name) || el.getAttribute(`w:${name}`) || "" : "";
}

/** «\* roman» y compañía, como los escribe Word. */
const SWITCH_FORMATS: Record<string, string> = {
  roman: "lowerRoman",
  ROMAN: "upperRoman",
  alphabetic: "lowerLetter",
  ALPHABETIC: "upperLetter",
  arabic: "decimal",
  Arabic: "decimal",
};

/** «PAGE \* MERGEFORMAT» → marca de página; NUMPAGES → total; SECTIONPAGES → páginas de la sección. */
function markFor(instruction: string): string | null {
  const word = instruction.trim().split(/\s+/)[0]?.toUpperCase();
  if (word === "PAGE") {
    const format = /\\\*\s*([A-Za-z]+)/.exec(instruction)?.[1];
    return pageMark(format ? SWITCH_FORMATS[format] : undefined);
  }
  if (word === "NUMPAGES") return PAGES_MARK;
  if (word === "SECTIONPAGES") return SECTION_PAGES_MARK;
  return null;
}

function newText(doc: XmlDocument, text: string): XmlElement {
  const t = doc.createElementNS(W, "w:t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  return t;
}

function newRun(doc: XmlDocument, text: string, like?: XmlElement | null): XmlElement {
  const run = doc.createElementNS(W, "w:r");
  const props = like ? children(like, "rPr")[0] : undefined;
  if (props) run.appendChild(props.cloneNode(true));
  run.appendChild(newText(doc, text));
  return run;
}

/** Deja el resultado del campo en una sola marca: el primer texto pasa a ser la marca y el resto se vacía. */
function replaceResult(doc: XmlDocument, runs: XmlElement[], mark: string, end: XmlElement) {
  let placed = false;
  for (const run of runs) {
    for (const t of children(run, "t")) {
      if (!placed) {
        while (t.firstChild) t.removeChild(t.firstChild);
        t.appendChild(doc.createTextNode(mark));
        t.setAttribute("xml:space", "preserve");
        placed = true;
      } else {
        run.removeChild(t);
      }
    }
  }
  if (!placed) end.parentNode?.insertBefore(newRun(doc, mark, end), end);
}

/** Las corridas de un párrafo en orden, también las que están dentro de un hipervínculo o un control de contenido. */
function runsOf(paragraph: XmlElement): XmlElement[] {
  const out: XmlElement[] = [];
  const walk = (node: XmlElement) => {
    for (const child of children(node)) {
      const name = localName(child);
      if (name === "r") out.push(child);
      else if (["hyperlink", "sdt", "sdtContent", "smartTag", "customXml", "ins", "fldSimple"].includes(name)) walk(child);
    }
  };
  walk(paragraph);
  return out;
}

function markComplexFields(doc: XmlDocument): number {
  let count = 0;
  for (const paragraph of descendants(doc, "p")) {
    let depth = 0;
    let instruction = "";
    let mark: string | null = null;
    let inResult = false;
    let result: XmlElement[] = [];
    for (const run of runsOf(paragraph)) {
      const chars = children(run, "fldChar");
      const instr = children(run, "instrText");
      if (depth === 1 && !chars.length && !instr.length && inResult && mark) result.push(run);
      for (const node of instr) if (depth === 1) instruction += node.textContent ?? "";
      for (const char of chars) {
        const type = attr(char, "fldCharType");
        if (type === "begin") {
          depth += 1;
          if (depth === 1) {
            instruction = "";
            mark = null;
            inResult = false;
            result = [];
          }
        } else if (type === "separate" && depth === 1) {
          mark = markFor(instruction);
          inResult = true;
        } else if (type === "end") {
          if (depth === 1) {
            mark = mark ?? markFor(instruction);
            if (mark) {
              replaceResult(doc, result, mark, run);
              count += 1;
            }
          }
          depth = Math.max(0, depth - 1);
        }
      }
    }
  }
  return count;
}

function markSimpleFields(doc: XmlDocument): number {
  let count = 0;
  for (const field of descendants(doc, "fldSimple")) {
    const mark = markFor(attr(field, "instr"));
    const parent = field.parentNode;
    if (!parent) continue;
    const runs = children(field, "r");
    if (mark) {
      count += 1;
      if (runs.length) replaceResult(doc, runs, mark, runs[runs.length - 1]!);
      else field.appendChild(newRun(doc, mark));
    }
    // El visor no dibuja fldSimple: se deja su resultado como texto común.
    while (field.firstChild) parent.insertBefore(field.firstChild, field);
    parent.removeChild(field);
  }
  return count;
}

/** Una marca justo antes de cada llamada a nota al pie, dentro de la misma corrida. */
function markFootnoteReferences(doc: XmlDocument) {
  for (const reference of descendants(doc, "footnoteReference")) {
    reference.parentNode?.insertBefore(newText(doc, FOOTNOTE_MARK), reference);
  }
}

/**
 * Solo el salto de página manual: las marcas de corte que Word deja guardadas (lastRenderedPageBreak) el visor ya no
 * las usa para cortar.
 */
function isPageBreak(node: XmlElement): boolean {
  return localName(node) === "br" && attr(node, "type") === "page";
}

/** Lo que se ve en el renglón: si hay algo de esto antes del salto, el salto no abre la sección. */
const VISIBLE = new Set(["t", "tab", "ptab", "drawing", "pict", "object", "sym", "noBreakHyphen", "footnoteReference", "endnoteReference"]);

function isVisible(node: XmlElement): boolean {
  const name = localName(node);
  if (name === "t") return (node.textContent ?? "") !== "";
  return VISIBLE.has(name);
}

/**
 * La marca de comienzo de sección va en el primer párrafo. Si ese párrafo empieza con un salto de página (sin nada
 * visible antes), la marca va después del salto: si no, el visor la deja en la última página de la sección anterior.
 * Si hay texto antes del salto (un «PRIMERA PARTE» y después el salto), la sección empieza con ese texto: la marca va
 * al comienzo del párrafo.
 */
function placeSectionMark(doc: XmlDocument, paragraph: XmlElement, mark: string) {
  for (const run of runsOf(paragraph)) {
    const parts = children(run);
    const breakAt = parts.findIndex(isPageBreak);
    if (breakAt >= 0 && !parts.slice(0, breakAt).some(isVisible)) {
      const after = parts[breakAt]!.nextSibling;
      const text = newText(doc, mark);
      if (after) run.insertBefore(text, after);
      else run.appendChild(text);
      return;
    }
    if (parts.some(isVisible)) break;
  }
  const props = children(paragraph, "pPr")[0];
  const run = newRun(doc, mark);
  if (props?.nextSibling) paragraph.insertBefore(run, props.nextSibling);
  else if (props) paragraph.appendChild(run);
  else paragraph.insertBefore(run, paragraph.firstChild);
}

/** Los bloques del cuerpo como los ve el visor: los controles de contenido y customXml se abren. */
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

/** Marca el comienzo de cada sección del cuerpo y devuelve cómo numera cada una (inicio y formato). */
function markSections(doc: XmlDocument): SectionNumbering[] {
  const body = descendants(doc, "body")[0];
  if (!body) return [];
  const sections: SectionNumbering[] = [];
  let opened = false;
  const readNumbering = (sectPr: XmlElement | undefined): SectionNumbering => {
    const pg = sectPr ? children(sectPr, "pgNumType")[0] : undefined;
    const start = attr(pg, "start");
    const format = attr(pg, "fmt");
    return { start: start ? Number.parseInt(start, 10) : undefined, format: format || undefined };
  };
  for (const block of bodyBlocks(body)) {
    if (!opened) {
      const paragraph = localName(block) === "p" ? block : descendants(block, "p")[0];
      if (paragraph) {
        placeSectionMark(doc, paragraph, sectionMark(sections.length));
        opened = true;
      }
    }
    const props = localName(block) === "p" ? children(block, "pPr")[0] : undefined;
    const sectPr = props ? children(props, "sectPr")[0] : undefined;
    if (sectPr) {
      sections.push(readNumbering(sectPr));
      opened = false;
    }
  }
  const last = children(body, "sectPr")[0];
  if (opened || last) sections.push(readNumbering(last));
  return sections;
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

/** Copia del .docx con marcas en los números de página, en el comienzo de cada sección y en las llamadas a notas. */
export async function markPageFields(input: Uint8Array): Promise<MarkedDocx> {
  const zip = await JSZip.loadAsync(input);
  let fields = 0;
  let sections: SectionNumbering[] = [];
  const parts = Object.keys(zip.files).filter((name) =>
    /^word\/(document|header\d*|footer\d*)\.xml$/.test(name),
  );
  for (const name of parts) {
    const xml = await zip.file(name)?.async("string");
    const doc = xml == null ? null : parse(xml);
    if (!doc) continue; // si una parte no se puede leer, queda igual que en el Word
    fields += markComplexFields(doc) + markSimpleFields(doc);
    if (name === "word/document.xml") {
      sections = markSections(doc);
      markFootnoteReferences(doc);
    }
    zip.file(name, serialize(doc));
  }
  // STORE: es una copia para el visor; las imágenes pasan igual.
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "STORE" });
  return { bytes, sections, fields };
}
