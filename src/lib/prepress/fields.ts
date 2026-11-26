import JSZip from "jszip";
import { DOMParser, XMLSerializer, type Document as XmlDocument, type Element as XmlElement } from "@xmldom/xmldom";

/**
 * Campos de número de página. El visor muestra el número que Word dejó guardado (el mismo en todas las páginas),
 * así que en una copia que solo sirve para componer se cambian por marcas que después se numeran página por página.
 * El Word limpio que se descarga no pasa por acá.
 */
export const MARK_OPEN = "";
export const MARK_CLOSE = "";
export const PAGE_MARK = `${MARK_OPEN}P${MARK_CLOSE}`;
export const PAGES_MARK = `${MARK_OPEN}N${MARK_CLOSE}`;
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

/** «PAGE \* MERGEFORMAT» → PAGE_MARK; «NUMPAGES» / «SECTIONPAGES» → PAGES_MARK. */
function markFor(instruction: string): string | null {
  const word = instruction.trim().split(/\s+/)[0]?.toUpperCase();
  if (word === "PAGE") return PAGE_MARK;
  if (word === "NUMPAGES" || word === "SECTIONPAGES") return PAGES_MARK;
  return null;
}

function newRun(doc: XmlDocument, text: string, like?: XmlElement | null): XmlElement {
  const run = doc.createElementNS(W, "w:r");
  const props = like ? children(like, "rPr")[0] : undefined;
  if (props) run.appendChild(props.cloneNode(true));
  const t = doc.createElementNS(W, "w:t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  run.appendChild(t);
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

function markComplexFields(doc: XmlDocument): number {
  let count = 0;
  for (const paragraph of descendants(doc, "p")) {
    const runs = children(paragraph).flatMap((node) => (localName(node) === "r" ? [node] : children(node, "r")));
    let depth = 0;
    let instruction = "";
    let mark: string | null = null;
    let inResult = false;
    let result: XmlElement[] = [];
    for (const run of runs) {
      const chars = children(run, "fldChar");
      const instr = children(run, "instrText");
      if (depth === 1 && !chars.length && !instr.length && inResult && mark) result.push(run);
      for (const node of instr) if (depth === 1) instruction += node.textContent ?? "";
      for (const char of chars) {
        const type = char.getAttributeNS(W, "fldCharType") || char.getAttribute("w:fldCharType");
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
    const mark = markFor(field.getAttributeNS(W, "instr") || field.getAttribute("w:instr") || "");
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

/** Marca el comienzo de cada sección del cuerpo y devuelve cómo numera cada una (inicio y formato). */
function markSections(doc: XmlDocument): SectionNumbering[] {
  const body = descendants(doc, "body")[0];
  if (!body) return [];
  const sections: SectionNumbering[] = [];
  let first: XmlElement | null = null;
  const readNumbering = (sectPr: XmlElement | undefined): SectionNumbering => {
    const pg = sectPr ? children(sectPr, "pgNumType")[0] : undefined;
    const start = pg?.getAttributeNS(W, "start") || pg?.getAttribute("w:start");
    const format = pg?.getAttributeNS(W, "fmt") || pg?.getAttribute("w:fmt");
    return { start: start ? Number.parseInt(start, 10) : undefined, format: format || undefined };
  };
  const open = (block: XmlElement) => {
    const paragraph = localName(block) === "p" ? block : descendants(block, "p")[0];
    if (!paragraph) return;
    const props = children(paragraph, "pPr")[0];
    const run = newRun(doc, sectionMark(sections.length));
    if (props?.nextSibling) paragraph.insertBefore(run, props.nextSibling);
    else if (props) paragraph.appendChild(run);
    else paragraph.insertBefore(run, paragraph.firstChild);
  };
  for (const block of children(body)) {
    const name = localName(block);
    if (name === "sectPr") continue;
    if (!first) {
      first = block;
      open(block);
    }
    const sectPr = name === "p" ? children(children(block, "pPr")[0] ?? block, "sectPr")[0] : undefined;
    if (sectPr) {
      sections.push(readNumbering(sectPr));
      first = null;
    }
  }
  if (first) sections.push(readNumbering(children(body, "sectPr")[0]));
  return sections;
}

function parse(xml: string): XmlDocument | null {
  const doc = new DOMParser({ onError: () => undefined }).parseFromString(xml, "text/xml");
  return doc.getElementsByTagName("parsererror").length ? null : doc;
}

function serialize(doc: XmlDocument): string {
  const out = new XMLSerializer().serializeToString(doc);
  return out.startsWith("<?xml") ? out : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
}

/** Copia del .docx con marcas en los números de página y en el comienzo de cada sección. */
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
    if (!doc) continue;
    fields += markComplexFields(doc) + markSimpleFields(doc);
    if (name === "word/document.xml") sections = markSections(doc);
    zip.file(name, serialize(doc));
  }
  // STORE: es una copia para el visor; las imágenes pasan igual.
  const bytes = await zip.generateAsync({ type: "uint8array", compression: "STORE" });
  return { bytes, sections, fields };
}
