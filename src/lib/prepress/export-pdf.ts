import type {
  PDFContext as PDFContextType,
  PDFDict as PDFDictType,
  PDFDocument as PDFDocumentType,
  PDFObject as PDFObjectType,
  PDFPage as PDFPageType,
  PDFRef as PDFRefType,
} from "pdf-lib";

export type PdfExport = {
  bytes: Uint8Array;
  /** true si el archivo cumple y se declara PDF/X-1a:2001. */
  pdfx: boolean;
  /** Lo que impide declararlo PDF/X-1a, en palabras para quien edita. */
  problems: string[];
};

type PdfLib = typeof import("pdf-lib");

/**
 * Cada sección ya paginada por el visor se vuelca a una página PDF del mismo
 * tamaño. El texto sigue siendo texto (fuentes incrustadas) y un JPEG del
 * Word entra como DCT, sin reencodear. Después se prepara para imprenta
 * (Lightning Source / IngramSpark): caja de corte, colores neutros en gris y,
 * si todo está en gris con las fuentes incrustadas, PDF/X-1a:2001.
 */
export async function exportPagesToPdf(
  sections: HTMLElement[],
  opts: { title: string; subset: boolean; cropMarks?: boolean; grayscale?: boolean },
): Promise<PdfExport> {
  const { toPdf } = await import("jsafa-pdf");
  const lib = await import("pdf-lib");
  const { PDFDocument, rgb } = lib;

  if (sections.length === 0) throw new Error("El documento no tiene páginas para exportar.");

  const merged = await PDFDocument.create();
  merged.setTitle(opts.title);
  merged.setCreator("Pliego");
  merged.setProducer("Pliego premedia");

  for (const section of sections) {
    const bytes = await toPdf(section, {
      subset: opts.subset,
      title: opts.title,
      autoFonts: true,
      waitForAssets: true,
    });
    const part = await PDFDocument.load(bytes);
    const copied = await merged.copyPages(part, part.getPageIndices());
    for (const page of copied) {
      const { width, height } = page.getSize();
      // El corte es la página del Word, con o sin marcas alrededor.
      page.setTrimBox(0, 0, width, height);
      page.setBleedBox(0, 0, width, height);
      if (opts.cropMarks) drawCropMarks(page, rgb);
      merged.addPage(page);
    }
  }

  const problems = prepareForPress(lib, merged, Boolean(opts.grayscale));
  shareFontPrograms(lib, merged);
  const pdfx = problems.length === 0;
  if (pdfx) declarePdfX1a(lib, merged, opts.title);
  const id = lib.PDFHexString.of(fileId(opts.title));
  merged.context.trailerInfo.ID = merged.context.obj([id, id]);
  infoDict(lib, merged).set(lib.PDFName.of("Trapped"), lib.PDFName.of("False"));

  const bytes = await merged.save({ useObjectStreams: false });
  // pdf-lib siempre escribe «%PDF-1.7»; PDF/X-1a:2001 pide 1.3. Mismo largo: no se corre ningún desplazamiento.
  setHeaderVersion(bytes, pdfx ? "1.3" : "1.4");
  return { bytes, pdfx, problems };
}

/** Marcas de corte en vectores, fuera del recorte. No tocan las imágenes. */
function drawCropMarks(
  page: import("pdf-lib").PDFPage,
  rgb: typeof import("pdf-lib").rgb,
) {
  const width = page.getWidth();
  const height = page.getHeight();
  const gap = 6;
  const len = 14;
  const pad = gap + len;
  page.setMediaBox(-pad, -pad, width + pad * 2, height + pad * 2);
  const ink = rgb(0, 0, 0);
  const lines: Array<[[number, number], [number, number]]> = [
    [[-pad, 0], [-gap, 0]],
    [[0, -pad], [0, -gap]],
    [[width + gap, 0], [width + pad, 0]],
    [[width, -pad], [width, -gap]],
    [[-pad, height], [-gap, height]],
    [[0, height + gap], [0, height + pad]],
    [[width + gap, height], [width + pad, height]],
    [[width, height + gap], [width, height + pad]],
  ];
  for (const [start, end] of lines) {
    page.drawLine({
      start: { x: start[0], y: start[1] },
      end: { x: end[0], y: end[1] },
      thickness: 0.4,
      color: ink,
    });
  }
}

const NUMBER = String.raw`-?(?:\d+\.?\d*|\.\d+)`;
/** «0 0 0 rg» → «0 g»: un color con los tres canales iguales es un gris. Así el negro es negro puro (100 % K). */
const NEUTRAL_FILL = new RegExp(String.raw`(^|\s)(${NUMBER})\s+\2\s+\2\s+rg(?=\s|$)`, "g");
const NEUTRAL_STROKE = new RegExp(String.raw`(^|\s)(${NUMBER})\s+\2\s+\2\s+RG(?=\s|$)`, "g");
const COLOR_LEFT = /(^|\s)(rg|RG|scn?|SCN?)(?=\s|$)/;
const ANY_RGB = new RegExp(String.raw`(^|\s)(${NUMBER})\s+(${NUMBER})\s+(${NUMBER})\s+(rg|RG)(?=\s|$)`, "g");

export function neutralToGray(content: string): string {
  return content.replace(NEUTRAL_FILL, "$1$2 g").replace(NEUTRAL_STROKE, "$1$2 G");
}

/** Interior en blanco y negro: cualquier color RGB pasa a su gris (luminancia), como lo imprimiría la imprenta. */
export function allToGray(content: string): string {
  return neutralToGray(content).replace(ANY_RGB, (_, lead: string, r: string, g: string, b: string, op: string) => {
    const y = 0.299 * Number(r) + 0.587 * Number(g) + 0.114 * Number(b);
    return `${lead}${Number(y.toFixed(4))} ${op === "rg" ? "g" : "G"}`;
  });
}

function lookup(lib: PdfLib, context: PDFContextType, value: PDFObjectType | undefined): PDFObjectType | undefined {
  return value instanceof lib.PDFRef ? context.lookup(value) : value;
}

function asDict(lib: PdfLib, context: PDFContextType, value: PDFObjectType | undefined): PDFDictType | undefined {
  const found = lookup(lib, context, value);
  if (found instanceof lib.PDFDict) return found;
  if (found instanceof lib.PDFRawStream) return found.dict;
  return undefined;
}

function nameOf(lib: PdfLib, context: PDFContextType, value: PDFObjectType | undefined): string | undefined {
  const found = lookup(lib, context, value);
  return found instanceof lib.PDFName ? found.decodeText() : undefined;
}

function numberOf(lib: PdfLib, context: PDFContextType, value: PDFObjectType | undefined): number | undefined {
  const found = lookup(lib, context, value);
  return found instanceof lib.PDFNumber ? found.asNumber() : undefined;
}

/** Un stream nuevo con el mismo diccionario (menos la compresión y el largo, que se recalculan). */
function replaceStream(lib: PdfLib, context: PDFContextType, ref: PDFRefType, dict: PDFDictType, bytes: Uint8Array, drop: string[] = []) {
  const entries: Record<string, PDFObjectType> = {};
  for (const [key, value] of dict.entries()) {
    const name = key.decodeText();
    if (["Filter", "DecodeParms", "Length", ...drop].includes(name)) continue;
    entries[name] = value;
  }
  context.assign(ref, context.flateStream(bytes, entries));
}

/** Reescribe un stream de contenido con los neutros en gris. Devuelve si quedó algún color que no es gris. */
function rewriteContent(lib: PdfLib, context: PDFContextType, ref: PDFRefType, grayscale: boolean): boolean {
  const stream = context.lookup(ref);
  if (!(stream instanceof lib.PDFRawStream)) return false;
  const raw = lib.decodePDFRawStream(stream).decode();
  let text = "";
  for (let i = 0; i < raw.length; i += 0x8000) text += String.fromCharCode(...raw.subarray(i, i + 0x8000));
  const next = grayscale ? allToGray(text) : neutralToGray(text);
  if (next !== text) {
    const bytes = new Uint8Array(next.length);
    for (let i = 0; i < next.length; i += 1) bytes[i] = next.charCodeAt(i) & 0xff;
    replaceStream(lib, context, ref, stream.dict, bytes);
  }
  return COLOR_LEFT.test(next);
}

function contentRefs(lib: PdfLib, page: PDFPageType): PDFRefType[] {
  const contents = page.node.get(lib.PDFName.of("Contents"));
  const list = contents instanceof lib.PDFRef ? page.doc.context.lookup(contents) : contents;
  if (list instanceof lib.PDFArray) return list.asArray().filter((x): x is PDFRefType => x instanceof lib.PDFRef);
  return contents instanceof lib.PDFRef ? [contents] : [];
}

/** Imagen RGB sin pérdida (Flate, 8 bits, sin predictor) → gris, también sin pérdida. Una transparencia se compone sobre blanco. */
function imageToGray(lib: PdfLib, context: PDFContextType, ref: PDFRefType): boolean {
  const stream = context.lookup(ref);
  if (!(stream instanceof lib.PDFRawStream)) return false;
  const dict = stream.dict;
  if (nameOf(lib, context, dict.get(lib.PDFName.of("Filter"))) !== "FlateDecode") return false;
  if (dict.get(lib.PDFName.of("DecodeParms"))) return false;
  if (nameOf(lib, context, dict.get(lib.PDFName.of("ColorSpace"))) !== "DeviceRGB") return false;
  if (numberOf(lib, context, dict.get(lib.PDFName.of("BitsPerComponent"))) !== 8) return false;
  const pixels = lib.decodePDFRawStream(stream).decode();
  let alpha: Uint8Array | null = null;
  const maskRef = dict.get(lib.PDFName.of("SMask"));
  if (maskRef instanceof lib.PDFRef) {
    const mask = context.lookup(maskRef);
    if (!(mask instanceof lib.PDFRawStream)) return false;
    if (nameOf(lib, context, mask.dict.get(lib.PDFName.of("Filter"))) !== "FlateDecode") return false;
    if (mask.dict.get(lib.PDFName.of("DecodeParms"))) return false;
    alpha = lib.decodePDFRawStream(mask).decode();
    if (alpha.length * 3 !== pixels.length) return false;
  }
  const gray = new Uint8Array(pixels.length / 3);
  for (let i = 0, j = 0; j < gray.length; i += 3, j += 1) {
    let y = Math.round(0.299 * pixels[i]! + 0.587 * pixels[i + 1]! + 0.114 * pixels[i + 2]!);
    if (alpha) y = Math.round((y * alpha[j]! + 255 * (255 - alpha[j]!)) / 255);
    gray[j] = y;
  }
  dict.set(lib.PDFName.of("ColorSpace"), lib.PDFName.of("DeviceGray"));
  replaceStream(lib, context, ref, dict, gray, ["SMask", "Decode"]);
  return true;
}

function fontEmbedded(lib: PdfLib, context: PDFContextType, font: PDFDictType): boolean {
  let target = font;
  if (nameOf(lib, context, font.get(lib.PDFName.of("Subtype"))) === "Type0") {
    const descendants = lookup(lib, context, font.get(lib.PDFName.of("DescendantFonts")));
    const first = descendants instanceof lib.PDFArray ? asDict(lib, context, descendants.get(0)) : undefined;
    if (!first) return false;
    target = first;
  }
  if (nameOf(lib, context, target.get(lib.PDFName.of("Subtype"))) === "Type3") return true;
  const descriptor = asDict(lib, context, target.get(lib.PDFName.of("FontDescriptor")));
  return Boolean(
    descriptor &&
      (descriptor.get(lib.PDFName.of("FontFile")) ||
        descriptor.get(lib.PDFName.of("FontFile2")) ||
        descriptor.get(lib.PDFName.of("FontFile3"))),
  );
}

/**
 * Deja el PDF listo para imprenta sin tocar el texto ni recomprimir fotos: neutros en gris, sin anotaciones y, en
 * blanco y negro, las imágenes sin pérdida pasadas a gris. Devuelve lo que impide declararlo PDF/X-1a.
 */
function prepareForPress(lib: PdfLib, doc: PDFDocumentType, grayscale: boolean): string[] {
  const context = doc.context;
  const missingFonts = new Set<string>();
  const colorImages = new Set<number>();
  let colorVectors = false;
  let transparency = false;
  const seenStreams = new Set<string>();
  const seenImages = new Set<string>();

  const walkResources = (resources: PDFDictType | undefined, pageNumber: number) => {
    if (!resources) return;
    const fonts = asDict(lib, context, resources.get(lib.PDFName.of("Font")));
    for (const [, value] of fonts?.entries() ?? []) {
      const font = asDict(lib, context, value);
      if (font && !fontEmbedded(lib, context, font)) {
        missingFonts.add(nameOf(lib, context, font.get(lib.PDFName.of("BaseFont"))) ?? "sin nombre");
      }
    }
    const states = asDict(lib, context, resources.get(lib.PDFName.of("ExtGState")));
    for (const [, value] of states?.entries() ?? []) {
      const state = asDict(lib, context, value);
      if (!state) continue;
      for (const key of ["CA", "ca"]) {
        const alpha = numberOf(lib, context, state.get(lib.PDFName.of(key)));
        if (alpha !== undefined && alpha < 1) transparency = true;
      }
      const softMask = state.get(lib.PDFName.of("SMask"));
      if (softMask && nameOf(lib, context, softMask) !== "None") transparency = true;
    }
    const objects = asDict(lib, context, resources.get(lib.PDFName.of("XObject")));
    for (const [, value] of objects?.entries() ?? []) {
      if (!(value instanceof lib.PDFRef)) continue;
      const stream = context.lookup(value);
      if (!(stream instanceof lib.PDFRawStream)) continue;
      const subtype = nameOf(lib, context, stream.dict.get(lib.PDFName.of("Subtype")));
      const key = value.toString();
      if (subtype === "Image") {
        if (seenImages.has(key)) continue;
        seenImages.add(key);
        if (grayscale) imageToGray(lib, context, value);
        const image = context.lookup(value);
        if (!(image instanceof lib.PDFRawStream)) continue;
        const space = nameOf(lib, context, image.dict.get(lib.PDFName.of("ColorSpace")));
        if (space !== "DeviceGray" && space !== "DeviceCMYK") colorImages.add(pageNumber);
        if (image.dict.get(lib.PDFName.of("SMask"))) transparency = true;
      } else if (subtype === "Form" && !seenStreams.has(key)) {
        seenStreams.add(key);
        if (stream.dict.get(lib.PDFName.of("Group"))) transparency = true;
        const resources = asDict(lib, context, stream.dict.get(lib.PDFName.of("Resources")));
        if (rewriteContent(lib, context, value, grayscale)) colorVectors = true;
        walkResources(resources, pageNumber);
      }
    }
  };

  doc.getPages().forEach((page, index) => {
    page.node.delete(lib.PDFName.of("Annots"));
    if (page.node.get(lib.PDFName.of("Group"))) transparency = true;
    for (const ref of contentRefs(lib, page)) {
      const key = ref.toString();
      if (seenStreams.has(key)) continue;
      seenStreams.add(key);
      if (rewriteContent(lib, context, ref, grayscale)) colorVectors = true;
    }
    walkResources(page.node.Resources(), index + 1);
  });

  const problems: string[] = [];
  if (missingFonts.size) {
    problems.push(
      `Fuentes sin incrustar: ${[...missingFonts].join(", ")}. La imprenta rechaza el archivo así: subí esas fuentes.`,
    );
  }
  if (colorImages.size) {
    const pages = [...colorImages].slice(0, 12).join(", ") + (colorImages.size > 12 ? "…" : "");
    problems.push(
      grayscale
        ? `Hay imágenes en color que no se pueden pasar a gris sin recomprimirlas (JPEG en color), en las páginas ${pages}. Pasalas a escala de grises en el Word.`
        : `Hay imágenes en color RGB (páginas ${pages}). Si el interior es en blanco y negro, encendé «Interior en blanco y negro».`,
    );
  }
  if (colorVectors) problems.push("Hay texto o filetes en color.");
  if (transparency) problems.push("Hay transparencias (por ejemplo un PNG con fondo transparente), que PDF/X-1a no admite.");
  return problems;
}

function declarePdfX1a(lib: PdfLib, doc: PDFDocumentType, title: string) {
  const context = doc.context;
  context.header = lib.PDFHeader.forVersion(1, 3);
  const intent = context.obj({
    Type: "OutputIntent",
    S: "GTS_PDFX",
    OutputConditionIdentifier: lib.PDFString.of("CGATS TR 001"),
    OutputCondition: lib.PDFString.of("SWOP (CGATS TR 001)"),
    RegistryName: lib.PDFString.of("http://www.color.org"),
    Info: lib.PDFString.of("U.S. Web Coated (SWOP) v2"),
  });
  doc.catalog.set(lib.PDFName.of("OutputIntents"), context.obj([intent]));
  const info = infoDict(lib, doc);
  info.set(lib.PDFName.of("GTS_PDFXVersion"), lib.PDFString.of("PDF/X-1a:2001"));
  info.set(lib.PDFName.of("GTS_PDFXConformance"), lib.PDFString.of("PDF/X-1a:2001"));
  info.set(lib.PDFName.of("Title"), lib.PDFHexString.fromText(title));
}

function setHeaderVersion(bytes: Uint8Array, version: string) {
  const want = `%PDF-${version}`;
  const head = String.fromCharCode(...bytes.subarray(0, 8));
  if (!/^%PDF-\d\.\d$/.test(head) || want.length !== 8) return;
  for (let i = 0; i < 8; i += 1) bytes[i] = want.charCodeAt(i);
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) if (a[i] !== b[i]) return false;
  return true;
}

/**
 * Cada página se arma como un PDF aparte, así que la misma fuente entera viene una vez por página. Los programas de
 * fuente idénticos (byte por byte) quedan uno solo, compartido: el archivo pesa mucho menos y nada cambia en lo impreso.
 */
function shareFontPrograms(lib: PdfLib, doc: PDFDocumentType) {
  const context = doc.context;
  const kept = new Map<string, Array<{ ref: PDFRefType; bytes: Uint8Array }>>();
  const drop: PDFRefType[] = [];
  for (const [, object] of context.enumerateIndirectObjects()) {
    if (!(object instanceof lib.PDFDict)) continue;
    if (nameOf(lib, context, object.get(lib.PDFName.of("Type"))) !== "FontDescriptor") continue;
    for (const key of ["FontFile", "FontFile2", "FontFile3"]) {
      const ref = object.get(lib.PDFName.of(key));
      if (!(ref instanceof lib.PDFRef)) continue;
      const stream = context.lookup(ref);
      if (!(stream instanceof lib.PDFRawStream)) continue;
      const bytes = stream.contents;
      let hash = 2166136261;
      const step = Math.max(1, Math.floor(bytes.length / 4096));
      for (let i = 0; i < bytes.length; i += step) hash = Math.imul(hash ^ bytes[i]!, 16777619) >>> 0;
      const bucketKey = `${bytes.length}:${hash}`;
      const bucket = kept.get(bucketKey) ?? [];
      const twin = bucket.find((item) => item.ref !== ref && sameBytes(item.bytes, bytes));
      if (twin) {
        object.set(lib.PDFName.of(key), twin.ref);
        drop.push(ref);
      } else if (!bucket.some((item) => item.ref === ref)) {
        bucket.push({ ref, bytes });
        kept.set(bucketKey, bucket);
      }
    }
  }
  for (const ref of drop) context.delete(ref);
}

/** El diccionario Info del archivo (título, fechas, datos PDF/X). */
function infoDict(lib: PdfLib, doc: PDFDocumentType): PDFDictType {
  const info = asDict(lib, doc.context, doc.context.trailerInfo.Info);
  if (info) return info;
  const fresh = doc.context.obj({});
  doc.context.trailerInfo.Info = doc.context.register(fresh);
  return fresh;
}

/** Identificador del archivo (PDF/X lo exige): 16 bytes en hexadecimal. */
function fileId(title: string): string {
  const seed = `${title}|${Date.now()}|${Math.random()}`;
  let out = "";
  for (let round = 0; round < 4; round += 1) {
    let hash = 2166136261 ^ round;
    for (let i = 0; i < seed.length; i += 1) hash = Math.imul(hash ^ seed.charCodeAt(i), 16777619) >>> 0;
    out += hash.toString(16).padStart(8, "0");
  }
  return out;
}
