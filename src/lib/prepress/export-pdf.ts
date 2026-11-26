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
  /**
   * true si quedó alguna fuente OpenType (OTF) entera y se pidió sin recortar: volviendo a exportar con
   * `subset: true` se arregla sola. Con `subset: true` ya pedido queda en false (el aviso sigue en `problems`).
   */
  needsSubset: boolean;
};

export type ExportOptions = {
  title: string;
  subset: boolean;
  cropMarks?: boolean;
  grayscale?: boolean;
  /** Sangrado de 0,125" (9 pt) arriba, abajo y en el borde de afuera (nunca en el lomo). */
  bleed?: boolean;
  /** Para PDF/X: las imágenes con transparencia (PNG sin pérdida) se componen sobre blanco, sin recomprimir. */
  flatten?: boolean;
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
  opts: ExportOptions,
): Promise<PdfExport> {
  const { toPdf } = await import("jsafa-pdf");
  const lib = await import("pdf-lib");
  const { PDFDocument, grayscale } = lib;

  if (sections.length === 0) throw new Error("El documento no tiene páginas para exportar.");

  const merged = await PDFDocument.create();
  merged.setTitle(opts.title);
  merged.setCreator("Pliego");
  merged.setProducer("Pliego premedia");

  let pageIndex = 0;
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
      // El corte es la página del Word; el sangrado y las marcas quedan alrededor, sin mover el contenido.
      const boxes = pressBoxes(pageIndex, width, height, { bleed: opts.bleed, cropMarks: opts.cropMarks });
      pageIndex += 1;
      const { media, trim, bleed } = boxes;
      if (media.x !== 0 || media.y !== 0 || media.width !== width || media.height !== height) {
        // Sin CropBox el visor muestra la MediaBox entera: así se ven el sangrado y las marcas.
        page.node.delete(lib.PDFName.of("CropBox"));
        page.setMediaBox(media.x, media.y, media.width, media.height);
      }
      page.setTrimBox(trim.x, trim.y, trim.width, trim.height);
      page.setBleedBox(bleed.x, bleed.y, bleed.width, bleed.height);
      if (boxes.marks.length) drawCropMarks(page, grayscale, boxes.marks);
      merged.addPage(page);
    }
  }

  const found = prepareForPress(lib, merged, Boolean(opts.grayscale), Boolean(opts.flatten), opts.subset);
  const problems = found.problems;
  shareFontPrograms(lib, merged);
  const pdfx = problems.length === 0;
  if (pdfx) declarePdfX1a(lib, merged, opts.title);
  const id = lib.PDFHexString.of(fileId(opts.title));
  merged.context.trailerInfo.ID = merged.context.obj([id, id]);
  infoDict(lib, merged).set(lib.PDFName.of("Trapped"), lib.PDFName.of("False"));

  const bytes = await merged.save({ useObjectStreams: false });
  // pdf-lib siempre escribe «%PDF-1.7»; PDF/X-1a:2001 pide 1.3. Mismo largo: no se corre ningún desplazamiento.
  setHeaderVersion(bytes, pdfx ? "1.3" : "1.4");
  return { bytes, pdfx, problems, needsSubset: found.wholeOtf && !opts.subset };
}

export type Box = { x: number; y: number; width: number; height: number };
export type MarkLine = [[number, number], [number, number]];

/** 0,125" en puntos: el sangrado que pide Lightning Source. */
export const BLEED = 9;

/**
 * Las cajas de una página, en el mismo sistema de coordenadas del contenido (que no se mueve): el corte es la
 * página del Word; el sangrado agrega 9 pt arriba, abajo y afuera (derecha en las impares — índice 0, 2… —,
 * izquierda en las pares), nunca en el lomo. Las marcas de corte van por fuera del sangrado.
 * Ojo: el sangrado queda vacío salvo que algo del Word llegue hasta el borde de la hoja; el visor no estira nada.
 */
export function pressBoxes(
  index: number,
  width: number,
  height: number,
  opts: { bleed?: boolean; cropMarks?: boolean } = {},
): { media: Box; trim: Box; bleed: Box; marks: MarkLine[] } {
  const trim: Box = { x: 0, y: 0, width, height };
  const recto = index % 2 === 0;
  const bleed: Box = opts.bleed
    ? { x: recto ? 0 : -BLEED, y: -BLEED, width: width + BLEED, height: height + BLEED * 2 }
    : { ...trim };
  if (!opts.cropMarks) return { media: { ...bleed }, trim, bleed, marks: [] };
  // Las marcas siguen las líneas de corte y empiezan donde termina el sangrado (más un respiro).
  const gap = opts.bleed ? BLEED + 3 : 6;
  const len = 14;
  const pad = gap + len;
  const media: Box = { x: -pad, y: -pad, width: width + pad * 2, height: height + pad * 2 };
  const marks: MarkLine[] = [
    [[-pad, 0], [-gap, 0]],
    [[0, -pad], [0, -gap]],
    [[width + gap, 0], [width + pad, 0]],
    [[width, -pad], [width, -gap]],
    [[-pad, height], [-gap, height]],
    [[0, height + gap], [0, height + pad]],
    [[width + gap, height], [width + pad, height]],
    [[width, height + gap], [width, height + pad]],
  ];
  return { media, trim, bleed, marks };
}

/** Marcas de corte en vectores, fuera del recorte (y del sangrado). No tocan las imágenes. */
function drawCropMarks(
  page: import("pdf-lib").PDFPage,
  grayscale: typeof import("pdf-lib").grayscale,
  lines: MarkLine[],
) {
  const ink = grayscale(0);
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
  // Lo que dibuja pdf-lib en esta etapa (marcas de corte, q/Q) no viene comprimido: solo se revisa.
  if (stream instanceof lib.PDFContentStream) return COLOR_LEFT.test(stream.getContentsString());
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
  const image = readFlateImage(lib, context, ref);
  const stream = context.lookup(ref);
  if (!image || image.channels !== 3 || !(stream instanceof lib.PDFRawStream)) return false;
  let gray = rgbToGray(image.pixels);
  if (image.alpha) gray = compositeOnWhite(gray, image.alpha, 1);
  const dict = stream.dict;
  dict.set(lib.PDFName.of("ColorSpace"), lib.PDFName.of("DeviceGray"));
  replaceStream(lib, context, ref, dict, gray, ["SMask", "Decode"]);
  dropMask(context, image);
  return true;
}

/** Aplana la transparencia de una imagen sin pérdida sobre blanco, sin cambiarle el color (RGB o gris). */
function flattenImage(lib: PdfLib, context: PDFContextType, ref: PDFRefType): boolean {
  const image = readFlateImage(lib, context, ref);
  const stream = context.lookup(ref);
  if (!image || !image.alpha || !(stream instanceof lib.PDFRawStream)) return false;
  const flat = compositeOnWhite(image.pixels, image.alpha, image.channels);
  replaceStream(lib, context, ref, stream.dict, flat, ["SMask"]);
  dropMask(context, image);
  return true;
}

/** Cada píxel sobre blanco según su alfa: lo transparente queda blanco, lo opaco queda igual. Sin pérdida (8 bits). */
export function compositeOnWhite(pixels: Uint8Array, alpha: Uint8Array, channels: number): Uint8Array {
  const out = new Uint8Array(pixels.length);
  for (let j = 0; j < alpha.length; j += 1) {
    const a = alpha[j]!;
    for (let c = 0; c < channels; c += 1) {
      const i = j * channels + c;
      out[i] = a === 255 ? pixels[i]! : Math.round((pixels[i]! * a + 255 * (255 - a)) / 255);
    }
  }
  return out;
}

/** RGB de 8 bits → gris por luminancia. */
export function rgbToGray(pixels: Uint8Array): Uint8Array {
  const gray = new Uint8Array(Math.floor(pixels.length / 3));
  for (let i = 0, j = 0; j < gray.length; i += 3, j += 1) {
    gray[j] = Math.round(0.299 * pixels[i]! + 0.587 * pixels[i + 1]! + 0.114 * pixels[i + 2]!);
  }
  return gray;
}

type FlateImage = { pixels: Uint8Array; channels: number; alpha: Uint8Array | null; maskRef: PDFRefType | null };

/**
 * Lee una imagen que se puede tocar sin riesgo: Flate, 8 bits, sin predictor, RGB o gris, y su transparencia
 * (SMask) también así. Si algo no cuadra devuelve undefined y la imagen queda como estaba.
 */
function readFlateImage(lib: PdfLib, context: PDFContextType, ref: PDFRefType): FlateImage | undefined {
  const stream = context.lookup(ref);
  if (!(stream instanceof lib.PDFRawStream)) return undefined;
  const dict = stream.dict;
  if (nameOf(lib, context, dict.get(lib.PDFName.of("Filter"))) !== "FlateDecode") return undefined;
  if (dict.get(lib.PDFName.of("DecodeParms")) || dict.get(lib.PDFName.of("Decode"))) return undefined;
  const space = nameOf(lib, context, dict.get(lib.PDFName.of("ColorSpace")));
  const channels = space === "DeviceRGB" ? 3 : space === "DeviceGray" ? 1 : 0;
  if (!channels) return undefined;
  if (numberOf(lib, context, dict.get(lib.PDFName.of("BitsPerComponent"))) !== 8) return undefined;
  const width = numberOf(lib, context, dict.get(lib.PDFName.of("Width"))) ?? 0;
  const height = numberOf(lib, context, dict.get(lib.PDFName.of("Height"))) ?? 0;
  const pixels = lib.decodePDFRawStream(stream).decode();
  if (pixels.length !== width * height * channels) return undefined;
  const maskValue = dict.get(lib.PDFName.of("SMask"));
  if (!maskValue) return { pixels, channels, alpha: null, maskRef: null };
  if (!(maskValue instanceof lib.PDFRef)) return undefined;
  const mask = context.lookup(maskValue);
  if (!(mask instanceof lib.PDFRawStream)) return undefined;
  const md = mask.dict;
  if (nameOf(lib, context, md.get(lib.PDFName.of("Filter"))) !== "FlateDecode") return undefined;
  if (md.get(lib.PDFName.of("DecodeParms")) || md.get(lib.PDFName.of("Matte"))) return undefined;
  if (numberOf(lib, context, md.get(lib.PDFName.of("BitsPerComponent"))) !== 8) return undefined;
  // pdf-lib escribe «Decode [0 1]», que es lo normal; un alfa invertido no se toca.
  const decode = lookup(lib, context, md.get(lib.PDFName.of("Decode")));
  if (decode instanceof lib.PDFArray) {
    if (numberOf(lib, context, decode.get(0)) !== 0 || numberOf(lib, context, decode.get(1)) !== 1) return undefined;
  } else if (decode) return undefined;
  const alpha = lib.decodePDFRawStream(mask).decode();
  if (alpha.length !== width * height) return undefined;
  return { pixels, channels, alpha, maskRef: maskValue };
}

/** El alfa ya no lo usa nadie (pdf-lib arma uno por imagen): se borra para no dejarlo suelto en el archivo. */
function dropMask(context: PDFContextType, image: FlateImage) {
  if (image.maskRef) context.delete(image.maskRef);
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

/** Una OTF (CFF) entera se incrusta como archivo OpenType completo donde PDF espera solo el programa CFF. */
function wholeOpenTypeCff(lib: PdfLib, context: PDFContextType, font: PDFDictType): boolean {
  let target = font;
  if (nameOf(lib, context, font.get(lib.PDFName.of("Subtype"))) === "Type0") {
    const descendants = lookup(lib, context, font.get(lib.PDFName.of("DescendantFonts")));
    const first = descendants instanceof lib.PDFArray ? asDict(lib, context, descendants.get(0)) : undefined;
    if (!first) return false;
    target = first;
  }
  const descriptor = asDict(lib, context, target.get(lib.PDFName.of("FontDescriptor")));
  const file = descriptor ? lookup(lib, context, descriptor.get(lib.PDFName.of("FontFile3"))) : undefined;
  if (!(file instanceof lib.PDFRawStream)) return false;
  const head = lib.decodePDFRawStream(file).decode().subarray(0, 4);
  return String.fromCharCode(...head) === "OTTO";
}

/**
 * Deja el PDF listo para imprenta sin tocar el texto ni recomprimir fotos: neutros en gris, sin anotaciones y, en
 * blanco y negro, las imágenes sin pérdida pasadas a gris. Devuelve lo que impide declararlo PDF/X-1a.
 */
function prepareForPress(
  lib: PdfLib,
  doc: PDFDocumentType,
  grayscale: boolean,
  flatten: boolean,
  subset: boolean,
): { problems: string[]; wholeOtf: boolean } {
  const context = doc.context;
  const missingFonts = new Set<string>();
  const wholeOtf = new Set<string>();
  const colorImages = new Set<number>();
  const deepImages = new Set<number>();
  const maskedImages = new Set<number>();
  let colorVectors = false;
  let opacity = false;
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
      } else if (font && wholeOpenTypeCff(lib, context, font)) {
        wholeOtf.add(nameOf(lib, context, font.get(lib.PDFName.of("BaseFont"))) ?? "sin nombre");
      }
    }
    const states = asDict(lib, context, resources.get(lib.PDFName.of("ExtGState")));
    for (const [, value] of states?.entries() ?? []) {
      const state = asDict(lib, context, value);
      if (!state) continue;
      for (const key of ["CA", "ca"]) {
        const alpha = numberOf(lib, context, state.get(lib.PDFName.of(key)));
        // La opacidad de CSS (jsafa la escribe así) no se aplana: solo se avisa.
        if (alpha !== undefined && alpha < 1) opacity = true;
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
        if (flatten) flattenImage(lib, context, value);
        const image = context.lookup(value);
        if (!(image instanceof lib.PDFRawStream)) continue;
        const space = nameOf(lib, context, image.dict.get(lib.PDFName.of("ColorSpace")));
        if (space !== "DeviceGray" && space !== "DeviceCMYK") colorImages.add(pageNumber);
        // Hoy jsafa escribe siempre 8 bits (los PNG pasan por RGBA de 8 bits); se revisa por si eso cambia.
        const bits = numberOf(lib, context, image.dict.get(lib.PDFName.of("BitsPerComponent")));
        if (bits !== undefined && bits > 8) deepImages.add(pageNumber);
        if (image.dict.get(lib.PDFName.of("SMask"))) maskedImages.add(pageNumber);
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
  if (wholeOtf.size) {
    const names = [...wholeOtf].join(", ");
    problems.push(
      subset
        ? `Fuentes OpenType (OTF) que no se pudieron recortar y quedaron enteras: ${names}. Así no cumplen PDF/X: probá con la versión TTF de esas fuentes.`
        : `Fuentes OpenType (OTF) incrustadas enteras: ${names}. Así no cumplen PDF/X: apagá «Incrustar la fuente entera».`,
    );
  }
  if (colorImages.size) {
    const pages = pageList(colorImages);
    problems.push(
      grayscale
        ? `Hay imágenes en color que no se pueden pasar a gris sin recomprimirlas (JPEG en color), en las páginas ${pages}. Pasalas a escala de grises en el Word.`
        : `Hay imágenes en color RGB (páginas ${pages}). Si el interior es en blanco y negro, encendé «Interior en blanco y negro».`,
    );
  }
  if (colorVectors) problems.push("Hay texto o filetes en color.");
  if (deepImages.size) {
    problems.push(
      `Hay imágenes de 16 bits por color (páginas ${pageList(deepImages)}), que PDF/X-1a no admite. Guardalas en 8 bits antes de ponerlas en el Word.`,
    );
  }
  if (maskedImages.size) {
    problems.push(
      flatten
        ? `Hay imágenes con transparencia que no se pudieron poner sobre blanco sin recomprimirlas (páginas ${pageList(maskedImages)}). Ponelas sobre fondo blanco en el Word.`
        : `Hay imágenes con transparencia, por ejemplo un PNG con fondo transparente (páginas ${pageList(maskedImages)}), que PDF/X-1a no admite.`,
    );
  }
  if (opacity) {
    problems.push("Hay texto, filetes o imágenes semitransparentes (con opacidad), que PDF/X-1a no admite. Dejalos opacos en el Word.");
  }
  if (transparency) problems.push("Hay transparencias (máscaras o grupos transparentes), que PDF/X-1a no admite.");
  return { problems, wholeOtf: wholeOtf.size > 0 };
}

/** Las primeras 12 páginas de la lista, para el aviso. */
function pageList(pages: Set<number>): string {
  return [...pages].slice(0, 12).join(", ") + (pages.size > 12 ? "…" : "");
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
  // ISO 15930-1: la versión es «PDF/X-1:2001» y la conformidad «PDF/X-1a:2001».
  info.set(lib.PDFName.of("GTS_PDFXVersion"), lib.PDFString.of("PDF/X-1:2001"));
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
