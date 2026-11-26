export type ParsedFace = {
  family: string;
  /** El nombre de familia de siempre (nameID 1), el que usa Word: «Calibri Light» aunque la familia sea «Calibri». */
  legacyFamily?: string;
  subfamily: string;
  weight: number;
  style: "normal" | "italic";
};

const WEIGHTS: Array<[RegExp, number]> = [
  [/extra.?light|ultra.?light/, 200],
  [/semi.?bold|demi.?bold|demibold/, 600],
  [/extra.?bold|ultra.?bold/, 800],
  [/thin|hairline/, 100],
  [/light/, 300],
  [/medium/, 500],
  [/black|heavy/, 900],
  [/bold/, 700],
];

export function fontKey(name: string): string {
  return name
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "");
}

export function classifySubfamily(subfamily: string): { weight: number; style: "normal" | "italic" } {
  const text = subfamily.toLowerCase();
  const style = /italic|oblique|kursiv|italica/.test(text) ? "italic" : "normal";
  const weight = WEIGHTS.find(([pattern]) => pattern.test(text))?.[1] ?? 400;
  return { weight, style };
}

function tag(bytes: Uint8Array, offset: number): string {
  return String.fromCharCode(bytes[offset] ?? 0, bytes[offset + 1] ?? 0, bytes[offset + 2] ?? 0, bytes[offset + 3] ?? 0);
}

function decodeName(platform: number, data: Uint8Array): string {
  if (platform === 3 || platform === 0) {
    let text = "";
    for (let i = 0; i + 1 < data.length; i += 2) text += String.fromCharCode((data[i]! << 8) | data[i + 1]!);
    return text.replace(/\0/g, "").trim();
  }
  return new TextDecoder("latin1").decode(data).replace(/\0/g, "").trim();
}

function parseSfnt(bytes: Uint8Array, base: number): ParsedFace {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const numTables = view.getUint16(base + 4);
  let nameOffset = 0;
  let nameLength = 0;
  let os2Offset = 0;
  for (let i = 0; i < numTables; i++) {
    const rec = base + 12 + i * 16;
    const id = tag(bytes, rec);
    const offset = view.getUint32(rec + 8);
    const length = view.getUint32(rec + 12);
    if (id === "name") {
      nameOffset = offset;
      nameLength = length;
    } else if (id === "OS/2") os2Offset = offset;
  }
  if (!nameOffset) throw new Error("El archivo no tiene tabla de nombre.");

  const count = view.getUint16(nameOffset + 2);
  const stringOffset = view.getUint16(nameOffset + 4);
  const names = new Map<number, string>();
  for (let i = 0; i < count; i++) {
    const rec = nameOffset + 6 + i * 12;
    const platform = view.getUint16(rec);
    const nameId = view.getUint16(rec + 6);
    const length = view.getUint16(rec + 8);
    const offset = view.getUint16(rec + 10);
    if (![1, 2, 16, 17].includes(nameId)) continue;
    const start = nameOffset + stringOffset + offset;
    if (start + length > nameOffset + nameLength) continue;
    const value = decodeName(platform, bytes.subarray(start, start + length));
    if (!value) continue;
    const rank = platform === 3 ? 2 : platform === 0 ? 1 : 0;
    const previous = names.get(nameId);
    if (!previous || rank >= 1) names.set(nameId, value);
  }

  const family = names.get(16) || names.get(1);
  const subfamily = names.get(17) || names.get(2) || "Regular";
  if (!family) throw new Error("No pude leer el nombre de la fuente.");

  let weight = classifySubfamily(subfamily).weight;
  let style = classifySubfamily(subfamily).style;
  if (os2Offset && os2Offset + 64 < bytes.byteLength) {
    const os2Weight = view.getUint16(os2Offset + 4);
    const selection = view.getUint16(os2Offset + 62);
    if (os2Weight >= 100 && os2Weight <= 1000) weight = os2Weight;
    if (selection & 1) style = "italic";
  }

  const legacy = names.get(1);
  return { family, subfamily, weight, style, ...(legacy && legacy !== family ? { legacyFamily: legacy } : {}) };
}

/**
 * Una cara de una colección (TTC/OTC, como «cambria.ttc», que trae Cambria y Cambria Math) como archivo TTF/OTF
 * propio: las mismas tablas, copiadas byte a byte, con su propio índice. El navegador y el PDF no siempre saben elegir
 * una cara dentro de una colección. Si no es una colección, devuelve los mismos bytes.
 */
export function extractCollectionFace(bytes: Uint8Array, index: number): Uint8Array {
  if (tag(bytes, 0) !== "ttcf") return bytes;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const count = view.getUint32(8);
  if (index < 0 || index >= count) throw new Error("La colección no tiene esa cara.");
  const base = view.getUint32(12 + index * 4);
  const numTables = view.getUint16(base + 4);
  const records: Array<{ offset: number; length: number; tag: string }> = [];
  for (let i = 0; i < numTables; i++) {
    const rec = base + 12 + i * 16;
    const offset = view.getUint32(rec + 8);
    const length = view.getUint32(rec + 12);
    if (offset + length > bytes.length) throw new Error("La colección está dañada.");
    records.push({ offset, length, tag: tag(bytes, rec) });
  }
  const header = 12 + numTables * 16;
  const size = records.reduce((sum, record) => sum + ((record.length + 3) & ~3), header);
  const out = new Uint8Array(size);
  const ov = new DataView(out.buffer);
  out.set(bytes.subarray(base, base + 12), 0);
  let at = header;
  let headAt = -1;
  records.forEach((record, i) => {
    const rec = 12 + i * 16;
    out.set(bytes.subarray(base + 12 + i * 16, base + 12 + i * 16 + 8), rec); // etiqueta y suma de la tabla
    ov.setUint32(rec + 8, at);
    ov.setUint32(rec + 12, record.length);
    out.set(bytes.subarray(record.offset, record.offset + record.length), at);
    if (record.tag === "head") headAt = at;
    at += (record.length + 3) & ~3;
  });
  // La suma de control del archivo entero (head.checkSumAdjustment) cambia con el archivo nuevo.
  if (headAt >= 0 && headAt + 12 <= out.length) {
    ov.setUint32(headAt + 8, 0);
    let sum = 0;
    for (let i = 0; i < out.length; i += 4) sum = (sum + ov.getUint32(i)) >>> 0;
    ov.setUint32(headAt + 8, (0xb1b0afba - sum) >>> 0);
  }
  return out;
}

/** Lee TTF, OTF y colecciones TTC. No recomprime ni convierte el archivo. */
export function parseFontFaces(bytes: Uint8Array): ParsedFace[] {
  if (bytes.byteLength < 12) throw new Error("Archivo de fuente vacío.");
  const kind = tag(bytes, 0);
  if (kind === "wOFF" || kind === "wOF2") {
    throw new Error("Sube la fuente en TTF u OTF, no en WOFF.");
  }
  if (kind === "ttcf") {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const count = view.getUint32(8);
    const faces: ParsedFace[] = [];
    for (let i = 0; i < count; i++) faces.push(parseSfnt(bytes, view.getUint32(12 + i * 4)));
    return faces;
  }
  return [parseSfnt(bytes, 0)];
}
