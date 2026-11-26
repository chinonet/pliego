import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { extractCollectionFace, fontKey, parseFontFaces } from "./font-face.ts";

const carlito = new Uint8Array(await readFile(new URL("../../../public/fonts/Carlito-Regular.ttf", import.meta.url)));
const regular = parseFontFaces(carlito);
assert.equal(regular.length, 1);
assert.equal(regular[0]?.family, "Carlito");
assert.equal(regular[0]?.style, "normal");
assert.ok((regular[0]?.weight ?? 0) >= 400 && (regular[0]?.weight ?? 0) < 500);

const boldItalic = parseFontFaces(
  new Uint8Array(await readFile(new URL("../../../public/fonts/Carlito-BoldItalic.ttf", import.meta.url))),
);
assert.equal(boldItalic[0]?.style, "italic");
assert.ok((boldItalic[0]?.weight ?? 0) >= 700);

assert.equal(fontKey("Calibri Light"), fontKey("calibri-light"));
assert.equal(fontKey("Minion Pro"), fontKey("MinionPro"));

// Una colección TTC (como cambria.ttc): cada cara se puede sacar como archivo propio, con las mismas tablas.
const bold = new Uint8Array(await readFile(new URL("../../../public/fonts/Carlito-Bold.ttf", import.meta.url)));
function tablesOf(font: Uint8Array, base = 0): Map<string, Uint8Array> {
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  const out = new Map<string, Uint8Array>();
  for (let i = 0; i < view.getUint16(base + 4); i++) {
    const rec = base + 12 + i * 16;
    const name = String.fromCharCode(...font.subarray(rec, rec + 4));
    out.set(name, font.subarray(view.getUint32(rec + 8), view.getUint32(rec + 8) + view.getUint32(rec + 12)));
  }
  return out;
}
function collection(fonts: Uint8Array[]): Uint8Array {
  // Cabecera ttcf, después cada fuente con sus desplazamientos corridos al lugar nuevo.
  const header = 12 + fonts.length * 4;
  const total = header + fonts.reduce((sum, font) => sum + font.length, 0);
  const out = new Uint8Array(total);
  const ov = new DataView(out.buffer);
  out.set([0x74, 0x74, 0x63, 0x66, 0, 1, 0, 0], 0);
  ov.setUint32(8, fonts.length);
  let at = header;
  fonts.forEach((font, i) => {
    ov.setUint32(12 + i * 4, at);
    out.set(font, at);
    const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
    for (let t = 0; t < view.getUint16(4); t++) ov.setUint32(at + 12 + t * 16 + 8, view.getUint32(12 + t * 16 + 8) + at);
    at += font.length;
  });
  return out;
}
const ttc = collection([carlito, bold]);
const inCollection = parseFontFaces(ttc);
assert.deepEqual(
  inCollection.map((face) => [face.family, face.weight >= 700]),
  [
    ["Carlito", false],
    ["Carlito", true],
  ],
);
const second = extractCollectionFace(ttc, 1);
assert.equal(String.fromCharCode(...second.subarray(0, 4)), String.fromCharCode(...bold.subarray(0, 4)));
assert.equal(parseFontFaces(second).length, 1, "ya no es una colección");
assert.ok(parseFontFaces(second)[0]!.weight >= 700);
const original = tablesOf(bold);
const copied = tablesOf(second);
assert.deepEqual([...copied.keys()].sort(), [...original.keys()].sort());
for (const [name, data] of original) {
  if (name === "head") continue; // cambia solo la suma de control del archivo
  assert.deepEqual(copied.get(name), data, name);
}
assert.deepEqual(copied.get("head")!.subarray(12), original.get("head")!.subarray(12));
// La suma de control del archivo nuevo da lo que pide el formato.
{
  const view = new DataView(second.buffer, second.byteOffset, second.byteLength);
  let sum = 0;
  for (let i = 0; i < second.length; i += 4) sum = (sum + view.getUint32(i)) >>> 0;
  assert.equal(sum, 0xb1b0afba);
}
assert.equal(extractCollectionFace(carlito, 0), carlito, "una fuente suelta queda igual");
assert.throws(() => extractCollectionFace(ttc, 2));

console.log("font-face ok", regular[0]);
