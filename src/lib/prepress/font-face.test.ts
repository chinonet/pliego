import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { fontKey, parseFontFaces } from "./font-face.ts";

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

console.log("font-face ok", regular[0]);
