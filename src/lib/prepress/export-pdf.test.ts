import assert from "node:assert/strict";
import { BLEED, compositeOnWhite, pressBoxes, rgbToGray } from "./export-pdf.ts";

// Sin sangrado ni marcas: todo es la página del Word.
const plain = pressBoxes(0, 432, 648);
assert.deepEqual(plain.trim, { x: 0, y: 0, width: 432, height: 648 });
assert.deepEqual(plain.bleed, plain.trim);
assert.deepEqual(plain.media, plain.trim);
assert.equal(plain.marks.length, 0);

// Impar (índice 0): sangra a la derecha, arriba y abajo; el lomo (izquierda) queda en 0.
const recto = pressBoxes(0, 432, 648, { bleed: true });
assert.equal(BLEED, 9);
assert.deepEqual(recto.trim, { x: 0, y: 0, width: 432, height: 648 });
assert.deepEqual(recto.bleed, { x: 0, y: -9, width: 441, height: 666 });
assert.deepEqual(recto.media, recto.bleed);

// Par (índice 1): sangra a la izquierda.
const verso = pressBoxes(1, 432, 648, { bleed: true });
assert.deepEqual(verso.bleed, { x: -9, y: -9, width: 441, height: 666 });
assert.deepEqual(verso.media, verso.bleed);
assert.deepEqual(pressBoxes(2, 432, 648, { bleed: true }).bleed, recto.bleed);

// Solo marcas: como antes (6 pt de respiro, 14 de largo).
const marked = pressBoxes(0, 432, 648, { cropMarks: true });
assert.deepEqual(marked.media, { x: -20, y: -20, width: 472, height: 688 });
assert.deepEqual(marked.bleed, marked.trim);
assert.equal(marked.marks.length, 8);
assert.deepEqual(marked.marks[0], [[-20, 0], [-6, 0]]);

// Sangrado y marcas: las marcas arrancan fuera del sangrado y la MediaBox contiene a la BleedBox.
const both = pressBoxes(1, 432, 648, { bleed: true, cropMarks: true });
const inside = (outer: { x: number; y: number; width: number; height: number }, inner: typeof outer) =>
  inner.x >= outer.x &&
  inner.y >= outer.y &&
  inner.x + inner.width <= outer.x + outer.width &&
  inner.y + inner.height <= outer.y + outer.height;
assert.ok(inside(both.media, both.bleed));
assert.ok(inside(both.bleed, both.trim));
for (const [start, end] of both.marks) {
  for (const [x, y] of [start, end]) {
    const outX = x <= -BLEED || x >= 432 + BLEED;
    const outY = y <= -BLEED || y >= 648 + BLEED;
    assert.ok(outX || outY, `marca dentro del sangrado: ${x},${y}`);
    assert.ok(x >= both.media.x && x <= both.media.x + both.media.width);
    assert.ok(y >= both.media.y && y <= both.media.y + both.media.height);
  }
}

// Transparencia sobre blanco, sin pérdida: opaco igual, transparente blanco, mitad a mitad.
assert.deepEqual(
  [...compositeOnWhite(new Uint8Array([10, 20, 30, 10, 20, 30, 0, 0, 0]), new Uint8Array([255, 0, 128]), 3)],
  [10, 20, 30, 255, 255, 255, 127, 127, 127],
);
assert.deepEqual([...compositeOnWhite(new Uint8Array([0, 200]), new Uint8Array([51, 255]), 1)], [204, 200]);

// RGB → gris por luminancia.
assert.deepEqual([...rgbToGray(new Uint8Array([255, 255, 255, 0, 0, 0, 255, 0, 0]))], [255, 0, 76]);

console.log("export-pdf ok");
