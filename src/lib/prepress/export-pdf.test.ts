import assert from "node:assert/strict";
import { BLEED, compositeOnWhite, drawnNames, imagesOverPaint, type PaintOp, pressBoxes, rgbToGray } from "./export-pdf.ts";

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

// Qué imágenes tienen algo pintado debajo (y no se pueden poner sobre blanco).
const white = { r: 1, g: 1, b: 1, a: 1 };
const gray = { r: 0.8, g: 0.8, b: 0.8, a: 1 };
const noSide = { width: 0, color: { r: 0, g: 0, b: 0, a: 0 } };
const rect = (x: number, y: number, width: number, height: number) => ({ x, y, width, height });
const box = (r: ReturnType<typeof rect>, background: typeof white | null, side = noSide): PaintOp => ({
  kind: "box",
  box: { rect: r, background, top: side, right: side, bottom: side, left: side },
});
const image = (r: ReturnType<typeof rect>, transform: number[] | null = null): PaintOp => ({
  kind: "image",
  image: { draw: r },
  transform,
});
assert.deepEqual(
  imagesOverPaint([
    box(rect(0, 0, 600, 800), white), // la hoja blanca no cuenta
    image(rect(10, 10, 100, 100)),
    box(rect(200, 200, 200, 200), gray), // celda sombreada
    image(rect(250, 250, 50, 50)), // encima de la celda
    image(rect(400, 10, 50, 50)), // fuera de todo
    { kind: "text", run: { rect: rect(10, 500, 300, 20), style: { color: { r: 0, g: 0, b: 0, a: 1 } } } },
    image(rect(100, 490, 40, 40)), // encima del texto
    image(rect(10, 110, 50, 50)), // apenas toca la primera imagen en el borde: no cuenta
    image(rect(60, 60, 30, 30)), // encima de la primera imagen
  ]),
  [false, true, false, true, false, true],
);
// Lo dibujado después (encima de la imagen) no cuenta; un borde de color sí, pero solo donde está el borde.
const framed = box(rect(0, 0, 200, 200), null, { width: 2, color: { r: 0, g: 0, b: 0, a: 1 } });
assert.deepEqual(imagesOverPaint([image(rect(0, 0, 10, 10)), framed, image(rect(50, 50, 20, 20)), image(rect(0, 50, 10, 10))]), [false, false, true]);
// Texto transparente no pinta; una imagen girada ocupa su caja girada.
assert.deepEqual(
  imagesOverPaint([
    { kind: "text", run: { rect: rect(0, 0, 100, 20), style: { color: { r: 0, g: 0, b: 0, a: 0 } } } },
    image(rect(0, 0, 50, 10)),
    box(rect(300, 0, 20, 100), gray),
    image(rect(0, 0, 100, 10), [0, 1, -1, 0, 310, 0]), // girada 90°: queda en x 300…310
  ]),
  [false, true],
);

// Los «Do» del contenido, en orden y con los nombres decodificados.
assert.deepEqual(drawnNames("q 1 0 0 1 0 0 cm /Image-12-0 Do Q\nq /GS-1 gs /Im#20b Do Q /X Do"), ["Image-12-0", "Im b", "X"]);
assert.deepEqual(drawnNames("BT <0012> Tj ET /Fx Dox"), []);

console.log("export-pdf ok");
