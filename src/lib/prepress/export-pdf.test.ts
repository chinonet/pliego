import assert from "node:assert/strict";
import jpeg from "jpeg-js";
import * as lib from "pdf-lib";
import {
  BLEED,
  compositeOnWhite,
  decodeJpegRgb,
  drawnNames,
  fileSize,
  imagesOverPaint,
  jpegsToGray,
  pageList,
  type PaintOp,
  prepareForPress,
  pressBoxes,
  rgbToGray,
} from "./export-pdf.ts";

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

// ---- Fotos JPEG en color → gris sin pérdida ----

/** Un JPEG RGB real (degradé de colores) de w × h. */
function colorJpeg(w: number, h: number, quality = 90): Uint8Array {
  const data = new Uint8Array(w * h * 4);
  for (let y = 0; y < h; y += 1) {
    for (let x = 0; x < w; x += 1) {
      const i = (y * w + x) * 4;
      data[i] = Math.round((255 * x) / (w - 1));
      data[i + 1] = Math.round((255 * y) / (h - 1));
      data[i + 2] = 255 - data[i]!;
      data[i + 3] = 255;
    }
  }
  return new Uint8Array(jpeg.encode({ width: w, height: h, data }, quality).data);
}

/** Agrega un segmento APP14 «Adobe» con ese transform justo después del SOI. */
function withAdobe(bytes: Uint8Array, transform: number): Uint8Array {
  const payload = [0x41, 0x64, 0x6f, 0x62, 0x65, 0, 100, 0, 0, 0, 0, transform];
  const segment = [0xff, 0xee, 0, payload.length + 2, ...payload];
  return new Uint8Array([...bytes.subarray(0, 2), ...segment, ...bytes.subarray(2)]);
}

const photo = colorJpeg(64, 48);
const decodedPhoto = (await decodeJpegRgb(photo))!;
assert.equal(decodedPhoto.width, 64);
assert.equal(decodedPhoto.height, 48);
assert.equal(decodedPhoto.rgb.length, 64 * 48 * 3);
// Lo mismo que jpeg-js con sus opciones normales: sin marcador Adobe es YCbCr.
const plainDecode = jpeg.decode(photo, { useTArray: true, formatAsRGBA: false });
assert.deepEqual(decodedPhoto.rgb, plainDecode.data);
// El color se parece al original (un JPEG al 90 % no se aleja mucho).
const corner = decodedPhoto.rgb.subarray((47 * 64 + 63) * 3, (47 * 64 + 63) * 3 + 3);
assert.ok(Math.abs(corner[0]! - 255) < 24 && Math.abs(corner[1]! - 255) < 24 && corner[2]! < 24, `esquina: ${[...corner]}`);
// Con marcador Adobe transform 1 (YCbCr) da lo mismo; con transform 0 los canales se leen tal cual (RGB).
assert.deepEqual((await decodeJpegRgb(withAdobe(photo, 1)))!.rgb, decodedPhoto.rgb);
const raw = (await decodeJpegRgb(withAdobe(photo, 0)))!.rgb;
assert.deepEqual(raw, jpeg.decode(withAdobe(photo, 0), { useTArray: true, formatAsRGBA: false, colorTransform: false }).data);
assert.notDeepEqual(raw, decodedPhoto.rgb);
// Lo que no es un JPEG RGB de 8 bits no se decodifica.
assert.equal(await decodeJpegRgb(new Uint8Array([1, 2, 3])), null);
assert.equal(await decodeJpegRgb(photo.subarray(0, photo.length - 200)), null);

/** Un PDF de dos páginas: la foto en la 1 y en un formulario de la 2, y otra foto distinta suelta en la 2. */
async function photoPdf() {
  const doc = await lib.PDFDocument.create();
  const first = await doc.embedJpg(photo);
  const second = await doc.embedJpg(colorJpeg(32, 32, 75));
  const page1 = doc.addPage([200, 200]);
  page1.drawImage(first, { x: 10, y: 10, width: 64, height: 48 });
  const page2 = doc.addPage([200, 200]);
  page2.drawImage(second, { x: 10, y: 10, width: 32, height: 32 });
  // Un formulario (XObject Form) en la página 2 que usa la misma foto de la página 1.
  const form = doc.context.formXObject([], {
    BBox: [0, 0, 100, 100],
    Resources: { XObject: { Fp: first.ref } },
  });
  const formRef = doc.context.register(form);
  const resources = page2.node.Resources()!;
  const objects = resources.lookup(lib.PDFName.of("XObject"), lib.PDFDict);
  objects.set(lib.PDFName.of("Fx"), formRef);
  // pdf-lib mete las fotos recién al guardar; en la app el PDF también llega ya guardado (de jsafa).
  const loaded = await lib.PDFDocument.load(await doc.save({ useObjectStreams: false }));
  return { doc: loaded, first: first.ref, second: second.ref };
}

const imageDict = (doc: lib.PDFDocument, ref: lib.PDFRef) => {
  const stream = doc.context.lookup(ref);
  assert.ok(stream instanceof lib.PDFRawStream);
  return stream;
};

{
  const { doc, first, second } = await photoPdf();
  const before = imageDict(doc, first).contents.length + imageDict(doc, second).contents.length;
  const done = await jpegsToGray(lib, doc);
  assert.deepEqual([...done.pages], [1, 2]);
  assert.equal(done.before, before);
  for (const ref of [first, second]) {
    const stream = imageDict(doc, ref);
    assert.equal(stream.dict.get(lib.PDFName.of("ColorSpace"))?.toString(), "/DeviceGray");
    assert.equal(stream.dict.get(lib.PDFName.of("Filter"))?.toString(), "/FlateDecode");
    assert.equal(stream.dict.get(lib.PDFName.of("Subtype"))?.toString(), "/Image");
  }
  assert.ok(done.after > 0);
  // Sin pérdida: el gris guardado es exactamente la luminancia de lo que se decodificó.
  const gray = lib.decodePDFRawStream(imageDict(doc, first)).decode();
  assert.deepEqual(gray, rgbToGray(decodedPhoto.rgb));
  assert.equal(gray.length, 64 * 48);
  // Ya en gris: prepareForPress no la cuenta como imagen en color.
  const press = prepareForPress(lib, doc, true, true, true, new Set(), { jpegToGray: true });
  assert.ok(!press.problems.some((p) => p.includes("color")), press.problems.join(" | "));
  // Y el PDF se guarda y se vuelve a abrir.
  const reopened = await lib.PDFDocument.load(await doc.save());
  assert.equal(reopened.getPageCount(), 2);
}

{
  // Con el mismo cache, la segunda exportación no decodifica nada y da exactamente el mismo gris.
  const cache = new Map<string, Uint8Array>();
  let calls = 0;
  const counting = async (bytes: Uint8Array) => {
    calls += 1;
    return decodeJpegRgb(bytes);
  };
  const one = await photoPdf();
  const firstRun = await jpegsToGray(lib, one.doc, counting, cache);
  assert.equal(calls, 2);
  assert.equal(cache.size, 2);
  const two = await photoPdf();
  const secondRun = await jpegsToGray(lib, two.doc, counting, cache);
  assert.equal(calls, 2, "la segunda vez sale del cache");
  assert.deepEqual([...secondRun.pages], [...firstRun.pages]);
  assert.equal(secondRun.after, firstRun.after);
  for (const ref of [two.first, two.second]) {
    const a = imageDict(one.doc, ref);
    const b = imageDict(two.doc, ref);
    assert.deepEqual(lib.decodePDFRawStream(b).decode(), lib.decodePDFRawStream(a).decode());
    assert.equal(b.dict.get(lib.PDFName.of("ColorSpace"))?.toString(), "/DeviceGray");
    assert.equal(b.dict.get(lib.PDFName.of("Filter"))?.toString(), "/FlateDecode");
    assert.equal(b.dict.get(lib.PDFName.of("Width"))?.toString(), a.dict.get(lib.PDFName.of("Width"))?.toString());
  }
  const reopened = await lib.PDFDocument.load(await two.doc.save());
  assert.equal(reopened.getPageCount(), 2);
}

{
  // Una máscara de color de 3 canales no vale en gris: esa foto queda como estaba.
  const { doc, first } = await photoPdf();
  imageDict(doc, first).dict.set(lib.PDFName.of("Mask"), doc.context.obj([250, 255, 250, 255, 250, 255]));
  const done = await jpegsToGray(lib, doc);
  assert.deepEqual([...done.pages], [2]);
  assert.equal(imageDict(doc, first).dict.get(lib.PDFName.of("ColorSpace"))?.toString(), "/DeviceRGB");
}

{
  // Sin la opción, la foto JPEG en color queda igual y la app sugiere encenderla.
  const { doc, first } = await photoPdf();
  const original = imageDict(doc, first).contents;
  const press = prepareForPress(lib, doc, true, true, true);
  assert.equal(imageDict(doc, first).contents, original);
  assert.ok(press.problems.some((p) => p.includes("«Pasar también los JPEG a gris»")), press.problems.join(" | "));
  // En color (sin blanco y negro) el aviso es el de siempre.
  assert.ok(prepareForPress(lib, doc, false, true, true).problems.some((p) => p.includes("«Interior en blanco y negro»")));
}

{
  // Lo que no se puede leer sin riesgo queda como estaba: CMYK, con Decode, un decodificador que no puede o que da otro tamaño.
  const { doc, first, second } = await photoPdf();
  imageDict(doc, first).dict.set(lib.PDFName.of("ColorSpace"), lib.PDFName.of("DeviceCMYK"));
  imageDict(doc, second).dict.set(lib.PDFName.of("Decode"), doc.context.obj([0, 1, 0, 1, 0, 1]));
  const done = await jpegsToGray(lib, doc);
  assert.equal(done.pages.size, 0);
  assert.equal(imageDict(doc, first).dict.get(lib.PDFName.of("Filter"))?.toString(), "/DCTDecode");
  assert.equal(imageDict(doc, second).dict.get(lib.PDFName.of("Filter"))?.toString(), "/DCTDecode");
}
{
  const { doc, first } = await photoPdf();
  const fails = await jpegsToGray(lib, doc, async () => null);
  assert.equal(fails.pages.size, 0);
  const wrong = await jpegsToGray(lib, doc, async () => ({ width: 2, height: 2, rgb: new Uint8Array(12) }));
  assert.equal(wrong.pages.size, 0);
  assert.equal(imageDict(doc, first).dict.get(lib.PDFName.of("ColorSpace"))?.toString(), "/DeviceRGB");
  // Con la opción encendida pero sin poder pasarla, el aviso lo dice.
  const press = prepareForPress(lib, doc, true, true, true, new Set(), { jpegToGray: true });
  assert.ok(press.problems.some((p) => p.includes("no se pudieron pasar a gris")), press.problems.join(" | "));
}

// Páginas y tamaños en palabras.
assert.equal(pageList(new Set([2])), "página 2");
assert.equal(pageList(new Set([9, 2, 5])), "páginas 2, 5, 9");
assert.equal(pageList(new Set(Array.from({ length: 14 }, (_, i) => i + 1))), "páginas 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12…");
assert.equal(fileSize(28751), "28 KB");
assert.equal(fileSize(100), "1 KB");
assert.equal(fileSize(2.5 * 1024 * 1024), "2,5 MB");
assert.equal(fileSize(18.4 * 1024 * 1024), "18 MB");

console.log("export-pdf ok");
