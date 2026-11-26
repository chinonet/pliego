import assert from "node:assert/strict";
import { layoutPadding, planBlanks, planFootnotes, wrongSide, type PlanPage } from "./paginate.ts";

// Márgenes: 72 px a la izquierda (interior), 48 a la derecha; encuadernación de 18 pt = 24 px.
const base = { left: 72, right: 48, top: 60 };
assert.deepEqual(layoutPadding(base, 0, 0, false, false), base);
assert.deepEqual(layoutPadding(base, 1, 0, true, false), { left: 48, right: 72, top: 60 }, "par con simétricos: se invierten");
assert.deepEqual(layoutPadding(base, 0, 18, true, false), { left: 96, right: 48, top: 60 }, "impar: encuadernación a la izquierda");
assert.deepEqual(layoutPadding(base, 1, 18, true, false), { left: 48, right: 96, top: 60 }, "par: encuadernación a la derecha");
assert.deepEqual(layoutPadding(base, 1, 18, false, false), { left: 96, right: 48, top: 60 }, "sin simétricos: siempre a la izquierda");
assert.deepEqual(layoutPadding(base, 1, 18, false, true), { left: 72, right: 48, top: 84 }, "encuadernación arriba");
assert.deepEqual(layoutPadding(base, 1, 18, true, true), { left: 48, right: 96, top: 60 }, "con simétricos manda el interior");

// Secciones: 0 = portada (1 página), 1 = capítulo impar (3 páginas), 2 = capítulo impar.
const pages = (sizes: number[]): PlanPage[] =>
  sizes.flatMap((size, section) => Array.from({ length: size }, (_, i) => ({ section, first: i === 0 })));
const odd = [{}, { type: "oddPage" as const }, { type: "oddPage" as const }];
// Portada = 1, capítulo 1 = 2 (par): blanco antes; 3, 4, 5; capítulo 2 = 6 (par): blanco antes.
assert.deepEqual(planBlanks(pages([1, 3, 1]), [], odd), [1, 4]);
// Con dos páginas de portada el capítulo cae en 3: no hace falta.
assert.deepEqual(planBlanks(pages([2, 3, 1]), [], odd), [5]);
// Sección que vuelve a empezar en 1: impar por número, sin blanco aunque caiga a la izquierda.
assert.deepEqual(planBlanks(pages([1, 2]), [{}, { start: 1 }], odd), []);
assert.deepEqual(wrongSide(pages([1, 2]), odd), [{ section: 1, page: 2 }]);
// Vuelve a empezar en 2 y pide impar: Word igual agrega el blanco.
assert.deepEqual(planBlanks(pages([1, 2]), [{}, { start: 2 }], odd), [1]);
// Par.
assert.deepEqual(planBlanks(pages([2, 1]), [], [{}, { type: "evenPage" }]), [2]);
assert.deepEqual(planBlanks(pages([1, 1]), [], [{}, { type: "evenPage" }]), []);
// La primera página del documento nunca lleva blanco antes.
assert.deepEqual(planBlanks(pages([1]), [{ start: 2 }], [{ type: "oddPage" }]), []);

// Notas: seguidas, por sección y por página; las de marca propia no cuentan.
const notes = [
  { section: 0, custom: [false, false] },
  { section: 0, custom: [false, true, false] },
  { section: 1, custom: [false] },
];
assert.deepEqual(planFootnotes(notes, "continuous", 1), [[1, 2], [3, null, 4], [5]]);
assert.deepEqual(planFootnotes(notes, "eachSect", 1), [[1, 2], [3, null, 4], [1]]);
assert.deepEqual(planFootnotes(notes, "eachPage", 1), [[1, 2], [1, null, 2], [1]]);
assert.deepEqual(planFootnotes(notes, "continuous", 5), [[5, 6], [7, null, 8], [9]]);
assert.deepEqual(planFootnotes(notes, "continuous", Number.NaN), [[1, 2], [3, null, 4], [5]]);

console.log("paginate-layout: ok");
