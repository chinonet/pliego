import assert from "node:assert/strict";
import {
  keepsSpaceBefore,
  lastFitting,
  layoutPadding,
  planBlanks,
  planFootnotes,
  spansAcross,
  wrongSide,
  type PlanPage,
} from "./paginate.ts";
import { DEFAULT_RENDER_INFO } from "./render-copy.ts";

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

// Cuánto entra de una vez: bordes de abajo de cada bloque, llamadas a nota por bloque y lugar que deja cada nota.
const bottoms = [100, 200, 300, 400, 500];
const at = (list: number[]) => (index: number) => list[index]!;
assert.equal(lastFitting(5, at(bottoms), () => 0, 0, 250, () => 0), 1, "sin notas: el último que no pasa el límite");
assert.equal(lastFitting(5, at(bottoms), () => 0, 0, 50, () => 0), -1, "ni el primero entra");
assert.equal(lastFitting(5, at(bottoms), () => 0, 0, 900, () => 0), 4, "entra todo");
// Las notas de los bloques que se van dejan lugar: con el bloque 2 adentro se va una nota (60 px) y entra.
const refs = [0, 0, 1, 0, 1];
assert.equal(lastFitting(5, at(bottoms), at(refs), 2, 250, (n) => n * 60), 2);
// Con una sola nota (la del bloque 4), el bloque 2 no entra ni sacándola.
assert.equal(lastFitting(5, at(bottoms), at([0, 0, 0, 0, 1]), 1, 230, (n) => n * 60), 1);

// Celdas combinadas en vertical al cortar una tabla: 3 filas quedan arriba.
assert.deepEqual(
  spansAcross(
    [
      { row: 0, col: 0, span: 5 },
      { row: 1, col: 1, span: 2 },
      { row: 2, col: 2, span: 1 },
      { row: 2, col: 3, span: 3 },
    ],
    3,
  ),
  [
    { row: 0, col: 0, span: 5, keep: 3, rest: 2 },
    { row: 2, col: 3, span: 3, keep: 1, rest: 2 },
  ],
  "la de la columna 1 termina antes del corte",
);

// Espacio anterior arriba de la página, como Word.
const word2013 = { ...DEFAULT_RENDER_INFO, compatibilityMode: 15 };
assert.equal(keepsSpaceBefore("section", false, word2013), true, "comienzo de sección: queda");
assert.equal(keepsSpaceBefore("natural", false, DEFAULT_RENDER_INFO), false, "salto natural: se saca");
assert.equal(keepsSpaceBefore("manual", false, DEFAULT_RENDER_INFO), true, "salto manual en Word viejo: queda");
assert.equal(keepsSpaceBefore("manual", false, word2013), false, "salto manual desde Word 2013: se saca");
assert.equal(keepsSpaceBefore("manual", false, { ...DEFAULT_RENDER_INFO, suppressSpBfAfterPgBrk: true }), false);
assert.equal(keepsSpaceBefore("section", true, DEFAULT_RENDER_INFO), false, "continuación de un párrafo cortado");
assert.equal(keepsSpaceBefore(undefined, false, DEFAULT_RENDER_INFO), true);

console.log("paginate-layout: ok");
