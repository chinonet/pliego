import assert from "node:assert/strict";
import {
  contentText,
  cropMatrix,
  cropTransform,
  formatCounter,
  markerText,
  pagesText,
  parseContent,
  parseCounters,
  parseCrop,
  splitInvisible,
  splitSmallCaps,
  splitWords,
  transformCase,
} from "./prepare-export.ts";

// Contadores como los devuelve getComputedStyle.
assert.deepEqual(parseCounters("docx-num-1-0 0 docx-num-1-1 3"), [
  { name: "docx-num-1-0", value: 0 },
  { name: "docx-num-1-1", value: 3 },
]);
assert.deepEqual(parseCounters("none"), []);
assert.deepEqual(parseCounters("docx-num-2-0"), [{ name: "docx-num-2-0", value: null }]);

// El content del ::before de la numeración del visor.
const parts = parseContent(String.raw`counter(docx-num-1-0) "." counter(docx-num-1-1, lower-roman) "\9 "`);
assert.deepEqual(parts, [
  { kind: "counter", name: "docx-num-1-0", style: "decimal" },
  { kind: "text", text: "." },
  { kind: "counter", name: "docx-num-1-1", style: "lower-roman" },
  { kind: "text", text: "\t" },
]);
const counters = new Map([
  ["docx-num-1-0", 3],
  ["docx-num-1-1", 4],
]);
assert.equal(contentText(parts!, counters), "3.iv\t");
assert.equal(parseContent("none"), null);
assert.equal(parseContent("normal"), null);
assert.equal(parseContent("attr(title)"), null, "lo que no se resuelve no se inventa");
assert.deepEqual(parseContent(String.raw`"\f0b7 "`), [{ kind: "text", text: "\uf0b7" }], "viñeta de Symbol escapada");
assert.deepEqual(parseContent(String.raw`'a\'b'`), [{ kind: "text", text: "a'b" }]);
assert.deepEqual(parseContent('"x" / "alt"'), [{ kind: "text", text: "x" }], "el texto alternativo no se dibuja");
assert.deepEqual(parseContent('counters(c, ".", upper-alpha)'), [{ kind: "counter", name: "c", style: "upper-alpha" }]);
assert.equal(contentText(parseContent('counter(nuevo) ")"')!, new Map()), "0)", "un contador que no existe vale 0");

// Estilos de número.
assert.equal(formatCounter(4, "upper-roman"), "IV");
assert.equal(formatCounter(1994, "lower-roman"), "mcmxciv");
assert.equal(formatCounter(4000, "upper-roman"), "4000", "fuera de rango: decimal, como el navegador");
assert.equal(formatCounter(28, "lower-alpha"), "ab");
assert.equal(formatCounter(26, "upper-latin"), "Z");
assert.equal(formatCounter(0, "lower-alpha"), "0");
assert.equal(formatCounter(7, "decimal-leading-zero"), "07");
assert.equal(formatCounter(2, "lower-greek"), "β");
assert.equal(formatCounter(5, "ordinalText"), "5", "un nombre que CSS no conoce sale en decimal");
assert.equal(formatCounter(5, "katakana"), null, "un estilo que acá no se calcula se avisa");
assert.equal(formatCounter(5, "disc"), "\u2022");
assert.equal(markerText(3, "decimal"), "3. ");
assert.equal(markerText(1, "disc"), "\u2022 ");
assert.equal(markerText(1, "none"), "");

// Mayúsculas de estilo.
assert.equal(transformCase("señor", "uppercase"), "SEÑOR");
assert.equal(transformCase("Straße", "uppercase", "", "de"), "STRASSE");
assert.equal(transformCase("ÁRBOL", "lowercase"), "árbol");
assert.equal(transformCase("el árbol viejo", "capitalize"), "El Árbol Viejo");
assert.equal(transformCase("bol viejo", "capitalize", "ár"), "bol Viejo", "la palabra empezó en el nodo anterior");
assert.equal(transformCase(" casa", "capitalize", "la"), " Casa");

// Versalitas: minúsculas a mayúsculas chicas, las mayúsculas y los signos quedan.
assert.deepEqual(splitSmallCaps("Hola, Mundo 2"), [
  { text: "H", small: false },
  { text: "OLA", small: true },
  { text: ", M", small: false },
  { text: "UNDO", small: true },
  { text: " 2", small: false },
]);
assert.deepEqual(splitSmallCaps("Ab", true), [{ text: "AB", small: true }], "all-small-caps achica también las mayúsculas");

// Trozos que jsafa dibuja por separado.
assert.deepEqual(splitWords("dos  palabras\u00A0más "), ["dos", "  ", "palabras", "\u00A0", "más", " "]);
assert.deepEqual(splitInvisible("pa\u00ADla\u200Bbra"), ["pa", "\u00AD", "la", "\u200B", "bra"]);
assert.deepEqual(splitInvisible("nada"), ["nada"]);

// Recortes del Word como los deja el visor.
assert.deepEqual(parseCrop("rect(10.00% 70.00% 95.00% 5.00%)"), { top: 0.1, right: 0.3, bottom: 0.05, left: 0.05 });
assert.deepEqual(parseCrop("inset(10% 30% 5% 5%)"), { top: 0.1, right: 0.3, bottom: 0.05, left: 0.05 });
assert.deepEqual(parseCrop("inset(10%)"), { top: 0.1, right: 0.1, bottom: 0.1, left: 0.1 });
assert.equal(parseCrop("rect(0.00% 100.00% 100.00% 0.00%)"), null, "sin recorte");
assert.equal(parseCrop("none"), null);
assert.equal(parseCrop("circle(50%)"), null);

// La parte visible ocupa justo la caja de la imagen (escala desde el centro más el corrimiento).
const crop = { left: 0.1, top: 0.2, right: 0.3, bottom: 0 };
const m = cropMatrix(crop, 100, 50);
const map = (x: number, y: number) => [50 + m.tx + m.sx * (x - 50), 25 + m.ty + m.sy * (y - 25)];
const near = (a: number[], b: number[]) => a.every((value, i) => Math.abs(value - b[i]!) < 1e-9);
assert.ok(near(map(10, 10), [0, 0]), "la esquina visible de arriba a la izquierda va al borde");
assert.ok(near(map(70, 50), [100, 50]), "la esquina visible de abajo a la derecha va al otro borde");
assert.equal(cropTransform({ left: 0.25, top: 0.25, right: 0.25, bottom: 0.25 }, 100, 100), "translate(0px, 0px) scale(2, 2)");
assert.match(cropTransform(crop, 100, 50, 90), /^rotate\(90deg\) translate\(/, "el giro queda alrededor del centro");

assert.equal(pagesText([3]), "la página 3");
assert.equal(pagesText([8, 3, 5, 3]), "las páginas 3, 5 y 8");
assert.equal(pagesText([]), "algunas páginas");
assert.match(pagesText(Array.from({ length: 20 }, (_, i) => i + 1)), /12…$/);

console.log("prepare-export ok");
