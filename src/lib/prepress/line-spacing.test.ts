import assert from "node:assert/strict";
import JSZip from "jszip";
import {
  CTX_SUFFIX,
  FontMetricsCache,
  fontMetricsFromBytes,
  normalizeLineRules,
  parseDeclaredLineHeight,
  prepareLineSpacing,
  readFontMetrics,
  readSpacingInfo,
  specificity,
  splitSelectors,
  styleClassKey,
  tagDirectContextual,
  WORD_FONT_METRICS,
  wordLinePx,
  wordLineRatio,
} from "./line-spacing.ts";

const close = (a: number, b: number, msg: string) => assert.ok(Math.abs(a - b) < 1e-3, `${msg}: ${a} ≠ ${b}`);

// Altura de renglón sencillo de Word.
close(wordLineRatio(WORD_FONT_METRICS["Times New Roman"]!), 2355 / 2048, "Times New Roman 1,15");
close(wordLineRatio(WORD_FONT_METRICS.Calibri!), 2500 / 2048, "Calibri 1,22 (la interlínea de hhea ya está en win)");
close(wordLineRatio({ unitsPerEm: 1000, ascent: 800, descent: -200, lineGap: 100 }), 1.1, "sin OS/2, hhea con interlínea");
close(wordLineRatio({ unitsPerEm: 1000, ascent: 800, descent: 200, lineGap: 0 }), 1, "descendente en positivo");

// Fuente mínima (head, hhea, OS/2) para el lector propio.
function fakeFont(o: { upm: number; asc: number; desc: number; gap: number; winA: number; winD: number }): Uint8Array {
  const tables = [
    { tag: "head", size: 54 },
    { tag: "hhea", size: 36 },
    { tag: "OS/2", size: 78 },
  ];
  const headerSize = 12 + tables.length * 16;
  const total = headerSize + tables.reduce((sum, t) => sum + t.size, 0);
  const bytes = new Uint8Array(total);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x00010000);
  view.setUint16(4, tables.length);
  let offset = headerSize;
  const at = new Map<string, number>();
  tables.forEach((t, i) => {
    const rec = 12 + i * 16;
    for (let k = 0; k < 4; k++) bytes[rec + k] = t.tag.charCodeAt(k);
    view.setUint32(rec + 8, offset);
    view.setUint32(rec + 12, t.size);
    at.set(t.tag, offset);
    offset += t.size;
  });
  view.setUint16(at.get("head")! + 18, o.upm);
  view.setInt16(at.get("hhea")! + 4, o.asc);
  view.setInt16(at.get("hhea")! + 6, o.desc);
  view.setInt16(at.get("hhea")! + 8, o.gap);
  view.setUint16(at.get("OS/2")! + 74, o.winA);
  view.setUint16(at.get("OS/2")! + 76, o.winD);
  return bytes;
}
const times = fakeFont({ upm: 2048, asc: 1825, desc: -443, gap: 87, winA: 1825, winD: 443 });
assert.deepEqual(readFontMetrics(times), WORD_FONT_METRICS["Times New Roman"]);
assert.deepEqual(await fontMetricsFromBytes(times), WORD_FONT_METRICS["Times New Roman"]);
assert.throws(() => readFontMetrics(new Uint8Array(4)));

// Prioridad: fuente subida > tabla de Windows > sustituto.
const cache = new FontMetricsCache();
close(wordLineRatio((await cache.metrics("'Times New Roman'"))!), 2355 / 2048, "la tabla vale aunque no haya archivo");
const odd = fakeFont({ upm: 1000, asc: 900, desc: -300, gap: 0, winA: 900, winD: 300 });
cache.addUrl("Times New Roman", "/fonts/no-se-usa.ttf");
close(wordLineRatio((await cache.metrics("Times New Roman"))!), 2355 / 2048, "el sustituto no pisa la tabla");
cache.addBytes("Times New Roman", odd, { weight: 700, style: "normal" });
close(wordLineRatio((await cache.metrics("Times New Roman"))!), 1.2, "la fuente subida manda");
cache.addBytes("Times New Roman", times, { weight: 400, style: "normal" });
close(wordLineRatio((await cache.metrics("times new roman"))!), 2355 / 2048, "la regular gana a la negrita");
assert.equal(await cache.metrics("Desconocida"), null);

// Lo que escribe el visor.
assert.deepEqual(parseDeclaredLineHeight(""), { kind: "single" });
assert.deepEqual(parseDeclaredLineHeight("1.15"), { kind: "auto", multiple: 1.15 });
assert.deepEqual(parseDeclaredLineHeight("calc(100% + 14pt)"), { kind: "atLeast", pt: 14 });
assert.deepEqual(parseDeclaredLineHeight("12pt"), { kind: "exact", pt: 12 });
assert.deepEqual(parseDeclaredLineHeight("20px"), { kind: "other" });

const r = 2355 / 2048;
const px12 = 16;
close(wordLinePx({ kind: "single" }, r, px12)!, r * 16, "sencillo");
close(wordLinePx({ kind: "auto", multiple: 2 }, r, px12)!, r * 32, "doble");
close(wordLinePx({ kind: "atLeast", pt: 12 }, r, px12)!, r * 16, "al menos 12 pt con letra de 12: manda la letra");
close(wordLinePx({ kind: "atLeast", pt: 24 }, r, px12)!, 32, "al menos 24 pt: manda el mínimo");
assert.equal(wordLinePx({ kind: "exact", pt: 12 }, r, px12), null, "exacto lo deja el visor");

assert.equal(specificity("p.docx_normal"), specificity(".docx p"));
assert.ok(specificity("table.docx_t p") > specificity("p.docx_normal"));
assert.ok(specificity("#a p") > specificity(".a .b .c p"));
assert.deepEqual(splitSelectors(".docx p, p.docx_normal"), [".docx p", "p.docx_normal"]);
assert.deepEqual(splitSelectors("p:not(.a, .b), span"), ["p:not(.a, .b)", "span"]);

// lineRule implícito.
const spacingXml = `<w:p><w:pPr><w:spacing w:after="120" w:line="360"/></w:pPr></w:p><w:spacing w:line="240" w:lineRule="exact"/><w:spacing w:before="0"/>`;
const normalized = normalizeLineRules(spacingXml);
assert.equal(normalized.changed, 1);
assert.match(normalized.xml, /<w:spacing w:lineRule="auto" w:after="120" w:line="360"\/>/);
assert.match(normalized.xml, /w:line="240" w:lineRule="exact"/, "el explícito no se toca");
assert.equal(normalizeLineRules("<w:p/>").xml, "<w:p/>");

// contextualSpacing.
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Vieta"><w:name w:val="Viñeta"/><w:basedOn w:val="ListParagraph"/></w:style>
<w:style w:type="paragraph" w:styleId="Cita"><w:name w:val="Cita"/><w:basedOn w:val="ListParagraph"/><w:pPr><w:contextualSpacing w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Lista"><w:name w:val="Lista"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
<w:style w:type="character" w:styleId="Fuerte"><w:name w:val="Fuerte"/></w:style>
</w:styles>`;
const info = readSpacingInfo(stylesXml);
assert.deepEqual(info.contextual.sort(), ["listparagraph", "vieta"], "se hereda por basedOn y se apaga con val=0");
assert.equal(info.defaultStyle, "normal");
assert.equal(styleClassKey("Heading 1.x&y"), "heading-1-xandy");

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
<w:p><w:pPr><w:contextualSpacing/></w:pPr><w:r><w:t>Normal con la opción a mano</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:contextualSpacing/></w:pPr><w:r><w:t>Igual que su estilo</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:contextualSpacing w:val="0"/></w:pPr><w:r><w:t>La apaga</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Lista"/><w:contextualSpacing/></w:pPr><w:r><w:t>Con numeración</w:t></w:r></w:p>
</w:body></w:document>`;
const tagged = tagDirectContextual(stylesXml, [documentXml]);
assert.equal(tagged.tagged, 2);
assert.equal(tagged.skipped, 1, "el estilo con numeración no se copia");
assert.match(tagged.parts[0]!, new RegExp(`<w:pPr><w:pStyle[^>]*w:val="Normal${CTX_SUFFIX}1"`), "pStyle nuevo al principio del pPr");
assert.match(tagged.parts[0]!, new RegExp(`w:val="ListParagraph${CTX_SUFFIX}0"`));
assert.match(tagged.parts[0]!, /<w:pStyle w:val="ListParagraph"\/><w:contextualSpacing\/>/, "el que coincide queda igual");
assert.match(tagged.stylesXml, new RegExp(`w:styleId="Normal${CTX_SUFFIX}1"`));
const copyTag = new RegExp(`<w:style[^>]*w:styleId="Normal${CTX_SUFFIX}1"[^>]*>`).exec(tagged.stylesXml)?.[0] ?? "";
assert.ok(copyTag && !/default/.test(copyTag), "la copia no es la predeterminada");
const taggedInfo = readSpacingInfo(tagged.stylesXml);
assert.ok(taggedInfo.contextual.includes(`normal${CTX_SUFFIX.toLowerCase()}1`));
assert.ok(!taggedInfo.contextual.includes(`listparagraph${CTX_SUFFIX.toLowerCase()}0`));
assert.equal(taggedInfo.defaultStyle, "normal", "la predeterminada sigue siendo la original");
assert.equal(tagDirectContextual(stylesXml, ["<w:p/>"]).parts[0], "<w:p/>", "sin la opción no se reescribe");

// Todo junto sobre el zip de la copia de composición.
const zip = new JSZip();
zip.file("word/styles.xml", stylesXml);
zip.file("word/document.xml", documentXml.replace("</w:body>", '<w:p><w:pPr><w:spacing w:line="480"/></w:pPr></w:p></w:body>'));
zip.file("word/media/image1.png", new Uint8Array([1, 2, 3]));
const prepared = await prepareLineSpacing(zip);
const body = await zip.file("word/document.xml")!.async("string");
assert.match(body, /w:lineRule="auto" w:line="480"/);
assert.match(body, new RegExp(`Normal${CTX_SUFFIX}1`));
assert.ok(prepared.spacing.contextual.includes("listparagraph"));
assert.equal(prepared.warnings.length, 1, "aviso por el párrafo que no se pudo reproducir");
assert.deepEqual(await zip.file("word/media/image1.png")!.async("uint8array"), new Uint8Array([1, 2, 3]));

console.log("line-spacing ok");
