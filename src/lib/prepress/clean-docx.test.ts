import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { cleanDocx, isFontName } from "./clean-docx.ts";

const original = new Uint8Array(await readFile(new URL("../../../public/muestra-premedia.docx", import.meta.url)));
const source = await JSZip.loadAsync(original);
const sourceImage = await source.file(/word\/media\//)[0]?.async("uint8array");
assert.ok(sourceImage && sourceImage.byteLength > 1000, "la muestra trae imagen");

const cleaned = await cleanDocx(original, {
  acceptRevisions: true,
  removeComments: true,
  keepEndnotes: true,
  removeHighlights: false,
});

const zip = await JSZip.loadAsync(cleaned.bytes);
assert.equal(zip.file("word/comments.xml"), null);
assert.equal(zip.file(/comments/i).length, 0);

const documentXml = await zip.file("word/document.xml")!.async("string");
assert.match(documentXml, /Tirada de 2\.000 ejemplares/);
assert.doesNotMatch(documentXml, /Borrador interno/);
assert.doesNotMatch(documentXml, /<w:ins[\s>]/);
assert.doesNotMatch(documentXml, /<w:del[\s>]/);
assert.doesNotMatch(documentXml, /commentReference/);
assert.doesNotMatch(documentXml, /commentRangeStart/);
assert.match(documentXml, /<w:footnoteReference\b/);
assert.match(documentXml, /<w:endnoteReference\b/);
assert.match(documentXml, /w:val="yellow"/);

const footnotes = await zip.file("word/footnotes.xml")!.async("string");
assert.match(footnotes, /papel estucado de 150 g/);

const image = await zip.file(/word\/media\//)[0]!.async("uint8array");
assert.equal(image.byteLength, sourceImage.byteLength);
assert.equal(Buffer.from(image).equals(Buffer.from(sourceImage)), true);

assert.equal(cleaned.report.commentsRemoved, 1);
assert.equal(cleaned.report.footnotesKept, 1);
assert.equal(cleaned.report.endnotesKept, 1);
assert.ok(cleaned.report.insertionsAccepted >= 1);
assert.ok(cleaned.report.deletionsRemoved >= 1);
assert.equal(cleaned.report.images.length, 1);
assert.equal(isFontName("Calibri"), true);
assert.equal(isFontName("MS Gothic"), true);
assert.equal(isFontName("inherit"), false);
assert.equal(isFontName("es-AR"), false);

const noMarks = await cleanDocx(original, {
  acceptRevisions: true,
  removeComments: true,
  keepEndnotes: false,
  removeHighlights: true,
});
const bare = await JSZip.loadAsync(noMarks.bytes);
const bareXml = await bare.file("word/document.xml")!.async("string");
assert.doesNotMatch(bareXml, /<w:endnoteReference\b/);
assert.doesNotMatch(bareXml, /w:val="yellow"/);
assert.match(bareXml, /<w:footnoteReference\b/);

// Revisiones de estructura y comentarios, con un Word armado a mano.
const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MS = "http://schemas.microsoft.com/office";
const del = '<w:del w:id="90" w:author="E" w:date="2026-01-01T00:00:00Z"/>';
const ins = '<w:ins w:id="91" w:author="E" w:date="2026-01-01T00:00:00Z"/>';
const cell = (text: string, props = "") =>
  `<w:tc>${props ? `<w:tcPr>${props}</w:tcPr>` : ""}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const para = (text: string, style: string, markDel = false) =>
  `<w:p><w:pPr><w:pStyle w:val="${style}"/>${markDel ? `<w:rPr>${del}</w:rPr>` : ""}</w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const bodyXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
<w:tr>${cell("Fila buena")}${cell("Celda buena")}</w:tr>
<w:tr><w:trPr>${del}</w:trPr>${cell("Fila borrada")}${cell("Fila borrada 2")}</w:tr>
<w:tr><w:trPr>${ins}</w:trPr>${cell("Fila nueva")}${cell("Celda borrada", '<w:cellDel w:id="92" w:author="E"/>')}</w:tr>
<w:tr>${cell("Celda nueva", '<w:cellIns w:id="93" w:author="E"/>')}${cell("Vecina")}</w:tr>
</w:tbl>
<w:tbl><w:tr><w:trPr>${del}</w:trPr>${cell("Tabla entera borrada")}</w:tr></w:tbl>
<w:tbl><w:tr><w:tc><w:tbl><w:tr><w:trPr>${del}</w:trPr>${cell("Tabla anidada borrada")}</w:tr></w:tbl></w:tc></w:tr></w:tbl>
${para("Uno, ", "Primero", true)}
<w:bookmarkEnd w:id="0"/>
${para("dos, ", "Segundo", true)}
${para("tres.", "Tercero")}
<w:p><w:pPr><w:pStyle w:val="Vacio"/><w:rPr>${del}</w:rPr></w:pPr><w:del w:id="94" w:author="E"><w:r><w:delText>Todo borrado</w:delText></w:r></w:del></w:p>
<w:tbl><w:tr>${cell("Después del vacío")}</w:tr></w:tbl>
${para("Antes de tabla", "Queda", true)}
<w:tbl><w:tr>${cell("Final")}</w:tr></w:tbl>
<w:p><w:r><w:t>Cierre</w:t><w:commentReference w:id="0"/></w:r></w:p>
<w:sectPr/>
</w:body></w:document>`;

const relsXml = (body: string) =>
  `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
const pkg = new JSZip();
pkg.file(
  "[Content_Types].xml",
  '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
    '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>' +
    '<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>' +
    '<Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/>' +
    '<Override PartName="/customXml/comments-notes.xml" ContentType="application/xml"/>' +
    "</Types>",
);
pkg.file("_rels/.rels", relsXml(`<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/>`));
pkg.file(
  "word/_rels/document.xml.rels",
  relsXml(
    `<Relationship Id="rId1" Type="${REL}/comments" Target="comments.xml"/>` +
      `<Relationship Id="rId2" Type="${MS}/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>` +
      `<Relationship Id="rId3" Type="${MS}/2016/09/relationships/commentsIds" Target="/word/commentsIds.xml"/>` +
      `<Relationship Id="rId4" Type="${MS}/2011/relationships/people" Target="./people.xml"/>` +
      `<Relationship Id="rId5" Type="${REL}/image" Target="media/people.jpg"/>` +
      `<Relationship Id="rId6" Type="${REL}/image" Target="media/comments.png"/>` +
      `<Relationship Id="rId7" Type="${REL}/customXml" Target="../customXml/comments-notes.xml"/>`,
  ),
);
pkg.file("word/document.xml", bodyXml);
pkg.file(
  "word/comments.xml",
  `<?xml version="1.0" encoding="UTF-8"?><w:comments ${W}><w:comment w:id="0" w:author="E"><w:p><w:r><w:t>Ojo</w:t></w:r></w:p></w:comment><w:comment w:id="1" w:author="E"><w:p/></w:comment></w:comments>`,
);
pkg.file("word/_rels/comments.xml.rels", relsXml(""));
pkg.file("word/commentsExtended.xml", `<?xml version="1.0" encoding="UTF-8"?><w15:commentsEx xmlns:w15="${MS}/word/2012/wordml"/>`);
pkg.file("word/commentsIds.xml", `<?xml version="1.0" encoding="UTF-8"?><w16cid:commentsIds xmlns:w16cid="${MS}/word/2016/wordml/cid"/>`);
pkg.file("word/people.xml", `<?xml version="1.0" encoding="UTF-8"?><w15:people xmlns:w15="${MS}/word/2012/wordml"/>`);
const peopleJpg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 1, 2, 3, 4, 0xff, 0xd9]);
pkg.file("word/media/people.jpg", peopleJpg);
pkg.file("word/media/comments.png", new Uint8Array([0x89, 0x50, 0x4e, 0x47, 9, 9]));
pkg.file("customXml/comments-notes.xml", '<?xml version="1.0" encoding="UTF-8"?><notas/>');
const handmade = await pkg.generateAsync({ type: "uint8array" });

const structural = await cleanDocx(handmade, {
  acceptRevisions: true,
  removeComments: true,
  keepEndnotes: true,
  removeHighlights: false,
});
const out = await JSZip.loadAsync(structural.bytes);
const outXml = await out.file("word/document.xml")!.async("string");

// Filas: la borrada se va entera; la insertada queda sin marca.
assert.match(outXml, /Fila buena/);
assert.doesNotMatch(outXml, /Fila borrada/);
assert.match(outXml, /Fila nueva/);
assert.doesNotMatch(outXml, /<w:trPr>\s*<w:ins\b/);
// Una tabla sin filas desaparece; si estaba dentro de una celda, la celda conserva un párrafo.
assert.doesNotMatch(outXml, /Tabla entera borrada/);
assert.doesNotMatch(outXml, /Tabla anidada borrada/);
assert.match(outXml, /<w:tc><w:p\/><\/w:tc>/);
// Celdas: la borrada se va; la insertada queda sin marca.
assert.doesNotMatch(outXml, /Celda borrada/);
assert.match(outXml, /Celda nueva/);
assert.match(outXml, /Vecina/);
assert.doesNotMatch(outXml, /cellDel|cellIns/);
assert.ok(structural.report.warnings.some((w) => /celda que estaba marcada como eliminada/.test(w)));

// Fin de párrafo borrado: se une con el siguiente, que pone sus propiedades (como en Word).
assert.match(
  outXml,
  /<w:pStyle w:val="Tercero"\/><\/w:pPr><w:r><w:t xml:space="preserve">Uno, <\/w:t><\/w:r><w:r><w:t xml:space="preserve">dos, <\/w:t><\/w:r><w:r><w:t xml:space="preserve">tres\.<\/w:t>/,
);
assert.doesNotMatch(outXml, /Primero|Segundo/);
assert.match(outXml, /<w:bookmarkEnd w:id="0"\/>/);
// Un párrafo borrado entero antes de una tabla se va.
assert.doesNotMatch(outXml, /Vacio|Todo borrado/);
assert.match(outXml, /Después del vacío/);
// Uno con texto antes de una tabla no se puede unir: queda, sin la marca, y avisa.
assert.match(outXml, /<w:pStyle w:val="Queda"\/>(<w:rPr\/>)?<\/w:pPr><w:r><w:t xml:space="preserve">Antes de tabla/);
assert.ok(structural.report.warnings.some((w) => /no había un párrafo siguiente/.test(w)));
assert.doesNotMatch(outXml, /<w:(ins|del)\b/);
assert.ok(!structural.report.warnings.some((w) => /Quedó alguna marca/.test(w)));
// 3 filas + 1 celda + 3 párrafos (dos unidos y uno vacío quitado).
assert.equal(structural.report.deletionsRemoved, 7);

// Comentarios: solo las partes que son comentarios según el tipo de relación.
for (const gone of [
  "word/comments.xml",
  "word/_rels/comments.xml.rels",
  "word/commentsExtended.xml",
  "word/commentsIds.xml",
  "word/people.xml",
]) {
  assert.equal(out.file(gone), null, `${gone} se quita`);
}
const keptJpg = await out.file("word/media/people.jpg")?.async("uint8array");
assert.ok(keptJpg, "word/media/people.jpg sobrevive");
assert.equal(Buffer.from(keptJpg).equals(Buffer.from(peopleJpg)), true);
assert.ok(out.file("word/media/comments.png"), "word/media/comments.png sobrevive");
assert.ok(out.file("customXml/comments-notes.xml"), "customXml/comments-notes.xml sobrevive");
const outRels = await out.file("word/_rels/document.xml.rels")!.async("string");
assert.doesNotMatch(outRels, /relationships\/comments"|commentsExtended|commentsIds|relationships\/people"/);
assert.match(outRels, /Target="media\/people\.jpg"/);
assert.match(outRels, /Target="media\/comments\.png"/);
assert.match(outRels, /Target="\.\.\/customXml\/comments-notes\.xml"/);
const outTypes = await out.file("[Content_Types].xml")!.async("string");
assert.doesNotMatch(outTypes, /\/word\/comments\.xml|commentsExtended|commentsIds|\/word\/people\.xml/);
assert.match(outTypes, /Extension="jpg"/);
assert.match(outTypes, /PartName="\/customXml\/comments-notes\.xml"/);
assert.match(outTypes, /PartName="\/word\/document\.xml"/);
assert.doesNotMatch(outXml, /commentReference/);
assert.equal(structural.report.commentsRemoved, 2);
assert.equal(structural.report.images.length, 2);

// Sin aceptar revisiones ni quitar comentarios, nada de esto se toca.
const untouched = await cleanDocx(handmade, {
  acceptRevisions: false,
  removeComments: false,
  keepEndnotes: true,
  removeHighlights: false,
});
const same = await JSZip.loadAsync(untouched.bytes);
const sameXml = await same.file("word/document.xml")!.async("string");
assert.match(sameXml, /Fila borrada/);
assert.match(sameXml, /Celda borrada/);
assert.match(sameXml, /Primero/);
assert.ok(same.file("word/comments.xml"));
assert.ok(same.file("word/people.xml"));
assert.equal(untouched.report.commentsRemoved, 0);

console.log("clean-docx ok", cleaned.report, structural.report);
