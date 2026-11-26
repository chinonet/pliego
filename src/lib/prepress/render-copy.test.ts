import assert from "node:assert/strict";
import JSZip from "jszip";
import { DOMParser, type Element as XmlElement } from "@xmldom/xmldom";
import { FOOTNOTE_MARK, markPageFields } from "./fields.ts";
import { cropToPercent, FOOTNOTE_CUSTOM_MARK, lengthToEmu, prepareRenderCopy } from "./render-copy.ts";

const WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS = [
  `xmlns:w="${WNS}"`,
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
].join(" ");
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const docXml = `${HEAD}<w:document ${NS}><w:body>
<w:p><w:r><w:t>Portada</w:t></w:r></w:p>
<w:p><w:pPr><w:sectPr>
  <w:headerReference w:type="default" r:id="rId1"/><w:footerReference w:type="default" r:id="rId2"/>
  <w:type w:val="oddPage"/><w:pgMar w:top="1440" w:left="1440" w:right="1080" w:gutter="360"/><w:titlePg/>
</w:sectPr></w:pPr><w:r><w:t>Fin de la portada</w:t></w:r></w:p>
<w:p><w:r><w:t>Anglo</w:t><w:noBreakHyphen/><w:t>sajón y lar</w:t><w:softHyphen/><w:t>go</w:t></w:r></w:p>
<w:p><w:r><w:t>Nota rara</w:t></w:r><w:r><w:footnoteReference w:customMarkFollows="1" w:id="2"/><w:t>*</w:t></w:r>
<w:r><w:t>y nota común</w:t><w:footnoteReference w:id="3"/></w:r></w:p>
<w:p><w:r><w:pict><v:shapetype id="_x0000_t75"/><v:shape id="foto" style="width:72pt;height:1in"><v:imagedata r:id="rId5" cropleft="32768f" o:title=""/></v:shape></w:pict></w:r></w:p>
<w:p><w:r><w:pict><v:shape style="position:absolute;width:300pt;height:400pt;z-index:-1"><v:imagedata r:id="rId5"/></v:shape></w:pict></w:r></w:p>
<w:p><w:r><w:t>Antes del recuadro</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><a:graphic><a:graphicData><w:txbxContent><w:p><w:r><w:t>Recuadro</w:t></w:r></w:p></w:txbxContent></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>
<mc:Fallback><w:pict><v:shape style="width:100pt;height:50pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Recuadro</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
<w:p><w:r><w:t>Ventas</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><a:graphic><a:graphicData><c:chart r:id="rId9"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:pPr><w:sectPr><w:headerReference w:type="first" r:id="rId3"/><w:type w:val="continuous"/></w:sectPr></w:pPr><w:r><w:t>Capítulo</w:t></w:r></w:p>
<w:p><w:r><w:t>Final</w:t></w:r></w:p>
<w:sectPr><w:footnotePr><w:numStart w:val="5"/></w:footnotePr><w:pgSz w:w="8640" w:h="12960"/></w:sectPr>
</w:body></w:document>`;

const settingsXml = `${HEAD}<w:settings ${NS}>
<w:mirrorMargins/><w:gutterAtTop w:val="0"/><w:autoHyphenation w:val="true"/><w:evenAndOddHeaders/>
<w:footnotePr><w:numRestart w:val="eachPage"/><w:numStart w:val="3"/></w:footnotePr>
</w:settings>`;

const relsXml = `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;

const typesXml = `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const zip = new JSZip();
zip.file("[Content_Types].xml", typesXml);
zip.file("word/document.xml", docXml);
zip.file("word/settings.xml", settingsXml);
zip.file("word/_rels/document.xml.rels", relsXml);
zip.file("word/header1.xml", `${HEAD}<w:hdr ${NS}><w:p><w:r><w:t>Título</w:t></w:r></w:p></w:hdr>`);
zip.file("word/header2.xml", `${HEAD}<w:hdr ${NS}><w:p><w:r><w:t>Primera</w:t><w:noBreakHyphen/></w:r></w:p></w:hdr>`);
zip.file("word/footer1.xml", `${HEAD}<w:ftr ${NS}><w:p><w:r><w:t>Pie</w:t></w:r></w:p></w:ftr>`);
zip.file("word/media/image1.png", new Uint8Array([9, 8, 7, 6]));
const input = await zip.generateAsync({ type: "uint8array" });

const marked = await markPageFields(input);
const { bytes, info, warnings } = await prepareRenderCopy(marked.bytes);
const out = await JSZip.loadAsync(bytes);
const body = await out.file("word/document.xml")!.async("string");

// Datos de diseño
assert.equal(info.mirrorMargins, true);
assert.equal(info.gutterAtTop, false, "w:val=0 apaga el interruptor");
assert.equal(info.autoHyphenation, true);
assert.equal(info.footnoteRestart, "eachPage", "de settings");
assert.equal(info.footnoteStart, 5, "el sectPr final manda sobre settings");
assert.equal(info.sections.length, marked.sections.length, "las mismas secciones que fields.ts");
assert.deepEqual(info.sections, [{ type: "oddPage", gutterPt: 18, leftPt: 72, rightPt: 54 }, { type: "continuous" }, {}]);

// Encabezados y pies vinculados al anterior
const parsed = new DOMParser().parseFromString(body, "text/xml");
const sectPrs = Array.from(parsed.getElementsByTagNameNS(WNS, "sectPr")) as XmlElement[];
const refs = (sectPr: XmlElement, kind: string) =>
  (Array.from(sectPr.getElementsByTagNameNS(WNS, kind)) as XmlElement[]).map(
    (ref) => `${ref.getAttribute("w:type")}=${ref.getAttribute("r:id")}`,
  ).sort();
assert.deepEqual(refs(sectPrs[0]!, "headerReference"), [
  "default=rId1",
  "even=rIdPliegoBlankHeader",
  "first=rIdPliegoBlankHeader",
], "primera página y pares sin encabezado propio: en blanco, como en Word");
assert.deepEqual(refs(sectPrs[1]!, "headerReference"), ["default=rId1", "even=rIdPliegoBlankHeader", "first=rId3"]);
assert.deepEqual(refs(sectPrs[1]!, "footerReference"), ["default=rId2", "even=rIdPliegoBlankFooter", "first=rIdPliegoBlankFooter"]);
assert.deepEqual(refs(sectPrs[2]!, "headerReference"), ["default=rId1", "even=rIdPliegoBlankHeader", "first=rId3"], "la última hereda todo");
const kinds = (sectPr: XmlElement) =>
  (Array.from(sectPr.childNodes) as XmlElement[]).filter((n) => n.nodeType === 1).map((n) => n.localName);
const order = kinds(sectPrs[1]!);
assert.ok(order.lastIndexOf("footerReference") < order.indexOf("type"), "las referencias van antes que el resto del sectPr");
const rels = await out.file("word/_rels/document.xml.rels")!.async("string");
assert.match(rels, /Id="rIdPliegoBlankHeader"[^>]*Target="pliego-blank-header.xml"|Target="pliego-blank-header.xml"[^>]*Id="rIdPliegoBlankHeader"/);
assert.ok(out.file("word/pliego-blank-header.xml") && out.file("word/pliego-blank-footer.xml"), "partes vacías creadas");
assert.match(await out.file("[Content_Types].xml")!.async("string"), /pliego-blank-footer\.xml/);

// Guiones
assert.ok(body.includes("Anglo</w:t><w:t xml:space=\"preserve\">‑</w:t>"), "guion de no separación visible");
assert.ok(body.includes("­"), "guion opcional");
assert.doesNotMatch(body, /noBreakHyphen|softHyphen/);
assert.ok((await out.file("word/header2.xml")!.async("string")).includes("‑"), "también en encabezados");

// Nota con símbolo propio: las dos marcas juntas, en el mismo texto, justo antes de la llamada
assert.ok(body.includes(`${FOOTNOTE_MARK}${FOOTNOTE_CUSTOM_MARK}</w:t><w:footnoteReference w:customMarkFollows="1"`));
assert.equal(body.split(FOOTNOTE_CUSTOM_MARK).length - 1, 1, "la nota común no lleva la marca");
assert.equal(FOOTNOTE_CUSTOM_MARK, "C");

// Imagen VML → imagen en línea con la misma relación y el mismo tamaño
assert.match(body, /r:embed="rId5"/);
assert.match(body, /<wp:extent cx="914400" cy="914400"\/>/, "72pt × 1in");
assert.match(body, /<a:srcRect l="50000"\/>/, "el recorte viaja");
assert.equal(body.split("<w:drawing").length - 1, 3, "la foto convertida, el gráfico y el recuadro de mc:Choice");
assert.doesNotMatch(body, /<v:imagedata r:id="rId5" cropleft/, "la foto ya no está en VML");
assert.match(body, /z-index:-1/, "la de detrás del texto queda como estaba");
assert.deepEqual(await out.file("word/media/image1.png")!.async("uint8array"), new Uint8Array([9, 8, 7, 6]), "la imagen no se toca");

// Avisos
assert.ok(warnings.some((w) => w.startsWith("Hay un cuadro de texto en el texto del libro («Recuadro»)")), "un solo aviso por cuadro (no cuenta mc:Choice)");
assert.ok(!warnings.some((w) => w.startsWith("Hay 2 cuadros")));
assert.ok(warnings.some((w) => w.includes("Hay un gráfico de Word en el texto del libro")));
assert.ok(warnings.some((w) => w.includes("detrás del texto")));
assert.ok(!warnings.some((w) => w.includes("forma o dibujo")), "la foto y el recuadro no cuentan como formas");

assert.equal(lengthToEmu("2.5cm"), 900000);
assert.equal(lengthToEmu("96"), 914400, "sin unidad: píxeles");
assert.equal(lengthToEmu("auto"), null);
assert.equal(cropToPercent("6554f"), 10001);
assert.equal(cropToPercent(".25"), 25000);

// Un docx sin settings ni secciones no se rompe
const tiny = new JSZip();
tiny.file("word/document.xml", `${HEAD}<w:document ${NS}><w:body><w:p><w:r><w:t>Hola</w:t></w:r></w:p></w:body></w:document>`);
const plain = await prepareRenderCopy(await tiny.generateAsync({ type: "uint8array" }));
assert.deepEqual(plain.info.sections, [{}]);
assert.equal(plain.info.footnoteStart, 1);
assert.deepEqual(plain.warnings, []);

console.log("render-copy ok");
