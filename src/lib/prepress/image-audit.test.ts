import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import JSZip from "jszip";
import { auditDocx, cropFromTransform, iccInfo, imageLevels, readJpeg, readPng, sniffImage } from "./image-audit.ts";
import {
  deobfuscateFont,
  embeddedFamilies,
  embeddedFonts,
  embeddedSkips,
  faceKey,
  missingFacesFrom,
  SUBSTITUTE_FACES,
} from "./substitutes.ts";

const enc = new TextEncoder();
const bytesOf = (...parts: Array<Uint8Array | number[] | string>) => {
  const chunks = parts.map((part) => (typeof part === "string" ? enc.encode(part) : part instanceof Uint8Array ? part : new Uint8Array(part)));
  const out = new Uint8Array(chunks.reduce((sum, chunk) => sum + chunk.length, 0));
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
};
const be16 = (n: number) => [(n >> 8) & 255, n & 255];
const be32 = (n: number) => [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255];

async function deflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([new Uint8Array(data)]).stream().pipeThrough(new CompressionStream("deflate"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Perfil ICC v2 mínimo con su etiqueta desc. */
function profile(description: string, space = "RGB "): Uint8Array {
  const text = enc.encode(`${description}\0`);
  const desc = bytesOf("desc", [0, 0, 0, 0], be32(text.length), text);
  const header = new Uint8Array(128);
  header.set(enc.encode(space), 16);
  header.set(enc.encode("acsp"), 36);
  const out = bytesOf(header, be32(1), "desc", be32(144), be32(desc.length), desc);
  out.set(be32(out.length), 0);
  return out;
}

function pngChunk(type: string, data: Uint8Array | number[]) {
  const body = data instanceof Uint8Array ? data : new Uint8Array(data);
  return bytesOf(be32(body.length), type, body, [0, 0, 0, 0]);
}

async function png(bitDepth: number, colorType: number, icc?: Uint8Array): Promise<Uint8Array> {
  const parts = [
    new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", [...be32(40), ...be32(20), bitDepth, colorType, 0, 0, 0]),
  ];
  if (icc) parts.push(pngChunk("iCCP", bytesOf("Adobe RGB\0", [0], await deflate(icc))));
  parts.push(pngChunk("IDAT", [0]), pngChunk("IEND", []));
  return bytesOf(...parts);
}

function jpeg(options: { orientation?: number; icc?: Uint8Array; components?: number; adobe?: boolean }): Uint8Array {
  const segments: Uint8Array[] = [new Uint8Array([0xff, 0xd8])];
  const segment = (marker: number, payload: Uint8Array) => segments.push(bytesOf([0xff, marker], be16(payload.length + 2), payload));
  if (options.orientation) {
    segment(0xe1, bytesOf("Exif\0\0", "MM\0*", be32(8), be16(1), be16(0x0112), be16(3), be32(1), be16(options.orientation), [0, 0], be32(0)));
  }
  if (options.icc) segment(0xe2, bytesOf("ICC_PROFILE\0", [1, 1], options.icc));
  if (options.adobe) segment(0xee, bytesOf("Adobe", [0, 100, 0, 0, 0, 0, 2]));
  const n = options.components ?? 3;
  const comps: number[] = [];
  for (let i = 0; i < n; i++) comps.push(i + 1, 0x11, 0);
  segment(0xc0, bytesOf([8], be16(30), be16(50), [n], comps));
  segments.push(new Uint8Array([0xff, 0xda, 0, 2, 0xff, 0xd9]));
  return bytesOf(...segments);
}

/* Formatos por los primeros bytes */
const emf = new Uint8Array(64);
emf.set([1, 0, 0, 0]);
emf.set(enc.encode(" EMF"), 40);
assert.equal(sniffImage(emf), "emf");
assert.equal(sniffImage(new Uint8Array([0xd7, 0xcd, 0xc6, 0x9a, 0, 0])), "wmf");
assert.equal(sniffImage(enc.encode("GIF89a......")), "gif");
assert.equal(sniffImage(enc.encode("BM\0\0\0\0")), "bmp");
assert.equal(sniffImage(enc.encode("II*\0....")), "tiff");
assert.equal(sniffImage(enc.encode("RIFF\0\0\0\0WEBPVP8 ")), "webp");
assert.equal(sniffImage(enc.encode('<?xml version="1.0"?>\n<svg xmlns="http://www.w3.org/2000/svg"/>')), "svg");
assert.equal(sniffImage(jpeg({})), "jpeg");

/* Lectura de JPEG y PNG */
const adobeRgb = profile("Adobe RGB (1998)");
assert.deepEqual(iccInfo(adobeRgb), { space: "RGB", description: "Adobe RGB (1998)" });
const rotated = readJpeg(jpeg({ orientation: 6, icc: adobeRgb }))!;
assert.equal(rotated.orientation, 6);
assert.equal(rotated.width, 50);
assert.equal(rotated.height, 30);
assert.equal(iccInfo(rotated.icc!)?.description, "Adobe RGB (1998)");
assert.equal(readJpeg(jpeg({ components: 4 }))?.adobe, false);
assert.equal(readJpeg(jpeg({ components: 4, adobe: true }))?.adobe, true);
const deep = readPng(await png(16, 6, adobeRgb))!;
assert.equal(deep.bitDepth, 16);
assert.equal(deep.alpha, true);
assert.equal(deep.iccName, "Adobe RGB");

/* Avisos del Word */
const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const pic = (id: string, extra = "", rect = "", link = false) =>
  `<w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip ${link ? "r:link" : "r:embed"}="${id}">${extra}</a:blip>${rect}<a:stretch/></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>
<w:p>${pic("rId1")}${pic("rId2", "<a:grayscl/>", '<a:srcRect l="10000"/>')}${pic("rId3")}</w:p>
<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId9"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p/></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>
<w:p>${pic("rId4")}${pic("rId5", "", "", true)}${pic("rId6")}${pic("rId7")}</w:p>
<w:sectPr/></w:body></w:document>`;
const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.jpeg"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image3.png"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image4.png"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="file:///C:/fotos/tapa.jpg" TargetMode="External"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image6.jpeg"/>
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image7.png"/>
<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>
</Relationships>`;
const zip = new JSZip();
zip.file("word/document.xml", documentXml);
zip.file("word/_rels/document.xml.rels", rels);
zip.file("word/charts/chart1.xml", "<c:chartSpace/>");
zip.file("word/media/image1.emf", emf);
zip.file("word/media/image2.jpeg", jpeg({ orientation: 6, icc: adobeRgb }));
zip.file("word/media/image3.png", await png(16, 6, adobeRgb));
// El nombre dice PNG, pero es un GIF: manda lo que hay adentro.
zip.file("word/media/image4.png", enc.encode("GIF89a\x01\x00\x01\x00"));
zip.file("word/media/image6.jpeg", jpeg({ components: 4 }));
zip.file("word/media/image7.png", await png(8, 2));
zip.file("word/media/huerfana.tif", enc.encode("II*\0...."));
const warnings = await auditDocx(await zip.generateAsync({ type: "uint8array" }));
const has = (pattern: RegExp, message: string) => assert.ok(warnings.some((line) => pattern.test(line)), `${message}\n${warnings.join("\n")}`);

has(/«image1\.emf».*EMF o WMF.*no va a salir/, "EMF: no sale");
has(/«image4\.png».*GIF, BMP, WebP o SVG.*se redibuja/, "GIF detectado por su contenido, no por el nombre");
has(/«image2\.jpeg».*EXIF.*girada/, "orientación EXIF");
has(/image2\.jpeg» \(«Adobe RGB \(1998\)».*no es sRGB/, "perfil de JPEG con su nombre");
has(/image3\.png» \(«Adobe RGB \(1998\)».*no es sRGB/, "perfil de PNG (iCCP descomprimido)");
has(/«image3\.png».*16 bits.*8 bits/, "PNG de 16 bits");
has(/«image3\.png».*transparencia.*blanco/, "PNG con alfa");
has(/«image6\.jpeg».*CMYK sin la marca de Adobe.*negativo/, "CMYK sin Adobe");
has(/vinculada.*«tapa\.jpg»|«tapa\.jpg».*vinculada/, "imagen vinculada");
has(/un gráfico de Word.*pegalo como imagen|un gráfico de Word.*Copialo y pegalo como imagen/i, "gráfico");
has(/un cuadro de texto o forma: puede no salir/, "el cuadro de texto se cuenta una vez (sin el Fallback)");
has(/efecto de Word \(escala de grises\).*no se aplica/, "efecto de imagen");
has(/recortada dentro del Word/, "recorte");
assert.ok(!warnings.some((line) => line.includes("image7.png")), "una PNG común no da avisos");
assert.ok(!warnings.some((line) => line.includes("huerfana")), "lo que no se usa no se revisa");

/* Resolución: ppp de la parte visible */
assert.deepEqual(
  { ...imageLevels({ naturalWidth: 3000, naturalHeight: 3000, widthIn: 10, heightIn: 10 }) },
  { ppi: 300, level: "ok", note: "sin recomprimir", visibleWidth: 3000, visibleHeight: 3000 },
);
const croppedLevel = imageLevels({ naturalWidth: 2000, naturalHeight: 2000, widthIn: 5, heightIn: 5, crop: { l: 25000, r: 25000 } });
assert.equal(croppedLevel.visibleWidth, 1000);
assert.equal(croppedLevel.ppi, 200, "el recorte baja los ppp: se mide lo que se ve");
assert.equal(croppedLevel.level, "baja");
assert.equal(imageLevels({ naturalWidth: 300, naturalHeight: 300, widthIn: 5, heightIn: 5 }).level, "rechazo");
assert.equal(imageLevels({ naturalWidth: 2000, naturalHeight: 2000, widthIn: 5, heightIn: 5, lineArt: true }).level, "linea");
assert.equal(imageLevels({ naturalWidth: 0, naturalHeight: 0, widthIn: 5, heightIn: 5 }).level, "rota");
assert.equal(imageLevels({ naturalWidth: 10, naturalHeight: 10, widthIn: 0, heightIn: 0 }).level, "sin-medida");
const fromCss = cropFromTransform("rotate(90deg) scale(2, 1.25)");
assert.equal(Math.round(fromCss.l ?? 0), 50000);
assert.equal(Math.round(fromCss.t ?? 0), 20000);

/* Sustitutas: Times y Helvetica con sus cuatro caras */
for (const family of ["Times", "Helvetica", "Courier"]) {
  const keys = new Set(SUBSTITUTE_FACES.filter((face) => face.family === family).map((face) => faceKey(face.family, face.weight, face.style)));
  assert.equal(keys.size, 4, `${family} tiene normal, negrita, cursiva y negrita cursiva`);
}

/* Caras que faltan */
const missing = missingFacesFrom(
  [
    { family: "Calibri", weight: 400, style: "normal", chars: 900, sample: "texto" },
    { family: "Calibri", weight: 700, style: "italic", chars: 12, sample: "énfasis" },
    { family: "Tahoma", weight: 400, style: "italic", chars: 30, sample: "nota" },
    { family: "Garamond", weight: 400, style: "normal", chars: 50, sample: "cita" },
  ],
  [
    { family: "Calibri", weight: 400, style: "normal" },
    { family: "Calibri", weight: 700, style: "normal" },
    { family: "Calibri", weight: 400, style: "italic" },
    { family: "Tahoma", weight: [100, 900], style: "normal" },
  ],
);
assert.deepEqual(
  missing.map((face) => [face.label, face.familyMissing]),
  [
    ["Garamond normal", true],
    ["Tahoma cursiva", false],
    ["Calibri negrita cursiva", false],
  ],
);

/* Fuentes incrustadas: se leen de fontTable y se descifran como docx-preview */
const carlito = new Uint8Array(await readFile(new URL("../../../public/fonts/Carlito-Regular.ttf", import.meta.url)));
const guid = "{0E2E7E6A-4C4B-4B5B-9F1E-2B3C4D5E6F70}";
const odttf = deobfuscateFont(carlito, guid);
assert.notDeepEqual(odttf.subarray(0, 32), carlito.subarray(0, 32), "cifrada");
assert.deepEqual(deobfuscateFont(odttf, guid), carlito, "descifrar es volver al original");
const fontZip = new JSZip();
fontZip.file(
  "word/fontTable.xml",
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:embedRegular r:id="rId1" w:fontKey="${guid}" w:subsetted="1"/><w:embedItalic r:id="rId2" w:fontKey="${guid}"/></w:font>
<w:font w:name="Times New Roman"/>
</w:fonts>`,
);
fontZip.file(
  "word/_rels/fontTable.xml.rels",
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.odttf"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font2.odttf"/>
</Relationships>`,
);
fontZip.file("word/fonts/font1.odttf", odttf);
fontZip.file("word/fonts/font2.odttf", new Uint8Array(64));
const fontDocx = await fontZip.generateAsync({ type: "uint8array" });
const refs = await embeddedFamilies(fontDocx);
assert.deepEqual(
  refs.map((ref) => [ref.family, ref.weight, ref.style, ref.subsetted, ref.path]),
  [
    ["Calibri", 400, "normal", true, "word/fonts/font1.odttf"],
    ["Calibri", 400, "italic", false, "word/fonts/font2.odttf"],
  ],
);
const embedded = await embeddedFonts(fontDocx);
assert.equal(embedded.faces.length, 1, "la que no es una fuente no se usa");
assert.deepEqual(embedded.faces[0]!.bytes, carlito);
assert.equal(embedded.faces[0]!.restricted, false);
assert.equal(embedded.warnings.length, 1);
assert.match(embedded.warnings[0]!, /«Calibri» \(cursiva\) no se pudo leer/);
const skips = embeddedSkips(embedded.faces);
assert.ok(skips.has(faceKey("Calibri", 400, "normal")), "la sustituta normal de Calibri se salta");
assert.ok(!skips.has(faceKey("Calibri", 700, "normal")), "la negrita sustituta sigue");

console.log("image-audit ok", warnings.length, "avisos");
