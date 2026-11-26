/**
 * jsafa-pdf incrusta las fuentes con pdf-lib sin decir qué «features» de OpenType usar, y pdf-lib (fontkit) aplica las
 * de siempre: ligaduras (con Cambria, «fi» pasa a ser un solo signo, que además al copiar el texto del PDF sale como
 * «ϐ») e interletraje (kerning). Word no usa ninguna de las dos salvo que el texto las pida, y la vista previa tampoco
 * (styles.css). Este plugin las apaga en las llamadas de jsafa que incrustan fuentes. Si jsafa cambia y no encuentra
 * esas llamadas, el build falla en vez de dejar pasar las ligaduras sin aviso.
 */

/**
 * Lo que se apaga: ligaduras comunes, contextuales, opcionales e históricas, interletraje y la descomposición de letras
 * acentuadas (ccmp: con Cambria, «Í» pasaba a «I» más un acento suelto que al copiar el texto del PDF salía «IƵ»; la
 * letra ya viene compuesta en el texto y la fuente la tiene entera).
 */
export const PDF_FONT_FEATURES = "{ liga: false, clig: false, dlig: false, hlig: false, kern: false, ccmp: false }";

const CALL = /\.embedFont\((\w+)\.bytes, \{ subset: ([^{}]+?) \}\)/g;

/** Devuelve el código de jsafa con las «features» puestas y cuántas llamadas cambió. */
export function patchJsafa(code) {
  let count = 0;
  const patched = code.replace(CALL, (_all, bytes, subset) => {
    count += 1;
    return `.embedFont(${bytes}.bytes, { subset: ${subset}, features: ${PDF_FONT_FEATURES} })`;
  });
  return { code: patched, count };
}

const TARGET = /[\\/]jsafa-pdf[\\/]dist[\\/]jsafa-pdf\.js$/;

/** Plugin de Vite (build y dev). */
export function jsafaFeaturesPlugin() {
  return {
    name: "pliego:jsafa-features",
    enforce: "pre",
    transform(code, id) {
      if (!TARGET.test(id.split("?")[0])) return null;
      const { code: patched, count } = patchJsafa(code);
      if (!count) {
        this.error(
          "jsafa-pdf cambió: no se encontraron las llamadas que incrustan las fuentes. Revisá scripts/jsafa-features-plugin.mjs (sin esto el PDF sale con ligaduras que Word no usa).",
        );
      }
      return { code: patched, map: null };
    },
  };
}
