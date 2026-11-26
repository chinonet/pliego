# Pliego

Convierte un Word (.docx) en un PDF para premedia, en el navegador.

- Quita marcas de revisión (también filas, celdas y fines de párrafo borrados) y comentarios; deja las notas al pie.
- No vuelve a comprimir las imágenes: las fotos pasan tal cual y los PNG con sus mismos píxeles.
- Usa las fuentes que el Word trae incrustadas y completa con las que ya trae la app (Calibri, Cambria, Arial, Times
  New Roman, Courier New, Tahoma, Candara, MS Gothic, Symbol y Wingdings); cada fuente va una sola vez en el PDF.
- Pagina como Word, renglón por renglón: interlineado de cada letra, párrafos partidos con viudas y huérfanas, tablas
  partidas por filas (con sus filas de título), sin espacio anterior arriba de la página, notas al pie con su llamada y
  numeradas como en el Word, márgenes simétricos y de encuadernación, capítulos en página impar o par, secciones
  continuas, encabezados y pies por sección y por página par o impar, números de página reales (`\* roman`).
- Lo que el PDF no vería por sí solo pasa a texto antes de exportar: números de lista, versalitas y mayúsculas de
  estilo, subrayados, puntos guía, guiones de corte; el justificado no estira las letras; las imágenes recortadas salen
  recortadas.
- Avisa lo que no puede pasar bien: imágenes EMF/WMF/TIFF, gráficos y SmartArt, cuadros de texto, fotos giradas o con
  otro perfil de color, variantes de letra que faltan, imágenes de poca resolución.
- Para IngramSpark / Lightning Source: hoja exacta al corte, caja de corte declarada, sangrado opcional de 0,125",
  negro en gris puro y, con «Interior en blanco y negro», imágenes PNG en gris sin pérdida. Si todo está en gris con las
  fuentes incrustadas, el PDF sale PDF/X-1a:2001; si no, la app dice qué falta.

## Deploy

En el VPS corre en https://pliego.srv1165239.hstgr.cloud; cómo instalarlo y actualizarlo está en [deploy/README.md](deploy/README.md).

## Uso

```bash
npm install
npm run dev
```

Abre la dirección que imprime Vite y suelta un `.docx`.
