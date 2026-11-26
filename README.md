# Pliego

Convierte un Word (.docx) en un PDF para premedia, en el navegador.

- Quita marcas de revisión y comentarios.
- Deja las notas al pie.
- No vuelve a comprimir las imágenes.
- Incrusta las fuentes completas que ya trae la app (Calibri, Cambria, Arial, Times New Roman, Courier New, Tahoma, Candara, MS Gothic, Symbol y Wingdings), una sola vez por archivo.
- Respeta los estilos del Word y el tamaño de la página: lo que no entra pasa a una página nueva, con sus notas al pie.
- Numera las páginas de verdad (inicio y formato de cada sección, `\* roman`).
- Para IngramSpark / Lightning Source: caja de corte declarada, negro en gris puro y, con «Interior en blanco y negro»,
  las imágenes PNG en escala de grises sin pérdida. Si todo está en gris con las fuentes incrustadas, el PDF sale
  PDF/X-1a:2001; si no, la app dice qué falta.

## Deploy

En el VPS corre en https://pliego.srv1165239.hstgr.cloud; cómo instalarlo y actualizarlo está en [deploy/README.md](deploy/README.md).

## Uso

```bash
npm install
npm run dev
```

Abre la dirección que imprime Vite y suelta un `.docx`.
