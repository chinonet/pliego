# Pliego

Convierte un Word (.docx) en un PDF para premedia, en el navegador.

- Quita marcas de revisión y comentarios.
- Deja las notas al pie.
- No vuelve a comprimir las imágenes.
- Incrusta las fuentes completas que ya trae la app (Calibri, Cambria, Arial, Times New Roman, Courier New, Tahoma, Candara, MS Gothic, Symbol y Wingdings).

## Deploy

En el VPS corre en https://pliego.srv1165239.hstgr.cloud; cómo instalarlo y actualizarlo está en [deploy/README.md](deploy/README.md).

## Uso

```bash
npm install
npm run dev
```

Abre la dirección que imprime Vite y suelta un `.docx`.
