# Deploy en el VPS

VPS Hostinger `root@31.97.90.21`, con el repositorio en `/opt/pliego`.
La app queda en https://pliego.srv1165239.hstgr.cloud.

Pliego convierte el Word en el navegador: no usa cuentas ni base de datos. En el VPS corre un solo
contenedor (`pliego-web-1`) con el servidor Node armado por Vite, detrás del Traefik compartido que ya
da HTTPS a publishing-ai y a los demás servicios (red `root_default`, certificados `mytlschallenge`).

## Primera instalación
```bash
git clone https://github.com/chinonet/pliego /opt/pliego
cd /opt/pliego/deploy
cp .env.example .env
chmod 600 .env
docker compose up -d --build
```
La primera vez tarda unos minutos (baja Node y compila). Después de levantar, Traefik responde 404
durante 30-50 segundos hasta que pasa el chequeo de salud; es normal.

## Actualizar a la última versión de GitHub
```bash
cd /opt/pliego && git pull && cd deploy && docker compose up -d --build
```

## Ver estado y logs
```bash
cd /opt/pliego/deploy
docker compose ps
docker compose logs -f web
```

## Reiniciar o apagar
```bash
cd /opt/pliego/deploy && docker compose restart
cd /opt/pliego/deploy && docker compose down
```
Los contenedores arrancan solos cuando se reinicia el VPS (`restart: always`).

## Cómo está armado
- `Dockerfile` (raíz del repo): compila con `VITE_AUTH_ENABLED=false` (sin cuentas) y
  `NITRO_PRESET=node-server` (servidor Node común en vez del formato de Vercel, que es el que usa
  `vite.config.ts` por defecto). Solo queda la carpeta `.output` en la imagen final.
- `deploy/docker-compose.yml`: el servicio `web` en el puerto 8080 interno, con las etiquetas de
  Traefik para el dominio de `APP_DOMAIN`.
- `deploy/.env`: solo `APP_DOMAIN`. Vive únicamente en el VPS (está en `.gitignore`).
