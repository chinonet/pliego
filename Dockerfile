# Pliego en Docker (VPS). Dos etapas: la primera compila con Vite, la segunda deja solo el servidor armado.
FROM node:22-bookworm-slim AS build
WORKDIR /app
ENV NPM_CONFIG_UPDATE_NOTIFIER=false
COPY package.json package-lock.json ./
RUN npm ci --omit=dev --no-audit --no-fund
COPY . .
# Sin cuentas ni base de datos (Pliego convierte todo en el navegador).
# node-server: servidor Node común en vez del formato de Vercel.
ENV VITE_AUTH_ENABLED=false \
    NITRO_PRESET=node-server \
    NODE_OPTIONS=--max-old-space-size=1536
RUN npm run build

FROM node:22-bookworm-slim
WORKDIR /app
ENV NODE_ENV=production HOST=0.0.0.0 PORT=8080
COPY --from=build /app/.output ./.output
USER node
EXPOSE 8080
CMD ["node", ".output/server/index.mjs"]
