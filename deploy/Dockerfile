# PonPower — Node/Express + Vite app for ponpower.propcortex.net.
# Runs the built-in Express server (server/index-prod.ts) which serves the SPA
# and the in-memory /api/share routes. Storage is in-memory (MemStorage), so no
# DATABASE_URL is needed here — that (Neon) is only used by the separate Vercel
# serverless function, which this container does not run. Trade-off: shared
# links are ephemeral across container restarts.
FROM node:22-bookworm-slim AS build
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY . .
# vite build -> dist/public, esbuild server -> dist/index.js; then drop dev deps.
RUN npm run build && npm prune --omit=dev

FROM node:22-bookworm-slim AS run
WORKDIR /app
ENV NODE_ENV=production \
    PORT=5000
RUN useradd -m -u 1001 app
COPY --from=build /app/node_modules ./node_modules
COPY --from=build /app/dist ./dist
COPY --from=build /app/package.json ./package.json
USER app
EXPOSE 5000
CMD ["node", "dist/index.js"]
