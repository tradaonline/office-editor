# ============================================================
# Global build arguments (declared before any FROM so they can
# be used in FROM lines; must be re-declared inside each stage
# to be visible there).
# ============================================================

# OnlyOffice DocumentServer version — controls both the source image
# tag AND the versioned asset directory prefix (/v<DS_VERSION>-<HASH>).
ARG DS_VERSION=9.3.1

# Revision counter. Bump this (--build-arg HASH=2) whenever you want
# to bust the browser cache for the OnlyOffice assets without changing
# the DocumentServer version itself.
ARG HASH=1

# Harness inclusion: 0 = production (no :81 site, no /srv-harness), 1 = include harness
ARG XERP_HARNESS=0

# Public source repository URL exposed to /xerp/about
ARG XERP_SOURCE_URL="https://github.com/tradaonline/office-editor"

# ============================================================
# Stage 1: OnlyOffice DocumentServer assets source
# ============================================================
FROM onlyoffice/documentserver:${DS_VERSION} AS documentserver

# AllFonts.js and themes.js are NOT present in the image — they are
# generated at container startup by documentserver-generate-allfonts.sh.
# We run that script here (passing `false` so it skips the data-container
# wait branch) so the files exist before the COPY in the final stage.
RUN documentserver-generate-allfonts.sh false

# ============================================================
# Stage 2: Next.js website builder
# ============================================================
FROM node:22-alpine AS builder

# Re-declare args inside this stage to make them visible here.
ARG DS_VERSION
ARG HASH

# XERP: origins allowed to embed /xerp and talk to it (comma-separated).
ARG XERP_PARENT_ORIGINS=""
ARG XERP_SOURCE_URL="https://github.com/tradaonline/office-editor"
ARG XERP_HARNESS=0

# Expose the versioned asset path to Next.js at build time.
ENV NEXT_PUBLIC_APP_ROOT=/v${DS_VERSION}-${HASH}
ENV NEXT_PUBLIC_XERP_PARENT_ORIGINS=${XERP_PARENT_ORIGINS}
ENV NEXT_PUBLIC_XERP_SOURCE_URL=${XERP_SOURCE_URL}

WORKDIR /app

# Install pnpm
# XERP: pinned; newer pnpm turns the ignored build scripts into an error.
RUN corepack enable && corepack prepare pnpm@10.17.1 --activate

# Copy dependency manifests first for better layer caching.
COPY package.json pnpm-lock.yaml ./

# Install dependencies (frozen lockfile for reproducibility).
RUN pnpm install --frozen-lockfile

# Copy the rest of the source code.
COPY . .

# Run the Next.js static export build. XERP: drop the upstream sample
# templates (about 45 MB) — the embed never offers them.
RUN pnpm build && rm -rf out/files \
 && mkdir -p /harness \
 && if [ "$XERP_HARNESS" = "1" ]; then cp -r xerp/harness/. /harness/; fi

# XERP: x2t ships brotli-encoded, and browsers accept brotli only over HTTPS.
# Keep each file plain with .br and .gz siblings; Caddy picks per request.
RUN apk add --no-cache brotli \
 && for d in out/x2t out/x2t-1; do for f in "$d/x2t.js" "$d/x2t.wasm"; do \
      if brotli -t "$f" 2>/dev/null; then mv "$f" "$f.br" && brotli -d "$f.br" -o "$f"; \
      else brotli -q 11 -k -f "$f"; fi; \
      gzip -k -9 -f "$f"; \
    done; done

# ============================================================
# Stage 3: Caddy production server
# ============================================================
FROM caddy:2-alpine AS final

# Re-declare args inside this stage.
ARG DS_VERSION
ARG HASH
ARG XERP_HARNESS=0

WORKDIR /srv

# Copy the Next.js static export output.
COPY --from=builder /app/out ./

# Copy OnlyOffice DocumentServer assets directly from the source stage
# into the versioned directory — assets never pass through the builder,
# so there is no redundant copy of the large asset tree.
COPY --from=documentserver /var/www/onlyoffice/documentserver/fonts         ./v${DS_VERSION}-${HASH}/fonts
COPY --from=documentserver /var/www/onlyoffice/documentserver/sdkjs         ./v${DS_VERSION}-${HASH}/sdkjs
COPY --from=documentserver /var/www/onlyoffice/documentserver/web-apps      ./v${DS_VERSION}-${HASH}/web-apps
# XERP: no sdkjs-plugins — the embed runs with plugins off.

# api.js is generated from a template at runtime in a full DocumentServer
# deployment, but here we serve it statically — copy the template as-is.
RUN cp "./v${DS_VERSION}-${HASH}/web-apps/apps/api/documents/api.js.tpl" \
       "./v${DS_VERSION}-${HASH}/web-apps/apps/api/documents/api.js"

# Copy Caddyfile.
COPY Caddyfile /etc/caddy/Caddyfile
COPY Caddyfile.harness /tmp/Caddyfile.harness
COPY --from=builder /harness /tmp/harness

# XERP: harness on :81 conditional on build arg XERP_HARNESS=1 (default 0).
RUN if [ "$XERP_HARNESS" = "1" ]; then \
      cp -r /tmp/harness /srv-harness && \
      cat /tmp/Caddyfile.harness >> /etc/caddy/Caddyfile; \
    fi && \
    rm -rf /tmp/harness /tmp/Caddyfile.harness

EXPOSE 80 443
