# syntax=docker/dockerfile:1
FROM node:22-bookworm-slim AS node-runtime

FROM node-runtime AS frontend
RUN apt-get update && apt-get install -y --no-install-recommends git ca-certificates \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app/web
COPY web/package.json web/package-lock.json ./
RUN npm ci
COPY web/ ./
RUN npm run build

FROM python:3.11-slim-bookworm AS python-dependencies
WORKDIR /build
COPY pyproject.toml README.md LICENSE THIRD_PARTY_NOTICES.md ./
COPY src/ ./src/
RUN pip install --no-cache-dir --prefix=/install '.[web]'

FROM python:3.11-slim-bookworm AS runtime
RUN apt-get update && apt-get install -y --no-install-recommends ca-certificates libstdc++6 \
    && rm -rf /var/lib/apt/lists/* \
    && groupadd --gid 10001 jakal \
    && useradd --uid 10001 --gid 10001 --no-create-home --home-dir /tmp jakal \
    && install -d -m 0700 -o jakal -g jakal /var/lib/jakal
WORKDIR /app
COPY --from=node-runtime /usr/local/bin/node /usr/local/bin/node
COPY --from=python-dependencies /install/ /usr/local/
COPY server/ ./server/
COPY LICENSE THIRD_PARTY_NOTICES.md ./
COPY --from=frontend /app/web/dist/ ./web/dist/
# The converter imports the package and WASM at runtime; Node is not build-only.
COPY --from=frontend /app/web/node_modules/@rhwp/core/ ./web/node_modules/@rhwp/core/
ENV JAKAL_ENV=production \
    JAKAL_HOST=0.0.0.0 \
    JAKAL_DATA_DIR=/var/lib/jakal \
    HOME=/tmp \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1
USER 10001:10001
RUN python -c "from server.runtime import check_runtime; check_runtime()" \
    && node --input-type=module -e "import fs from 'node:fs'; import {initSync} from './web/node_modules/@rhwp/core/rhwp.js'; initSync({module: fs.readFileSync('./web/node_modules/@rhwp/core/rhwp_bg.wasm')});"
EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
    CMD ["python", "-m", "server.healthcheck"]
CMD ["python", "-m", "server"]
