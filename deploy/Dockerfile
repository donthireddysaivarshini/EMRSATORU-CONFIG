# Satoru EMR tools: Excel exports + forms panel, served under /exports/.
FROM node:22-bookworm-slim AS node

FROM python:3.12-slim-bookworm

# Node (for cht-conf and the scoring tests) + Chromium (runs the scoring tests
# in CHT's real form engine). Chromium only runs while a form is being checked.
COPY --from=node /usr/local/bin/node /usr/local/bin/node
COPY --from=node /usr/local/lib/node_modules /usr/local/lib/node_modules
RUN ln -s /usr/local/lib/node_modules/npm/bin/npm-cli.js /usr/local/bin/npm \
 && ln -s /usr/local/lib/node_modules/npm/bin/npx-cli.js /usr/local/bin/npx \
 && apt-get update \
 && apt-get install -y --no-install-recommends chromium fonts-noto-core xsltproc \
 && rm -rf /var/lib/apt/lists/* \
 && npm install -g cht-conf@6.6.1 --no-audit --no-fund \
 && npm cache clean --force

WORKDIR /app

COPY scripts/requirements.txt /app/scripts/requirements.txt
RUN pip install --no-cache-dir -r /app/scripts/requirements.txt

ENV CHROMIUM_PATH=/usr/bin/chromium \
    CHROMIUM_NO_SANDBOX=1 \
    PUPPETEER_SKIP_DOWNLOAD=1 \
    PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=1

COPY tests/forms/package.json tests/forms/package-lock.json /app/tests/forms/
RUN cd /app/tests/forms && npm ci --no-audit --no-fund && npm cache clean --force

COPY tests/forms/        /app/tests/forms/
RUN cd /app/tests/forms && node setup.js

# the exporter reads translations at runtime; form XMLs are read live from
# CouchDB, the copies here are only the fallback for command-line use
COPY scripts/            /app/scripts/
COPY forms/app/          /app/forms/app/
COPY translations/       /app/translations/

ENV PYTHONUNBUFFERED=1
EXPOSE 8095

# COUCH_URL is supplied by compose
CMD ["sh", "-c", "python scripts/export_server.py --url \"$COUCH_URL\" --host 0.0.0.0 --port 8095 --base-path /exports"]
