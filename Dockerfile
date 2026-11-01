# syntax=docker/dockerfile:1
FROM --platform=$BUILDPLATFORM node:24-bookworm-slim AS frontend
WORKDIR /src/frontend
COPY frontend/package*.json ./
RUN --mount=type=secret,id=build_ca,mode=0444 \
    if [ -f /run/secrets/build_ca ]; then export NODE_EXTRA_CA_CERTS=/run/secrets/build_ca; fi; npm ci --no-fund --no-audit
COPY frontend/ ./
RUN npm run build

FROM python:3.11-slim-bookworm AS wheels
RUN --mount=type=secret,id=build_ca,mode=0444 \
    if [ -f /run/secrets/build_ca ]; then export SSL_CERT_FILE=/run/secrets/build_ca; fi; \
    sed -i 's|http://deb.debian.org|https://deb.debian.org|g' /etc/apt/sources.list.d/debian.sources && apt-get -o Acquire::https::CaInfo="${SSL_CERT_FILE:-/etc/ssl/certs/ca-certificates.crt}" update && apt-get -o Acquire::https::CaInfo="${SSL_CERT_FILE:-/etc/ssl/certs/ca-certificates.crt}" install -y --no-install-recommends build-essential autoconf automake libtool pkg-config && rm -rf /var/lib/apt/lists/*
WORKDIR /src
COPY pyproject.toml README.md LICENSE ./
COPY .github/ci-constraints.txt .github/ci-constraints.txt
COPY edf_bill_fetcher/ edf_bill_fetcher/
COPY --from=frontend /src/edf_bill_fetcher/web/static/ edf_bill_fetcher/web/static/
RUN --mount=type=secret,id=build_ca,mode=0444 \
    if [ -f /run/secrets/build_ca ]; then export PIP_CERT=/run/secrets/build_ca; fi; pip wheel --wheel-dir /wheels -c .github/ci-constraints.txt '.[web,msg]'

FROM python:3.11-slim-bookworm
RUN --mount=type=secret,id=build_ca,mode=0444 \
    if [ -f /run/secrets/build_ca ]; then export SSL_CERT_FILE=/run/secrets/build_ca; fi; \
    sed -i 's|http://deb.debian.org|https://deb.debian.org|g' /etc/apt/sources.list.d/debian.sources && apt-get -o Acquire::https::CaInfo="${SSL_CERT_FILE:-/etc/ssl/certs/ca-certificates.crt}" update && apt-get -o Acquire::https::CaInfo="${SSL_CERT_FILE:-/etc/ssl/certs/ca-certificates.crt}" install -y --no-install-recommends tesseract-ocr tesseract-ocr-eng libgl1 libegl1 libopengl0 libxkbcommon0 libdbus-1-3 libfontconfig1 fonts-dejavu-core && rm -rf /var/lib/apt/lists/*
COPY --from=wheels /wheels /wheels
RUN pip install --no-cache-dir --no-index --find-links=/wheels 'edf-bill-fetcher[web,msg]' && rm -rf /wheels && useradd --create-home --uid 1000 edf && mkdir /data && chown edf:edf /data
USER edf
WORKDIR /data
ENV PYTHONUNBUFFERED=1
EXPOSE 8000
ENTRYPOINT ["edf-web", "--data", "/data"]
CMD ["--host", "0.0.0.0", "--port", "8000"]
