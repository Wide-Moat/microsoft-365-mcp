# microsoft-365-mcp: one image, three listeners (MCP on 8000, signed links on 8001, metrics and probes on 9090).
# The sealing key and the store's passwords are files staged by the chart's init container, never the environment.
#
# BASE_IMAGE: a registry mirror or proxy may stand in for Docker Hub; keep the digest, so the base cannot change
# under the same name.
ARG BASE_IMAGE=python:3.12-slim-bookworm@sha256:6a0e06f5ba0b373d051aabd3fffb9b873f27f9a2a3a14bcbe37dee6f0e3bbc70

FROM ${BASE_IMAGE} AS build
WORKDIR /src
COPY requirements.lock ./
RUN pip install --no-cache-dir --require-hashes --target /deps -r requirements.lock
# The package is pure Python: copied as source, so no unlocked build backend runs.
COPY src/m365_mcp /app/m365_mcp

FROM ${BASE_IMAGE}
RUN apt-get update && apt-get upgrade -y --no-install-recommends && rm -rf /var/lib/apt/lists/* \
 && useradd --uid 10001 --no-create-home --shell /usr/sbin/nologin app
COPY --from=build /deps /opt/deps
COPY --from=build /app /opt/app
ENV PYTHONPATH=/opt/app:/opt/deps PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PORT=8000 LINKS_PORT=8001
USER 10001
EXPOSE 8000 8001 9090
ENTRYPOINT ["python", "-m", "m365_mcp.main"]
