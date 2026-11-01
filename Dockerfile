# m365-mcp: one image, two listeners (MCP on 8000, signed links on 8001).
FROM python:3.15-slim-bookworm@sha256:4e03260adeeebafccd084f3584e6b84823aba5651cfb53ef8a078ed441ddf9cd AS build
WORKDIR /src
COPY requirements.lock ./
RUN pip install --no-cache-dir --require-hashes --target /deps -r requirements.lock
COPY pyproject.toml README.md ./
COPY src ./src
RUN pip install --no-cache-dir --no-deps --target /app .

FROM python:3.15-slim-bookworm@sha256:4e03260adeeebafccd084f3584e6b84823aba5651cfb53ef8a078ed441ddf9cd
RUN apt-get update && apt-get upgrade -y --no-install-recommends && rm -rf /var/lib/apt/lists/* \
 && useradd --uid 10001 --no-create-home --shell /usr/sbin/nologin app
COPY --from=build /deps /opt/deps
COPY --from=build /app /opt/app
ENV PYTHONPATH=/opt/app:/opt/deps PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PORT=8000 LINKS_PORT=8001
USER 10001
EXPOSE 8000 8001
ENTRYPOINT ["python", "-m", "m365_mcp.app"]
