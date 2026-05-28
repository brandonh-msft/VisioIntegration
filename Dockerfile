FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    VISIO_MCP_FORCE_DRAWIO=1

WORKDIR /app

COPY pyproject.toml /app/pyproject.toml
COPY src /app/src

RUN pip install --upgrade pip \
    && pip install -e . \
    # Keep Chromium and its Linux dependencies in the image so import_pricing_estimate works in-container.
    && python -m playwright install --with-deps chromium

ENTRYPOINT ["python", "-m", "visio_mcp.server"]
