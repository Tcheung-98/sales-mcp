FROM python:3.12-slim

WORKDIR /app

# LibreOffice headless: PPTX → PDF for vision QA (B1 render_slides).
# poppler-utils: PDF pages → PNG. fonts: avoid blank/missing glyphs in renders.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    fonts-liberation \
    fonts-dejavu-core \
    fontconfig \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Install uv for dependency management
RUN pip install --no-cache-dir uv

# Copy dep files first for Docker layer caching
COPY pyproject.toml uv.lock* ./

# Install third-party deps only (project needs server.py / ingestion, copied next)
RUN uv sync --frozen --no-dev --no-install-project

# Copy application code, then install the local package
COPY . .
RUN uv sync --frozen --no-dev

# FastMCP SSE listens here
EXPOSE 8000

# Production ASGI server
CMD ["uv", "run", "uvicorn", "server:app", \
     "--host", "0.0.0.0", \
     "--port", "8000", \
     "--proxy-headers", \
     "--forwarded-allow-ips", "*"]
