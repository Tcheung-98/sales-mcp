FROM python:3.12-slim

WORKDIR /app

# Install uv for dependency management
RUN pip install --no-cache-dir uv

# Copy dep files first for Docker layer caching
COPY pyproject.toml uv.lock* ./

# Install dependencies (production only)
RUN uv sync --frozen --no-dev

# Copy application code
COPY . .

# FastMCP SSE listens here
EXPOSE 8000

# Production ASGI server
CMD ["uv", "run", "uvicorn", "server:app", \
     "--host", "0.0.0.0", \
     "--port", "8000", \
     "--proxy-headers", \
     "--forwarded-allow-ips", "*"]