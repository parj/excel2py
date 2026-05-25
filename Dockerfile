# Multi-stage Dockerfile for excel2py
# Build stage uses uv for fast dependency resolution
# Runtime stage uses distroless Python for minimal attack surface

# Build stage
FROM ghcr.io/astral-sh/uv:python3.14-trixie AS builder

WORKDIR /app
COPY pyproject.toml uv.lock ./

# Install dependencies without project source for better layer caching
RUN uv sync --frozen --no-dev --no-install-project

# Copy application code after deps are installed
COPY src/ ./src/

# Runtime stage
FROM nvcr.io/nvidia/distroless/python:3.14-v4.0.3

# Copy only the site-packages (not venv binaries)
COPY --from=builder /app/.venv/lib/python3.14/site-packages /usr/local/lib/python3.14/site-packages

# Copy application code
COPY --from=builder /app/src /app/src

WORKDIR /app
ENV PYTHONPATH=/usr/local/lib/python3.14/site-packages:/app/src

USER nonroot
ENTRYPOINT ["python3.14", "-m", "excel2py"]
