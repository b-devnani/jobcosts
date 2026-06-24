# Portable image for the Job Cost Projection tool.
# Works on Render (Docker runtime), Fly.io, Railway, or any Docker host.
FROM python:3.11-slim

WORKDIR /app

# Install dependencies first for better layer caching.
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App code (template + seed CSV + static frontend are included).
COPY . .

# Persist the SQLite database on a mounted volume at /data.
ENV JOBCOSTS_DB=/data/jobcosts.db
RUN mkdir -p /data
VOLUME ["/data"]

EXPOSE 8000

# Honour the platform-provided $PORT, defaulting to 8000 locally.
CMD ["sh", "-c", "uvicorn backend.app:app --host 0.0.0.0 --port ${PORT:-8000}"]
