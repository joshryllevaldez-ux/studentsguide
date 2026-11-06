# Production container for the shared StudentsGUIDE service.
FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=10000

WORKDIR /app

COPY requirements.txt ./requirements.txt
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY app.py ./app.py
COPY static ./static
COPY assets ./assets

EXPOSE 10000

# One worker keeps the lightweight in-memory classroom rate limiter consistent.
# Render terminates TLS at its proxy; Uvicorn trusts forwarded headers from Render.
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-10000} --proxy-headers --forwarded-allow-ips='*' --workers 1"]
