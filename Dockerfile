FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY pyproject.toml README.md ./
COPY src ./src
COPY storage ./storage

RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir .

CMD ["sh", "-c", "uvicorn a3presentation.main:app --host 0.0.0.0 --port ${PORT:-8000}"]
