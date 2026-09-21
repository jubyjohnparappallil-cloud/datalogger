FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    CLOUD_MODE=1 \
    DATA_DIR=/data

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /data/output /data/uploads

EXPOSE 10000
CMD gunicorn app:app --workers 1 --threads 8 --timeout 600 --bind 0.0.0.0:${PORT:-10000}
