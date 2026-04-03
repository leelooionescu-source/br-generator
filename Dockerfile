FROM python:3.12-slim

# Install Tesseract OCR with Romanian language
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-ron \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Create upload/output dirs
RUN mkdir -p uploads output

EXPOSE 10000

CMD gunicorn --bind 0.0.0.0:${PORT:-10000} --timeout 300 --workers 2 app:app
