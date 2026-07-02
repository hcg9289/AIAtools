FROM python:3.11-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-noto-cjk \
    fontconfig \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-chi-tra \
    && fc-cache -f \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .
COPY templates/ templates/
COPY assets/ assets/

# Level 3: 建立非 root 專屬用戶 (UID 1008)
RUN groupadd -g 1008 appgroup && useradd -u 1008 -g appgroup -s /bin/sh -d /app -M appuser \
    && mkdir -p uploads outputs \
    && chown -R appuser:appgroup /app

USER appuser

ENV FLASK_APP=app.py
ENV FLASK_ENV=production

EXPOSE 5008

CMD ["gunicorn", "--worker-class", "gthread", "--workers", "1", "--threads", "4", "--timeout", "600", "--bind", "0.0.0.0:5008", "app:app"]
