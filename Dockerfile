FROM python:3.11-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    chromium \
    fonts-noto-cjk \
    fontconfig \
    && fc-cache -f \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

# Level 3: 建立非 root 專屬用戶 (UID 1008)
RUN groupadd -g 1008 appgroup && useradd -u 1008 -g appgroup -s /bin/sh -d /app -M appuser \
    && mkdir -p uploads outputs \
    && chown -R appuser:appgroup /app

USER appuser

ENV PORT=5008
ENV CHROME_EXECUTABLE_PATH=/usr/bin/chromium

EXPOSE 5008

CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "5008", "--workers", "1", "--timeout-keep-alive", "60"]
