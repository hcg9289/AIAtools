FROM python:3.11-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .
COPY templates/ templates/

RUN mkdir -p uploads outputs && chmod 755 uploads outputs

ENV FLASK_APP=app.py
ENV FLASK_ENV=production

EXPOSE 5008

CMD ["python", "app.py"]
