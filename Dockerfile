FROM python:3.12-slim

RUN apt-get update && apt-get install -y \
    wget curl gnupg ca-certificates \
    libglib2.0-0 libnss3 libnspr4 libatk1.0-0 \
    libatk-bridge2.0-0 libcups2 libdrm2 libdbus-1-3 \
    libxkbcommon0 libx11-6 libxcomposite1 libxdamage1 \
    libxext6 libxfixes3 libxrandr2 libgbm1 libpango-1.0-0 \
    libcairo2 libasound2 libxshmfence1 fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# Aumentar límite de file descriptors para Chromium
RUN echo "* soft nofile 65536" >> /etc/security/limits.conf && \
    echo "* hard nofile 65536" >> /etc/security/limits.conf

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN playwright install chromium
RUN playwright install-deps chromium

COPY . .

RUN mkdir -p /app/uploads /app/output

ENV HEADLESS=true
ENV VERCEL_URL=*

EXPOSE 8000

# Usar shell para poder setear ulimit antes de arrancar uvicorn
CMD ["/bin/sh", "-c", "ulimit -n 65536 && uvicorn main:app --host 0.0.0.0 --port 8000"]
