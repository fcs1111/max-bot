FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice \
        fontconfig \
        fonts-dejavu \
        fonts-liberation \
        fonts-noto-core \
        fonts-crosextra-carlito \
        fonts-crosextra-caladea \
        fonts-freefont-ttf \
        fonts-noto \
        fonts-noto-cjk \
    && for pkg in \
        fonts-lato \
        fonts-open-sans \
        fonts-roboto \
        fonts-montserrat \
        fonts-pt-sans \
        fonts-pt-serif \
        fonts-cantarell \
        fonts-firacode \
        fonts-inter \
        fonts-ebgaramond; \
       do apt-get install -y --no-install-recommends "$pkg" || true; done \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
RUN mkdir -p /usr/local/share/fonts/custom \
    && if [ -d /app/fonts ]; then cp -r /app/fonts/* /usr/local/share/fonts/custom/ 2>/dev/null || true; fi \
    && mkdir -p /etc/fonts/conf.d \
    && cp /app/fontconfig/99-custom-font-fallbacks.conf /etc/fonts/conf.d/99-custom-font-fallbacks.conf 2>/dev/null || true \
    && fc-cache -f -v

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
