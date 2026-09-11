# worker.Dockerfile — Worker Celery (Enlaces MAO)
# Misma imagen que el Dockerfile del web, pero con CMD de Celery.
# Incluye Chromium de Playwright porque el worker automatiza el SIG.
# En Coolify: Custom Dockerfile Location = worker.Dockerfile

FROM python:3.14-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Dependencias de Python primero (aprovecha la caché de capas de Docker)
COPY requirements.txt ./
RUN pip install --upgrade pip && pip install -r requirements.txt

# Chromium de Playwright + dependencias de sistema (--with-deps instala el resto)
RUN playwright install --with-deps chromium

# Código de la aplicación
COPY . .

# Las operaciones SIG (Playwright) deben correr serializadas: concurrency=1
CMD ["celery", "-A", "enlacesMAO", "worker", "--loglevel=info", "--concurrency=1"]