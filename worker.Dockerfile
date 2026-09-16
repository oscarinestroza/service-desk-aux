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

# Misma imagen para ambos workers. La cola y el Beat se eligen por env vars:
#   - worker de enlaces (no-programadas): sin config extra (default).
#   - worker de tickets (programadas):     env CELERY_QUEUES=programadas + CELERY_BEAT=1
# Las operaciones SIG (Playwright) deben correr serializadas: concurrency=1.
# OJO: --prefetch-multiplier con GUIÓN (la CLI de Celery no acepta guion bajo).
# --beat funciona solo en Linux (dentro del contenedor); en Windows local corre
# aparte: python -m celery -A enlacesMAO beat --loglevel=info
CMD ["sh", "-c", "exec celery -A enlacesMAO worker -Q ${CELERY_QUEUES:-no-programadas} --loglevel=info --concurrency=1 --pool=solo --prefetch-multiplier=1 ${CELERY_BEAT:+--beat}"]