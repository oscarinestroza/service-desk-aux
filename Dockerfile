# Dockerfile — Enlaces MAO (Django)
# Una misma imagen sirve para el contenedor web (gunicorn) y el worker (Celery).
# Incluye Chromium de Playwright porque el worker automatiza el SIG.

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

EXPOSE 7070

# Web por defecto: entrypoint corre migraciones y collectstatic y luego gunicorn.
# En Coolify el worker reemplaza CMD con: celery -A enlacesMAO worker ...
CMD ["sh", "/app/entrypoint.sh", "gunicorn", "enlacesMAO.wsgi:application", \
     "--bind", "0.0.0.0:7070", "--workers", "2", "--timeout", "120", "--access-logfile", "-"]