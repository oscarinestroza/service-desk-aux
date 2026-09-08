#!/bin/sh
set -e

# Entrada del contenedor web: aplica migraciones, recolecta estáticos y
# ejecuta el comando recibido (gunicorn). El worker NO usa este script:
# Coolify reemplaza CMD con el comando de celery directamente.
python manage.py migrate --noinput
python manage.py collectstatic --noinput

exec "$@"