# Enlaces MAO — Directorio de Enlaces del CCG Honduras

Sistema web (Django) para la gestión del directorio de enlaces autorizados del
Centro Cívico Gubernamental (CCG) de Honduras. Permite administrar edificios,
instituciones y enlaces, además de sincronizar usuarios con el sistema externo
SIG mediante automatización Playwright.

## Requisitos

- Python 3.14+ (probado con 3.14)
- Git
- Docker + Docker Compose (para PostgreSQL y Redis; las tareas en segundo
  plano requieren Redis/Celery)
- Playwright (con navegador Chromium instalado)

## Instalación

1. **Clonar el repositorio**

   ```bash
   git clone https://github.com/TU-USUARIO/enlacesMAO.git
   cd enlacesMAO
   ```

2. **Crear y activar entorno virtual**

   ```bash
   python -m venv .venv
   .\.venv\Scripts\activate    # Windows
   source .venv/bin/activate   # Linux/macOS
   ```

3. **Instalar dependencias**

   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```

4. **Configurar variables de entorno**

   Copia `.env.example` a `.env` y completa los valores:

   ```bash
   copy .env.example .env    # Windows
   cp .env.example .env      # Linux/macOS
   ```

   Por defecto `.env.example` ya apunta a Postgres y Redis locales (ver paso 5).
   ⚠️ Genera un `SECRET_KEY` nuevo: `python -c "from django.core.management.utils import get_random_secret_key; print(get_random_secret_key())"`.

5. **Levantar PostgreSQL y Redis (Docker Compose)**

   ```bash
   docker compose up -d db redis
   ```

   Esto crea los contenedores `enlacesmao_db` (postgres:16) y `enlacesmao_redis`
   (redis:7) con volúmenes persistentes. Para detener: `docker compose down`.

6. **Aplicar migraciones y crear superusuario**

   ```bash
   python manage.py migrate
   python manage.py createsuperuser
   ```

7. **Levantar el worker de Celery (tareas SIG en segundo plano)**

   En una terminal aparte, activa el venv y ejecuta:

   ```bash
   celery -A enlacesMAO worker --loglevel=info --concurrency=1 --pool=solo
   ```

   `--concurrency=1` + `--pool=solo` garantiza que las operaciones SIG
   (Playwright) corran serializadas, una a la vez. Es obligatorio en Windows
   (evita `PermissionError: [WinError 5]` con el pool prefork por defecto).

8. **Levantar el servidor**

   ```bash
   python manage.py runserver
   ```

   Visita `http://127.0.0.1:8000/`

## Migrar datos de SQLite a PostgreSQL

Si ya tenías datos en SQLite y quieres llevarlos a Postgres (con `.env` apuntando
a SQLite primero y luego a Postgres):

```bash
# 1) Con SQLite como default, exporta los datos
python manage.py dumpdata --natural-foreign --natural-primary -o datos.json
python manage.py dumpdata --natural-foreign --natural-primary --exclude contenttypes --exclude auth.permission -o datos_sin_ct.json

# 2) Cambia .env a Postgres (ver .env.example) y aplica migraciones
python manage.py migrate

# 3) Carga los datos
python manage.py loaddata datos_sin_ct.json
```

Se recomienda excluir `contenttypes` y `auth.permission` para evitar
conflictos de tipos de contenido.

## Roles y permisos

El sistema usa grupos de Django con capacidades administrables:

| Rol | Capacidades |
|-----|-------------|
| **Administrador** | directorio, editar, importar, tickets, revisiones, admin |
| **Operador MAO** | directorio, editar, tickets, revisiones |
| **Encargado de Servicio** | directorio, tickets |

Los permisos de cada grupo se gestionan desde la app (sesión **Grupos y permisos**)
o desde el admin (`/admin/enlaces_ccg/permisogrupo/`).

## Configuración desde el panel de admin

- **Configuración SIG**: credenciales del login automatizado (Playwright) para
  sincronizar con el sistema externo SIG. Sin estos datos la sincronización se omite.
- **Configuración de correo**: cuenta SMTP y correos de revisión/notificación.
- **Importación Excel**: los enlaces importados **no sincronizan automáticamente** con el SIG.
  Se usa un flag interno (`_sig_sync_desactivado`) para que el signal `post_save` lo ignore,
  por lo que los datos quedan solo en la BD. La sincronización SIG solo se dispara al crear
  o editar un enlace manualmente desde el formulario web.

## Estructura

```
enlacesMAO/           Configuración del proyecto Django (incluye celery.py)
enlaces_ccg/          App principal (modelos, vistas, roles, admin)
enlaces_ccg/sig_sync/   Cliente Playwright y tareas Celery de sincronización SIG
media/                Archivos subidos por los usuarios
docker-compose.yml    Infraestructura local (PostgreSQL + Redis)
```

## Notas de seguridad

- `db.sqlite3`, `.env`, `media/` y archivos de secretos están ignorados por git.
- No subas credenciales reales. Usa `.env.example` como plantilla.
- `docker-compose.yml` usa credenciales de desarrollo; para producción usa
  secretos gestionados por tu proveedor (p. ej. Coolify).