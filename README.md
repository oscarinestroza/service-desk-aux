# Enlaces MAO — Directorio de Enlaces del CCG Honduras

Sistema web (Django) para la gestión del directorio de enlaces autorizados del
Centro Cívico Gubernamental (CCG) de Honduras. Permite administrar edificios,
instituciones y enlaces, además de sincronizar usuarios con el sistema externo
SIG mediante automatización Playwright.

## Requisitos

- Python 3.14+ (probado con 3.14)
- Git
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

   Copia `.env.example` a `.env` y completa los valores.

   ```bash
   copy .env.example .env    # Windows
   cp .env.example .env      # Linux/macOS
   ```

   ⚠️ Genera un `SECRET_KEY` nuevo: `python -c "from django.core.management.utils import get_random_secret_key; print(get_random_secret_key())"`.

5. **Aplicar migraciones y crear superusuario**

   ```bash
   python manage.py migrate
   python manage.py createsuperuser
   ```

6. **Levantar el servidor**

   ```bash
   python manage.py runserver
   ```

   Visita `http://127.0.0.1:8000/`

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

## Estructura

```
enlacesMAO/          Configuración del proyecto Django
enlaces_ccg/         App principal (modelos, vistas, roles, admin)
enlaces_ccg/sig_sync/  Cliente Playwright y tareas de sincronización SIG
media/               Archivos subidos por los usuarios
```

## Notas de seguridad

- `db.sqlite3`, `.env`, `media/` y archivos de secretos están ignorados por git.
- No subas credenciales reales. Usa `.env.example` como plantilla.