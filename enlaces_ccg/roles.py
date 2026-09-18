"""
Roles de usuario del sistema.

Se usan los Groups de django.contrib.auth (permisos por grupo).
Cada grupo representa un rol. Las capacidades por rol se definen aquí.

Grupos:
    - Administrador        → acceso total (directorio, editar, tickets, revisiones, admin)
    - Operador MAO         → directorio (ver/exportar), tickets, revisiones
    - Encargado de Servicio→ directorio (ver/exportar), revisiones

Capacidades (permisos lógicos):
    - directorio : ver listados de edificios/instituciones/enlaces + exportar
    - editar     : crear/editar enlaces e instituciones del directorio
    - tickets    : sección Seguimiento de tickets
    - revisiones : sección Modulo Operativo MAO
    - admin      : panel de administración de Django
"""

from functools import wraps

from django.contrib.auth.views import redirect_to_login
from django.shortcuts import render

# ---------------------------------------------------------------------------
# Definición de roles
# ---------------------------------------------------------------------------
ROL_ADMIN = "Administrador"
ROL_OPERADOR = "Operador MAO"
ROL_ENCARGADO = "Encargado de Servicio"

ROLES = (ROL_ADMIN, ROL_OPERADOR, ROL_ENCARGADO)

# Capacidades (permisos lógicos)
CAP_DIRECTORIO = "directorio"
CAP_EDITAR = "editar"
CAP_IMPORTAR = "importar"
CAP_TICKETS = "tickets"
CAP_ATENDER = "atender"
CAP_REVISIONES = "revisiones"
CAP_ADMIN = "admin"

CAPACIDADES_POR_ROL = {
    ROL_ADMIN: {CAP_DIRECTORIO, CAP_EDITAR, CAP_IMPORTAR, CAP_TICKETS, CAP_ATENDER, CAP_REVISIONES, CAP_ADMIN},
    ROL_OPERADOR: {CAP_DIRECTORIO, CAP_EDITAR, CAP_TICKETS, CAP_REVISIONES},
    ROL_ENCARGADO: {CAP_DIRECTORIO, CAP_TICKETS, CAP_ATENDER},
}

# Roles con acceso de staff (is_staff) — p. ej. que ven el panel admin
STAFF_ROLES = {ROL_ADMIN}


def rol_de(user):
    """Devuelve el nombre del rol del usuario (o None si no tiene/n o no autenticado)."""
    if not user or not user.is_authenticated:
        return None
    if user.is_superuser:
        return ROL_ADMIN
    # Preferir un rol de staff si existe
    for g in user.groups.all():
        if g.name in ROLES:
            return g.name
    return None


def rol_display(user):
    """Etiqueta legible del rol para la interfaz."""
    r = rol_de(user)
    return r or ("Autorizado" if user and user.is_authenticated else None)


def es_rol(user, *roles):
    return rol_de(user) in roles


def capacidades_de(grupo):
    """
    Devuelve el conjunto de capacidades de un grupo, leyendo la BD
    (`PermisoGrupo`) con respaldo en la definición estática.
    """
    nombre = getattr(grupo, "name", None)
    try:
        from .models import PermisoGrupo

        permiso = PermisoGrupo.objects.filter(grupo=grupo).first()
    except Exception:
        permiso = None
    if permiso is not None:
        return set(permiso.capacidades or [])
    return set(CAPACIDADES_POR_ROL.get(nombre, set()))


def tiene(user, capacidad):
    """¿El usuario tiene una capacidad? (los superusuarios siempre sí)."""
    if not user or not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    for g in user.groups.all():
        if capacidad in capacidades_de(g):
            return True
    return False


def requiere(rol_or_capacidad):
    """
    Decorador para vistas. Acepta el nombre de un rol o una capacidad.

    - No autenticado → redirige al login.
    - Autenticado sin permiso → página 403 descriptiva.
    """
    def decorador(vista):
        @wraps(vista)
        def _wrapper(request, *args, **kwargs):
            user = request.user
            if not user.is_authenticated:
                return redirect_to_login(request.get_full_path())
            if rol_or_capacidad in ROLES:
                tiene_permiso = es_rol(user, rol_or_capacidad)
            else:
                tiene_permiso = tiene(user, rol_or_capacidad)
            if not tiene_permiso:
                return render(
                    request,
                    "enlaces_ccg/403.html",
                    {"rol": rol_display(user)},
                    status=403,
                )
            return vista(request, *args, **kwargs)
        return _wrapper
    return decorador


# ---------------------------------------------------------------------------
# Catálogo de vistas del sistema (para la UI de permisos)
#
# Cada vista indica qué capacidades necesita para "ver" y para "editar".
# La administración de permisos se hace por sección/capacidad (no por vista).
# 'ver' y 'editar' disponibles en cada una para que la matriz sea clara.
# ---------------------------------------------------------------------------
VISTAS = [
    {"clave": "edificios", "etiqueta": "Edificios", "seccion": CAP_DIRECTORIO,
     "ver": (CAP_DIRECTORIO,), "editar": (CAP_EDITAR,)},
    {"clave": "instituciones", "etiqueta": "Instituciones", "seccion": CAP_DIRECTORIO,
     "ver": (CAP_DIRECTORIO,), "editar": (CAP_EDITAR,)},
    {"clave": "enlaces", "etiqueta": "Enlaces", "seccion": CAP_DIRECTORIO,
     "ver": (CAP_DIRECTORIO,), "editar": (CAP_EDITAR,)},
    {"clave": "tickets", "etiqueta": "Seguimiento de tickets", "seccion": CAP_TICKETS,
     "ver": (CAP_TICKETS,)},
    {"clave": "dashboard", "etiqueta": "Dashboard", "seccion": CAP_TICKETS,
     "ver": (CAP_TICKETS,)},
    {"clave": "indicadores", "etiqueta": "Indicadores de mejora", "seccion": CAP_TICKETS,
     "ver": (CAP_TICKETS,)},
    {"clave": "importar_deductivas", "etiqueta": "Importar deductivas", "seccion": CAP_IMPORTAR,
     "ver": (CAP_IMPORTAR,)},
    {"clave": "importar_enlaces", "etiqueta": "Importar enlaces", "seccion": CAP_IMPORTAR,
     "ver": (CAP_IMPORTAR,)},
    {"clave": "importar_instituciones", "etiqueta": "Importar instituciones", "seccion": CAP_IMPORTAR,
     "ver": (CAP_IMPORTAR,)},
    {"clave": "revision_tickets", "etiqueta": "Revisión de Tickets", "seccion": CAP_REVISIONES,
     "ver": (CAP_REVISIONES,)},
    {"clave": "revision_correos", "etiqueta": "Revisión de Correos", "seccion": CAP_REVISIONES,
     "ver": (CAP_REVISIONES,)},
    {"clave": "cierre_operador", "etiqueta": "Cierre de Operador", "seccion": CAP_REVISIONES,
     "ver": (CAP_REVISIONES,)},
    {"clave": "comentarios_operador", "etiqueta": "Comentarios de Operador", "seccion": CAP_REVISIONES,
     "ver": (CAP_REVISIONES,)},
    {"clave": "recordatorios", "etiqueta": "Recordatorios", "seccion": CAP_REVISIONES,
     "ver": (CAP_REVISIONES,)},
    {"clave": "plantillas_respuesta", "etiqueta": "Plantillas de respuesta", "seccion": CAP_REVISIONES,
     "ver": (CAP_REVISIONES,)},
    {"clave": "documentos", "etiqueta": "Documentos y carpetas", "seccion": CAP_DIRECTORIO,
     "ver": (CAP_DIRECTORIO,)},
]


SECCIONES = [
    {"clave": CAP_DIRECTORIO, "etiqueta": "Directorio de Enlaces"},
    {"clave": CAP_TICKETS, "etiqueta": "Seguimiento de tickets"},
    {"clave": CAP_IMPORTAR, "etiqueta": "Importaciones"},
    {"clave": CAP_REVISIONES, "etiqueta": "Modulo Operativo MAO"},
]


def nombres_capacidades():
    """Etiquetas humanas de cada capacidad (para la UI)."""
    return {
        CAP_DIRECTORIO: "Directorio de Enlaces",
        CAP_EDITAR: "Editar directorio",
        CAP_TICKETS: "Seguimiento de tickets",
        CAP_ATENDER: "Atender tickets",
        CAP_IMPORTAR: "Importaciones",
        CAP_REVISIONES: "Modulo Operativo MAO",
        CAP_ADMIN: "Panel de administración",
    }


def nombre_capacidad(cap):
    """Etiqueta legible de una clave de capacidad."""
    return nombres_capacidades().get(cap, cap)
