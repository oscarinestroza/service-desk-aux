from .models import Institucion
from .roles import (
    CAP_ADMIN,
    CAP_ATENDER,
    CAP_DIRECTORIO,
    CAP_EDITAR,
    CAP_IMPORTAR,
    CAP_REVISIONES,
    CAP_TICKETS,
    VISTAS,
    modulos_visibles,
    rol_display,
    tiene,
)


def constantes_globales(request):
    """Expone constantes y permisos de la app a todas las plantillas."""
    modulos = modulos_visibles(request.user)
    return {
        "NIVEL_CHOICES": Institucion.NIVEL_CHOICES,
        "rol_actual": rol_display(request.user),
        # Permisos por capacidad (usables directamente en {% if %})
        "puede_directorio": tiene(request.user, CAP_DIRECTORIO),
        "puede_editar": tiene(request.user, CAP_EDITAR),
        "puede_importar": tiene(request.user, CAP_IMPORTAR),
        "puede_tickets": tiene(request.user, CAP_TICKETS),
        "puede_atender": tiene(request.user, CAP_ATENDER),
        "puede_revisiones": tiene(request.user, CAP_REVISIONES),
        "puede_admin": tiene(request.user, CAP_ADMIN) or request.user.is_staff,
        # Módulos del menú lateral visibles para el usuario
        "modulos_visibles": modulos,
        # Secciones (claves de capacidad) con al menos un módulo visible
        "secciones_visibles": {v["seccion"] for v in VISTAS if v["clave"] in modulos},
    }
