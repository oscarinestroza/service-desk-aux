from .models import Institucion
from .roles import (
    CAP_ADMIN,
    CAP_ATENDER,
    CAP_DIRECTORIO,
    CAP_EDITAR,
    CAP_IMPORTAR,
    CAP_REVISIONES,
    CAP_TICKETS,
    rol_display,
    tiene,
)


def constantes_globales(request):
    """Expone constantes y permisos de la app a todas las plantillas."""
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
    }
