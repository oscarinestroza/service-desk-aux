"""Lógica del módulo Atención de Llamadas.

Obtiene los tickets del mes con recepción efectiva «Vía telefónica» (incluye
los corregidos a ese tipo desde Revisión de Tickets), cuyo servicio esté
activo y cuya descripción contractual no sea «Solicitud Adicional».
"""

from .models import Ticket
from .revision_tickets import formatear_diferencia, rango_mes  # noqa: F401


def tickets_telefonicos(mes):
    """Tickets del mes con recepción efectiva «Vía telefónica».

    Requiere que el servicio del ticket esté activo y que su descripción
    contractual no sea «Solicitud Adicional».
    """
    inicio, fin = rango_mes(mes)
    qs = (
        Ticket.objects.select_related("revision", "servicio")
        .filter(
            fecha__gte=inicio,
            fecha__lt=fin,
            servicio__activo=True,
        )
        .exclude(servicio__descripcion_contractual="solicitud_adicional")
        .order_by("-fecha")
    )
    return [
        t for t in qs if t.tipo_recepcion_mostrado == Ticket.TIPO_RECEPCION_TELEFONO
    ]
