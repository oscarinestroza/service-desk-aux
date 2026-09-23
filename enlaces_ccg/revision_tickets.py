"""Lógica del módulo Revisión de Tickets.

Detecta tickets del mes que incumplen los criterios de recepción y reconcilia
los que el operador marcó como "Validación SIG" cuando el SIG ya los actualizó.
"""

from datetime import timedelta

from django.utils import timezone

from .models import (
    RevisionTicket,
    Ticket,
    es_anomalia_recepcion,
)


def diagnosticos_recepcion(ticket):
    """Mensajes de posible diagnóstico según los criterios de recepción."""
    mensajes = []
    tipo = ticket.tipo_recepcion_mostrado
    recepcion = ticket.fecha_recepcion_mostrada
    apertura = ticket.fecha

    if tipo == Ticket.TIPO_RECEPCION_AUTOMATICA:
        mensajes.append(
            "El ticket se cargó sin indicar el medio por el que fue recibido "
            "ni la hora de recepción."
        )
    if recepcion is not None and apertura is not None:
        diferencia = apertura - recepcion
        if diferencia < timedelta(0):
            mensajes.append(
                "Se debe revisar la hora de recepción del ticket ya que indica "
                "que es posterior a la apertura del mismo."
            )
        elif (
            tipo == Ticket.TIPO_RECEPCION_CORREO
            and diferencia > timedelta(hours=2)
        ) or (
            tipo == Ticket.TIPO_RECEPCION_TELEFONO
            and diferencia > timedelta(minutes=5)
        ):
            mensajes.append(
                "El tiempo de apertura del ticket superó el rango establecido "
                "de 5 minutos para llamadas y 2 horas para correos."
            )
    return mensajes


def recepcion_resuelta_vals(tipo, apertura, recepcion):
    """True si (tipo, apertura, recepción) cumplen los criterios de recepción.

    Requiere las 3 condiciones:
      1. El tipo de recepción no es «Vía automática» (ni vacío).
      2. Tiene hora de recepción.
      3. La diferencia apertura − recepción no supera el umbral.
    """
    if tipo not in (Ticket.TIPO_RECEPCION_CORREO, Ticket.TIPO_RECEPCION_TELEFONO):
        return False
    if recepcion is None or apertura is None:
        return False
    return not es_anomalia_recepcion(tipo, apertura, recepcion)


def recepcion_resuelta(ticket):
    """True si el ticket ya cumple los criterios de recepción."""
    return recepcion_resuelta_vals(
        ticket.tipo_recepcion_mostrado,
        ticket.fecha,
        ticket.fecha_recepcion_mostrada,
    )


def puede_editar_revision(user, revision):
    """¿El usuario puede editar esta revisión?

    Las revisiones en estado «Corregido» solo las puede editar el usuario que
    hizo la corrección, un superusuario o alguien del grupo Administrador.
    """
    from .roles import ROL_ADMIN

    if revision is None or revision.estado != RevisionTicket.ESTADO_CORREGIDO:
        return True
    if not user or not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    if revision.corregido_por_id == user.pk:
        return True
    return user.groups.filter(name=ROL_ADMIN).exists()


def rango_mes(mes):
    """Devuelve (inicio, fin) aware del mes local: 'actual' o 'pasado'."""
    ahora = timezone.localtime()
    if mes == "pasado":
        inicio = ahora.replace(
            day=1, hour=0, minute=0, second=0, microsecond=0
        )
        if inicio.month == 1:
            inicio = inicio.replace(year=inicio.year - 1, month=12)
        else:
            inicio = inicio.replace(month=inicio.month - 1)
        if inicio.month == 12:
            fin = inicio.replace(year=inicio.year + 1, month=1)
        else:
            fin = inicio.replace(month=inicio.month + 1)
        return inicio, fin
    inicio = ahora.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    if inicio.month == 12:
        fin = inicio.replace(year=inicio.year + 1, month=1)
    else:
        fin = inicio.replace(month=inicio.month + 1)
    return inicio, fin


def tickets_del_mes(mes):
    """Tickets cuya apertura (`fecha`) cae en el mes indicado."""
    inicio, fin = rango_mes(mes)
    return (
        Ticket.objects.select_related("revision")
        .filter(fecha__gte=inicio, fecha__lt=fin)
        .order_by("-fecha")
    )


def tickets_detectados(mes):
    """Tickets del mes que incumplen los criterios de recepción."""
    return [t for t in tickets_del_mes(mes) if t.anomalia_recepcion()]


def reconciliar(mes):
    """Marca como 'corregido' los 'validacion_sig' cuya apertura ya cambió en el SIG.

    Se compara la apertura actual del ticket con la que tenía al marcar la
    validación (`apertura_antes`). Mientras no cambie, el ticket permanece en
    «Validación SIG» para que el operador vea que sigue pendiente.
    """
    inicio, fin = rango_mes(mes)
    pendientes = (
        RevisionTicket.objects.filter(
            estado=RevisionTicket.ESTADO_VALIDACION_SIG,
            ticket__fecha__gte=inicio,
            ticket__fecha__lt=fin,
        ).select_related("ticket")
    )
    corregidos = 0
    for revision in pendientes:
        ticket = revision.ticket
        if revision.apertura_antes is None:
            # Registros previos sin marca: se inicializa para no cerrarlos solos.
            revision.apertura_antes = ticket.fecha
            revision.save(update_fields=["apertura_antes", "actualizado_en"])
            continue
        if ticket.fecha != revision.apertura_antes and recepcion_resuelta(ticket):
            revision.estado = RevisionTicket.ESTADO_CORREGIDO
            revision.corregido_en = timezone.now()
            revision.save(
                update_fields=["estado", "corregido_en", "actualizado_en"]
            )
            corregidos += 1
    return corregidos


def formatear_diferencia(delta):
    """Formatea un timedelta como HH:MM (horas totales)."""
    if delta is None:
        return "—"
    total_min = int(delta.total_seconds() // 60)
    if total_min < 0:
        return "—"
    horas, minutos = divmod(total_min, 60)
    return f"{horas:02d}:{minutos:02d}"
