"""
Tareas de sincronización en background para el SIG (Celery).

Las operaciones SIG (Playwright/Selenium) se encolan en Celery con Redis como
broker. Se configuran para correr serializadas (una a la vez): se recomienda
levantar el worker con --concurrency=1 y worker_prefetch_multiplier=1.

Semántica de usuario_sig: es el NOMBRE DE USUARIO en el SIG (lo que el
usuario escribe en el formulario). No es un ID numérico. El éxito de la
creación NO depende del ID: SyncLog.usuario_sig_id es solo informativo y
opcional (el SIG redirige al catálogo tras crear y la URL no conserva el ID).
"""

import logging

from celery import shared_task

logger = logging.getLogger("sig_sync")


def _enviar_correo_creacion(enlace):
    """Envía correo de notificación con credenciales del nuevo usuario SIG.

    - En PRUEBA (settings.DEBUG=True): el correo va SOLO a los correos
      configurados en ConfiguracionCorreo.correo_notificacion (nunca al
      correo_principal del enlace, porque ya se les notifica manualmente).
    - En PRODUCCIÓN (settings.DEBUG=False): se envía al correo_principal
      del enlace creado, añadiendo los correos configurados como adicionales.

    Se evitan destinatarios duplicados.
    """
    from django.conf import settings
    from django.core.mail import send_mail
    from enlaces_ccg.utils import aplicar_emisor_a_settings, obtener_config_correo

    extras = obtener_config_correo()["correo_notificacion"]

    if getattr(settings, "DEBUG", False):
        # Fase de prueba: solo a los correos configurados
        destinatarios = [c for c in extras if c]
    else:
        # Producción: al correo del usuario + adicionales configurados
        destinatarios = []
        if enlace.correo_principal:
            destinatarios.append(enlace.correo_principal)
        for extra in extras:
            if extra and extra not in destinatarios:
                destinatarios.append(extra)

    if not destinatarios:
        logger.warning("Sin destinatarios de notificación — saltando correo")
        return

    password = enlace.password_sig or getattr(settings, "SIG_DEFAULT_PASSWORD", "")

    asunto = f"Nuevo usuario SIG creado — {enlace.usuario_sig}"
    mensaje = (
        f"Se ha creado un nuevo usuario en el SIG para el enlace autorizado.\n\n"
        f"Datos del enlace:\n"
        f"  Nombre: {enlace.nombre_completo}\n"
        f"  Institución: {enlace.institucion.nombre if enlace.institucion else 'Sin especificar'}\n"
        f"  Estado: {enlace.estado}\n\n"
        f"Credenciales SIG:\n"
        f"  URL: https://sig.gia.mx/webapp/\n"
        f"  Usuario: {enlace.usuario_sig}\n"
        f"  Contraseña: {password}\n"
        f"  PIN: {enlace.pin_sig or '—'}\n\n"
        f"Este correo fue generado automáticamente por el sistema de Enlaces CCG."
    )

    try:
        aplicar_emisor_a_settings()
        send_mail(
            subject=asunto,
            message=mensaje,
            from_email=getattr(settings, "DEFAULT_FROM_EMAIL", None),
            recipient_list=destinatarios,
            fail_silently=False,
        )
        logger.info("Correo enviado a %s para enlace %s", destinatarios, enlace.pk)
    except Exception as e:
        logger.error("Error enviando correo para enlace %s: %s", enlace.pk, e)


def _enviar_correo_review(enlace, motivo):
    """Envía correo de revisión cuando falla la creación automática en el SIG,
    para que el destinatario haga la creación manualmente.

    Pueden especificarse varios correos separados por ';' en
    ConfiguracionCorreo.correo_review (o settings.SIG_REVIEW_EMAIL).
    """
    from django.conf import settings
    from django.core.mail import send_mail
    from enlaces_ccg.utils import aplicar_emisor_a_settings, obtener_config_correo

    destinatarios = obtener_config_correo()["correo_review"]
    if not destinatarios:
        logger.warning("Sin correos de revisión configurados — saltando correo de revisión")
        return

    password = enlace.password_sig or getattr(settings, "SIG_DEFAULT_PASSWORD", "")

    asunto = f"[REVISIÓN] Falló creación SIG — {enlace.usuario_sig}"
    mensaje = (
        f"La creación automática del usuario en el SIG falló y requiere "
        f"revisión para crearse manualmente.\n\n"
        f"Motivo del error:\n  {motivo}\n\n"
        f"Datos del enlace:\n"
        f"  Nombre: {enlace.nombre_completo}\n"
        f"  Institución: {enlace.institucion.nombre if enlace.institucion else 'Sin especificar'}\n"
        f"  Estado: {enlace.estado}\n\n"
        f"Credenciales SIG propuestas:\n"
        f"  URL: https://sig.gia.mx/webapp/\n"
        f"  Usuario: {enlace.usuario_sig}\n"
        f"  Contraseña: {password}\n"
        f"  PIN: {enlace.pin_sig or '—'}\n\n"
        f"Por favor cree el usuario manualmente en el SIG y luego marque el "
        f"enlace como sincronizado.\n"
        f"Este correo fue generado automáticamente por el sistema de Enlaces CCG."
    )

    try:
        aplicar_emisor_a_settings()
        send_mail(
            subject=asunto,
            message=mensaje,
            from_email=getattr(settings, "DEFAULT_FROM_EMAIL", None),
            recipient_list=destinatarios,
            fail_silently=False,
        )
        logger.info("Correo de revisión enviado a %s para enlace %s", destinatarios, enlace.pk)
    except Exception as e:
        logger.error("Error enviando correo de revisión para enlace %s: %s", enlace.pk, e)


def _ya_creado_en_sig(enlace_id):
    """True si ya existe un SyncLog exitoso de creación para el enlace."""
    from enlaces_ccg.models import SyncLog

    return SyncLog.objects.filter(
        enlace_id=enlace_id, accion="crear", estado="exitoso"
    ).exists()


# ------------------------------------------------------------------
# Ejecutores (corren dentro del worker thread)
# ------------------------------------------------------------------
@shared_task
def crear_enlace_sig(enlace_id):
    """Crea un usuario en el SIG para el enlace dado (tarea Celery)."""
    from django.conf import settings
    from enlaces_ccg.models import EnlaceAutorizado, SyncLog
    from .client import SIGClient, sig_usuario_configurado

    if not sig_usuario_configurado():
        logger.warning("Usuario SIG no configurado - saltando sync")
        return

    try:
        enlace = EnlaceAutorizado.objects.select_related(
            "institucion", "institucion__edificio"
        ).get(pk=enlace_id)
    except EnlaceAutorizado.DoesNotExist:
        logger.error("Enlace %s no existe", enlace_id)
        return

    if not enlace.usuario_sig:
        logger.info("Enlace %s sin usuario_sig, saltando crear en SIG", enlace_id)
        return

    if _ya_creado_en_sig(enlace_id):
        logger.info("Enlace %s ya fue creado en SIG, saltando", enlace_id)
        return

    log = SyncLog.objects.create(enlace=enlace, accion="crear", estado="pendiente")

    try:
        client = SIGClient()
        result = client.sincronizar(enlace, accion="crear")
        # logout() se hace dentro de sincronizar()

        if result["exitoso"]:
            log.estado = "exitoso"
            log.mensaje = result["mensaje"]
            log.usuario_sig_id = result.get("usuario_sig_id", "")
            log.save(update_fields=["estado", "mensaje", "usuario_sig_id"])
            logger.info("Enlace %s creado en SIG", enlace_id)
            _enviar_correo_creacion(enlace)
        else:
            log.estado = "error"
            log.mensaje = result.get("mensaje", "Error desconocido")
            log.save(update_fields=["estado", "mensaje"])
            _enviar_correo_review(enlace, log.mensaje)

    except Exception as e:
        log.estado = "error"
        log.mensaje = str(e)
        log.save(update_fields=["estado", "mensaje"])
        logger.exception("Error creando usuario SIG para enlace %s", enlace_id)
        _enviar_correo_review(enlace, str(e))


@shared_task
def deshabilitar_enlace_sig(enlace_id):
    """Deshabilita un usuario en el SIG (tarea Celery)."""
    from django.conf import settings
    from enlaces_ccg.models import EnlaceAutorizado, SyncLog
    from .client import SIGClient, sig_usuario_configurado

    if not sig_usuario_configurado():
        logger.warning("Usuario SIG no configurado - saltando sync")
        return

    try:
        enlace = EnlaceAutorizado.objects.get(pk=enlace_id)
    except EnlaceAutorizado.DoesNotExist:
        return

    if not enlace.usuario_sig:
        return

    log = SyncLog.objects.create(enlace=enlace, accion="deshabilitar", estado="pendiente")

    try:
        client = SIGClient()
        client.login()
        client.deshabilitar_usuario(enlace.usuario_sig)
        client.logout()

        log.estado = "exitoso"
        log.mensaje = f"Usuario {enlace.usuario_sig} deshabilitado"
        log.save(update_fields=["estado", "mensaje"])

    except Exception as e:
        log.estado = "error"
        log.mensaje = str(e)
        log.save(update_fields=["estado", "mensaje"])
        logger.exception("Error deshabilitando enlace %s", enlace_id)


@shared_task
def reactivar_enlace_sig(enlace_id):
    """Reactiva un usuario en el SIG (cambia estatus a Activo) (tarea Celery)."""
    from django.conf import settings
    from enlaces_ccg.models import EnlaceAutorizado, SyncLog
    from .client import SIGClient, sig_usuario_configurado

    if not sig_usuario_configurado():
        logger.warning("Usuario SIG no configurado - saltando sync")
        return

    try:
        enlace = EnlaceAutorizado.objects.get(pk=enlace_id)
    except EnlaceAutorizado.DoesNotExist:
        return

    if not enlace.usuario_sig:
        return

    log = SyncLog.objects.create(enlace=enlace, accion="reactivar", estado="pendiente")

    try:
        client = SIGClient()
        client.login()
        client.reactivar_usuario(enlace.usuario_sig)
        client.logout()

        log.estado = "exitoso"
        log.mensaje = f"Usuario {enlace.usuario_sig} reactivado"
        log.save(update_fields=["estado", "mensaje"])

    except Exception as e:
        log.estado = "error"
        log.mensaje = str(e)
        log.save(update_fields=["estado", "mensaje"])
        logger.exception("Error reactivando enlace %s", enlace_id)


@shared_task
def actualizar_correo_enlace_sig(enlace_id):
    """Actualiza el correo de un usuario en el SIG (tarea Celery)."""
    from django.conf import settings
    from enlaces_ccg.models import EnlaceAutorizado, SyncLog
    from .client import SIGClient, sig_usuario_configurado

    if not sig_usuario_configurado():
        logger.warning("Usuario SIG no configurado - saltando sync")
        return

    try:
        enlace = EnlaceAutorizado.objects.get(pk=enlace_id)
    except EnlaceAutorizado.DoesNotExist:
        return

    if not enlace.usuario_sig:
        return

    log = SyncLog.objects.create(
        enlace=enlace, accion="actualizar_correo", estado="pendiente"
    )

    try:
        client = SIGClient()
        client.login()
        client.cambiar_correo(enlace.usuario_sig, enlace.correo_principal or "")
        client.logout()

        log.estado = "exitoso"
        log.mensaje = f"Correo de {enlace.usuario_sig} actualizado a {enlace.correo_principal or '(vacío)'}"
        log.save(update_fields=["estado", "mensaje"])

    except Exception as e:
        log.estado = "error"
        log.mensaje = str(e)
        log.save(update_fields=["estado", "mensaje"])
        logger.exception("Error actualizando correo del enlace %s", enlace_id)


# ------------------------------------------------------------------
# API pública — encola tareas en Celery (.delay)
# Estas funciones mantienen los mismos nombres que usaban con la cola
# en memoria, de modo que signals.py no necesita cambios.
# ------------------------------------------------------------------
def sincronizar_enlace(enlace_id):
    """Encola la creación de usuario SIG para un enlace."""
    crear_enlace_sig.delay(enlace_id)
    logger.info("Enlace %s encolado para crear en SIG (Celery)", enlace_id)


def deshabilitar_enlace(enlace_id):
    """Encola la deshabilitación de usuario SIG para un enlace."""
    deshabilitar_enlace_sig.delay(enlace_id)
    logger.info("Enlace %s encolado para deshabilitar en SIG (Celery)", enlace_id)


def reactivar_enlace(enlace_id):
    """Encola la reactivación de usuario SIG para un enlace."""
    reactivar_enlace_sig.delay(enlace_id)
    logger.info("Enlace %s encolado para reactivar en SIG (Celery)", enlace_id)


def actualizar_correo_enlace(enlace_id):
    """Encola la actualización de correo de un usuario en el SIG."""
    actualizar_correo_enlace_sig.delay(enlace_id)
    logger.info("Enlace %s encolado para actualizar correo en SIG (Celery)", enlace_id)
