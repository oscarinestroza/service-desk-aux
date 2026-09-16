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
import mimetypes

from celery import shared_task
from django.utils.html import escape

logger = logging.getLogger("sig_sync")


def _plantilla_email_html(titulo, header_bg, intro, secciones, footer_note=None):
    """Devuelve un correo HTML autocontenido (inline styles para Outlook).

    `secciones` es una lista de dicts:
        - {"titulo": str, "filas": [("clave", "valor"), ...]}  → tabla de datos
        - {"texto": str}                                         → párrafo simple
    Los valores deben llegar ya escapados con `escape()`.
    """
    bloques = []
    for sec in secciones:
        if "texto" in sec:
            bloques.append(f'<p style="margin:0 0 16px 0;">{sec["texto"]}</p>')
        elif "filas" in sec:
            filas_html = "".join(
                f"""
                <tr>
                  <td width="140" style="padding:8px 16px;color:#71717a;font-size:14px;vertical-align:top;">{k}</td>
                  <td style="padding:8px 16px;color:#18181b;font-size:14px;vertical-align:top;word-break:break-all;">{v}</td>
                </tr>"""
                for k, v in sec["filas"]
            )
            bloques.append(
                f"""
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:0 0 22px 0;background-color:#fafafa;border:1px solid #e4e4e7;border-radius:6px;">
                  <tr>
                    <td colspan="2" style="background-color:#ffffff;border-bottom:1px solid #e4e4e7;padding:8px 16px;font-size:12px;font-weight:bold;letter-spacing:.05em;color:#71717a;text-transform:uppercase;">{sec["titulo"]}</td>
                  </tr>
                  {filas_html}
                </table>"""
            )

    footer = footer_note or (
        "Este correo fue generado automáticamente por el sistema de Enlaces CCG. "
        "Por favor no responda a este mensaje."
    )

    return f"""
<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background-color:#f4f4f5;font-family:Arial,Helvetica,sans-serif;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f4f4f5;padding:24px 0;">
  <tr>
    <td align="center">
      <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff;border-radius:8px;overflow:hidden;border:1px solid #e4e4e7;">
        <tr>
          <td style="padding:26px 32px;background-color:{header_bg};">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
              <tr>
                <td><span style="color:#ffffff;font-size:18px;font-weight:bold;">{escape(titulo)}</span></td>
                <td align="right"><span style="color:#ffffff;font-size:13px;opacity:.9;">Enlaces CCG</span></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="padding:28px 32px;color:#18181b;font-size:15px;line-height:1.6;">
            <p style="margin:0 0 16px 0;">{intro}</p>
            {''.join(bloques)}
            <p style="margin:0;color:#71717a;font-size:12px;border-top:1px solid #e4e4e7;padding-top:16px;">
              {escape(footer)}
            </p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>"""


def _adjuntos_carpeta_nuevos_enlaces():
    """Devuelve lista de (nombre_archivo, contenido bytes, mimetype) para los
    documentos de la carpeta 'Adjuntos: Nuevos Enlaces MAO'."""
    from enlaces_ccg.models import DocumentoCarpeta

    adjuntos = []
    carpeta = DocumentoCarpeta.objects.filter(
        slug="adjuntos-nuevos-enlaces"
    ).first()
    if not carpeta:
        return adjuntos
    for doc in carpeta.documentos.all():
        nombre = doc.nombre_archivo
        try:
            with doc.archivo.open("rb") as f:
                contenido = f.read()
        except Exception as e:
            logger.error("No se pudo leer adjunto %s: %s", doc.pk, e)
            continue
        mimetype, _ = mimetypes.guess_type(nombre)
        adjuntos.append((nombre, contenido, mimetype or "application/octet-stream"))
    return adjuntos


def _plantilla_correo_mao(
    cuerpo_dinamico,
    titulo_encabezado="NOTIFICACIÓN DE SERVICIO",
    introduccion=(
        "Buen día Estimado Usuario,<br><br>"
        "Por este medio compartimos la información indicada por el servicio:"
    ),
    despedida="Saludos cordiales,",
    mostrar_privacidad=False,
):
    """Devuelve un correo HTML con el formato corporativo MAO.

    Réplica en Python de GenerarPlantillaCorreo (VBA) usado en la apertura de
    tickets: tarjeta principal al 100% de ancho, encabezado azul corporativo,
    texto de introducción, cuerpo dinámico en caja gris clara con acento azul,
    firma SOPORTE MAO y aviso de privacidad opcional.

    `cuerpo_dinamico` debe llegar ya construido/escapado por el llamador
    (misma convención que `_plantilla_email_html`).
    """
    html = (
        "<table border='0' cellpadding='0' cellspacing='0' width='100%' "
        "style='width: 100%; font-family: Calibri, Arial, sans-serif; border: "
        "1px solid #dcdcdc; border-radius: 6px; overflow: hidden; "
        "background-color: #ffffff; text-align: left;'>"
        # Encabezado azul
        "<tr><td style='background-color: #1A365D; color: #ffffff; padding: "
        "18px 24px; text-align: left;'>"
        "<h2 style='margin: 0; font-family: Calibri, Arial, sans-serif; "
        "font-size: 16pt; font-weight: bold; letter-spacing: 0.5px; "
        "color: #ffffff; text-transform: uppercase;'>"
        f"{escape(titulo_encabezado)}</h2></td></tr>"
        # Texto estático de introducción (Calibri 11pt)
        "<tr><td style='padding: 20px 24px 10px 24px; color: #333333; "
        "font-family: Calibri, Arial, sans-serif; font-size: 11pt; "
        "line-height: 1.5;'>"
        f"<font face='Calibri' size='3' style='font-size: 11pt; color: "
        f"#333333;'>{introduccion}</font></td></tr>"
        # Cuerpo dinámico (caja gris clara con borde izquierdo azul)
        "<tr><td style='padding: 10px 24px 20px 24px;'>"
        "<table border='0' cellpadding='0' cellspacing='0' width='100%' "
        "style='width: 100%; background-color: #F7FAFC; border-left: 4px "
        "solid #2B6CB0; border-radius: 4px;'>"
        "<tr><td style='padding: 16px; color: #2D3748; font-family: Calibri, "
        "Arial, sans-serif; font-size: 11pt; line-height: 1.5;'>"
        f"<font face='Calibri' size='3' style='font-size: 11pt; color: "
        f"#2D3748;'>{cuerpo_dinamico}</font></td></tr></table></td></tr>"
        # Pie de página y firma (Calibri 11pt)
        "<tr><td style='padding: 10px 24px 20px 24px; color: #4A5568; "
        "font-family: Calibri, Arial, sans-serif; font-size: 11pt; "
        "line-height: 1.5; border-top: 1px solid #EDF2F7;'>"
        f"<font face='Calibri' size='3' style='font-size: 11pt; color: "
        f"#4A5568;'>{despedida}<br><strong style='color: #1A365D; "
        f"font-size: 11pt;'>SOPORTE MAO</strong></font></td></tr>"
    )

    # Aviso de privacidad / disclaimer (Calibri 9pt)
    if mostrar_privacidad:
        html += (
            "<tr><td style='padding: 15px 24px; background-color: #F8FAFC; "
            "border-top: 1px solid #E2E8F0; color: #718096; font-family: "
            "Calibri, Arial, sans-serif; font-size: 9pt; line-height: 1.4; "
            "text-align: justify;'>"
            "<font face='Calibri' size='1' style='font-size: 9pt; color: "
            "#718096;'><strong>Aviso de Confidencialidad:</strong> Este mensaje "
            "y sus anexos contienen información confidencial destinada "
            "exclusivamente a su destinatario. Si usted ha recibido este correo "
            "por error, se le notifica que cualquier revisión, divulgación, "
            "copia o distribución del mismo está estrictamente prohibida. Por "
            "favor, notifique inmediatamente al remitente y elimine este mensaje "
            "de su sistema.</font></td></tr>"
        )

    html += "</table>"

    return (
        "<!DOCTYPE html>\n"
        '<html lang="es">\n'
        '<head><meta charset="utf-8"></head>\n'
        '<body style="margin:0;padding:0;background-color:#f4f4f5;'
        'font-family:Calibri,Arial,sans-serif;">\n'
        '<table role="presentation" width="100%" cellpadding="0" '
        'cellspacing="0" border="0" style="background-color:#f4f4f5;'
        'padding:24px 0;">\n'
        "<tr><td align='left'>\n"
        f"{html}\n"
        "</td></tr>\n"
        "</table>\n"
        "</body>\n"
        "</html>"
    )


def _enviar_correo_creacion(enlace):
    """Envía correo de notificación con credenciales del nuevo usuario SIG.

    - Si settings.MAIL_NOTIFICAR_USUARIO es false: el correo va SOLO a los
      correos configurados en ConfiguracionCorreo.correo_notificacion (nunca
      al correo_principal del enlace). Útil durante pruebas sin depender de
      DEBUG (así no se exponen secretos en páginas de error).
    - Si es true (default): se envía al correo_principal del enlace creado,
      añadiendo los correos configurados como adicionales.

    Se evitan destinatarios duplicados.
    """
    from django.conf import settings
    from enlaces_ccg.utils import aplicar_emisor_a_settings, obtener_config_correo

    extras = obtener_config_correo()["correo_notificacion"]

    if not getattr(settings, "MAIL_NOTIFICAR_USUARIO", True):
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

    from ..models import ConfiguracionSIG
    cfg_sig = ConfiguracionSIG.cargar()
    password = enlace.password_sig or cfg_sig.default_password_valor()

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

    url_fila = (
        "<strong>URL:</strong> "
        '<a href="https://sig.gia.mx/webapp/" style="color:#2B6CB0;">'
        "https://sig.gia.mx/webapp/</a><br>"
        "<strong>Empresa:</strong> Centro Cívico Gubernamental<br><br>"
    )
    filas = [
        ("Usuario", escape(enlace.usuario_sig)),
        ("Contraseña", f"<strong>{escape(password)}</strong>"),
    ]
    if enlace.pin_sig and enlace.pin_sig != "1":
        filas.append(("PIN", escape(enlace.pin_sig)))

    cuerpo_dinamico = url_fila + "<br>".join(
        f"<strong>{escape(k)}:</strong> {v}" for k, v in filas
    )

    mensaje_html = _plantilla_correo_mao(
        cuerpo_dinamico=cuerpo_dinamico,
        titulo_encabezado="NUEVO USUARIO SIG CREADO",
        introduccion=(
            f"Buen día <strong>{escape(enlace.nombre_completo)}</strong>,<br><br>"
            f"Por este medio compartimos sus credenciales para seguimiento de "
            f"solicitudes en la MAO (Plataforma SIG):"
        ),
        despedida=(
            "Favor su apoyo confirmando si pudo acceder correctamente.<br><br>"
            "De igual manera, le compartimos la presentación utilizada en la "
            "última capacitación para su referencia.<br>"
            "Cualquier duda o consulta estamos a la orden.<br><br>"
            "Saludos cordiales,"
        ),
    )

    try:
        aplicar_emisor_a_settings()
        from django.core.mail import EmailMultiAlternatives

        email = EmailMultiAlternatives(
            subject=asunto,
            body=mensaje,
            from_email=getattr(settings, "DEFAULT_FROM_EMAIL", None),
            to=destinatarios,
        )
        email.attach_alternative(mensaje_html, "text/html")
        for nombre, contenido, mimetype in _adjuntos_carpeta_nuevos_enlaces():
            email.attach(nombre, contenido, mimetype)
        email.send(fail_silently=False)
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

    from ..models import ConfiguracionSIG
    cfg_sig = ConfiguracionSIG.cargar()
    password = enlace.password_sig or cfg_sig.default_password_valor()

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
        mensaje_html = _plantilla_email_html(
            titulo="[REVISIÓN] Falló la creación SIG",
            header_bg="#dc2626",
            intro=(
                "La creación automática del usuario en el SIG falló y requiere "
                "revisión para crearse manualmente."
            ),
            secciones=[
                {
                    "titulo": "Motivo del error",
                    "filas": [("Detalle", escape(motivo))],
                },
                {
                    "titulo": "Datos del enlace",
                    "filas": [
                        ("Nombre", escape(enlace.nombre_completo)),
                        ("Institución", escape(enlace.institucion.nombre if enlace.institucion else "Sin especificar")),
                        ("Estado", escape(f"{enlace.estado}")),
                    ],
                },
                {
                    "titulo": "Credenciales SIG propuestas",
                    "filas": [
                        ("URL", '<a href="https://sig.gia.mx/webapp/" style="color:#dc2626;">https://sig.gia.mx/webapp/</a>'),
                        ("Usuario", escape(enlace.usuario_sig)),
("Contraseña", escape(password)),
                        ("PIN", escape(enlace.pin_sig or "—")),
                    ],
                },
                {
                    "texto": "Por favor cree el usuario manualmente en el SIG y "
                             "luego marque el enlace como sincronizado.",
                },
            ],
        )
        send_mail(
            subject=asunto,
            message=mensaje,
            from_email=getattr(settings, "DEFAULT_FROM_EMAIL", None),
            recipient_list=destinatarios,
            html_message=mensaje_html,
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
@shared_task(queue="no-programadas")
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
            "institucion"
        ).prefetch_related("institucion__edificio").get(pk=enlace_id)
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
            # Marcar enlace como sincronizado sin re-disparar el signal
            enlace._sig_sync_desactivado = True
            enlace.sincronizado = True
            enlace.save(update_fields=["sincronizado", "actualizado_en"])
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


@shared_task(queue="no-programadas")
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
        # Si falló la deshabilitación en el SIG, el estado quedó inconsistente:
        # desmarcar sincronizado para que se note y se pueda reintentar.
        enlace._sig_sync_desactivado = True
        enlace.sincronizado = False
        enlace.save(update_fields=["sincronizado", "actualizado_en"])


@shared_task(queue="no-programadas")
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
        # Si falló la reactivación en el SIG, el estado quedó inconsistente:
        # desmarcar sincronizado para que se note y se pueda reintentar.
        enlace._sig_sync_desactivado = True
        enlace.sincronizado = False
        enlace.save(update_fields=["sincronizado", "actualizado_en"])


@shared_task(queue="no-programadas")
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
