"""
Utilidades de la app enlaces_ccg.

Funciones auxiliares que no encajan en views ni models.
"""

import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from django.conf import settings

logger = logging.getLogger(__name__)


def enviar_correo_notificacion(
    asunto: str,
    mensaje: str,
    destinatarios: list[str],
    *,
    html: bool = False,
) -> bool:
    """
    Envía un correo electrónico usando smtplib con soporte TLS (puerto 587).

    Lee las credenciales de settings:
        - EMAIL_HOST          (servidor SMTP)
        - EMAIL_PORT          (puerto, default 587)
        - EMAIL_HOST_USER     (usuario / remitente)
        - EMAIL_HOST_PASSWORD (contraseña o app-password)
        - DEFAULT_FROM_EMAIL  (dirección From)

    Args:
        asunto:       Asunto del correo.
        mensaje:      Cuerpo del mensaje (texto plano o HTML).
        destinatarios: Lista de direcciones de correo destino.
        html:         Si es True, el body se envía como text/html;
                      de lo contrario como text/plain.

    Returns:
        True si el correo se envió exitosamente, False en caso contrario.
    """
    if not destinatarios:
        logger.warning("enviar_correo_notificacion: sin destinatarios, se omite envío.")
        return False

    # --- Configuración: prioriza la DB (editable en admin) sobre settings ---
    host = getattr(settings, "EMAIL_HOST", "smtp.gmail.com")
    port = getattr(settings, "EMAIL_PORT", 587)
    base = _config_basico()
    user = base["correo_emisor"]
    password = base["password_emisor"]
    from_email = getattr(settings, "DEFAULT_FROM_EMAIL", user)

    if not user or not password:
        logger.error(
            "enviar_correo_notificacion: EMAIL_HOST_USER o EMAIL_HOST_PASSWORD "
            "no están configurados en settings."
        )
        return False

    # --- Construcción del mensaje MIME ---
    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"] = from_email
    msg["To"] = ", ".join(destinatarios)

    subtype = "html" if html else "plain"
    msg.attach(MIMEText(mensaje, subtype, "utf-8"))

    # --- Envío con smtplib + TLS ---
    try:
        with smtplib.SMTP(host, port) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(user, password)
            server.sendmail(from_email, destinatarios, msg.as_string())
        logger.info("Correo enviado exitosamente a: %s", destinatarios)
        return True

    except smtplib.SMTPAuthenticationError:
        logger.error(
            "Error de autenticación SMTP. Verifique EMAIL_HOST_USER / EMAIL_HOST_PASSWORD."
        )
        return False
    except smtplib.SMTPConnectError:
        logger.error("No se pudo conectar al servidor SMTP %s:%s", host, port)
        return False
    except smtplib.SMTPException as exc:
        logger.error("Error SMTP durante el envío: %s", exc)
        return False
    except OSError as exc:
        logger.error("Error de red al enviar correo: %s", exc)
        return False


def obtener_config_correo():
    """Devuelve la configuración de correo efectiva.

    Los valores definidos en el modelo singleton `ConfiguracionCorreo`
    (editables desde el admin) tienen prioridad; si no están, se usa el
    fallback de `settings`.

    Returns:
        dict con:
          - correo_emisor
          - password_emisor
          - correo_review (lista, antes separadas por ';')
          - correo_notificacion (lista)
    """
    from .models import ConfiguracionCorreo

    cfg = ConfiguracionCorreo.cargar()

    from django.conf import settings as s
    return {
        "correo_emisor": cfg.emisor(
            fallback=getattr(s, "EMAIL_HOST_USER", "")
        ),
        "password_emisor": cfg.emisor_password(
            fallback=getattr(s, "EMAIL_HOST_PASSWORD", "")
        ),
        "correo_review": cfg.review_lista(
            fallback=[getattr(s, "SIG_REVIEW_EMAIL", "")] if getattr(s, "SIG_REVIEW_EMAIL", "") else []
        ),
        # Correo(s) adicional(es) de notificación de éxito (fase de prueba).
        # En producción la notificación va al correo_principal del enlace.
        "correo_notificacion": cfg.notificacion_lista(fallback=[]),
    }


def aplicar_emisor_a_settings():
    """Aplica el correo emisor/password de la DB a settings para que
    `send_mail` (backend SMTP de Django) use el remitente configurado."""
    config = _config_basico()
    settings.EMAIL_HOST_USER = config["correo_emisor"]
    settings.EMAIL_HOST_PASSWORD = config["password_emisor"]
    settings.DEFAULT_FROM_EMAIL = config["correo_emisor"]


def _config_basico():
    """Config básica (emisor/password) sin armar listas, para aplicar a settings."""
    from .models import ConfiguracionCorreo

    cfg = ConfiguracionCorreo.cargar()
    return {
        "correo_emisor": cfg.emisor(fallback=settings.EMAIL_HOST_USER),
        "password_emisor": cfg.emisor_password(fallback=settings.EMAIL_HOST_PASSWORD),
    }
