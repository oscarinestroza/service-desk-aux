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

    Lee las credenciales desde ConfiguracionCorreo (admin):
        - smtp_host         (servidor SMTP)
        - smtp_port         (puerto, default 587)
        - smtp_use_tls      (usar TLS)
        - correo_emisor     (usuario / remitente)
        - password_emisor   (contraseña o app-password)

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

    from .models import ConfiguracionCorreo
    cfg = ConfiguracionCorreo.cargar()

    host = cfg.smtp_host_valor()
    port = cfg.smtp_port_valor()
    use_tls = cfg.smtp_use_tls_valor()
    user = cfg.correo_emisor_valor()
    password = cfg.password_emisor_valor()
    from_email = cfg.default_from_email_valor() or user

    if not user or not password:
        logger.error(
            "enviar_correo_notificacion: correo_emisor o password_emisor "
            "no están configurados en ConfiguracionCorreo (admin)."
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
            if use_tls:
                server.starttls()
                server.ehlo()
            server.login(user, password)
            server.sendmail(from_email, destinatarios, msg.as_string())
        logger.info("Correo enviado exitosamente a: %s", destinatarios)
        return True

    except smtplib.SMTPAuthenticationError:
        logger.error(
            "Error de autenticación SMTP. Verifique correo_emisor / password_emisor en admin."
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
    """Devuelve la configuración de correo efectiva desde ConfiguracionCorreo (admin).

    Returns:
        dict con:
          - correo_emisor
          - password_emisor
          - correo_review (lista, antes separadas por ';')
          - correo_notificacion (lista)
    """
    from .models import ConfiguracionCorreo

    cfg = ConfiguracionCorreo.cargar()
    return {
        "correo_emisor": cfg.correo_emisor_valor(),
        "password_emisor": cfg.password_emisor_valor(),
        "correo_review": cfg.review_lista(),
        "correo_notificacion": cfg.notificacion_lista(),
    }


def aplicar_emisor_a_settings():
    """Aplica el correo emisor/password de la BD a settings para que
    `send_mail` (backend SMTP de Django) use el remitente configurado."""
    from .models import ConfiguracionCorreo

    cfg = ConfiguracionCorreo.cargar()
    settings.EMAIL_HOST_USER = cfg.correo_emisor_valor()
    settings.EMAIL_HOST_PASSWORD = cfg.password_emisor_valor()
    settings.DEFAULT_FROM_EMAIL = cfg.default_from_email_valor() or cfg.correo_emisor_valor()
    settings.EMAIL_HOST = cfg.smtp_host_valor()
    settings.EMAIL_PORT = cfg.smtp_port_valor()
    settings.EMAIL_USE_TLS = cfg.smtp_use_tls_valor()
