"""Descarga del reporte de tickets del SIG (Excel) y pipeline de tickets.

Fase 1: solo la descarga del Excel. El parser a BD llega en la Fase 2.

Dos modos de descarga (misma función, pasos opcionales):
  - MODO_RAPIDA:   pasos 1 y 5 (ir a la URL + apretar descargar). Sin filtros.
                   Timeout de descarga: 3 minutos.
  - MODO_COMPLETA: pasos 1 a 5 (además activa filtros avanzados, llena la
                   fecha 'desde' y aplica). Timeout de descarga: 10 minutos.

Las credenciales del worker de tickets se leen del entorno
(SIG_TICKETS_URL / SIG_TICKETS_USUARIO / SIG_TICKETS_PASSWORD), a diferencia
de la cuenta de enlaces que vive en ConfiguracionSIG (admin).
"""

import json
import logging
import os
import time
from datetime import datetime
from pathlib import Path

from celery import shared_task
from django.conf import settings
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from ..models import ConfiguracionTickets
from .client import SIGClient, SIGClientError

logger = logging.getLogger("sig_sync_tickets")

MODO_RAPIDA = "rapida"
MODO_COMPLETA = "completa"
MODO_VALORES = (MODO_RAPIDA, MODO_COMPLETA)

# Timeouts de descarga (segundos)
TIMEOUT_DESCARGA_RAPIDA = 3 * 60     # aborta a los 3 min
TIMEOUT_DESCARGA_COMPLETA = 10 * 60  # aborta a los 10 min

REPORTE_URL_DEFAULT = "/admin/Solicitud/SeguimientoAtencion"

# Selectores provisionales del reporte (calibrar en la primera prueba real con
# capturas; se evita el hash jss... de MUI cuando es posible).
SEL_TICKETS = {
    "filtros_avanzados_switch": (
        "input.MuiSwitch-input[type='checkbox']",
        "input[type='checkbox'][value='']",
    ),
    "fecha_desde": (
        "div:has(> label:has-text('Inicio')) input",
        "input[aria-label*='desde' i]",
        "input[placeholder*='desde' i]",
    ),
    "aplicar_filtros": (
        "button:has-text('Aplicar')",
        "button:has-text('Buscar')",
        "button:has-text('Filtrar')",
        "button[type='submit']",
    ),
    "descargar_excel": (
        "#btnSolicitudesExcel",
        "button:has-text('Excel')",
        "button:has-text('Descargar')",
        "a:has-text('Excel')",
        "a:has-text('Descargar')",
        "button[class*='excel' i]",
    ),
}


def _directorio_descargas() -> Path:
    path = Path(settings.MEDIA_ROOT) / "tickets" / "descargas"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _normalizar_fecha_desde(valor) -> str:
    """Convierte el valor 'desde' a formato dd/mm/aaaa (como lo pide el SIG).

    Acepta '01/01/2020' (dd/mm/aaaa) o '2020-01-01' (ISO).
    """
    valor = (valor or "").strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(valor, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return "01/01/2020"


def _normalizar_base_url(url) -> str:
    """Deja la URL base del SIG en su forma canónica (termina en /webapp).

    Tolera que llegue completa (https://sig.gia.mx/webapp/seguridad/entrar)
    o ya base (https://sig.gia.mx/webapp) para no depender del formato.
    """
    url = (url or "").strip().rstrip("/")
    idx = url.find("/webapp")
    if idx != -1:
        return url[: idx + len("/webapp")]
    return url


def _credenciales_tickets():
    """Devuelve (url, usuario, password) de la cuenta del worker de tickets.

    Usa el entorno (env vars); la URL cae a ConfiguracionSIG si falta.
    """
    from ..models import ConfiguracionSIG

    url = os.environ.get("SIG_TICKETS_URL", "").strip()
    if not url:
        try:
            url = ConfiguracionSIG.cargar().url_valor()
        except Exception:  # noqa: BLE001
            url = ""
    url = _normalizar_base_url(url)
    usuario = os.environ.get("SIG_TICKETS_USUARIO", "").strip()
    password = os.environ.get("SIG_TICKETS_PASSWORD", "").strip()
    if not usuario or not password:
        raise SIGClientError(
            "Credenciales del worker de tickets no configuradas: "
            "SIG_TICKETS_USUARIO / SIG_TICKETS_PASSWORD"
        )
    return url, usuario, password


def _fecha_desde_config() -> str:
    try:
        return ConfiguracionTickets.cargar().fecha_desde_valor()
    except Exception:  # noqa: BLE001
        return "01/01/2020"


def _url_reporte_config() -> str:
    try:
        return ConfiguracionTickets.cargar().url_reporte_valor()
    except Exception:  # noqa: BLE001
        return REPORTE_URL_DEFAULT


def _tamano_humano(path: Path) -> str:
    size = path.stat().st_size
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024:
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"


def _escribir_metadatos(datos) -> Path:
    """Guarda un JSON con datos del intento de descarga (para debug)."""
    datos = dict(datos)
    datos.setdefault("fecha", datetime.now().isoformat(timespec="seconds"))
    path = _directorio_descargas() / f"descarga_{datetime.now():%Y%m%d_%H%M%S}.json"
    try:
        path.write_text(
            json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:  # noqa: BLE001
        logger.warning("No se pudo escribir metadatos: %s", e)
    return path


class TicketsSIGClient(SIGClient):
    """Cliente Playwright para descargar el reporte de tickets del SIG."""

    def __init__(self, modo=MODO_RAPIDA, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.modo = modo if modo in MODO_VALORES else MODO_RAPIDA
        self.timeout_descarga_ms = (
            TIMEOUT_DESCARGA_RAPIDA if self.modo == MODO_RAPIDA
            else TIMEOUT_DESCARGA_COMPLETA
        ) * 1000

    def descargar_excel(self, report_url=None, desde=None):
        """Pasos 1 y 5 (+ 2, 3, 4 en modo completo). Devuelve dict con la ruta."""
        url = self._url(report_url or REPORTE_URL_DEFAULT)
        inicio = time.monotonic()
        page = self.page
        timeout_ms = self.timeout * 1000

        logger.info("Tickets (%s): navegando a %s", self.modo, url)
        page.goto(url, timeout=timeout_ms, wait_until="domcontentloaded")

        if self.modo == MODO_COMPLETA:
            self._aplicar_filtros_completa(desde or _fecha_desde_config())

        with page.expect_download(timeout=self.timeout_descarga_ms) as dl_info:
            self._disparar_descarga()

        download = dl_info.value
        nombre = download.suggested_filename or f"tickets_{self.modo}.xlsx"
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        destino = _directorio_descargas() / f"{stamp}_{nombre}"
        download.save_as(str(destino))

        segundos = round(time.monotonic() - inicio, 1)
        logger.info(
            "Tickets (%s): descargado %s (%s) en %s s",
            self.modo, nombre, _tamano_humano(destino), segundos,
        )
        return {
            "modo": self.modo,
            "path": str(destino),
            "nombre": nombre,
            "tamano_bytes": destino.stat().st_size,
            "segundos": segundos,
        }

    def _aplicar_filtros_completa(self, desde):
        """Pasos 2, 3 y 4: filtros avanzados + fecha 'desde' + aplicar."""
        if not self._click_primero(SEL_TICKETS["filtros_avanzados_switch"]):
            raise SIGClientError(
                "No se encontró el switch de filtros avanzados (calibrar selector)"
            )
        if not self._fill_primero(SEL_TICKETS["fecha_desde"], desde):
            raise SIGClientError(
                "No se encontró el campo de fecha 'desde' (calibrar selector)"
            )
        if not self._click_primero(SEL_TICKETS["aplicar_filtros"]):
            raise SIGClientError(
                "No se encontró el botón de aplicar filtros (calibrar selector)"
            )
        self.page.wait_for_timeout(1500)
        logger.info("Tickets (completa): filtros aplicados (desde=%s)", desde)

    def _disparar_descarga(self):
        if not self._click_primero(SEL_TICKETS["descargar_excel"]):
            raise SIGClientError(
                "No se encontró el botón de descarga del Excel (calibrar selector)"
            )

    def _click_primero(self, selectores):
        for sel in selectores:
            try:
                loc = self.page.locator(sel).first
                loc.wait_for(state="visible", timeout=8000)
                loc.click(timeout=8000)
                return True
            except PlaywrightTimeoutError:
                continue
            except Exception:  # noqa: BLE001
                continue
        return False

    def _fill_primero(self, selectores, valor):
        for sel in selectores:
            try:
                loc = self.page.locator(sel).first
                loc.wait_for(state="visible", timeout=8000)
                loc.fill(str(valor))
                return True
            except PlaywrightTimeoutError:
                continue
            except Exception:  # noqa: BLE001
                continue
        return False


def descargar_excel_tickets(modo=MODO_RAPIDA, desde=None, report_url=None):
    """Orquesta la descarga: credenciales -> login -> pasos -> guardado.

    Regresa dict con los datos de la descarga. Si algo falla guarda screenshot
    y metadatos de error, y re-lanza la excepción.
    """
    base_url, usuario, password = _credenciales_tickets()
    desde = _normalizar_fecha_desde(desde or _fecha_desde_config())
    url_reporte = report_url or _url_reporte_config()

    client = None
    try:
        client = TicketsSIGClient(
            modo=modo, base_url=base_url, username=usuario, password=password
        )
        client.login()
        resultado = client.descargar_excel(
            report_url=url_reporte, desde=desde
        )
        _escribir_metadatos(resultado)
        return resultado
    except Exception as e:
        logger.exception("Descarga de tickets (%s) falló", modo)
        _escribir_metadatos({"modo": modo, "error": str(e)})
        if client is not None:
            try:
                client._screenshot(f"tickets_{modo}_error")
            except Exception:  # noqa: BLE001
                pass
        raise
    finally:
        if client is not None:
            try:
                client.logout()
            except Exception:  # noqa: BLE001
                try:
                    client.close()
                except Exception:  # noqa: BLE001
                    pass


@shared_task(queue="programadas")
def descargar_excel_tickets_task(modo=MODO_RAPIDA, desde=None, report_url=None):
    """Tarea Celery del worker de tickets (cola programadas)."""
    resultado = descargar_excel_tickets(
        modo=modo, desde=desde, report_url=report_url
    )
    logger.info("Tarea descargar_excel_tickets_task terminada: %s", resultado)
    return resultado