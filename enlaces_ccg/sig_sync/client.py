"""
Cliente Playwright para sincronizar enlaces con el SIG.

Migrado desde Selenium a Playwright (sync_api) conservando la misma API
pública (SIGClient.login/logout/crear_usuario/buscar_usuario/cambiar_estatus/
deshabilitar_usuario/reactivar_usuario/cambiar_correo/sincronizar) para que
tasks.py no requiera cambios.

Selectors basados en el HTML del SIG:
  - Login: React app en https://sig.gia.mx/webapp/
  - Crear: /webapp/admin/Seguridad/Usuario/nuevo
  - Editar: /webapp/admin/Seguridad/Usuario/<id>
  - Catálogo: /webapp/admin/Seguridad/Usuario
"""

import hashlib
import logging
import re
import time

from django.conf import settings
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, sync_playwright

from ..models import ConfiguracionSIG

logger = logging.getLogger("sig_sync")

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
# Rutas relativas a base_url (que ya incluye /webapp).
# OJO: base_url termina en ".../webapp", por eso NO deben empezar con "/webapp".
LOGIN_URL = "/seguridad/entrar"
CATALOGO_URL = "/admin/Seguridad/Usuario"
NUEVO_URL = "/admin/Seguridad/Usuario/nuevo"

# Selectores del formulario SIG
SEL = {
    # Login
    "login_usuario": "#usaurio",
    "login_password": "#combinacion",
    "login_rol_select": "#select-simpleSelect",
    "login_boton": "//button[.//span[text()='INGRESAR']]",
    # Formulario de usuario
    "usuario": "#usuario",
    "password": "#password",
    "nombre": "#nombre",
    "apellido_paterno": "#apellidoPaterno",
    "apellido_materno": "#apellidoMaterno",
    "correo": "#correo",
    "celular": "#celular",
    "pin": "#pin",
    "estatus_select": "#select-simpleSelect",  # MUI Select (div)
    "tipo_usuario_select": "#selectTipoUsuario",
    "sexo_select": "#selectSexo",
    "confirmar": "button:has(span.MuiButton-label:contains('Guardar')), button:has(span:contains('Guardar')), #confirmar",
    # React-select (requieren click + input + Enter)
    "grupo_permisos": "div[name='grupo'] input",
    "puesto": "div[name='puesto'] input",
    "area": "div[name='area'] input",
    "departamento": "div[name='departamento'] input",
    "cuadrilla": "div[name='cuadrilla'] input",
}

# Valores por defecto para campos SIG (catálogo)
DEFAULT_SIG_VALUES = {
    "grupo_permisos": "46",      # Solicitante CCG
    "puesto": "13",              # Solicitante
    "area": "2",                 # Unidad funcional
    "departamento": "1",         # NO APLICA
    "tipo_usuario": "2",         # Solicitante
    "estatus_activo": "1",       # Activo
    "estatus_inactivo": "2",     # Inactivo
}


def _sha256(text):
    """SHA-256 de un string (requerido por el SIG para passwords)."""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def config_sig():
    """Devuelve el config SIG (registro singleton o None) sin romper."""
    try:
        return ConfiguracionSIG.cargar()
    except Exception:  # noqa: BLE001 - la BD puede no estar lista
        return None


def sig_usuario_configurado():
    """True si hay usuario SIG configurado (en BD o settings)."""
    cfg = config_sig()
    if cfg and cfg.usuario:
        return True
    return bool(getattr(settings, "SIG_USER", ""))


class SIGClientError(Exception):
    pass


class SIGClient:
    """Interfaz Playwright para el SIG."""

    def __init__(
        self,
        base_url=None,
        username=None,
        password=None,
        headless=True,
        timeout=None,
    ):
        self.base_url = (base_url or getattr(settings, "SIG_URL", "")).rstrip("/")
        self.username = username or getattr(settings, "SIG_USER", "")
        self.password = password or getattr(settings, "SIG_PASSWORD", "")
        self.timeout = timeout or getattr(settings, "SIG_TIMEOUT", 30)
        try:
            cfg = ConfiguracionSIG.cargar()
            self.base_url = base_url or cfg.url_valor(self.base_url)
            self.username = username or cfg.usuario_valor(self.username)
            self.password = password or cfg.password_valor(self.password)
            self.timeout = timeout or cfg.timeout_valor(self.timeout)
            # Normalizar base_url (no debe terminar en '/')
            if self.base_url:
                self.base_url = self.base_url.rstrip("/")
        except Exception as exc:  # noqa: BLE001 - la BD puede no estar lista
            logger.warning(
                "No se pudo leer ConfiguracionSIG (%s); usando settings.", exc
            )
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self._headless = headless
        self._sweetalert_confirmado = False

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------
    def _start_driver(self):
        if self.page:
            return
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(headless=self._headless)
        self.context = self.browser.new_context(
            viewport={"width": 1600, "height": 900},
        )
        self.context.set_default_timeout(self.timeout * 1000)
        self.context.set_default_navigation_timeout(self.timeout * 1000)
        self.page = self.context.new_page()

    def close(self):
        try:
            if self.browser:
                self.browser.close()
        except Exception:
            pass
        try:
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass
        self.page = None
        self.context = None
        self.browser = None
        self.playwright = None

    def __enter__(self):
        self._start_driver()
        return self

    def __exit__(self, *args):
        self.close()

    # ------------------------------------------------------------------
    # Debug
    # ------------------------------------------------------------------
    def _screenshot(self, nombre="debug"):
        """Guarda un screenshot para debug."""
        import os
        from django.conf import settings
        debug_dir = os.path.join(settings.BASE_DIR, "sig_debug")
        os.makedirs(debug_dir, exist_ok=True)
        path = os.path.join(debug_dir, f"{nombre}.png")
        try:
            self.page.screenshot(path=path)
            logger.info("Screenshot guardado: %s", path)
        except Exception as e:
            logger.warning("Error al tomar screenshot: %s", e)
        return path

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _url(self, path):
        """Construye una URL absoluta a partir del base_url sin duplicar rutas.

        base_url termina en ".../webapp". Si `path` empieza con el último
        segmento del base (p.ej. "/webapp/..."), se deja igual para no
        duplicarlo. Ejemplo:
          _url("/webapp/seguridad/entrar")          -> .../webapp/seguridad/entrar
          _url("/seguridad/entrar")                 -> .../webapp/seguridad/entrar
        """
        base = self.base_url.rstrip("/")
        base_tail = "/" + base.split("/")[-1]  # p.ej. "/webapp"
        path = path if path.startswith("/") else "/" + path
        if base_tail != "/" and path.startswith(base_tail + "/"):
            path = path[len(base_tail):]
        return f"{base}{path}"

    def _locator(self, selector):
        return self.page.locator(selector)

    def _wait(self, selector, timeout=None, state="attached"):
        """Espera a que un elemento exista/sea visible según Playwright."""
        loc = self.page.locator(selector)
        t = (timeout or self.timeout) * 1000
        try:
            loc.first.wait_for(state=state, timeout=t)
        except PlaywrightTimeoutError:
            pass
        return loc

    def _wait_clickable(self, selector, timeout=None):
        """Espera a que un elemento sea visible y retorna su locator."""
        loc = self.page.locator(selector)
        t = (timeout or self.timeout) * 1000
        try:
            loc.first.wait_for(state="visible", timeout=t)
        except PlaywrightTimeoutError:
            pass
        return loc.first

    def _safe_send_keys(self, selector, value, clear=True):
        """Llena un input, con manejo de errores."""
        try:
            loc = self._wait_clickable(selector, timeout=10)
            if clear:
                loc.fill("")
            loc.fill(str(value))
            return True
        except Exception as e:
            logger.warning("No se pudo escribir en %s: %s", selector, e)
            return False

    def _select_by_visible(self, selector, text):
        """Selecciona un <select> nativo por texto visible."""
        loc = self._wait(selector, timeout=10)
        loc.first.select_option(label=str(text))

    def _select_by_value(self, selector, value):
        """Selecciona un <select> nativo por valor."""
        loc = self._wait(selector, timeout=10)
        loc.first.select_option(value=str(value))

    def _mui_select(self, select_id, option_text):
        """
        Interactúa con un MUI Select (renderizado como <div>, no <select>):
        1. Click en el contenedor del select para abrir el dropdown
        2. Busca la opción por texto y hace click
        """
        try:
            container = self.page.locator(f"#{select_id}")
            container.first.click(timeout=10000)
            time.sleep(0.5)
            try:
                option = self.page.locator(
                    f"//div[@role='listbox']//li[contains(., '{option_text}')]"
                ).first
                option.click(timeout=5000)
            except Exception:
                options = self.page.locator(
                    "div[role='option'], li[role='option'], .MuiMenuItem-root"
                )
                count = options.count()
                for i in range(count):
                    opt = options.nth(i)
                    if option_text.lower() in (opt.inner_text() or "").lower():
                        opt.click()
                        break
                else:
                    logger.warning(
                        "Opción '%s' no encontrada en MUI Select #%s",
                        option_text,
                        select_id,
                    )
                    return False
            time.sleep(0.3)
            return True
        except Exception as e:
            logger.warning("No se pudo seleccionar en MUI Select #%s: %s", select_id, e)
            return False

    def _mui_select_por_input_id(self, input_id, option_text):
        """
        Abre un MUI Select cuyo input hidden asociado tiene `input_id`, y
        selecciona la opción por texto visible.

        El MUI Select se renderiza como un <div class="MuiSelect-root"
        role="button"> dentro del MISMO FormControl que contiene el input
        hidden (por ejemplo #selectSexo). Este helper localiza ese div
        clickeable, abre el dropdown y elige la opción.
        """
        try:
            hidden_input = self.page.locator(f"#{input_id}")
            hidden_input.first.wait_for(state="attached", timeout=10000)
            select_div = hidden_input.locator(
                "xpath=./ancestor::div[contains(@class,'MuiFormControl-root')]"
                "//div[contains(@class,'MuiSelect-root') and @role='button']"
            ).first
            select_div.click()
            time.sleep(0.5)
            try:
                option = self.page.locator(
                    f"//div[@role='listbox']//li[contains(., '{option_text}')]"
                ).first
                option.click(timeout=5000)
            except Exception:
                options = self.page.locator(
                    "div[role='option'], li[role='option'], .MuiMenuItem-root"
                )
                count = options.count()
                for i in range(count):
                    opt = options.nth(i)
                    if option_text.lower() in (opt.inner_text() or "").lower():
                        opt.click()
                        break
                else:
                    logger.warning(
                        "Opción '%s' no encontrada en MUI Select (input #%s)",
                        option_text,
                        input_id,
                    )
                    return False
            time.sleep(0.3)
            return True
        except Exception as e:
            logger.warning(
                "No se pudo seleccionar en MUI Select (input #%s): %s", input_id, e
            )
            return False

    def _marcar_checkbox_departamento(self, texto_departamento, espera=1.0):
        """
        Marca el checkbox de permisos correspondiente al departamento
        (texto_departamento) en la tabla que aparece al final del formulario.

        La tabla de permisos de departamento solo aparece después de
        seleccionar un departamento; por eso se espera `espera` segundos.
        """
        try:
            time.sleep(espera)
            tr = self.page.locator(
                f"xpath=//tr[contains(., '{texto_departamento}')]"
            ).first
            tr.wait_for(state="attached", timeout=10000)

            # Checkbox real: input[type=checkbox] dentro de la fila.
            # En MUI el input suele estar oculto (tabindex=-1) y por encima se
            # dibuja un ícono SVG, por lo que marcamos el input con force=True
            # y, como respaldo, hacemos click en el contenedor clickeable.
            checkbox = tr.locator("xpath=.//input[@type='checkbox']").first
            try:
                checkbox.wait_for(state="attached", timeout=5000)
            except PlaywrightTimeoutError:
                logger.warning(
                    "No se encontró input[type=checkbox] en fila %s",
                    texto_departamento,
                )
                return False

            if not checkbox.is_checked():
                try:
                    checkbox.check(force=True)
                except Exception:
                    # Respaldo: click en el ícono/área del checkbox (MUI)
                    try:
                        tr.locator(
                            "xpath=.//*[@data-testid='CheckBoxOutlineBlank'] | "
                            ".//*[contains(@class,'Checkbox')]"
                        ).first.click()
                    except Exception:
                        checkbox.click(force=True)
                time.sleep(0.3)
            logger.info("Checkbox de departamento marcado: %s", texto_departamento)
            return True
        except Exception as e:
            logger.warning(
                "No se pudo marcar checkbox de departamento %s: %s",
                texto_departamento,
                e,
            )
            return False

    def _fill_react_select_by_name(self, name, option_label):
        """
        Llena un React-Select identificado por el input hidden `name`.

        React-Select renderiza un input hidden `<input name="X">value</input>`
        y un contenedor clickeable `div.select__control`. El flujo:
        1. localizar el control que contiene el input hidden `name`
        2. click en el value-container para abrir el menú
        3. escribir `option_label` en el input visible (filtra las opciones)
        4. Enter para seleccionar la primera coincidencia
        """
        try:
            hidden = self.page.locator(f"input[name='{name}']").first
            hidden.wait_for(state="attached", timeout=10000)
            container = hidden.locator(
                "xpath=./ancestor::div[contains(@class,'basic-single')]"
            )
            control = container.locator(
                "xpath=.//div[contains(@class,'select__control')]"
            ).first
            value_container = control.locator(
                "xpath=.//div[contains(@class,'select__value-container')]"
            ).first
            value_container.click()
            time.sleep(0.4)
            input_el = control.locator(
                "xpath=.//input[contains(@id,'react-select')]"
            ).first
            input_el.fill(str(option_label))
            time.sleep(1.0)
            input_el.press("Enter")
            time.sleep(0.4)
            return True
        except Exception as e:
            logger.warning("No se pudo llenar React-Select %s: %s", name, e)
            return False

    def _mui_select_by_value(self, select_id, value):
        """
        Interactúa con un MUI Select buscando por data-value attribute.
        """
        try:
            select_container = self.page.locator(f"#{select_id}").first
            select_container.click(timeout=10000)
            time.sleep(0.5)
            try:
                option = self.page.locator(
                    f"div[data-value='{value}'], li[data-value='{value}']"
                ).first
                option.click(timeout=5000)
            except Exception:
                options = self.page.locator(
                    "div[role='option'], li[role='option'], .MuiMenuItem-root"
                )
                count = options.count()
                selected = False
                for i in range(count):
                    opt = options.nth(i)
                    dv = opt.get_attribute("data-value")
                    if dv == str(value):
                        opt.click()
                        selected = True
                        break
                if not selected:
                    for i in range(count):
                        opt = options.nth(i)
                        if str(value) in (opt.inner_text() or ""):
                            opt.click()
                            selected = True
                            break
                if not selected:
                    logger.warning(
                        "Valor '%s' no encontrado en MUI Select #%s", value, select_id
                    )
                    return False
            time.sleep(0.3)
            return True
        except Exception as e:
            logger.warning("No se pudo seleccionar valor en MUI Select #%s: %s", select_id, e)
            return False

    def _fill_react_select(self, container_selector, search_text):
        """
        Interactúa con un React-Select:
        1. Click en el contenedor para abrir el dropdown
        2. Escribe texto de búsqueda
        3. Espera resultados y selecciona el primero con Enter
        """
        try:
            container = self.page.locator(container_selector).first
            container.click(timeout=10000)
            time.sleep(0.3)
            input_el = container.locator("input").first
            input_el.fill(str(search_text))
            time.sleep(0.8)
            input_el.press("Enter")
            time.sleep(0.3)
            return True
        except Exception as e:
            logger.warning("No se pudo llenar React-Select %s: %s", container_selector, e)
            return False

    def _handle_sweetalert(self):
        """Confirma el popup de SweetAlert2 si aparece.

        El criterio de éxito de una creación en el SIG es la aparición de
        este popup: aun cuando falle la extracción del ID o cualquier paso
        posterior, el usuario queda creado en el SIG con solo aparecer la
        confirmación.

        Returns:
            True si apareció un popup (confirmado), False si no apareció.
        """
        try:
            confirm_btn = self.page.locator(
                ".swal2-confirm, .swal2-actions button:first-child"
            ).first
            confirm_btn.wait_for(state="visible", timeout=8000)
            confirm_btn.click()
            time.sleep(1)
            return True
        except Exception:
            logger.info("No apareció SweetAlert2 popup")
            return False

    def logout(self):
        """Cierra sesión en el SIG usando la UI (icono superior → Salir)."""
        try:
            user_menu = self.page.locator(
                "header button[aria-label], "
                "header .MuiIconButton-root, "
                "header [class*='avatar'], "
                "header [class*='user'], "
                "nav button:last-child"
            ).first
            user_menu.click(timeout=8000)
            time.sleep(1)
            try:
                salir_btn = self.page.locator(
                    "//li[contains(.,'Salir')] | //a[contains(.,'Salir')] | "
                    "//button[contains(.,'Salir')]"
                ).first
                salir_btn.click(timeout=5000)
            except Exception:
                opts = self.page.locator(
                    "//*[contains(text(),'Salir') or contains(text(),'salir') "
                    "or contains(text(),'Cerrar sesión')]"
                )
                count = opts.count()
                for i in range(count):
                    try:
                        opts.nth(i).click()
                        break
                    except Exception:
                        continue
            time.sleep(2)
            logger.info("Logout exitoso en SIG")
        except Exception as e:
            logger.warning("Error al hacer logout en SIG: %s", e)
        finally:
            self.close()

    # ------------------------------------------------------------------
    # Login
    # ------------------------------------------------------------------
    def login(self):
        """Login al SIG — flujo exacto de referencia que funciona."""
        self._start_driver()
        logger.info("Iniciando sesión en SIG como %s", self.username)

        page = self.page
        timeout_ms = self.timeout * 1000

        # Contexto nuevo = caché y cookies limpios (equivale al clear previo)

        login_url = self._url(LOGIN_URL)
        page.goto(login_url, timeout=timeout_ms)

        # Usuario
        user = page.locator("#usaurio")
        user.fill(self.username)

        # Contraseña
        pwd = page.locator("#combinacion")
        pwd.fill(self.password)

        # Dropdown rol
        dropdown = page.locator("#select-simpleSelect").first
        dropdown.click(timeout=timeout_ms)
        time.sleep(1)

        # Buscar y click en la opción
        TARGET = "Centro Cívico Gubernamental de Honduras"
        opciones = page.locator(
            "//*[@role='option'] | //li | //div[contains(@class,'MuiMenuItem')]"
        )
        count = opciones.count()
        clicked = False
        for i in range(count):
            op = opciones.nth(i)
            if TARGET.lower() in (op.inner_text() or "").lower():
                op.click()
                clicked = True
                break
        if not clicked:
            logger.error("Rol '%s' no encontrado en login", TARGET)
            raise SIGClientError(f"Rol '{TARGET}' no encontrado en login")

        # Botón INGRESAR
        boton = page.locator("//button[.//span[text()='INGRESAR']]").first
        boton.click(timeout=timeout_ms)

        logger.info("Sesion iniciada en SIG como %s", self.username)

        # Esperar cambio de URL (sale del login)
        try:
            page.wait_for_url(
                "**/*", timeout=timeout_ms, wait_until="domcontentloaded"
            )
        except Exception:
            pass
        try:
            page.wait_for_function(
                "() => !location.href.includes('seguridad/entrar')",
                timeout=timeout_ms,
            )
        except Exception:
            pass

        time.sleep(1)

        logger.info("Primer wait")

        # Verificar que no estemos en página de login
        current = page.url
        if "seguridad/entrar" in current.lower() or "login" in current.lower():
            raise SIGClientError(f"Login falló — seguimos en página de login: {current}")

        logger.info("Login exitoso: %s", current)
        return True

    # ------------------------------------------------------------------
    # Crear usuario
    # ------------------------------------------------------------------
    def crear_usuario(self, enlace):
        """
        Crea un usuario en el SIG a partir de un EnlaceAutorizado.

        Retorna el ID del usuario creado (extraído de la URL post-guardado).
        """
        logger.info("Creando usuario SIG: %s", enlace.usuario_sig)
        self.page.goto(self._url(NUEVO_URL), timeout=self.timeout * 1000)
        time.sleep(2)

        # ---- Campos de texto ----
        self._safe_send_keys(SEL["usuario"], enlace.usuario_sig)
        self._safe_send_keys(
            SEL["password"], _sha256(enlace.password_sig or settings.SIG_DEFAULT_PASSWORD)
        )
        self._safe_send_keys(SEL["nombre"], enlace.nombres)
        self._safe_send_keys(SEL["apellido_paterno"], enlace.primer_apellido)
        self._safe_send_keys(SEL["apellido_materno"], enlace.segundo_apellido or "")
        self._safe_send_keys(SEL["correo"], enlace.correo_principal or "")
        self._safe_send_keys(SEL["celular"], enlace.telefono_principal)
        self._safe_send_keys(SEL["pin"], enlace.pin_sig or "0000")

        # ---- MUI Selects ----
        self._mui_select_por_input_id("simple-select", "1")
        self._mui_select_por_input_id("selectDuracionSesion", "1")
        self._mui_select_por_input_id("selectEstatus", "Activo")
        self._mui_select_por_input_id("selectTipoUsuario", "Solicitante")
        genero_map = {"M": "Masculino", "F": "Femenino", "O": "Otro"}
        genero_texto = genero_map.get(enlace.genero, "")
        if genero_texto:
            self._mui_select_por_input_id("selectSexo", genero_texto)
        else:
            logger.warning(
                "Género '%s' no mapeado; se deja Sexo sin tocar.", enlace.genero
            )

        # ---- React-Selects (valores fijos por defecto) ----
        self._fill_react_select_by_name("grupo", "Solicitante CCG")
        self._fill_react_select_by_name("puesto", "Solicitante CCGH")
        self._fill_react_select_by_name("area", "Oficinas Técnicos, Administrativos")
        self._fill_react_select_by_name("unidad", "Área de oficinas")
        self._fill_react_select_by_name("cuadrilla", "No aplica")
        self._fill_react_select_by_name("departamento", "USUARIOS CCG")

        # ---- Checkbox de permisos de departamento ----
        self._marcar_checkbox_departamento("USUARIOS CCG", espera=1.0)

        time.sleep(0.5)

        # ---- Guardar ----
        try:
            confirmar = self._wait_clickable(SEL["confirmar"], timeout=10)
            confirmar.click()
        except Exception:
            buttons = self.page.locator("button")
            count = buttons.count()
            for i in range(count):
                b = buttons.nth(i)
                txt = (b.inner_text() or "").upper()
                if "CONFIRMAR" in txt or "GUARDAR" in txt:
                    b.click()
                    break

        sweetalert_ok = self._handle_sweetalert()
        time.sleep(2)

        sig_id = self._extract_user_id_from_url()
        if sig_id:
            logger.info("Usuario SIG creado con ID: %s", sig_id)
        else:
            logger.info("Usuario SIG creado (ID no recuperado; no es crítico)")

        self._sweetalert_confirmado = sweetalert_ok
        return sig_id

    def _extract_user_id_from_url(self):
        """Extrae el ID del usuario de la URL actual después de guardar."""
        time.sleep(2)
        url = self.page.url
        match = re.search(r"/Usuario/(\d+)", url)
        if match:
            return match.group(1)

        try:
            match = re.search(r'"id"\s*:\s*(\d+)', self.page.content())
            if match:
                return match.group(1)
        except Exception:
            pass

        return None

    # ------------------------------------------------------------------
    # Buscar usuario en catálogo
    # ------------------------------------------------------------------
    def buscar_usuario(self, usuario_sig):
        """
        Busca un usuario en el catálogo del SIG por su nombre de usuario.
        Usa el flujo que funciona: segundo input search -> buscar -> doble clic en celda.
        Retorna el ID del usuario encontrado o None.
        """
        logger.info("Buscando usuario SIG: %s", usuario_sig)
        page = self.page
        timeout_ms = self.timeout * 1000

        page.goto(self._url(CATALOGO_URL), timeout=timeout_ms)
        time.sleep(2)

        xpath_segundo_input = "xpath=(//input[@type='search'])[2]"
        input_usuario = page.locator(xpath_segundo_input).first
        input_usuario.click(timeout=timeout_ms)
        input_usuario.fill(usuario_sig)
        time.sleep(1)

        xpath_celda_usuario = f"xpath=//tbody//td[contains(text(), '{usuario_sig}')]"
        celda_usuario = page.locator(xpath_celda_usuario).first
        celda_usuario.wait_for(state="visible", timeout=timeout_ms)
        celda_usuario.dblclick()
        logger.info("Doble clic en celda del usuario: %s", usuario_sig)
        time.sleep(2)

        sig_id = self._extract_user_id_from_url()
        if sig_id:
            logger.info("Usuario encontrado con ID: %s", sig_id)
        return sig_id

    # ------------------------------------------------------------------
    # Cambiar estatus (unificado)
    # ------------------------------------------------------------------
    def cambiar_estatus(self, usuario_sig, nuevo_estatus):
        """
        Cambia el estatus de un usuario en el SIG.
        Flujo: buscar_usuario -> doble clic -> cambiar dropdown select-simpleSelect (Estatus) -> Guardar -> Aceptar (SweetAlert).

        Args:
            usuario_sig: nombre de usuario en SIG
            nuevo_estatus: "Activo" o "Cancelado"
        """
        if nuevo_estatus not in ("Activo", "Cancelado"):
            raise ValueError(
                f"Estatus inválido: {nuevo_estatus}. Use 'Activo' o 'Cancelado'"
            )

        logger.info("Cambiando estatus de %s a %s", usuario_sig, nuevo_estatus)

        sig_id = self.buscar_usuario(usuario_sig)
        if not sig_id:
            raise SIGClientError(
                f"No se pudo encontrar el usuario {usuario_sig} en el SIG"
            )

        time.sleep(3)

        page = self.page
        timeout_ms = self.timeout * 1000
        label_estatus = page.locator(
            "//label[contains(text(),'Estatus') or contains(text(),'estatus') or contains(text(),'Status')]"
        ).first
        label_estatus.wait_for(state="attached", timeout=timeout_ms)
        select_div = label_estatus.locator(
            "xpath=./ancestor::div[contains(@class,'MuiFormControl-root')]"
            "//div[contains(@class,'MuiSelect-root') and @role='button']"
        ).first
        select_div.click()
        time.sleep(0.5)

        opciones = page.locator(
            "//li[@role='option'] | //div[@role='option'] | //div[contains(@class,'MuiMenuItem')]"
        )
        count = opciones.count()
        for i in range(count):
            op = opciones.nth(i)
            if nuevo_estatus.lower() in (op.inner_text() or "").lower():
                op.click()
                break
        else:
            raise SIGClientError(f"Opción '{nuevo_estatus}' no encontrada en dropdown")

        time.sleep(0.5)

        # 3. Click en botón "Guardar"
        try:
            btn = self._wait_clickable("//button[.//span[contains(text(),'Guardar')]]", timeout=10)
            btn.click()
        except Exception:
            buttons = page.locator("button")
            count = buttons.count()
            for i in range(count):
                if "GUARDAR" in (buttons.nth(i).inner_text() or "").upper():
                    buttons.nth(i).click()
                    break

        self._handle_sweetalert()
        time.sleep(2)

        logger.info("Usuario %s actualizado a %s exitosamente", usuario_sig, nuevo_estatus)
        return True

    # ------------------------------------------------------------------
    # Deshabilitar usuario (usa cambiar_estatus)
    # ------------------------------------------------------------------
    def deshabilitar_usuario(self, usuario_sig):
        """Deshabilita un usuario en el SIG: cambia estatus a Cancelado."""
        return self.cambiar_estatus(usuario_sig, "Cancelado")

    # ------------------------------------------------------------------
    # Reactivar usuario (usa cambiar_estatus)
    # ------------------------------------------------------------------
    def reactivar_usuario(self, usuario_sig):
        """Reactiva un usuario en el SIG: cambia estatus a Activo."""
        return self.cambiar_estatus(usuario_sig, "Activo")

    # ------------------------------------------------------------------
    # Cambiar correo
    # ------------------------------------------------------------------
    def cambiar_correo(self, usuario_sig, nuevo_correo):
        """
        Cambia el correo de un usuario en el SIG.
        Flujo: buscar_usuario -> doble clic -> editar campo correo -> Guardar -> Aceptar (SweetAlert).

        Args:
            usuario_sig:   nombre de usuario en SIG
            nuevo_correo:  nuevo correo a asignar ("" para vaciarlo)
        """
        logger.info("Cambiando correo de %s a %s", usuario_sig, nuevo_correo or "(vacío)")

        sig_id = self.buscar_usuario(usuario_sig)
        if not sig_id:
            raise SIGClientError(
                f"No se pudo encontrar el usuario {usuario_sig} en el SIG"
            )

        time.sleep(3)

        correo_input = self._wait_clickable(SEL["correo"], timeout=10)
        correo_input.fill("")
        correo_input.fill(nuevo_correo or "")
        time.sleep(0.5)

        try:
            btn = self._wait_clickable(
                "//button[.//span[contains(text(),'Guardar')]]", timeout=10
            )
            btn.click()
        except Exception:
            buttons = self.page.locator("button")
            count = buttons.count()
            for i in range(count):
                if "GUARDAR" in (buttons.nth(i).inner_text() or "").upper():
                    buttons.nth(i).click()
                    break

        self._handle_sweetalert()
        time.sleep(2)

        logger.info("Correo de %s actualizado exitosamente", usuario_sig)
        return True

    # ------------------------------------------------------------------
    # Orquestador
    # ------------------------------------------------------------------
    def sincronizar(self, enlace, accion=None):
        """
        Orquesta la sincronización de un enlace con el SIG.

        La acción se puede pasar explícitamente ("crear", "reactivar",
        "deshabilitar"). Si no se pasa, se infiere.

        Retorna un dict con el resultado.
        """
        result = {"accion": accion, "exitoso": False, "usuario_sig_id": "", "mensaje": ""}

        try:
            self._start_driver()
            self.login()

            if accion is None:
                if enlace.estado == "INACTIVO" and enlace.usuario_sig:
                    accion = "deshabilitar"
                else:
                    accion = "crear"
                result["accion"] = accion

            if accion == "crear":
                sig_id = self.crear_usuario(enlace)
                if self._sweetalert_confirmado:
                    result["exitoso"] = True
                    result["usuario_sig_id"] = sig_id or ""
                    result["mensaje"] = (
                        f"Usuario creado exitosamente{(' con ID ' + sig_id) if sig_id else ''}"
                    )
                else:
                    result["mensaje"] = (
                        "No apareció confirmación (SweetAlert) tras guardar el usuario"
                    )

            elif accion == "reactivar":
                self.reactivar_usuario(enlace.usuario_sig)
                result["exitoso"] = True
                result["mensaje"] = "Usuario reactivado exitosamente"

            elif accion == "deshabilitar":
                self.deshabilitar_usuario(enlace.usuario_sig)
                result["exitoso"] = True
                result["mensaje"] = "Usuario deshabilitado exitosamente"

            else:
                result["mensaje"] = f"Acción desconocida: {accion}"

        except Exception as e:
            result["mensaje"] = str(e)
            logger.error("Error sincronizando enlace %s: %s", enlace.pk, e)

        finally:
            self.logout()

        return result
