"""
Models para la gestión de Enlaces Autorizados del CCG Honduras.

Relación jerárquica:
    Edificio ↔ Institución (M2M through InstitucionEdificio) → EnlaceAutorizado
"""

from datetime import datetime, time, timedelta

from django.conf import settings
from django.db import models
from django.utils import timezone


# ---------------------------------------------------------------------------
# Edificio
# ---------------------------------------------------------------------------
class Edificio(models.Model):
    """Representa cada torre o cuerpo del Centro Cívico Gubernamental."""

    nombre = models.CharField(
        max_length=100,
        unique=True,
        help_text="Nombre del edificio (ej. Torre 1, Cuerpo Bajo A)",
    )
    siglas = models.CharField(
        max_length=20,
        blank=True,
        default="",
        help_text="Siglas abreviadas (ej. T1, CBA)",
    )
    descripcion = models.TextField(
        blank=True,
        default="",
        help_text="Descripción opcional del edificio",
    )
    visible = models.BooleanField(
        default=True,
        help_text="Marque si desea que este edificio aparezca en la lista pública",
        verbose_name="Visible en directorio",
    )
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["nombre"]
        verbose_name = "Edificio"
        verbose_name_plural = "Edificios"

    def __str__(self):
        return self.nombre

    # Mapeo de pisos por edificio (siglas → cantidad de niveles desde PB)
    PISOS_POR_EDIFICIO = {
        "CBA": 5,    # PB a Nivel 5
        "CBB": 7,    # PB a Nivel 7
        "CBC": 7,    # PB a Nivel 7
        "T1": 23,    # PB a Nivel 23
        "T2": 24,    # PB a Nivel 24
    }

    def niveles_disponibles(self):
        """Devuelve lista de niveles disponibles para este edificio."""
        siglas_upper = (self.siglas or "").upper().strip()
        max_nivel = self.PISOS_POR_EDIFICIO.get(siglas_upper)
        if max_nivel is None:
            return ["PB"]
        niveles = ["PB"]
        for i in range(1, max_nivel + 1):
            niveles.append(f"Nivel {i}")
        return niveles


# ---------------------------------------------------------------------------
# Institución
# ---------------------------------------------------------------------------
class Institucion(models.Model):
    """Institución gubernamental radicada dentro de un Edificio."""

    ESTADO_CHOICES = [
        ("ACTIVO", "Activo"),
        ("INACTIVO", "Inactivo"),
    ]

    NIVEL_CHOICES = [("", "Sin nivel"), ("PB", "PB")] + [
        (f"Nivel {i}", f"Nivel {i}") for i in range(1, 25)
    ]

    edificio = models.ManyToManyField(
        Edificio,
        blank=True,
        through="InstitucionEdificio",
        related_name="instituciones",
        help_text="Edificio(s) donde radica la institución",
    )
    nombre = models.CharField(
        max_length=200,
        help_text="Nombre completo de la institución",
    )
    siglas = models.CharField(
        max_length=60,
        blank=True,
        default="",
        help_text="Siglas de la institución (ej. SESAL, SEPLAN)",
    )
    # Niveles disponibles por edificio (siglas del edificio -> lista de niveles)
    NIVELES_POR_EDIFICIO = {
        "CBA": ["PB"] + [f"Nivel {i}" for i in range(1, 6)],
        "CBB": ["PB"] + [f"Nivel {i}" for i in range(1, 8)],
        "CBC": ["PB"] + [f"Nivel {i}" for i in range(1, 8)],
        "T1": ["PB"] + [f"Nivel {i}" for i in range(1, 24)],
        "T2": ["PB"] + [f"Nivel {i}" for i in range(1, 25)],
    }

    estado = models.CharField(
        max_length=10,
        choices=ESTADO_CHOICES,
        default="ACTIVO",
        help_text="Estado actual de la institución",
    )
    contacto_nombre = models.CharField(
        max_length=150,
        blank=True,
        default="",
        help_text="Nombre del contacto institucional",
    )
    contacto_correo = models.EmailField(
        blank=True,
        default="",
        help_text="Correo electrónico del contacto",
    )
    contacto_telefono = models.CharField(
        max_length=20,
        blank=True,
        default="",
        help_text="Teléfono del contacto",
    )
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["nombre"]
        verbose_name = "Institución"
        verbose_name_plural = "Instituciones"

    def __str__(self):
        label = self.siglas if self.siglas else self.nombre
        return label

    @property
    def nombre_completo(self):
        """Devuelve las siglas entre paréntesis si existen."""
        if self.siglas:
            return f"{self.nombre} ({self.siglas})"
        return self.nombre


# ---------------------------------------------------------------------------
# InstitucionEdificio (tabla intermedia M2M)
# ---------------------------------------------------------------------------
class InstitucionEdificio(models.Model):
    """Relación Institución ↔ Edificio con niveles propios por edificio."""

    institucion = models.ForeignKey(
        Institucion,
        on_delete=models.CASCADE,
        related_name="instituciones_edificios",
    )
    edificio = models.ForeignKey(
        Edificio,
        on_delete=models.CASCADE,
        related_name="instituciones_edificios",
    )
    nivel = models.JSONField(
        default=list,
        blank=True,
        help_text='Niveles/pisos en este edificio (ej. ["PB", "Nivel 1", "Nivel 2"])',
    )

    class Meta:
        unique_together = ("institucion", "edificio")
        verbose_name = "Institución – Edificio"
        verbose_name_plural = "Instituciones – Edificios"

    def __str__(self):
        return f"{self.institucion} → {self.edificio}"


# ---------------------------------------------------------------------------
# Enlace Autorizado
# ---------------------------------------------------------------------------
class EnlaceAutorizado(models.Model):
    """Funcionario autorizado de una Institución dentro del CCG."""

    ESTADO_CHOICES = [
        ("ACTIVO", "Activo"),
        ("INACTIVO", "Inactivo"),
    ]

    GENERO_CHOICES = [
        ("M", "Masculino"),
        ("F", "Femenino"),
    ]

    # ---- Relaciones ----
    institucion = models.ForeignKey(
        Institucion,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="enlaces",
    )
    nivel_referencia = models.CharField(
        max_length=150,
        blank=True,
        default="",
        help_text="Nivel de referencia del enlace",
    )

    # ---- Datos personales ----
    nombres = models.CharField(
        max_length=150,
        help_text="Nombres completos del enlace",
    )
    primer_apellido = models.CharField(
        max_length=100,
        help_text="Primer apellido",
    )
    segundo_apellido = models.CharField(
        max_length=100,
        blank=True,
        default="",
        help_text="Segundo apellido (opcional)",
    )
    genero = models.CharField(
        max_length=1,
        choices=GENERO_CHOICES,
        blank=True,
        default="",
        help_text="Género del enlace",
    )

    # ---- Correo electrónico ----
    correo_principal = models.EmailField(
        blank=True,
        null=True,
        default="",
        unique=True,
        help_text="Correo institucional principal (único; el SIG no permite duplicados)",
    )
    correo_secundario = models.EmailField(
        blank=True,
        default="",
        help_text="Correo secundario (opcional)",
    )

    # ---- Teléfono ----
    telefono_principal = models.CharField(
        max_length=20,
        blank=True,
        default="",
        help_text="Teléfono principal de contacto",
    )
    telefono_secundario = models.CharField(
        max_length=20,
        blank=True,
        default="",
        help_text="Teléfono secundario (opcional)",
    )
    extension_telefonica = models.CharField(
        max_length=10,
        blank=True,
        default="",
        help_text="Extensión telefónica interna",
    )

    # ---- Sistema SIG ----
    usuario_sig = models.CharField(
        max_length=50,
        blank=True,
        default="",
        help_text="Usuario en el Sistema de Información Gerencial",
    )
    password_sig = models.CharField(
        max_length=128,
        blank=True,
        default="",
        help_text="Contraseña por defecto en el SIG",
    )
    nombre_sig = models.CharField(
        max_length=100,
        blank=True,
        default="",
        help_text="Nombre registrado en el SIG",
    )
    pin_sig = models.CharField(
        max_length=10,
        blank=True,
        default="",
        help_text="PIN de acceso al SIG",
    )
    sincronizado = models.BooleanField(
        default=False,
        help_text="Indica si el usuario fue creado/sincronizado en el SIG "
                  "(se marca automáticamente al crear con éxito, o manualmente).",
        verbose_name="Sincronizado con SIG",
    )

    # ---- Estado ----
    estado = models.CharField(
        max_length=10,
        choices=ESTADO_CHOICES,
        default="ACTIVO",
        help_text="Estado actual del enlace (Estatus en Dynamics)",
    )

    # ---- Alta ----
    fecha_alta = models.DateField(
        null=True,
        blank=True,
        help_text="Fecha de alta del enlace",
    )
    oficio_alta = models.CharField(
        max_length=100,
        blank=True,
        default="",
        help_text="Número de oficio de alta",
    )
    observaciones_alta = models.TextField(
        blank=True,
        default="",
        help_text="Observaciones del oficio de alta",
    )

    # ---- Seguimiento ----
    fecha_seguimiento = models.DateField(
        null=True,
        blank=True,
        help_text="Fecha de seguimiento",
    )
    oficio_seguimiento = models.CharField(
        max_length=100,
        blank=True,
        default="",
        help_text="Número de oficio de seguimiento",
    )
    observaciones_seguimiento = models.TextField(
        blank=True,
        default="",
        help_text="Observaciones del oficio de seguimiento",
    )

    # ---- Baja ----
    fecha_baja = models.DateField(
        null=True,
        blank=True,
        help_text="Fecha de baja del enlace",
    )
    oficio_baja = models.CharField(
        max_length=100,
        blank=True,
        default="",
        help_text="Número de oficio de baja",
    )
    observaciones_baja = models.TextField(
        blank=True,
        default="",
        help_text="Observaciones del oficio de baja",
    )

    # ---- Comentarios ----
    comentarios = models.TextField(
        blank=True,
        default="",
        help_text="Observaciones o comentarios sobre el enlace",
    )

    # ---- Timestamps ----
    creado_en = models.DateTimeField(auto_now_add=True)
    actualizado_en = models.DateTimeField(auto_now=True)

    def clean(self):
        """Validaciones adicionales del modelo (ejecutadas por full_clean)."""
        super().clean()
        if self.usuario_sig:
            qs = EnlaceAutorizado.objects.filter(
                usuario_sig__iexact=self.usuario_sig.strip()
            )
            if self.pk:
                qs = qs.exclude(pk=self.pk)
            if qs.exists():
                from django.core.exceptions import ValidationError

                raise ValidationError({
                    "usuario_sig": (
                        f"Ya existe un enlace con el usuario SIG "
                        f"'{self.usuario_sig}'. El usuario SIG no puede duplicarse."
                    )
                })

    def save(self, *args, **kwargs):
        # Nombre SIG = unión de nombres + apellidos (separado por espacios)
        if not self.nombre_sig:
            self.nombre_sig = self.nombre_completo
        # Normalizar correo vacío a None para que el unique no choque con ""
        if self.correo_principal is not None and not str(self.correo_principal).strip():
            self.correo_principal = None
        super().save(*args, **kwargs)

    class Meta:
        ordering = ["primer_apellido", "segundo_apellido", "nombres"]
        verbose_name = "Enlace Autorizado"
        verbose_name_plural = "Enlaces Autorizados"

    def __str__(self):
        return f"{self.nombres} {self.primer_apellido} {self.segundo_apellido}".strip()

    @property
    def nombre_completo(self):
        """Nombre completo formateado."""
        partes = [self.nombres, self.primer_apellido, self.segundo_apellido]
        return " ".join(p for p in partes if p).strip()

    @property
    def esta_activo(self):
        return self.estado == "ACTIVO"

    @property
    def edificio(self):
        """Devuelve el primer edificio de la institución (compatibilidad)."""
        if self.institucion_id:
            ed = self.institucion.edificio.first()
            return ed
        return None

    @property
    def edificios(self):
        """Devuelve todos los edificios de la institución."""
        if self.institucion_id:
            return self.institucion.edificio.all()
        return []

    @property
    def nombre_completo_admin(self):
        """Nombre completo para mostrar en admin."""
        return self.nombre_completo


# ---------------------------------------------------------------------------
# Adjuntos
# ---------------------------------------------------------------------------
def ruta_adjunto(instance, filename):
    """Ruta de almacenamiento: adjuntos/<enlace_id>/<archivo>"""
    return f"adjuntos/{instance.enlace.pk}/{filename}"


class Adjunto(models.Model):
    """Documento adjunto asociado a un EnlaceAutorizado."""

    enlace = models.ForeignKey(
        EnlaceAutorizado,
        on_delete=models.CASCADE,
        related_name="adjuntos",
    )
    archivo = models.FileField(
        upload_to=ruta_adjunto,
        help_text="Archivo adjunto (PDF, imagen, documento, etc.)",
    )
    nombre = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Nombre descriptivo del adjunto",
    )
    descripcion = models.TextField(
        blank=True,
        default="",
        help_text="Descripción opcional del documento",
    )
    subido_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-subido_en"]
        verbose_name = "Adjunto"
        verbose_name_plural = "Adjuntos"

    def __str__(self):
        return self.nombre or self.archivo.name

    @property
    def tamano_legible(self):
        """Devuelve el tamaño del archivo en formato legible."""
        try:
            bytes_val = self.archivo.size
        except Exception:
            return "—"
        if bytes_val < 1024:
            return f"{bytes_val} B"
        elif bytes_val < 1024 * 1024:
            return f"{bytes_val / 1024:.1f} KB"
        return f"{bytes_val / (1024 * 1024):.1f} MB"

    @property
    def icono(self):
        """Icono FA según tipo de archivo."""
        nombre = (self.archivo.name or "").lower()
        if nombre.endswith((".pdf",)):
            return "fas fa-file-pdf"
        if nombre.endswith((".doc", ".docx", ".odt")):
            return "fas fa-file-word"
        if nombre.endswith((".xls", ".xlsx", ".ods")):
            return "fas fa-file-excel"
        if nombre.endswith((".jpg", ".jpeg", ".png", ".gif", ".webp")):
            return "fas fa-file-image"
        if nombre.endswith((".mp4", ".avi", ".mov")):
            return "fas fa-file-video"
        return "fas fa-file"


# ---------------------------------------------------------------------------
# SyncLog — Registro de sincronizaciones con el SIG
# ---------------------------------------------------------------------------
class SyncLog(models.Model):
    """Registra cada intento de sincronización con el SIG."""

    ACTION_CHOICES = [
        ("crear", "Crear usuario"),
        ("deshabilitar", "Deshabilitar usuario"),
        ("reactivar", "Reactivar usuario"),
        ("actualizar_correo", "Actualizar correo"),
    ]
    STATUS_CHOICES = [
        ("exitoso", "Exitoso"),
        ("error", "Error"),
        ("pendiente", "Pendiente"),
    ]

    enlace = models.ForeignKey(
        EnlaceAutorizado,
        on_delete=models.CASCADE,
        related_name="sync_logs",
    )
    accion = models.CharField(max_length=20, choices=ACTION_CHOICES)
    estado = models.CharField(
        max_length=20, choices=STATUS_CHOICES, default="pendiente",
    )
    mensaje = models.TextField(blank=True, default="")
    usuario_sig_id = models.CharField(
        max_length=20, blank=True, default="",
        help_text="ID del usuario en el SIG (informativo; la creación no depende de él)",
    )
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-creado_en"]
        verbose_name = "Log de sincronización SIG"
        verbose_name_plural = "Logs de sincronización SIG"

    def __str__(self):
        return f"{self.get_accion_display()} → {self.get_estado_display()} ({self.creado_en:%d/%m/%Y %H:%M})"


class ConfiguracionCorreo(models.Model):
    """Configuración del envío de correo, editable desde el admin.

    Modelo singleton: solo debe existir UN registro (se gestiona con
    ConfiguracionCorreoAdmin.get_solo o el patrón get_or_create).
    """

    # --- Servidor SMTP ---
    smtp_host = models.CharField(
        max_length=200,
        blank=True,
        default="smtp.office365.com",
        help_text="Servidor SMTP (ej. smtp.office365.com).",
        verbose_name="Servidor SMTP",
    )
    smtp_port = models.PositiveIntegerField(
        blank=True,
        null=True,
        default=587,
        help_text="Puerto SMTP (ej. 587 para TLS).",
        verbose_name="Puerto SMTP",
    )
    smtp_use_tls = models.BooleanField(
        default=True,
        help_text="Usar TLS (STARTTLS) para la conexión.",
        verbose_name="Usar TLS",
    )

    # --- Cuenta emisora ---
    correo_emisor = models.EmailField(
        blank=True,
        default="",
        help_text="Cuenta SMTP que envía los correos (usuario/remitente).",
        verbose_name="Correo emisor",
    )
    password_emisor = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Contraseña de aplicación de la cuenta emisora.",
        verbose_name="Contraseña emisor",
    )
    default_from_email = models.EmailField(
        blank=True,
        default="",
        help_text="Dirección From por defecto (opcional, si difiere del emisor).",
        verbose_name="From por defecto",
    )

    # --- Destinatarios ---
    correo_review = models.CharField(
        max_length=500,
        blank=True,
        default="",
        help_text=(
            "Correo(s) de revisión para creación manual de usuarios SIG "
            "cuando falla la automatización. Separe varios con ';'."
        ),
        verbose_name="Correo(s) de revisión",
    )
    correo_notificacion = models.CharField(
        max_length=500,
        blank=True,
        default="",
        help_text=(
            "Correo(s) de notificación con credenciales cuando un usuario "
            "SIG se crea exitosamente. Separe varios con ';'."
        ),
        verbose_name="Correo(s) de notificación",
    )

    actualizado_en = models.DateTimeField(auto_now=True)
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Configuración de correo"
        verbose_name_plural = "Configuración de correo"

    def __str__(self):
        return "Configuración de correo"

    # --- Helpers sin fallback a settings ---
    def smtp_host_valor(self) -> str:
        return (self.smtp_host or "").strip() or "smtp.office365.com"

    def smtp_port_valor(self) -> int:
        return self.smtp_port or 587

    def smtp_use_tls_valor(self) -> bool:
        return self.smtp_use_tls

    def correo_emisor_valor(self) -> str:
        return (self.correo_emisor or "").strip()

    def password_emisor_valor(self) -> str:
        return (self.password_emisor or "").strip()

    def default_from_email_valor(self) -> str:
        return (self.default_from_email or "").strip()

    def review_lista(self) -> list[str]:
        """Lista de correos de revisión (separados por ';')."""
        valor = (self.correo_review or "").strip()
        if not valor:
            return []
        return [c.strip() for c in valor.split(";") if c.strip()]

    def notificacion_lista(self) -> list[str]:
        """Lista de correos de notificación (separados por ';')."""
        valor = (self.correo_notificacion or "").strip()
        if not valor:
            return []
        return [c.strip() for c in valor.split(";") if c.strip()]

    @classmethod
    def cargar(cls):
        """Devuelve el registro singleton de configuración o uno vacío."""
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj


class PermisoGrupo(models.Model):
    """Capacidades por grupo de usuario, administrables desde la app.

    La lista de capacidades se guarda como JSON y reemplaza (con respaldo)
    a la definición estática de `roles.CAPACIDADES_POR_ROL`.
    """

    grupo = models.OneToOneField(
        "auth.Group",
        on_delete=models.CASCADE,
        related_name="permiso_grupo",
        verbose_name="Grupo",
    )
    capacidades = models.JSONField(
        default=list,
        blank=True,
        verbose_name="Capacidades",
        help_text="Lista de claves de capacidad (directorio, editar, importar, tickets, revisiones, admin).",
    )

    class Meta:
        verbose_name = "Permisos de grupo"
        verbose_name_plural = "Permisos de grupos"

    def __str__(self):
        return f"Permisos de {self.grupo.name}"


class ConfiguracionSIG(models.Model):
    """Credenciales/parámetros del SIG para sincronización (Playwright).

    Modelo singleton: solo debe existir UN registro (ver ConfiguracionSIGAdmin).
    La configuración se gestiona exclusivamente desde el admin UI.
    """

    url = models.CharField(
        max_length=500,
        blank=True,
        default="",
        help_text="URL base del SIG (p. ej. https://sig.gia.mx/webapp/).",
        verbose_name="URL del SIG",
    )
    usuario = models.CharField(
        max_length=150,
        blank=True,
        default="",
        help_text="Usuario para el login automatizado del SIG.",
        verbose_name="Usuario",
    )
    password = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Contraseña para el login automatizado del SIG.",
        verbose_name="Contraseña",
    )
    default_password = models.CharField(
        max_length=100,
        blank=True,
        default="",
        help_text="Contraseña por defecto para usuarios SIG creados.",
        verbose_name="Contraseña por defecto",
    )
    timeout = models.PositiveIntegerField(
        blank=True,
        null=True,
        default=30,
        help_text="Tiempo de espera (segundos) para Playwright.",
        verbose_name="Timeout (segundos)",
    )
    actualizado_en = models.DateTimeField(auto_now=True)
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Configuración SIG"
        verbose_name_plural = "Configuración SIG"

    def __str__(self):
        return "Configuración SIG"

    # --- Helpers sin fallback a settings ---
    def url_valor(self) -> str:
        return (self.url or "").strip()

    def usuario_valor(self) -> str:
        return (self.usuario or "").strip()

    def password_valor(self) -> str:
        return (self.password or "").strip()

    def default_password_valor(self) -> str:
        return (self.default_password or "").strip()

    def timeout_valor(self) -> int:
        return self.timeout or 30

    @classmethod
    def cargar(cls):
        """Devuelve el registro singleton de configuración o uno vacío."""
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj


class ConfiguracionTickets(models.Model):
    """Parámetros del reporte de tickets del SIG (descarga y sincronización).

    Modelo singleton: solo debe existir UN registro (ver ConfiguracionTicketsAdmin).
    La cuenta SIG del worker de tickets NO vive aquí: se lee del entorno
    (SIG_TICKETS_URL / SIG_TICKETS_USUARIO / SIG_TICKETS_PASSWORD).
    """

    MODO_PARCIAL = "parcial"
    MODO_COMPLETO = "completo"
    MODO_CHOICES = (
        (MODO_PARCIAL, "Parcial (rápida)"),
        (MODO_COMPLETO, "Completo (histórico)"),
    )

    # --- Descarga (F1) ---
    url_reporte = models.CharField(
        max_length=500,
        blank=True,
        default="/admin/Solicitud/SeguimientoAtencion",
        help_text="Ruta del reporte de seguimiento de solicitudes en el SIG.",
        verbose_name="URL del reporte",
    )
    hoja_excel = models.CharField(
        max_length=100,
        blank=True,
        default="",
        help_text="Nombre de la hoja del Excel. Vacío = auto-detectar la primera.",
        verbose_name="Hoja del Excel",
    )
    fecha_desde_completa = models.CharField(
        max_length=20,
        blank=True,
        default="01/01/2020",
        help_text=(
            "Fecha 'desde' para la descarga completa (dd/mm/aaaa). El SIG asume "
            "que 'hasta' es hoy. En la descarga rápida no se usa."
        ),
        verbose_name="Fecha desde (completa)",
    )

    # --- Sincronización (F2) ---
    habilitado = models.BooleanField(
        default=True,
        help_text="Si desmarcado, el tick automático no descarga ni sincroniza.",
        verbose_name="Sincronización habilitada",
    )
    modo_default = models.CharField(
        max_length=10,
        choices=MODO_CHOICES,
        default=MODO_PARCIAL,
        help_text="Modo usado por el tick y por 'Sincronizar ahora' sin indicar modo.",
        verbose_name="Modo por defecto",
    )
    intervalo_minutos = models.PositiveSmallIntegerField(
        default=5,
        help_text="Cadencia mínima (minutos) entre sincronizaciones parciales.",
        verbose_name="Intervalo parcial (minutos)",
    )
    descarga_completa_horas = models.JSONField(
        default=list,
        blank=True,
        help_text="Horas del día para la descarga completa, ej. [\"07:00\", \"19:00\"].",
        verbose_name="Horas de descarga completa",
    )
    hora_inicio = models.TimeField(
        null=True,
        blank=True,
        verbose_name="Hora inicio (parcial)",
        help_text=(
            "Inicio del horario en que corre la sincronización parcial (hora local). "
            "Vacío = sin límite. La descarga completa no usa este horario."
        ),
    )
    hora_fin = models.TimeField(
        null=True,
        blank=True,
        verbose_name="Hora fin (parcial)",
        help_text=(
            "Fin del horario en que corre la sincronización parcial (hora local). "
            "Vacío = sin límite."
        ),
    )
    dias_semana = models.JSONField(
        default=list,
        blank=True,
        verbose_name="Días permitidos (parcial)",
        help_text=(
            "Días en que corre la sincronización parcial: 0=lunes … 6=domingo. "
            "Vacío = todos los días."
        ),
    )

    # --- Última sincronización (información) ---
    ultima_sincronizacion = models.DateTimeField(
        null=True,
        blank=True,
        editable=False,
        verbose_name="Última sincronización",
        help_text="Última corrida EXITOSA. No se actualiza si la sincronización falla.",
    )
    ultimo_intento = models.DateTimeField(
        null=True,
        blank=True,
        editable=False,
        verbose_name="Último intento",
        help_text="Última corrida (exitosa o fallida). Se usa para el intervalo.",
    )
    ultimo_modo = models.CharField(
        max_length=10, blank=True, default="", editable=False
    )
    ultimo_estado = models.CharField(
        max_length=20,
        blank=True,
        default="",
        editable=False,
        verbose_name="Último estado",
    )
    ultimo_mensaje = models.TextField(
        blank=True, default="", editable=False, verbose_name="Último mensaje"
    )
    total_tickets = models.PositiveIntegerField(
        default=0, editable=False, verbose_name="Total de tickets"
    )

    actualizado_en = models.DateTimeField(auto_now=True)
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Configuración de tickets"
        verbose_name_plural = "Configuración de tickets"

    def __str__(self):
        return "Configuración de tickets"

    def url_reporte_valor(self) -> str:
        return (self.url_reporte or "/admin/Solicitud/SeguimientoAtencion").strip()

    def hoja_valor(self) -> str:
        return (self.hoja_excel or "").strip()

    def fecha_desde_valor(self) -> str:
        return (self.fecha_desde_completa or "01/01/2020").strip()

    def modo_default_valor(self) -> str:
        return self.modo_default or self.MODO_PARCIAL

    @staticmethod
    def _siguiente_hora(horas, ahora):
        """Próxima ocurrencia (datetime) de una hora "HH:MM" del día (hora local)."""
        tz = timezone.get_current_timezone()
        fecha_local = timezone.localtime(ahora).date()
        candidatas = []
        if not isinstance(horas, list):
            return candidatas
        for hora in horas:
            try:
                hh, mm = (int(x) for x in str(hora).split(":")[:2])
                candidata = timezone.make_aware(
                    datetime.combine(fecha_local, time(hh, mm)), tz
                )
            except (ValueError, TypeError):
                continue
            if candidata <= ahora:
                candidata += timedelta(days=1)
            candidatas.append(candidata)
        return candidatas

    def en_horario(self, ahora=None) -> bool:
        """True si la hora local cae dentro del horario/días configurados.

        Sin `hora_inicio`/`hora_fin` devuelve True (sin límite). Soporta
        ventanas que cruzan medianoche (p. ej. 22:00 → 06:00).
        """
        if self.hora_inicio is None or self.hora_fin is None:
            return True
        if ahora is None:
            ahora = timezone.now()
        local = timezone.localtime(ahora)
        if self.dias_semana and local.weekday() not in self.dias_semana:
            return False
        t = local.time()
        if self.hora_inicio <= self.hora_fin:
            return self.hora_inicio <= t < self.hora_fin
        return t >= self.hora_inicio or t < self.hora_fin

    def proximo_inicio_horario(self, ahora=None):
        """Próximo datetime (aware) en que arranca la ventana de sincronización."""
        if ahora is None:
            ahora = timezone.now()
        if self.hora_inicio is None or self.hora_fin is None:
            return ahora
        tz = timezone.get_current_timezone()
        local = timezone.localtime(ahora)
        for delta in range(0, 8):
            fecha = (local + timedelta(days=delta)).date()
            if self.dias_semana and fecha.weekday() not in self.dias_semana:
                continue
            candidata = timezone.make_aware(
                datetime.combine(fecha, self.hora_inicio), tz
            )
            if candidata > ahora:
                return candidata
        return ahora

    def proxima_sincronizacion(self, ahora=None):
        """Momento previsto de la próxima sincronización automática (o None).

        Refleja la lógica de tick_sync_tickets_task: descarga completa cuando
        coincida una hora de descarga_completa_horas (sin horario); si no,
        parcial cada intervalo_minutos, solo dentro del horario configurado.
        """
        if not self.habilitado:
            return None
        if ahora is None:
            ahora = timezone.now()
        candidatas = []
        if self.en_horario(ahora):
            referencia = self.ultimo_intento or self.ultima_sincronizacion
            if referencia:
                parcial = referencia + timedelta(
                    minutes=self.intervalo_minutos or 5
                )
            else:
                parcial = ahora
            # Si el parcial ya venció, la próxima es inmediata.
            if parcial < ahora:
                parcial = ahora
            candidatas.append(parcial)
        else:
            candidatas.append(self.proximo_inicio_horario(ahora))
        # Las horas de descarga completa no dependen del horario.
        candidatas.extend(self._siguiente_hora(self.descarga_completa_horas, ahora))
        return min(candidatas) if candidatas else ahora

    @classmethod
    def cargar(cls):
        """Devuelve el registro singleton de configuración o uno vacío."""
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj


class Servicio(models.Model):
    """Catálogo de servicios de atención (columna `servicio` del SIG)."""

    nombre = models.CharField(max_length=255, unique=True, verbose_name="Servicio")
    activo = models.BooleanField(default=True, verbose_name="Activo")
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["nombre"]
        verbose_name = "Servicio"
        verbose_name_plural = "Servicios"

    def __str__(self):
        return self.nombre


class Falla(models.Model):
    """Catálogo de fallas (columna `falla_descripcion` del SIG)."""

    descripcion = models.CharField(
        max_length=400, unique=True, verbose_name="Descripción"
    )
    activo = models.BooleanField(default=True, verbose_name="Activo")
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["descripcion"]
        verbose_name = "Falla"
        verbose_name_plural = "Fallas"

    def __str__(self):
        return self.descripcion


class Nivel(models.Model):
    """Catálogo de niveles reportados en el SIG (columna `grupo`, ej. Nivel 8)."""

    nombre = models.CharField(max_length=60, unique=True, verbose_name="Nivel")
    activo = models.BooleanField(default=True, verbose_name="Activo")
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["nombre"]
        verbose_name = "Nivel"
        verbose_name_plural = "Niveles"

    def __str__(self):
        return self.nombre


class ResponsableAtencion(models.Model):
    """Catálogo de responsables de atención reportados en el SIG."""

    nombre = models.CharField(max_length=255, unique=True, verbose_name="Nombre")
    activo = models.BooleanField(default=True, verbose_name="Activo")
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["nombre"]
        verbose_name = "Responsable de atención"
        verbose_name_plural = "Responsables de atención"

    def __str__(self):
        return self.nombre


class Ticket(models.Model):
    """Solicitud / ticket de seguimiento del SIG (snapshot sincronizado).

    Estatus: se deriva en cada sincronización de la columna `cerro_fecha`
    (ABIERTO si no tiene fecha de cierre, CERRADO si la tiene). En F4 las
    acciones de cierre lo cambian localmente aunque el SIG no lo sepa aún.
    """

    ESTATUS_ABIERTO = "ABIERTO"
    ESTATUS_CERRADO = "CERRADO"
    ESTATUS_CHOICES = (
        (ESTATUS_ABIERTO, "Abierto"),
        (ESTATUS_CERRADO, "Cerrado"),
    )

    ticket_id = models.CharField(
        max_length=100,
        unique=True,
        help_text="Clave única del SIG: columna 'ID Solicitud servicio'.",
        verbose_name="ID de solicitud",
    )
    numero = models.CharField(
        max_length=60,
        blank=True,
        default="",
        help_text="Número visible en el SIG (ej. SS20-0459), sin sufijo.",
        verbose_name="Número",
    )
    numero_display = models.CharField(
        max_length=60,
        blank=True,
        default="",
        db_index=True,
        help_text="Número visible con sufijo -N para duplicados (SS20-0459-2).",
        verbose_name="Número visible",
    )
    solicitud_tipo = models.CharField(
        max_length=100, blank=True, default="", verbose_name="Tipo de solicitud"
    )
    fecha = models.DateTimeField(
        null=True, blank=True, db_index=True, verbose_name="Fecha"
    )
    fecha_cierre = models.DateTimeField(
        null=True, blank=True, verbose_name="Fecha de cierre"
    )
    estatus = models.CharField(
        max_length=8,
        choices=ESTATUS_CHOICES,
        default=ESTATUS_ABIERTO,
        db_index=True,
        verbose_name="Estatus",
    )
    descripcion = models.TextField(
        blank=True, default="", verbose_name="Descripción / asunto"
    )
    # ---- Vínculos con el directorio (Fase 3) ----
    torre = models.ForeignKey(
        Edificio,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="tickets",
        verbose_name="Edificio / torre",
    )
    institucion = models.ForeignKey(
        Institucion,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="tickets",
        verbose_name="Institución",
    )
    solicitante = models.ForeignKey(
        EnlaceAutorizado,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="tickets",
        verbose_name="Solicitante",
    )
    servicio = models.ForeignKey(
        Servicio,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="tickets",
        verbose_name="Servicio",
    )
    falla = models.ForeignKey(
        Falla,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="tickets",
        verbose_name="Falla",
    )
    nivel = models.ForeignKey(
        Nivel,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="tickets",
        verbose_name="Nivel",
        help_text="Nivel del reporte (columna grupo del SIG, ej. Nivel 8).",
    )
    responsable_atencion = models.ForeignKey(
        ResponsableAtencion,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="tickets",
        verbose_name="Responsable de atención",
    )
    solicitante_nombre = models.CharField(
        max_length=200,
        blank=True,
        default="",
        verbose_name="Solicitante (nombre SIG)",
        help_text="Nombre tal cual viene del SIG cuando no empata con un enlace.",
    )
    archivado = models.BooleanField(
        default=False, db_index=True, verbose_name="Archivado"
    )
    ultima_sync = models.DateTimeField(
        null=True, blank=True, verbose_name="Última sincronización"
    )
    raw_data = models.JSONField(
        default=dict, blank=True, verbose_name="Datos crudos del SIG"
    )
    creado_en = models.DateTimeField(auto_now_add=True)
    actualizado_en = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Ticket"
        verbose_name_plural = "Tickets"
        ordering = ("-fecha", "ticket_id")
        indexes = [
            models.Index(fields=("estatus", "fecha"), name="tick_estatus_fecha"),
            models.Index(fields=("numero", "ticket_id"), name="tick_numero_id"),
        ]

    def __str__(self):
        return self.numero_display or self.ticket_id

    def fases_proceso(self):
        """Fases del proceso de atención (timeline del detalle)."""
        r = self.raw_data or {}
        servicio = str(r.get("servicio") or "").strip()
        responsable = str(r.get("Responsable_atencion") or "").strip()
        numero = self.numero_display or self.numero or self.ticket_id
        cierre_inst = getattr(self, "cierre", None)
        actividades = str(
            (cierre_inst.actividades if cierre_inst and cierre_inst.actividades else "")
            or (r.get("Actividades") or "")
        ).strip()
        cerrado = self.esta_cerrado
        pendiente_sig = cierre_inst is not None and not self.cierre_sincronizado
        seguimiento_pendiente = pendiente_sig and not cerrado
        fecha_cierre_mostrada = (
            cierre_inst.fecha_cierre
            if cierre_inst and cierre_inst.fecha_cierre
            else self.fecha_cierre
        )
        detalle_cierre = (
            timezone.localtime(fecha_cierre_mostrada).strftime("%d/%m/%Y %H:%M")
            if fecha_cierre_mostrada
            else ""
        )
        if self.fecha:
            fecha_local = timezone.localtime(self.fecha)
            detalle_creado = (
                f"Ticket {numero} creado el {fecha_local.strftime('%d/%m/%Y')} "
                f"a las {fecha_local.strftime('%H:%M')}"
            )
        else:
            detalle_creado = numero
        return [
            {
                "clave": "creado",
                "nombre": "Ticket creado",
                "estado": "completado",
                "detalle": detalle_creado,
            },
            {
                "clave": "canalizado",
                "nombre": "Canalizado con el servicio",
                "estado": "completado" if servicio else "pendiente",
                "detalle": responsable or "Sin responsable de atención",
            },
            {
                "clave": "proceso",
                "nombre": "En proceso de atención",
                "estado": "completado" if cerrado else "activo",
                "pendiente_sig": seguimiento_pendiente,
                "detalle": (
                    "Atención terminada"
                    if cerrado
                    else (
                        "Trabajo en curso · Seguimiento pendiente de "
                        "sincronizar en el SIG"
                        if seguimiento_pendiente
                        else "Trabajo en curso"
                    )
                ),
            },
            {
                "clave": "finalizado",
                "nombre": "Finalizado",
                "estado": "completado" if cerrado else "pendiente",
                "pendiente_sig": pendiente_sig,
                "detalle": (
                    f"Cerrado el {detalle_cierre}"
                    + (
                        " · Pendiente de sincronizar en el SIG"
                        if pendiente_sig
                        else ""
                    )
                    if cerrado
                    else "Pendiente de cierre"
                ),
            },
            {
                "clave": "completado",
                "nombre": "Completado",
                "estado": "completado" if actividades else "pendiente",
                "detalle": (
                    actividades[:200] if actividades else "Sin actividades registradas"
                ),
            },
        ]

    def accion_atencion(self):
        """Etiqueta del botón principal según la fase en la que esté el ticket.

        - En proceso de atención → "Atender"
        - Finalizado (cerrado sin actividades) → "Completar"
        - Completado (con actividades) → "Modificar"
        """
        fases = {f["clave"]: f for f in self.fases_proceso()}
        if fases.get("proceso", {}).get("estado") == "activo":
            return "Atender"
        if (
            fases.get("finalizado", {}).get("estado") == "completado"
            and fases.get("completado", {}).get("estado") != "completado"
        ):
            return "Completar"
        return "Modificar"

    def comentarios_servicio(self):
        """Observaciones del SIG y una respuesta sugerida para el enlace.

        La sugerencia es el mensaje que el operador puede enviar al enlace
        (solicitante) para informarle cómo va su reporte y en qué estado está.
        """
        r = self.raw_data or {}
        observaciones = str(r.get("Observaciones") or "").strip()
        observaciones_usuario = str(r.get("ObservacionesUsuario") or "").strip()
        actividades = str(r.get("Actividades") or "").strip()
        solicitante = str(r.get("solicitud_solicitante") or "").strip()
        responsable = str(r.get("Responsable_atencion") or "").strip()
        numero = self.numero_display or self.numero or self.ticket_id
        saludo = f"Estimado(a) {solicitante}:" if solicitante else "Estimado(a):"
        if self.estatus == self.ESTATUS_CERRADO:
            sugerencia = (
                f"{saludo} su reporte {numero} ha sido finalizado. "
                "Agradecemos su paciencia; si el inconveniente persiste puede "
                "generar un nuevo reporte."
            )
        elif actividades:
            sugerencia = (
                f"{saludo} su reporte {numero} ya cuenta con actividades registradas "
                "por el servicio de atención. Estamos gestionando el cierre y le "
                "notificaremos en cuanto finalice."
            )
        else:
            responsable_txt = (
                f" con el responsable de atención {responsable}" if responsable else ""
            )
            sugerencia = (
                f"{saludo} su reporte {numero} fue recibido y canalizado"
                f"{responsable_txt}. Se encuentra en proceso de atención; le "
                "informaremos en cuanto haya avances."
            )
        return {
            "observaciones": observaciones,
            "observaciones_usuario": observaciones_usuario,
            "sugerencia": sugerencia,
        }

    @property
    def cierre_sincronizado(self):
        """True si el cierre local ya coincide con lo reportado por el SIG.

        Análogo a `EnlaceAutorizado.sincronizado`: el cierre quedó capturado
        localmente (el ticket se muestra cerrado), pero para considerarlo
        actualizado en el SIG la fecha de cierre del sistema debe coincidir
        (por día) con el `cerro_fecha` del SIG y los campos Diagnóstico,
        Actividades y Observaciones deben ser idénticos a los reportados.
        """
        cierre = getattr(self, "cierre", None)
        if cierre is None:
            return True
        r = self.raw_data or {}
        sig_fecha = self._fecha_sig_raw(r.get("cerro_fecha"))
        local_fecha = (
            timezone.localtime(cierre.fecha_cierre).date()
            if cierre.fecha_cierre
            else None
        )
        if not sig_fecha or not local_fecha or sig_fecha != local_fecha:
            return False
        norm = lambda v: " ".join(str(v or "").split())
        return all(
            (
                norm(r.get("Diagnostico")) == norm(cierre.diagnostico),
                norm(r.get("Actividades")) == norm(cierre.actividades),
                norm(r.get("Observaciones")) == norm(cierre.observaciones),
            )
        )

    @property
    def esta_cerrado(self):
        """El ticket se muestra cerrado si el SIG lo cerró o hay un cierre local
        con fecha de cierre, aunque el SIG aún no lo haya confirmado.

        Un seguimiento (registro local sin fecha de cierre) no cierra el ticket.
        """
        cierre = getattr(self, "cierre", None)
        if cierre is not None and cierre.fecha_cierre is not None:
            return True
        return self.estatus == self.ESTATUS_CERRADO

    @staticmethod
    def _fecha_sig_raw(valor):
        """Normaliza el valor crudo de `cerro_fecha` del SIG a una fecha."""
        from datetime import date, datetime

        if isinstance(valor, datetime):
            return valor.date()
        if isinstance(valor, date):
            return valor
        if valor:
            try:
                return datetime.strptime(str(valor)[:10], "%Y-%m-%d").date()
            except ValueError:
                pass
        return None


class TicketLog(models.Model):
    """Bitácora de cada corrida de sincronización de tickets."""

    MODO_PARCIAL = "parcial"
    MODO_COMPLETO = "completo"
    ESTADO_OK = "OK"
    ESTADO_ERROR = "ERROR"
    ESTADO_SALTADO = "SALTADO"

    modo = models.CharField(
        max_length=10,
        choices=ConfiguracionTickets.MODO_CHOICES,
        default=MODO_PARCIAL,
        verbose_name="Modo",
    )
    estado = models.CharField(
        max_length=10, blank=True, default="", verbose_name="Estado"
    )
    mensaje = models.TextField(blank=True, default="", verbose_name="Mensaje")
    creados = models.PositiveIntegerField(default=0, verbose_name="Creados")
    actualizados = models.PositiveIntegerField(
        default=0, verbose_name="Actualizados"
    )
    archivados = models.PositiveIntegerField(default=0, verbose_name="Archivados")
    errores = models.PositiveIntegerField(default=0, verbose_name="Errores")
    creado_en = models.DateTimeField(auto_now_add=True, verbose_name="Fecha")

    class Meta:
        verbose_name = "Bitácora de tickets"
        verbose_name_plural = "Bitácora de tickets"
        ordering = ("-creado_en",)

    def __str__(self):
        return f"TicketLog #{self.pk} [{self.modo} {self.estado}]"


class TicketRegistro(models.Model):
    """Consulta de seguimiento capturada por el operador en el detalle.

    Se crea un registro cada vez que un usuario consulta el estado y marca
    la casilla "Registrar" con una breve descripción de lo que pasó.
    """

    ticket = models.ForeignKey(
        Ticket,
        on_delete=models.CASCADE,
        related_name="registros",
        verbose_name="Ticket",
    )
    usuario = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        related_name="registros_tickets",
        verbose_name="Usuario",
    )
    descripcion = models.CharField(
        max_length=500,
        blank=True,
        default="",
        verbose_name="Descripción de lo sucedido",
    )
    creado_en = models.DateTimeField(auto_now_add=True, verbose_name="Fecha")

    class Meta:
        verbose_name = "Registro de seguimiento"
        verbose_name_plural = "Registros de seguimiento"
        ordering = ("-creado_en",)

    def __str__(self):
        return f"{self.ticket_id} · {self.creado_en:%d/%m/%Y %H:%M}"


# ---------------------------------------------------------------------------
# Cierre de tickets (captura local, se valida contra el SIG)
# ---------------------------------------------------------------------------
def ruta_ticket_adjunto(instance, filename):
    """Ruta de almacenamiento: tickets_adjuntos/<ticket_id>/<archivo>"""
    return f"tickets_adjuntos/{instance.ticket_id}/{filename}"


class TicketCierre(models.Model):
    """Cierre de un ticket capturado localmente (página de cierre).

    Guarda los datos que después se reflejan en el SIG (Diagnóstico,
    Actividades, Observaciones y fechas). Se muestra como CERRADO aunque el
    SIG aún no lo confirme; `Ticket.cierre_sincronizado` indica si el SIG ya
    reporta los mismos valores.
    """

    ticket = models.OneToOneField(
        Ticket,
        on_delete=models.CASCADE,
        related_name="cierre",
        verbose_name="Ticket",
    )
    fecha_inicio = models.DateTimeField(
        null=True, blank=True, verbose_name="Fecha de inicio de atención"
    )
    fecha_cierre = models.DateTimeField(
        null=True, blank=True, verbose_name="Fecha de cierre"
    )
    diagnostico = models.TextField(
        blank=True, default="", verbose_name="Diagnóstico"
    )
    actividades = models.TextField(
        blank=True, default="", verbose_name="Actividades realizadas"
    )
    observaciones = models.TextField(
        blank=True, default="", verbose_name="Observaciones"
    )
    observaciones_usuario = models.TextField(
        blank=True, default="", verbose_name="Observaciones del usuario"
    )
    creado_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        related_name="cierres_creados",
        verbose_name="Creado por",
    )
    actualizado_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        related_name="cierres_actualizados",
        verbose_name="Actualizado por",
    )
    creado_en = models.DateTimeField(auto_now_add=True, verbose_name="Creado en")
    actualizado_en = models.DateTimeField(auto_now=True, verbose_name="Actualizado en")

    class Meta:
        verbose_name = "Cierre de ticket"
        verbose_name_plural = "Cierres de ticket"

    def __str__(self):
        return f"Cierre de {self.ticket_id}"


class TicketAdjunto(models.Model):
    """Adjunto del cierre de un ticket."""

    ticket = models.ForeignKey(
        Ticket,
        on_delete=models.CASCADE,
        related_name="cierre_adjuntos",
        verbose_name="Ticket",
    )
    archivo = models.FileField(
        upload_to=ruta_ticket_adjunto,
        help_text="Archivo adjunto (PDF, imagen, documento, etc.)",
        verbose_name="Archivo",
    )
    nombre = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Nombre descriptivo del adjunto",
        verbose_name="Nombre",
    )
    subido_en = models.DateTimeField(auto_now_add=True, verbose_name="Subido en")
    subido_por = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        related_name="adjuntos_tickets",
        verbose_name="Subido por",
    )

    class Meta:
        ordering = ("-subido_en",)
        verbose_name = "Adjunto de cierre"
        verbose_name_plural = "Adjuntos de cierre"

    def __str__(self):
        return self.nombre or self.archivo.name

    @property
    def tamano_legible(self):
        """Devuelve el tamaño del archivo en formato legible."""
        try:
            bytes_val = self.archivo.size
        except Exception:
            return "—"
        if bytes_val < 1024:
            return f"{bytes_val} B"
        elif bytes_val < 1024 * 1024:
            return f"{bytes_val / 1024:.1f} KB"
        return f"{bytes_val / (1024 * 1024):.1f} MB"

    @property
    def icono(self):
        """Icono FA según tipo de archivo."""
        nombre = (self.archivo.name or "").lower()
        if nombre.endswith((".pdf",)):
            return "fas fa-file-pdf"
        if nombre.endswith((".doc", ".docx", ".odt")):
            return "fas fa-file-word"
        if nombre.endswith((".xls", ".xlsx", ".ods")):
            return "fas fa-file-excel"
        if nombre.endswith((".jpg", ".jpeg", ".png", ".gif", ".webp")):
            return "fas fa-file-image"
        if nombre.endswith((".mp4", ".avi", ".mov")):
            return "fas fa-file-video"
        return "fas fa-file"


# ---------------------------------------------------------------------------
# Documentos y carpetas (documentos compartidos / adjuntos de correo)
# ---------------------------------------------------------------------------
def ruta_documento(instance, filename):
    """Ruta de almacenamiento: documentos/<carpeta_slug>/<archivo>"""
    return f"documentos/{instance.carpeta.slug}/{filename}"


class DocumentoCarpeta(models.Model):
    """Carpeta que agrupa documentos (visibles para el directorio).

    La carpeta con slug `adjuntos-nuevos-enlaces` (nombre "Adjuntos:
    Nuevos Enlaces MAO") se adjunta automáticamente al correo de
    creación de usuarios del SIG.
    """

    slug = models.SlugField(
        max_length=100,
        unique=True,
        help_text="Identificador corto y único (solo letras, números, guiones).",
        verbose_name="Identificador (slug)",
    )
    nombre = models.CharField(
        max_length=200,
        help_text="Nombre visible de la carpeta.",
        verbose_name="Nombre",
    )
    descripcion = models.TextField(
        blank=True,
        default="",
        help_text="Descripción opcional de la carpeta.",
        verbose_name="Descripción",
    )
    seccion = models.ForeignKey(
        "Seccion",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="carpetas",
        verbose_name="Sección",
        help_text="Cada carpeta usa una sola sección del catálogo (una sección puede usarse en varias carpetas).",
    )
    creado_en = models.DateTimeField(auto_now_add=True)
    actualizado_en = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["nombre"]
        verbose_name = "Carpeta de documentos"
        verbose_name_plural = "Carpetas de documentos"

    def __str__(self):
        return self.nombre


class Seccion(models.Model):
    """Catálogo global de secciones de documentos.

    Las secciones se crean una sola vez y pueden usarse en varias carpetas.
    """

    nombre = models.CharField(
        max_length=200,
        unique=True,
        help_text="Nombre de la sección (catálogo compartido entre carpetas).",
        verbose_name="Nombre",
    )

    class Meta:
        ordering = ["nombre"]
        verbose_name = "Sección de documentos"
        verbose_name_plural = "Secciones de documentos"

    def __str__(self):
        return self.nombre


class Documento(models.Model):
    """Archivo dentro de una carpeta de documentos (DocumentoCarpeta)."""

    carpeta = models.ForeignKey(
        DocumentoCarpeta,
        on_delete=models.CASCADE,
        related_name="documentos",
        verbose_name="Carpeta",
    )
    archivo = models.FileField(
        upload_to=ruta_documento,
        help_text="Archivo (PDF, imagen, documento, etc.).",
        verbose_name="Archivo",
    )
    nombre = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Nombre descriptivo del documento.",
        verbose_name="Nombre",
    )
    subido_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-subido_en"]
        verbose_name = "Documento"
        verbose_name_plural = "Documentos"

    def __str__(self):
        return self.nombre or self.archivo.name

    @property
    def nombre_archivo(self):
        return self.archivo.name.split("/")[-1]

    @property
    def tamano_legible(self):
        """Devuelve el tamaño del archivo en formato legible."""
        try:
            bytes_val = self.archivo.size
        except Exception:
            return "—"
        if bytes_val < 1024:
            return f"{bytes_val} B"
        elif bytes_val < 1024 * 1024:
            return f"{bytes_val / 1024:.1f} KB"
        return f"{bytes_val / (1024 * 1024):.1f} MB"

    @property
    def icono(self):
        """Icono FA según tipo de archivo."""
        nombre = (self.archivo.name or "").lower()
        if nombre.endswith((".pdf",)):
            return "fas fa-file-pdf"
        if nombre.endswith((".doc", ".docx", ".odt")):
            return "fas fa-file-word"
        if nombre.endswith((".xls", ".xlsx", ".ods")):
            return "fas fa-file-excel"
        if nombre.endswith((".jpg", ".jpeg", ".png", ".gif", ".webp")):
            return "fas fa-file-image"
        if nombre.endswith((".mp4", ".avi", ".mov")):
            return "fas fa-file-video"
        return "fas fa-file"


# ---------------------------------------------------------------------------
class ComunicadoCC(models.Model):
    """Persona de contacto para los comunicados del Centro Cívico Gubernamental."""

    nombre = models.CharField(
        max_length=200,
        help_text="Nombre completo de la persona",
    )
    institucion = models.ForeignKey(
        Institucion,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="comunicados_cc",
        help_text="Institución a la que pertenece",
    )
    cargo = models.CharField(
        max_length=120,
        blank=True,
        default="",
        help_text="Cargo dentro de la institución",
    )
    correo = models.EmailField(
        blank=True,
        null=True,
        default="",
        help_text="Correo institucional de contacto",
    )
    creado_en = models.DateTimeField(auto_now_add=True)
    actualizado_en = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["nombre"]
        verbose_name = "Comunicado CC"
        verbose_name_plural = "Comunicados CC"

    def __str__(self):
        return self.nombre
