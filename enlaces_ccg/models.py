"""
Models para la gestión de Enlaces Autorizados del CCG Honduras.

Relación jerárquica:
    Edificio → Institución → EnlaceAutorizado
"""

from django.db import models


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
    imagen = models.ImageField(
        upload_to="edificios/",
        blank=True,
        null=True,
        help_text="Foto o imagen del edificio",
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

    edificio = models.ForeignKey(
        Edificio,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="instituciones",
        help_text="Edificio donde radica la institución",
    )
    nombre = models.CharField(
        max_length=200,
        help_text="Nombre completo de la institución",
    )
    siglas = models.CharField(
        max_length=30,
        blank=True,
        default="",
        help_text="Siglas de la institución (ej. SESAL, SEPLAN)",
    )
    nivel = models.JSONField(
        default=list,
        blank=True,
        help_text="Niveles/pisos dentro del edificio (ej. [\"PB\", \"Nivel 1\", \"Nivel 2\"])",
    )
    # Niveles disponibles por edificio (siglas del edificio -> lista de niveles)
    NIVELES_POR_EDIFICIO = {
        "CBA": ["PB"] + [f"Nivel {i}" for i in range(1, 6)],
        "CBB": ["PB"] + [f"Nivel {i}" for i in range(1, 8)],
        "CBC": ["PB"] + [f"Nivel {i}" for i in range(1, 8)],
        "TORRE 1": ["PB"] + [f"Nivel {i}" for i in range(1, 24)],
        "TORRE 2": ["PB"] + [f"Nivel {i}" for i in range(1, 25)],
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
        """Devuelve el edificio de la institución a la que pertenece."""
        return self.institucion.edificio

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

    correo_review = models.CharField(
        max_length=500,
        blank=True,
        default="",
        help_text=(
            "Correo(s) de revisión para creación manual de usuarios SIG "
            "cuando falla la automatización. Separe varios correos con ';'."
        ),
        verbose_name="Correo(s) de revisión (SIG_REVIEW_EMAIL)",
    )
    correo_notificacion = models.CharField(
        max_length=500,
        blank=True,
        default="",
        help_text=(
            "Correo(s) de notificación con credenciales cuando un usuario "
            "SIG se crea exitosamente. Separe varios correos con ';'."
        ),
        verbose_name="Correo(s) de notificación (SIG_NOTIFY_EMAIL)",
    )
    correo_emisor = models.EmailField(
        blank=True,
        default="",
        help_text="Cuenta SMTP que envía los correos (EMAIL_HOST_USER).",
        verbose_name="Correo emisor (EMAIL_HOST_USER)",
    )
    password_emisor = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Contraseña de aplicación de la cuenta emisora (EMAIL_HOST_PASSWORD).",
        verbose_name="Contraseña emisor (EMAIL_HOST_PASSWORD)",
    )
    actualizado_en = models.DateTimeField(auto_now=True)
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Configuración de correo"
        verbose_name_plural = "Configuración de correo"

    def __str__(self):
        return "Configuración de correo"

    def _aplicar_por_campo(self, campo, fallback):
        valor = (getattr(self, campo) or "").strip()
        return valor if valor else fallback

    def emisor(self, fallback=None):
        return self._aplicar_por_campo("correo_emisor", fallback or "")

    def emisor_password(self, fallback=None):
        return self._aplicar_por_campo("password_emisor", fallback or "")

    def review_lista(self, fallback=None):
        """Lista de correos de revisión (separados por ';')."""
        return self._separar("correo_review", fallback)

    def notificacion_lista(self, fallback=None):
        """Lista de correos de notificación (separados por ';')."""
        return self._separar("correo_notificacion", fallback)

    def _separar(self, campo, fallback):
        valor = (getattr(self, campo) or "").strip()
        if not valor:
            return list(fallback) if fallback else []
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

    Modelo singleton: solo debe existir UN registro (ver
    ConfiguracionSIGAdmin). Cada valor vacío cae al fallback de settings.
    """

    url = models.CharField(
        max_length=500,
        blank=True,
        default="",
        help_text="URL base del SIG (p. ej. https://sig.gia.mx/webapp/).",
        verbose_name="URL del SIG (SIG_URL)",
    )
    usuario = models.CharField(
        max_length=150,
        blank=True,
        default="",
        help_text="Usuario para el login automatizado del SIG (SIG_USER).",
        verbose_name="Usuario (SIG_USER)",
    )
    password = models.CharField(
        max_length=255,
        blank=True,
        default="",
        help_text="Contraseña para el login automatizado del SIG (SIG_PASSWORD).",
        verbose_name="Contraseña (SIG_PASSWORD)",
    )
    timeout = models.PositiveIntegerField(
        blank=True,
        null=True,
        help_text="Tiempo de espera (segundos) para Playwright (SIG_TIMEOUT).",
        verbose_name="Timeout en segundos (SIG_TIMEOUT)",
    )
    actualizado_en = models.DateTimeField(auto_now=True)
    creado_en = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Configuración SIG"
        verbose_name_plural = "Configuración SIG"

    def __str__(self):
        return "Configuración SIG"

    def url_valor(self, fallback=None):
        return (self.url or "").strip() or fallback

    def usuario_valor(self, fallback=None):
        return (self.usuario or "").strip() or fallback

    def password_valor(self, fallback=None):
        return (self.password or "").strip() or fallback

    def timeout_valor(self, fallback=None):
        if self.timeout:
            return self.timeout
        return fallback

    @classmethod
    def cargar(cls):
        """Devuelve el registro singleton de configuración o uno vacío."""
        obj, _ = cls.objects.get_or_create(pk=1)
        return obj
