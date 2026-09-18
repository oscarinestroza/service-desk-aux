"""
Vistas basadas en funciones para la app enlaces_ccg.

Rutas:
    /                           → lista_edificios
    /edificio/<int:pk>/         → detalle_edificio
    /institucion/<int:pk>/      → perfil_institucion
    /institucion/<int:pk>/exportar/ → exportar_enlaces
"""

import json
from io import BytesIO
from datetime import datetime, timedelta

from django.contrib import auth, messages
from django.contrib.auth.decorators import login_required
from django.core.paginator import Paginator
from django.db import models
from django.db.models import Count, F, Q
from django.http import Http404, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from urllib.parse import quote
from django.urls import reverse
from django.views.decorators.http import require_POST
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

from .models import (
    Adjunto, ComunicadoCC, ConfiguracionTickets, Documento, DocumentoCarpeta,
    Edificio,
    EnlaceAutorizado,
    Falla,
    Institucion, InstitucionEdificio, Nivel, ResponsableAtencion, Seccion,
    Servicio, Ticket,
    TicketRegistro,
)
from .roles import (
    CAP_ADMIN,
    CAP_DIRECTORIO,
    CAP_EDITAR,
    CAP_IMPORTAR,
    CAP_REVISIONES,
    CAP_TICKETS,
    requiere,
    tiene,
)


# ---------------------------------------------------------------------------
# Login / Logout
# ---------------------------------------------------------------------------
def login_vista(request):
    if request.user.is_authenticated:
        return redirect("enlaces_ccg:lista_edificios")

    error = None
    if request.method == "POST":
        usuario = request.POST.get("username", "").strip()
        contrasena = request.POST.get("password", "")
        user = auth.authenticate(request, username=usuario, password=contrasena)
        if user is not None:
            auth.login(request, user)
            destino = request.GET.get("next", "")
            return redirect(destino or "enlaces_ccg:lista_edificios")
        error = "Usuario o contraseña incorrectos."

    return render(request, "enlaces_ccg/login.html", {"error": error})


def logout_vista(request):
    auth.logout(request)
    return redirect("enlaces_ccg:login")


# ---------------------------------------------------------------------------
# Colores y estilos reutilizables para Excel
# ---------------------------------------------------------------------------
_HEADER_FILL = PatternFill(start_color="0D1B3E", end_color="0D1B3E", fill_type="solid")
_HEADER_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
_DATA_FONT = Font(name="Calibri", size=10)
_TITLE_FONT = Font(name="Calibri", bold=True, size=14, color="0D1B3E")
_SUBTITLE_FONT = Font(name="Calibri", bold=False, size=10, color="666666")
_THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def _aplicar_estilos_encabezado(ws, num_columnas):
    """Aplica estilo a la fila de encabezados de una hoja."""
    for col in range(1, num_columnas + 1):
        celda = ws.cell(row=3, column=col)
        celda.fill = _HEADER_FILL
        celda.font = _HEADER_FONT
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = _THIN_BORDER


def _escribir_titulo(ws, titulo, subtitulo, num_columnas):
    """Escribe título y subtítulo en las primeras filas de la hoja."""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_columnas)
    celda_titulo = ws.cell(row=1, column=1, value=titulo)
    celda_titulo.font = _TITLE_FONT
    celda_titulo.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_columnas)
    celda_sub = ws.cell(row=2, column=1, value=subtitulo)
    celda_sub.font = _SUBTITLE_FONT
    celda_sub.alignment = Alignment(horizontal="left")
    ws.row_dimensions[2].height = 20


# ---------------------------------------------------------------------------
# Plantilla 1: Todas las columnas, todos los enlaces
# ---------------------------------------------------------------------------
COLUMNAS_COMPLETAS = [
    "Nombres",
    "Primer Apellido",
    "Segundo Apellido",
    "Genero",
    "Correo Principal",
    "Correo Secundario",
    "Telefono Principal",
    "Telefono Secundario",
    "Extension",
    "Estado",
    "Nivel Referencia",
    "Usuario SIG",
    "Password SIG",
    "Nombre SIG",
    "PIN SIG",
    "Institucion",
    "Siglas Inst.",
    "Edificio",
    "Fecha Alta",
    "Oficio Alta",
    "Observ. Alta",
    "Fecha Seguimiento",
    "Oficio Seguimiento",
    "Observ. Seguimiento",
    "Fecha Baja",
    "Oficio Baja",
    "Observ. Baja",
    "Creado",
]


def _fila_completa(enlace):
    """Devuelve una tupla con todos los campos de un enlace."""
    return (
        enlace.nombres,
        enlace.primer_apellido,
        enlace.segundo_apellido or "",
        enlace.get_genero_display() if enlace.genero else "",
        enlace.correo_principal,
        enlace.correo_secundario or "",
        enlace.telefono_principal,
        enlace.telefono_secundario or "",
        enlace.extension_telefonica or "",
        enlace.estado,
        enlace.nivel_referencia or "",
        enlace.usuario_sig or "",
        enlace.password_sig or "",
        enlace.nombre_sig or "",
        enlace.pin_sig or "",
        enlace.institucion.nombre if enlace.institucion else "",
        enlace.institucion.siglas or "" if enlace.institucion else "",
        ", ".join(ed.nombre for ed in enlace.institucion.edificio.all()) if enlace.institucion else "",
        enlace.fecha_alta.strftime("%d/%m/%Y") if enlace.fecha_alta else "",
        enlace.oficio_alta or "",
        enlace.observaciones_alta or "",
        enlace.fecha_seguimiento.strftime("%d/%m/%Y") if enlace.fecha_seguimiento else "",
        enlace.oficio_seguimiento or "",
        enlace.observaciones_seguimiento or "",
        enlace.fecha_baja.strftime("%d/%m/%Y") if enlace.fecha_baja else "",
        enlace.oficio_baja or "",
        enlace.observaciones_baja or "",
        enlace.creado_en.strftime("%d/%m/%Y %H:%M") if enlace.creado_en else "",
    )


# ---------------------------------------------------------------------------
# Plantilla 2: Solo enlaces activos, columnas esenciales
# ---------------------------------------------------------------------------
COLUMNAS_ACTIVOS = [
    "Nombres Completos",
    "Genero",
    "Correo Principal",
    "Telefono",
    "Extension",
    "Estado",
    "Institucion",
    "Siglas",
    "Edificio",
    "Nivel Ref.",
    "Usuario SIG",
    "Fecha Alta",
    "Oficio Alta",
]


def _fila_activo(enlace):
    """Devuelve una tupla con columnas esenciales de un enlace activo."""
    nombre_completo = f"{enlace.nombres} {enlace.primer_apellido} {enlace.segundo_apellido}".strip()
    return (
        nombre_completo,
        enlace.get_genero_display() if enlace.genero else "",
        enlace.correo_principal,
        enlace.telefono_principal,
        enlace.extension_telefonica or "",
        enlace.estado,
        enlace.institucion.nombre if enlace.institucion else "",
        enlace.institucion.siglas or "" if enlace.institucion else "",
        ", ".join(ed.nombre for ed in enlace.institucion.edificio.all()) if enlace.institucion else "",
        enlace.nivel_referencia or "",
        enlace.usuario_sig or "",
        enlace.fecha_alta.strftime("%d/%m/%Y") if enlace.fecha_alta else "",
        enlace.oficio_alta or "",
    )


# ---------------------------------------------------------------------------
# Plantilla 3: Enlaces activos — columnas solicitadas por el usuario
# ---------------------------------------------------------------------------
COLUMNAS_SIMPLES = [
    "Edificio",
    "Institución",
    "Nombre Completo",
    "Correo",
    "Teléfono",
    "Estado",
]


def _fila_simple(enlace):
    """Devuelve una tupla con las 6 columnas esenciales para exportación simple."""
    nombre_completo = f"{enlace.nombres} {enlace.primer_apellido} {enlace.segundo_apellido}".strip()
    return (
        ", ".join(ed.nombre for ed in enlace.institucion.edificio.all()) if enlace.institucion else "",
        enlace.institucion.nombre if enlace.institucion else "",
        nombre_completo,
        enlace.correo_principal,
        enlace.telefono_principal,
        enlace.estado,
    )


# ---------------------------------------------------------------------------
# Vistas públicas (solo lectura)
# ---------------------------------------------------------------------------
@requiere(CAP_DIRECTORIO)
def lista_edificios(request):
    """Muestra todos los edificios del CCG."""
    edificios = Edificio.objects.filter(visible=True).annotate(
        num_instituciones=models.Count("instituciones", distinct=True),
        num_enlaces_activos=models.Count(
            "instituciones__enlaces",
            filter=models.Q(instituciones__enlaces__estado="ACTIVO"),
            distinct=True,
        ),
    ).order_by("nombre")

    return render(
        request,
        "enlaces_ccg/lista_edificios.html",
        {"edificios": edificios},
    )


@requiere(CAP_DIRECTORIO)
def comunicados_cc(request):
    """Vista de Comunicados CC: personas de contacto para comunicados."""
    comunicados = ComunicadoCC.objects.select_related("institucion").order_by(
        F("institucion__nombre").desc(nulls_last=True), "nombre"
    )
    instituciones = Institucion.objects.order_by("nombre")
    return render(
        request,
        "enlaces_ccg/comunicados_cc.html",
        {"comunicados": comunicados, "instituciones": instituciones},
    )


@requiere(CAP_EDITAR)
@require_POST
def crear_comunicado_cc(request):
    """Crea un nuevo comunicado/contacto desde el modal."""
    nombre = request.POST.get("nombre", "").strip()
    institucion_id = request.POST.get("institucion")
    cargo = request.POST.get("cargo", "").strip()
    correo = request.POST.get("correo", "").strip()

    if not nombre:
        return JsonResponse({"ok": False, "error": "El nombre es obligatorio."}, status=400)

    institucion = None
    if institucion_id and institucion_id.isdigit():
        institucion = Institucion.objects.filter(pk=int(institucion_id)).first()

    comunicado = ComunicadoCC.objects.create(
        nombre=nombre,
        institucion=institucion,
        cargo=cargo,
        correo=correo or None,
    )
    return JsonResponse({"ok": True, "id": comunicado.pk})


@requiere(CAP_EDITAR)
@require_POST
def editar_comunicado_cc(request, pk):
    """Actualiza un comunicado/contacto desde el modal."""
    comunicado = get_object_or_404(ComunicadoCC, pk=pk)
    nombre = request.POST.get("nombre", "").strip()
    institucion_id = request.POST.get("institucion")
    cargo = request.POST.get("cargo", "").strip()
    correo = request.POST.get("correo", "").strip()

    if not nombre:
        return JsonResponse({"ok": False, "error": "El nombre es obligatorio."}, status=400)

    institucion = None
    if institucion_id and institucion_id.isdigit():
        institucion = Institucion.objects.filter(pk=int(institucion_id)).first()

    comunicado.nombre = nombre
    comunicado.institucion = institucion
    comunicado.cargo = cargo
    comunicado.correo = correo or None
    comunicado.save()
    return JsonResponse({"ok": True, "id": comunicado.pk})


@requiere(CAP_EDITAR)
@require_POST
def eliminar_comunicado_cc(request, pk):
    """Elimina un comunicado/contacto."""
    comunicado = get_object_or_404(ComunicadoCC, pk=pk)
    comunicado.delete()
    return JsonResponse({"ok": True})


@requiere(CAP_DIRECTORIO)
def copiar_correos_comunicados(request):
    """Devuelve JSON con los correos de los comunicados (formato CCO)."""
    correos = []
    vistos = set()
    for c in ComunicadoCC.objects.exclude(correo__isnull=True).exclude(correo=""):
        correo = c.correo.strip().lower()
        if correo and correo not in vistos:
            vistos.add(correo)
            correos.append(correo)
    return JsonResponse({
        "ok": True,
        "correos": correos,
        "total": len(correos),
        "formato": "semicolon",
    })


def _entero_o_none(valor):
    """Convierte a int un valor de querystring (None si no es válido)."""
    try:
        return int(valor)
    except (TypeError, ValueError):
        return None


def _filtrar_tickets(qs, params):
    """Aplica los filtros del listado de tickets (compartido por la lista
    general y la pestaña de tickets de la ficha de enlace).

    Devuelve (queryset_filtrado, contexto) donde `contexto` trae los valores
    normalizados para repintar el formulario de filtros.
    """
    estatus = params.get("estatus", "").strip()
    if estatus in (Ticket.ESTATUS_ABIERTO, Ticket.ESTATUS_CERRADO):
        qs = qs.filter(estatus=estatus)

    incluir_archivados = params.get("incluir", "") == "1"
    if not incluir_archivados:
        qs = qs.filter(archivado=False)

    # Filtros por vínculo (Fase 3): torre / servicio / falla / nivel (por PK).
    torre_id = _entero_o_none(params.get("torre", ""))
    if torre_id:
        qs = qs.filter(torre_id=torre_id)
    servicio_id = _entero_o_none(params.get("servicio", ""))
    if servicio_id:
        qs = qs.filter(servicio_id=servicio_id)
    falla_id = _entero_o_none(params.get("falla", ""))
    if falla_id:
        qs = qs.filter(falla_id=falla_id)
    nivel_id = _entero_o_none(params.get("nivel", ""))
    if nivel_id:
        qs = qs.filter(nivel_id=nivel_id)

    # Responsable de atención (catálogo, se filtra por PK).
    responsable_id = _entero_o_none(params.get("responsable", ""))
    if responsable_id:
        qs = qs.filter(responsable_atencion_id=responsable_id)

    # Solicitante (enlace del directorio, se filtra por PK).
    solicitante_id = _entero_o_none(params.get("solicitante", ""))
    if solicitante_id:
        qs = qs.filter(solicitante_id=solicitante_id)

    institucion_id = _entero_o_none(params.get("institucion", ""))
    if institucion_id:
        qs = qs.filter(institucion_id=institucion_id)

    q = params.get("q", "").strip()
    if q:
        qs = qs.filter(
            Q(numero_display__icontains=q)
            | Q(numero__icontains=q)
            | Q(descripcion__icontains=q)
            | Q(raw_data__solicitud_descripcion__icontains=q)
            | Q(falla__descripcion__icontains=q)
            | Q(raw_data__falla__icontains=q)
            | Q(raw_data__falla_descripcion__icontains=q)
            | Q(raw_data__Observaciones__icontains=q)
            | Q(raw_data__ObservacionesUsuario__icontains=q)
        )

    desde = params.get("desde", "").strip()
    hasta = params.get("hasta", "").strip()
    if desde:
        try:
            qs = qs.filter(fecha__gte=datetime.strptime(desde, "%Y-%m-%d"))
        except ValueError:
            desde = ""
    if hasta:
        try:
            qs = qs.filter(
                fecha__lt=datetime.strptime(hasta, "%Y-%m-%d") + timedelta(days=1)
            )
        except ValueError:
            hasta = ""

    return qs, {
        "estatus": estatus,
        "incluir_archivados": incluir_archivados,
        "torre_id": torre_id,
        "servicio_id": servicio_id,
        "falla_id": falla_id,
        "nivel_id": nivel_id,
        "responsable_id": responsable_id,
        "solicitante_id": solicitante_id,
        "institucion_id": institucion_id,
        "q": q,
        "desde": desde,
        "hasta": hasta,
    }


@requiere(CAP_TICKETS)
def seguimiento_tickets(request):
    """Lista paginada de tickets con filtros (servidor-side).

    Rendimiento: consulta con índices, sin raw_data en la tabla, paginación
    de 50 registros y badge con total_tickets sin contar por petición.
    """
    cfg = ConfiguracionTickets.cargar()
    # Paginado servidor-side: SQL LIMIT/OFFSET; cada página trae solo 50 filas
    # (incluye raw_data para las columnas de contexto, sin consultas extra).
    qs = Ticket.objects.select_related(
        "torre", "institucion", "solicitante", "servicio", "falla", "nivel"
    )

    qs, filtros_ctx = _filtrar_tickets(qs, request.GET)
    estatus = filtros_ctx["estatus"]
    incluir_archivados = filtros_ctx["incluir_archivados"]
    torre_id = filtros_ctx["torre_id"]
    servicio_id = filtros_ctx["servicio_id"]
    falla_id = filtros_ctx["falla_id"]
    nivel_id = filtros_ctx["nivel_id"]
    responsable_id = filtros_ctx["responsable_id"]
    solicitante_id = filtros_ctx["solicitante_id"]
    institucion_id = filtros_ctx["institucion_id"]
    q = filtros_ctx["q"]
    desde = filtros_ctx["desde"]
    hasta = filtros_ctx["hasta"]

    paginator = Paginator(qs, 50)
    try:
        pagina = int(request.GET.get("page", "1"))
    except (TypeError, ValueError):
        pagina = 1
    page_obj = paginator.get_page(pagina)

    filtros = "&".join(
        f"{k}={quote(v)}"
        for k, v in request.GET.items()
        if k != "page" and v.strip() != ""
    )

    proxima = cfg.proxima_sincronizacion()
    if proxima:
        inicio = cfg.ultima_sincronizacion or (
            proxima - timedelta(minutes=cfg.intervalo_minutos or 5)
        )
        inicio_ts, proxima_ts = int(inicio.timestamp()), int(proxima.timestamp())
    else:
        inicio_ts = proxima_ts = 0

    return render(
        request,
        "enlaces_ccg/seguimiento_tickets.html",
        {
            "page_obj": page_obj,
            "filtros": filtros,
            "estatus": estatus,
            "q": q,
            "desde": desde,
            "hasta": hasta,
            "incluir_archivados": incluir_archivados,
            "torre_id": torre_id,
            "servicio_id": servicio_id,
            "falla_id": falla_id,
            "nivel_id": nivel_id,
            "responsable_id": responsable_id,
            "solicitante_id": solicitante_id,
            "institucion_id": institucion_id,
            "edificios": Edificio.objects.all(),
            "servicios": Servicio.objects.filter(activo=True),
            "fallas": Falla.objects.filter(activo=True),
            "niveles": Nivel.objects.filter(activo=True),
            "responsables": ResponsableAtencion.objects.filter(activo=True),
            "solicitantes": EnlaceAutorizado.objects.filter(
                tickets__isnull=False
            ).distinct().order_by("nombre_sig"),
            "instituciones": Institucion.objects.order_by("nombre"),
            "total_tickets": cfg.total_tickets,
            "ultima_sincronizacion": cfg.ultima_sincronizacion,
            "proxima_sincronizacion": proxima,
            "sync_inicio_ts": inicio_ts,
            "sync_proxima_ts": proxima_ts,
            "ultima_sync_ts": (
                int(cfg.ultima_sincronizacion.timestamp())
                if cfg.ultima_sincronizacion
                else 0
            ),
            "ultimo_intento_ts": (
                int(cfg.ultimo_intento.timestamp()) if cfg.ultimo_intento else 0
            ),
            "cfg": cfg,
        },
    )


@requiere(CAP_TICKETS)
def ticket_detalle(request, pk):
    """Detalle canónico por PK (único, sin ambigüedad)."""
    ticket = get_object_or_404(Ticket, pk=pk)
    return render(
        request,
        "enlaces_ccg/ticket_detalle.html",
        {"ticket": ticket, "registros": ticket.registros.all()[:5]},
    )


@requiere(CAP_TICKETS)
def ticket_json(request, pk):
    """Datos del ticket para el modal de acciones (JSON)."""
    t = get_object_or_404(Ticket, pk=pk)
    return JsonResponse(
        {
            "pk": t.pk,
            "numero": t.numero_display,
            "numero_base": t.numero,
            "ticket_id": t.ticket_id,
            "estatus": t.estatus,
            "fecha": t.fecha.strftime("%d/%m/%Y %H:%M") if t.fecha else "",
            "fecha_cierre": (
                t.fecha_cierre.strftime("%d/%m/%Y %H:%M") if t.fecha_cierre else ""
            ),
            "solicitud_tipo": t.solicitud_tipo,
            "descripcion": t.descripcion,
            "archivado": t.archivado,
            "ultima_sync": (
                t.ultima_sync.strftime("%d/%m/%Y %H:%M") if t.ultima_sync else ""
            ),
            "raw_data": t.raw_data,
            "url_detalle": reverse("enlaces_ccg:ticket_detalle", args=[t.pk]),
        }
    )


@requiere(CAP_TICKETS)
def ticket_por_numero(request, numero):
    """Detalle por número visible; acepta el sufijo -N de duplicados."""
    ticket = (
        Ticket.objects.filter(numero_display=numero).first()
        or Ticket.objects.filter(numero=numero).first()
    )
    if ticket is None:
        raise Http404(f"No existe el ticket {numero}")
    return render(
        request,
        "enlaces_ccg/ticket_detalle.html",
        {"ticket": ticket, "registros": ticket.registros.all()[:5]},
    )


@login_required
@require_POST
@requiere(CAP_TICKETS)
def registrar_seguimiento(request, pk):
    """Registra una consulta de seguimiento (solo descripción)."""
    ticket = get_object_or_404(Ticket, pk=pk)
    descripcion = request.POST.get("descripcion", "").strip()
    if not descripcion:
        messages.warning(request, "Escribe una breve descripción de lo sucedido.")
    else:
        TicketRegistro.objects.create(
            ticket=ticket, usuario=request.user, descripcion=descripcion
        )
        messages.success(request, "Seguimiento registrado correctamente.")
    return redirect("enlaces_ccg:ticket_detalle", pk=pk)


@login_required
@require_POST
@requiere(CAP_EDITAR)
def sincronizar_tickets_ahora(request):
    """Encola una DESCARGA RÁPIDA (parcial) de tickets en segundo plano."""
    from .sig_sync.tickets import MODO_SYNC_PARCIAL, sincronizar_tickets_task

    sincronizar_tickets_task.delay(modo=MODO_SYNC_PARCIAL)
    messages.success(
        request,
        "Descarga rápida de tickets encolada en segundo plano. "
        "Se reflejará al terminar en 'Última sincronización'.",
    )
    return redirect("enlaces_ccg:seguimiento_tickets")


@requiere(CAP_TICKETS)
def sincronizacion_estado(request):
    """Estado de la sincronización en JSON (para la barra de la lista)."""
    from .sig_sync.tickets import _lock_activo

    cfg = ConfiguracionTickets.cargar()
    ultima = cfg.ultima_sincronizacion
    if cfg.habilitado:
        proxima = cfg.proxima_sincronizacion()
        proxima_ts = proxima.timestamp() if proxima else None
    else:
        proxima_ts = None
    return JsonResponse(
        {
            "ultima_sincronizacion": int(ultima.timestamp()) if ultima else None,
            "ultimo_intento": (
                int(cfg.ultimo_intento.timestamp()) if cfg.ultimo_intento else None
            ),
            "proxima_sincronizacion": proxima_ts,
            "sincronizando": _lock_activo(),
            "habilitado": cfg.habilitado,
            "en_horario": cfg.en_horario(),
            "ultimo_estado": cfg.ultimo_estado or "",
            "ultimo_mensaje": cfg.ultimo_mensaje or "",
        }
    )


@requiere(CAP_TICKETS)
def dashboard(request):
    """Vista Dashboard de seguimiento de tickets (en construcción)."""
    return render(request, "enlaces_ccg/dashboard.html")


@requiere(CAP_TICKETS)
def indicadores_mejora(request):
    """Vista de Indicadores de mejora continua (en construcción)."""
    return render(request, "enlaces_ccg/indicadores_mejora.html")


@requiere(CAP_IMPORTAR)
def importar_deductivas(request):
    """Vista para importar deductivas (en construcción)."""
    return render(request, "enlaces_ccg/importar_deductivas.html")


@requiere(CAP_REVISIONES)
def revision_tickets(request):
    """Vista de Revisión de Tickets (en construcción)."""
    return render(request, "enlaces_ccg/revision_tickets.html")


@requiere(CAP_REVISIONES)
def revision_correos(request):
    """Vista de Revisión de Correos (en construcción)."""
    return render(request, "enlaces_ccg/revision_correos.html")


@requiere(CAP_REVISIONES)
def cierre_operador(request):
    """Vista de Cierre de Operador (en construcción)."""
    return render(request, "enlaces_ccg/cierre_operador.html")


@requiere(CAP_REVISIONES)
def comentarios_operador(request):
    """Vista de Comentarios de Operador (en construcción)."""
    return render(request, "enlaces_ccg/comentarios_operador.html")


@requiere(CAP_REVISIONES)
def recordatorios(request):
    """Vista de Recordatorios (en construcción)."""
    return render(request, "enlaces_ccg/recordatorios.html")


@requiere(CAP_REVISIONES)
def plantillas_respuesta(request):
    """Vista de Plantillas de respuesta (en construcción)."""
    return render(request, "enlaces_ccg/plantillas_respuesta.html")


@requiere(CAP_REVISIONES)
def aprendizaje(request):
    """Vista de Aprendizaje (en construcción)."""
    return render(request, "enlaces_ccg/aprendizaje.html")


@requiere(CAP_ADMIN)
def grupos(request):
    """Lista los grupos (roles) y permite crear uno nuevo."""
    from django.contrib.auth.models import Group

    from .models import PermisoGrupo
    from .roles import CAP_ADMIN, ROL_ADMIN, SECCIONES, capacidades_de

    if request.method == "POST":
        nombre = request.POST.get("nombre", "").strip()
        if not nombre:
            messages.error(request, "El nombre del grupo no puede estar vacío.")
        elif Group.objects.filter(name__iexact=nombre).exists():
            messages.error(request, f"Ya existe un grupo llamado «{nombre}».")
        else:
            g = Group.objects.create(name=nombre)
            PermisoGrupo.objects.get_or_create(grupo=g, defaults={"capacidades": []})
            messages.success(request, f"Grupo «{g.name}» creado.")
        return redirect("enlaces_ccg:grupos")

    grupos = list(
        Group.objects.prefetch_related("permiso_grupo").order_by("name")
    )

    # Etiquetas de capacidad por sección para mostrar el resumen
    etiqueta_seccion = {sec["clave"]: sec["etiqueta"] for sec in SECCIONES}
    filas = []
    for g in grupos:
        caps = capacidades_de(g)
        etiquetas = []
        if CAP_ADMIN in caps or g.name == ROL_ADMIN:
            etiquetas.append("Admin")
        for clave, etiq in etiqueta_seccion.items():
            if clave in caps:
                etiquetas.append(etiq)
                if clave == "directorio" and "editar" in caps and "Editar" not in etiquetas:
                    etiquetas.append("Editar")
        filas.append({"grupo": g, "etiquetas": etiquetas})

    return render(
        request,
        "enlaces_ccg/grupos.html",
        {"filas": filas},
    )


@requiere(CAP_ADMIN)
def permisos_grupo(request, pk):
    """Asigna permisos (secciones y vistas) a un grupo concreto."""
    from django.contrib.auth.models import Group

    from .models import PermisoGrupo
    from .roles import ROL_ADMIN, SECCIONES, VISTAS, capacidades_de

    grupo = Group.objects.filter(pk=pk).first()
    if grupo is None:
        messages.error(request, "El grupo no existe.")
        return redirect("enlaces_ccg:grupos")

    caps_por_seccion = {
        "directorio": ["directorio", "editar"],
        "tickets": ["tickets"],
        "importar": ["importar"],
        "revisiones": ["revisiones"],
    }

    if request.method == "POST":
        permiso = PermisoGrupo.objects.get_or_create(grupo=grupo)[0]
        nuevas = []
        for sec in SECCIONES:
            for cap in caps_por_seccion.get(sec["clave"], []):
                if request.POST.get(f"cap_{cap}"):
                    nuevas.append(cap)
        if grupo.name == ROL_ADMIN and CAP_ADMIN not in nuevas:
            nuevas.append(CAP_ADMIN)
        permiso.capacidades = nuevas
        permiso.save()
        messages.success(request, f"Permisos de «{grupo.name}» actualizados.")
        return redirect("enlaces_ccg:permisos_grupo", pk=grupo.pk)

    caps_actuales = capacidades_de(grupo)

    secciones = []
    for sec in SECCIONES:
        secciones.append(
            {
                "clave": sec["clave"],
                "etiqueta": sec["etiqueta"],
                "vistas": [v for v in VISTAS if v["seccion"] == sec["clave"]],
                "caps": caps_por_seccion.get(sec["clave"], []),
            }
        )

    return render(
        request,
        "enlaces_ccg/permisos_grupo.html",
        {
            "grupo": grupo,
            "secciones": secciones,
            "caps_actuales": caps_actuales,
        },
    )


@requiere(CAP_DIRECTORIO)
def lista_instituciones(request):
    """Muestra las instituciones del CCG con pestañas por estado y filtro por edificio."""
    vista = request.GET.get("vista", "activas")
    if vista not in ("activas", "inactivas", "todas"):
        vista = "activas"
    edificio_id = request.GET.get("edificio", "")
    niveles = request.GET.getlist("niveles")
    qs = (
        Institucion.objects.annotate(num_enlaces=Count("enlaces", filter=Q(enlaces__estado="ACTIVO")))
        .order_by("nombre")
    )
    if vista == "activas":
        qs = qs.filter(estado="ACTIVO")
    elif vista == "inactivas":
        qs = qs.filter(estado="INACTIVO")
    if edificio_id.isdigit():
        qs = qs.filter(edificio__id=edificio_id)
        if niveles:
            # Filtra instituciones que tengan los niveles en el edificio seleccionado.
            # `contains` (jsonb @>) no lo soporta SQLite, así que se evalúa del lado Python.
            inst_ids = [
                ie.institucion_id
                for ie in InstitucionEdificio.objects.filter(edificio_id=int(edificio_id))
                if any(n in (ie.nivel or []) for n in niveles)
            ]
            qs = qs.filter(pk__in=inst_ids)
    edificios = Edificio.objects.filter(visible=True).order_by("nombre")
    return render(
        request,
        "enlaces_ccg/lista_instituciones.html",
        {
            "instituciones": qs,
            "vista": vista,
            "edificio_id": edificio_id,
            "niveles": niveles,
            "qs_niveles": "".join(
                "&niveles={}".format(quote(n)) for n in niveles
            ),
            "edificios": edificios,
            "NIVEL_CHOICES": Institucion.NIVEL_CHOICES,
        },
    )


# ---------------------------------------------------------------------------
# Helper: filtrado común de enlaces (lista global y captura de correos)
# ---------------------------------------------------------------------------
def _filtrar_enlaces(params, institucion=None):
    """
    Devuelve queryset de EnlaceAutorizado filtrado según parámetros.
    params: request.GET (QueryDict)
    institucion: Instancia Institucion (None para lista global)
    """
    if institucion:
        # Perfil institución: filtro por vista_enlaces (venlaces)
        qs = institucion.enlaces.all()
        vista = params.get("venlaces", "activos")
        if vista == "activos":
            qs = qs.filter(estado="ACTIVO")
        elif vista == "inactivos":
            qs = qs.filter(estado="INACTIVO")
        # "todos" = sin filtro de estado
        return qs.order_by("primer_apellido", "segundo_apellido", "nombres")

    # Lista global
    qs = EnlaceAutorizado.objects.select_related("institucion").prefetch_related(
        "institucion__edificio"
    )

    vista = params.get("vista", "activos")
    if vista == "activos":
        qs = qs.filter(estado="ACTIVO").exclude(
            institucion__nombre__icontains="OPERADORA CC"
        )
    elif vista == "inactivos":
        qs = qs.filter(estado="INACTIVO")
    elif vista == "pendientes":
        # Activos pendientes de sincronizar con el SIG
        qs = qs.filter(estado="ACTIVO", sincronizado=False).exclude(
            institucion__nombre__icontains="OPERADORA CC"
        )
    # vista == "todos" -> sin filtro de estado

    busqueda = params.get("q", "").strip()
    col = params.get("col", "todas").strip()
    edificio_id = params.get("edificio", "")
    niveles = params.getlist("niveles")

    if edificio_id.isdigit():
        qs = qs.filter(institucion__edificio__id=edificio_id)
        if niveles:
            inst_ids = [
                ie.institucion_id
                for ie in InstitucionEdificio.objects.filter(edificio_id=int(edificio_id))
                if any(n in (ie.nivel or []) for n in niveles)
            ]
            qs = qs.filter(institucion_id__in=inst_ids)

    if busqueda:
        COLUMNAS_BUSQUEDA = {
            "todas": None,
            "nombre": ["nombres", "primer_apellido", "segundo_apellido"],
            "institucion": ["institucion__nombre"],
            "siglas": ["institucion__siglas"],
            "correo": ["correo_principal", "correo_secundario"],
            "telefono": ["telefono_principal", "telefono_secundario"],
            "nivel": ["nivel_referencia"],
        }
        if col in COLUMNAS_BUSQUEDA and COLUMNAS_BUSQUEDA[col]:
            q_objects = Q()
            for campo in COLUMNAS_BUSQUEDA[col]:
                q_objects |= Q(**{f"{campo}__icontains": busqueda})
            qs = qs.filter(q_objects)
        elif col == "todas":
            qs = qs.filter(
                Q(nombres__icontains=busqueda)
                | Q(primer_apellido__icontains=busqueda)
                | Q(segundo_apellido__icontains=busqueda)
                | Q(correo_principal__icontains=busqueda)
                | Q(correo_secundario__icontains=busqueda)
                | Q(telefono_principal__icontains=busqueda)
                | Q(telefono_secundario__icontains=busqueda)
                | Q(usuario_sig__icontains=busqueda)
                | Q(nivel_referencia__icontains=busqueda)
                | Q(institucion__nombre__icontains=busqueda)
                | Q(institucion__siglas__icontains=busqueda)
            )

    return qs.order_by("institucion__nombre", "primer_apellido", "segundo_apellido", "nombres")


@requiere(CAP_DIRECTORIO)
def lista_enlaces(request):
    """Muestra todos los enlaces autorizados del CCG con buscador."""
    # Validación de parámetros para el template
    vista = request.GET.get("vista", "activos")
    if vista not in ("activos", "inactivos", "todos", "pendientes"):
        vista = "activos"
    col = request.GET.get("col", "todas").strip()
    COLUMNAS_BUSQUEDA = [
        ("todas", "Todas las columnas"),
        ("nombre", "Nombre"),
        ("institucion", "Institución"),
        ("siglas", "Siglas de institución"),
        ("correo", "Correo (principal o secundario)"),
        ("telefono", "Teléfono (principal o secundario)"),
        ("nivel", "Nivel de referencia"),
    ]
    if col not in dict(COLUMNAS_BUSQUEDA):
        col = "todas"

    # Usar helper de filtrado compartido
    enlaces = _filtrar_enlaces(request.GET, institucion=None)

    edificios = Edificio.objects.filter(visible=True).order_by("nombre")
    qs_niveles = "".join("&niveles={}".format(quote(n)) for n in request.GET.getlist("niveles"))
    return render(
        request,
        "enlaces_ccg/lista_enlaces.html",
        {
            "enlaces": enlaces,
            "vista": vista,
            "busqueda": request.GET.get("q", "").strip(),
            "col": col,
            "columnas_busqueda": COLUMNAS_BUSQUEDA,
            "edificios": edificios,
            "edificio_id": request.GET.get("edificio", ""),
            "niveles": request.GET.getlist("niveles"),
            "qs_niveles": qs_niveles,
            "NIVEL_CHOICES": Institucion.NIVEL_CHOICES,
        },
    )


@requiere(CAP_DIRECTORIO)
def copiar_correos_enlaces(request):
    """Devuelve JSON con correos primarios y secundarios de los enlaces filtrados.

    Parámetros GET (mismos que lista_enlaces o perfil_institucion):
    - institucion=<pk> (opcional): si se proporciona, filtra por esa institución
    - Los demás parámetros (vista, q, col, edificio, niveles, venlaces)
      se pasan al helper _filtrar_enlaces.

    Respuesta JSON:
    {
        "ok": true,
        "correos": ["a@b.com", "c@d.com", ...],
        "total": 42,
        "formato": "semicolon"
    }
    """
    institucion_pk = request.GET.get("institucion")
    institucion = None
    if institucion_pk and institucion_pk.isdigit():
        institucion = get_object_or_404(Institucion, pk=institucion_pk)

    qs = _filtrar_enlaces(request.GET, institucion=institucion)

    # Extraer correos (principal + secundario) sin sorted(), solo deduplicar
    correos = []
    vistos = set()
    for e in qs:
        if e.correo_principal:
            c = e.correo_principal.strip().lower()
            if c not in vistos:
                vistos.add(c)
                correos.append(c)
        if e.correo_secundario:
            c = e.correo_secundario.strip().lower()
            if c not in vistos:
                vistos.add(c)
                correos.append(c)

    return JsonResponse({
        "ok": True,
        "correos": correos,
        "total": len(correos),
        "formato": "semicolon"
    })


@requiere(CAP_DIRECTORIO)
def detalle_edificio(request, pk):
    """Detalle de un edificio con sus instituciones (con pestañas Activos/Inactivos/Todos)."""
    edificio = get_object_or_404(Edificio, pk=pk)
    vista = request.GET.get("vista", "activas")
    if vista not in ("activas", "inactivas", "todas"):
        vista = "activas"

    qs = edificio.instituciones.prefetch_related("instituciones_edificios__edificio").annotate(
        num_enlaces_activos=Count(
            "enlaces",
            filter=Q(enlaces__estado="ACTIVO"),
            distinct=True,
        ),
        num_enlaces_total=Count("enlaces", distinct=True),
    ).order_by("nombre")
    if vista == "activas":
        qs = qs.filter(estado="ACTIVO")
    elif vista == "inactivas":
        qs = qs.filter(estado="INACTIVO")

    return render(
        request,
        "enlaces_ccg/detalle_edificio.html",
        {"edificio": edificio, "instituciones": qs, "vista": vista},
    )


@requiere(CAP_DIRECTORIO)
def perfil_institucion(request, pk):
    """Perfil de una institución con sus enlaces autorizados."""
    institucion = get_object_or_404(
        Institucion.objects.prefetch_related("edificio", "instituciones_edificios__edificio"), pk=pk
    )
    vista_enlaces = request.GET.get("venlaces", "activos")
    if vista_enlaces not in ("activos", "inactivos", "todos"):
        vista_enlaces = "activos"

    enlaces = institucion.enlaces.all()
    if vista_enlaces == "activos":
        enlaces = enlaces.filter(estado="ACTIVO")
    elif vista_enlaces == "inactivos":
        enlaces = enlaces.filter(estado="INACTIVO")
    enlaces = enlaces.order_by("primer_apellido", "segundo_apellido", "nombres")

    ver_tickets = request.GET.get("ver_tickets")
    page_obj = None
    total_tickets = abiertos = cerrados = 0
    filtros_ctx = {}
    filtros = ""

    if ver_tickets and tiene(request.user, CAP_TICKETS):
        qs = (
            Ticket.objects.filter(institucion=institucion)
            .select_related("torre", "solicitante", "servicio", "falla", "nivel")
            .order_by("-fecha", "ticket_id")
        )
        # Conteos sin filtros (totales reales de la institución)
        total_tickets = qs.count()
        abiertos = qs.filter(estatus=Ticket.ESTATUS_ABIERTO).count()
        cerrados = total_tickets - abiertos

        # Aplicar filtros del formulario
        qs, filtros_ctx = _filtrar_tickets(qs, request.GET)

        paginator = Paginator(qs, 50)
        try:
            pagina = int(request.GET.get("page", "1"))
        except (TypeError, ValueError):
            pagina = 1
        page_obj = paginator.get_page(pagina)

        # Cadena de parámetros activos (para los enlaces de paginación)
        filtros = "&".join(
            f"{k}={quote(v)}"
            for k, v in request.GET.items()
            if k != "page" and v.strip() != ""
        )

    return render(
        request,
        "enlaces_ccg/perfil_institucion.html",
        {
            "institucion": institucion,
            "enlaces": enlaces,
            "vista_enlaces": vista_enlaces,
            "ver_tickets": ver_tickets,
            "page_obj": page_obj,
            "total_tickets": total_tickets,
            "tickets_abiertos": abiertos,
            "tickets_cerrados": cerrados,
            "filtros": filtros,
            "edificios": Edificio.objects.all(),
            "servicios": Servicio.objects.filter(activo=True),
            "fallas": Falla.objects.filter(activo=True),
            "niveles": Nivel.objects.filter(activo=True),
            **(filtros_ctx or {}),
        },
    )


# ---------------------------------------------------------------------------
# Descarga Excel
# ---------------------------------------------------------------------------
@requiere(CAP_DIRECTORIO)
def exportar_enlaces(request, pk):
    """
    Genera y descarga un archivo Excel con los enlaces de una institución.

    GET params:
        plantilla=completa   → todas las columnas, todos los enlaces
        plantilla=activos    → columnas esenciales, solo activos (default)
        todos=1              → incluir inactivos (solo para plantilla=completa)
    """
    institucion = get_object_or_404(
        Institucion.objects.prefetch_related("edificio"), pk=pk
    )

    plantilla = request.GET.get("plantilla", "activos")
    mostrar_todos = request.GET.get("todos") == "1"

    # ---- Seleccionar plantilla de columnas y función de fila ----
    if plantilla == "completa":
        columnas = COLUMNAS_COMPLETAS
        funcion_fila = _fila_completa
        enlaces = institucion.enlaces.all()
        if not mostrar_todos:
            enlaces = enlaces.filter(estado="ACTIVO")
        titulo_hoja = "Enlaces Completos"
    else:
        columnas = COLUMNAS_ACTIVOS
        funcion_fila = _fila_activo
        enlaces = institucion.enlaces.filter(estado="ACTIVO")
        titulo_hoja = "Enlaces Activos"

    enlaces = enlaces.order_by("primer_apellido", "segundo_apellido", "nombres")

    # ---- Construir workbook ----
    wb = Workbook()
    ws = wb.active
    ws.title = titulo_hoja[:31]  # Excel limita a 31 caracteres

    num_columnas = len(columnas)
    subtitulo = f"{institucion.nombre_completo} — Generado {datetime.now().strftime('%d/%m/%Y %H:%M')}"

    _escribir_titulo(ws, titulo_hoja, subtitulo, num_columnas)

    # Encabezados (fila 3)
    for col_idx, nombre_col in enumerate(columnas, start=1):
        ws.cell(row=3, column=col_idx, value=nombre_col)
    _aplicar_estilos_encabezado(ws, num_columnas)

    # Datos (fila 4 en adelante)
    for row_idx, enlace in enumerate(enlaces, start=4):
        for col_idx, valor in enumerate(funcion_fila(enlace), start=1):
            celda = ws.cell(row=row_idx, column=col_idx, value=valor)
            celda.font = _DATA_FONT
            celda.border = _THIN_BORDER
            celda.alignment = Alignment(vertical="center", wrap_text=True)

    # Ajustar ancho de columnas automáticamente
    for col_idx in range(1, num_columnas + 1):
        max_len = len(str(columnas[col_idx - 1]))
        for row_idx in range(4, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[ws.cell(row=3, column=col_idx).column_letter].width = min(max_len + 4, 40)

    # Fila de totales
    total_row = ws.max_row + 2
    ws.cell(row=total_row, column=1, value=f"Total de registros: {enlaces.count()}").font = Font(
        name="Calibri", bold=True, size=10, color="333333"
    )

    # ---- Generar respuesta HTTP ----
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"enlaces_{institucion.siglas or 'inst'}_{plantilla}_{timestamp}.xlsx"

    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{nombre_archivo}"'
    return response


# ---------------------------------------------------------------------------
# Exportación global (todos los enlaces del CCG)
# ---------------------------------------------------------------------------
@requiere(CAP_DIRECTORIO)
def exportar_enlaces_global(request):
    """
    Genera y descarga un Excel con TODOS los enlaces del CCG.

    GET params:
        plantilla=activos    → solo activos, 6 columnas (default)
        plantilla=completa   → todas las columnas
        todos=1              → incluir inactivos (solo para plantilla=completa)
    """
    plantilla = request.GET.get("plantilla", "activos")
    mostrar_todos = request.GET.get("todos") == "1"

    enlaces = EnlaceAutorizado.objects.select_related("institucion").prefetch_related(
        "institucion__edificio"
    )

    if plantilla == "completa":
        columnas = COLUMNAS_COMPLETAS
        funcion_fila = _fila_completa
        if not mostrar_todos:
            enlaces = enlaces.filter(estado="ACTIVO")
        titulo_hoja = "Enlaces Completos CCG"
    else:
        columnas = COLUMNAS_SIMPLES
        funcion_fila = _fila_simple
        enlaces = enlaces.filter(estado="ACTIVO").exclude(
            institucion__nombre__icontains="OPERADORA CC"
        )
        titulo_hoja = "Enlaces Activos CCG"

    enlaces = enlaces.order_by(
        "institucion__nombre",
        "nombres",
        "primer_apellido",
        "segundo_apellido",
    )

    # ---- Construir workbook ----
    wb = Workbook()
    ws = wb.active
    ws.title = titulo_hoja[:31]

    num_columnas = len(columnas)
    subtitulo = f"Centro Cívico Gubernamental — Generado {datetime.now().strftime('%d/%m/%Y %H:%M')}"

    _escribir_titulo(ws, titulo_hoja, subtitulo, num_columnas)

    for col_idx, nombre_col in enumerate(columnas, start=1):
        ws.cell(row=3, column=col_idx, value=nombre_col)
    _aplicar_estilos_encabezado(ws, num_columnas)

    for row_idx, enlace in enumerate(enlaces, start=4):
        for col_idx, valor in enumerate(funcion_fila(enlace), start=1):
            celda = ws.cell(row=row_idx, column=col_idx, value=valor)
            celda.font = _DATA_FONT
            celda.border = _THIN_BORDER
            celda.alignment = Alignment(vertical="center", wrap_text=True)

    for col_idx in range(1, num_columnas + 1):
        max_len = len(str(columnas[col_idx - 1]))
        for row_idx in range(4, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[ws.cell(row=3, column=col_idx).column_letter].width = min(max_len + 4, 40)

    total_row = ws.max_row + 2
    ws.cell(row=total_row, column=1, value=f"Total de registros: {enlaces.count()}").font = Font(
        name="Calibri", bold=True, size=10, color="333333"
    )

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"enlaces_ccg_{plantilla}_{timestamp}.xlsx"

    response = HttpResponse(
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="{nombre_archivo}"'
    return response


# ---------------------------------------------------------------------------
# Detalle de enlace (AJAX) y guardar comentario
# ---------------------------------------------------------------------------
@requiere(CAP_DIRECTORIO)
def detalle_enlace_json(request, pk):
    """Devuelve los datos de un enlace en formato JSON (para el modal)."""
    enlace = get_object_or_404(
        EnlaceAutorizado.objects.select_related("institucion").prefetch_related(
            "institucion__edificio"
        ),
        pk=pk,
    )
    edificios_nombre = ""
    edificios_siglas = ""
    primer_edificio_id = None
    if enlace.institucion_id:
        edificios = list(enlace.institucion.edificio.all())
        edificios_nombre = ", ".join(ed.nombre for ed in edificios)
        edificios_siglas = ", ".join(ed.siglas or "" for ed in edificios if ed.siglas)
        if edificios:
            primer_edificio_id = edificios[0].pk
    return JsonResponse({
        "id": enlace.pk,
        "nombres": enlace.nombres,
        "primer_apellido": enlace.primer_apellido,
        "segundo_apellido": enlace.segundo_apellido or "",
        "nombre_completo": enlace.nombre_completo,
        "genero": enlace.genero or "",
        "genero_display": enlace.get_genero_display() if enlace.genero else "",
        "correo_principal": enlace.correo_principal or "",
        "correo_secundario": enlace.correo_secundario or "",
        "telefono_principal": enlace.telefono_principal or "",
        "telefono_secundario": enlace.telefono_secundario or "",
        "extension_telefonica": enlace.extension_telefonica or "",
        "estado": enlace.estado,
        "nivel_referencia": enlace.nivel_referencia or "",
        "usuario_sig": enlace.usuario_sig or "",
        "nombre_sig": enlace.nombre_sig or "",
        "pin_sig": enlace.pin_sig or "",
        "password_sig": enlace.password_sig or "",
        "institucion_id": enlace.institucion.pk if enlace.institucion else None,
        "institucion": enlace.institucion.nombre if enlace.institucion else "",
        "siglas": enlace.institucion.siglas or "" if enlace.institucion else "",
        "edificio_id": primer_edificio_id,
        "edificio": edificios_nombre,
        "edificio_siglas": edificios_siglas,
        "sincronizado": enlace.sincronizado,
        "comentarios": enlace.comentarios or "",
        "fecha_alta": enlace.fecha_alta.isoformat() if enlace.fecha_alta else "",
        "oficio_alta": enlace.oficio_alta or "",
        "observaciones_alta": enlace.observaciones_alta or "",
        "fecha_seguimiento": enlace.fecha_seguimiento.isoformat() if enlace.fecha_seguimiento else "",
        "oficio_seguimiento": enlace.oficio_seguimiento or "",
        "observaciones_seguimiento": enlace.observaciones_seguimiento or "",
        "fecha_baja": enlace.fecha_baja.isoformat() if enlace.fecha_baja else "",
        "oficio_baja": enlace.oficio_baja or "",
        "observaciones_baja": enlace.observaciones_baja or "",
    })


@requiere(CAP_DIRECTORIO)
def enlace_detalle(request, pk):
    """Ficha completa de un enlace con pestañas (datos, ubicación, contacto,
    SIG, alta/seguimiento/baja y tickets vinculados)."""
    enlace = get_object_or_404(
        EnlaceAutorizado.objects.select_related("institucion").prefetch_related(
            "institucion__edificio",
            "adjuntos",
            "sync_logs",
        ),
        pk=pk,
    )

    niveles_edificios = []
    if enlace.institucion_id:
        niveles_edificios = list(
            InstitucionEdificio.objects.filter(institucion=enlace.institucion)
            .select_related("edificio")
        )

    tickets = Ticket.objects.none()
    page_obj = None
    total_tickets = abiertos = cerrados = 0
    filtros_ctx = {}
    filtros = ""
    if tiene(request.user, CAP_TICKETS):
        tickets = (
            Ticket.objects.filter(solicitante=enlace)
            .select_related("torre", "institucion", "servicio", "falla", "nivel")
            .order_by("-fecha")
        )
        total_tickets = tickets.count()
        abiertos = tickets.filter(estatus=Ticket.ESTATUS_ABIERTO).count()
        cerrados = total_tickets - abiertos

        tickets, filtros_ctx = _filtrar_tickets(tickets, request.GET)
        paginator = Paginator(tickets, 50)
        try:
            pagina = int(request.GET.get("page", "1"))
        except (TypeError, ValueError):
            pagina = 1
        page_obj = paginator.get_page(pagina)
        filtros = "&".join(
            f"{k}={quote(v)}"
            for k, v in request.GET.items()
            if k != "page" and v.strip() != ""
        )

    return render(
        request,
        "enlaces_ccg/enlace_detalle.html",
        {
            "enlace": enlace,
            "niveles_edificios": niveles_edificios,
            "page_obj": page_obj,
            "total_tickets": total_tickets,
            "tickets_abiertos": abiertos,
            "tickets_cerrados": cerrados,
            "filtros": filtros,
            "edificios": Edificio.objects.all(),
            "servicios": Servicio.objects.filter(activo=True),
            "fallas": Falla.objects.filter(activo=True),
            "niveles": Nivel.objects.filter(activo=True),
            **(filtros_ctx or {}),
        },
    )


@require_POST
@requiere(CAP_EDITAR)
def guardar_comentario(request, pk):
    """Guarda el comentario de un enlace (solo usuarios autenticados)."""
    if not request.user.is_authenticated:
        return JsonResponse({"ok": False, "error": "No autenticado"}, status=403)

    enlace = get_object_or_404(EnlaceAutorizado, pk=pk)
    comentario = request.POST.get("comentarios", "").strip()
    enlace.comentarios = comentario
    enlace.save(update_fields=["comentarios"])

    return JsonResponse({"ok": True, "comentarios": enlace.comentarios})


@require_POST
@requiere(CAP_EDITAR)
def editar_enlace(request, pk):
    """Edita todos los campos de un enlace (solo usuarios autenticados)."""
    if not request.user.is_authenticated:
        return JsonResponse({"ok": False, "error": "No autenticado"}, status=403)

    enlace = get_object_or_404(EnlaceAutorizado, pk=pk)
    data = request.POST

    campos = {
        "nombres": data.get("nombres", "").strip(),
        "primer_apellido": data.get("primer_apellido", "").strip(),
        "segundo_apellido": data.get("segundo_apellido", "").strip(),
        "genero": data.get("genero", ""),
        "correo_principal": data.get("correo_principal", "").strip(),
        "correo_secundario": data.get("correo_secundario", "").strip(),
        "telefono_principal": data.get("telefono_principal", "").strip(),
        "telefono_secundario": data.get("telefono_secundario", "").strip(),
        "extension_telefonica": data.get("extension_telefonica", "").strip(),
        "usuario_sig": data.get("usuario_sig", "").strip(),
        "password_sig": data.get("password_sig", "").strip(),
        "nombre_sig": data.get("nombre_sig", "").strip(),
        "pin_sig": data.get("pin_sig", "").strip(),
        "nivel_referencia": data.get("nivel_referencia", "").strip(),
        "sincronizado": data.get("sincronizado") == "1",
        "estado": data.get("estado", enlace.estado),
        "comentarios": data.get("comentarios", "").strip(),
        "oficio_alta": data.get("oficio_alta", "").strip(),
        "observaciones_alta": data.get("observaciones_alta", "").strip(),
        "oficio_seguimiento": data.get("oficio_seguimiento", "").strip(),
        "observaciones_seguimiento": data.get("observaciones_seguimiento", "").strip(),
        "oficio_baja": data.get("oficio_baja", "").strip(),
        "observaciones_baja": data.get("observaciones_baja", "").strip(),
    }

    for f in ("fecha_alta", "fecha_seguimiento", "fecha_baja"):
        val = data.get(f, "").strip()
        campos[f] = val if val else None

    institucion_id = data.get("institucion_id", "").strip()
    if institucion_id:
        try:
            campos["institucion"] = Institucion.objects.get(pk=int(institucion_id))
        except (Institucion.DoesNotExist, ValueError, TypeError):
            campos["institucion"] = None
    else:
        # Institución es opcional: se puede dejar en blanco
        campos["institucion"] = None

    for campo, valor in campos.items():
        setattr(enlace, campo, valor)

    try:
        enlace.full_clean()
        enlace.save()
    except Exception as e:
        return JsonResponse({"ok": False, "error": str(e)}, status=400)

    return JsonResponse({"ok": True})


@require_POST
@requiere(CAP_EDITAR)
def crear_enlace(request):
    """Crea un nuevo enlace autorizado (solo usuarios autenticados)."""
    if not request.user.is_authenticated:
        return JsonResponse({"ok": False, "error": "No autenticado"}, status=403)

    data = request.POST

    campos = {
        "nombres": data.get("nombres", "").strip(),
        "primer_apellido": data.get("primer_apellido", "").strip(),
        "segundo_apellido": data.get("segundo_apellido", "").strip(),
        "genero": data.get("genero", ""),
        "correo_principal": data.get("correo_principal", "").strip(),
        "correo_secundario": data.get("correo_secundario", "").strip(),
        "telefono_principal": data.get("telefono_principal", "").strip(),
        "telefono_secundario": data.get("telefono_secundario", "").strip(),
        "extension_telefonica": data.get("extension_telefonica", "").strip(),
        "usuario_sig": data.get("usuario_sig", "").strip(),
        "password_sig": data.get("password_sig", "").strip(),
        "nombre_sig": data.get("nombre_sig", "").strip(),
        "pin_sig": data.get("pin_sig", "").strip(),
        "nivel_referencia": data.get("nivel_referencia", "").strip(),
        "estado": data.get("estado", "ACTIVO"),
        "comentarios": data.get("comentarios", "").strip(),
        "oficio_alta": data.get("oficio_alta", "").strip(),
        "observaciones_alta": data.get("observaciones_alta", "").strip(),
        "oficio_seguimiento": data.get("oficio_seguimiento", "").strip(),
        "observaciones_seguimiento": data.get("observaciones_seguimiento", "").strip(),
        "oficio_baja": data.get("oficio_baja", "").strip(),
        "observaciones_baja": data.get("observaciones_baja", "").strip(),
    }

    if not campos["nombres"] or not campos["primer_apellido"]:
        return JsonResponse({"ok": False, "error": "Nombres y primer apellido son obligatorios."}, status=400)

    for f in ("fecha_alta", "fecha_seguimiento", "fecha_baja"):
        val = data.get(f, "").strip()
        campos[f] = val if val else None

    institucion_id = data.get("institucion_id", "").strip()
    if institucion_id:
        try:
            campos["institucion"] = Institucion.objects.get(pk=int(institucion_id))
        except (Institucion.DoesNotExist, ValueError, TypeError):
            campos["institucion"] = None
    else:
        campos["institucion"] = None

    try:
        enlace = EnlaceAutorizado(**campos)
        enlace.full_clean()
        enlace.save()
    except Exception as e:
        return JsonResponse({"ok": False, "error": str(e)}, status=400)

    return JsonResponse({"ok": True, "id": enlace.pk})


@requiere(CAP_DIRECTORIO)
def instituciones_json(request):
    """Devuelve lista de instituciones para el select del modal de edición."""
    term = request.GET.get("q", "").strip()
    qs = Institucion.objects.prefetch_related("edificio").order_by("nombre")
    if term:
        qs = qs.filter(Q(nombre__icontains=term) | Q(siglas__icontains=term))
    resultados = []
    for inst in qs[:50]:
        edificio_parts = []
        for ed in inst.edificio.all():
            edificio_parts.append(
                f"{ed.nombre} ({ed.siglas})" if ed.siglas else ed.nombre
            )
        edificio = ", ".join(edificio_parts)
        resultados.append({
            "id": inst.pk,
            "nombre": inst.nombre,
            "siglas": inst.siglas or "",
            "edificio": edificio,
        })
    return JsonResponse({"instituciones": resultados})


# ---------------------------------------------------------------------------
# Detalle de institución (AJAX) para el modal de edición
# ---------------------------------------------------------------------------
@requiere(CAP_DIRECTORIO)
def detalle_institucion_json(request, pk):
    """Devuelve los datos de una institución en formato JSON (para el modal)."""
    if not request.user.is_authenticated:
        return JsonResponse({"ok": False, "error": "No autenticado"}, status=403)

    inst = get_object_or_404(
        Institucion.objects.prefetch_related("instituciones_edificios__edificio"),
        pk=pk,
    )
    edificios = list(inst.instituciones_edificios.select_related("edificio"))
    edificio_ids = [ie.edificio_id for ie in edificios]
    niveles_edificios = {
        ie.edificio_id: ie.nivel or []
        for ie in edificios
    }
    return JsonResponse({
        "id": inst.pk,
        "nombre": inst.nombre,
        "siglas": inst.siglas or "",
        "estado": inst.estado,
        "edificio_ids": edificio_ids,
        "niveles_edificios": niveles_edificios,
        "contacto_nombre": inst.contacto_nombre or "",
        "contacto_correo": inst.contacto_correo or "",
        "contacto_telefono": inst.contacto_telefono or "",
    })


@requiere(CAP_DIRECTORIO)
def edificios_json(request):
    """Devuelve lista de edificios para el select del modal de institución."""
    if not request.user.is_authenticated:
        return JsonResponse({"ok": False, "error": "No autenticado"}, status=403)

    edificios = []
    for ed in Edificio.objects.order_by("nombre"):
        label = ed.nombre
        if ed.siglas:
            label = f"{label} ({ed.siglas})"
        edificios.append({"id": ed.pk, "nombre": label, "siglas": ed.siglas or ""})
    return JsonResponse({"edificios": edificios})


@require_POST
@requiere(CAP_EDITAR)
def crear_institucion(request):
    """Crea una nueva institución (solo usuarios autenticados)."""
    if not request.user.is_authenticated:
        return JsonResponse({"ok": False, "error": "No autenticado"}, status=403)

    data = request.POST
    inst = Institucion()
    _aplicar_datos_institucion(inst, data)

    try:
        inst.full_clean()
        inst.save()
        _aplicar_edificios_niveles(inst)
    except Exception as e:
        return JsonResponse({"ok": False, "error": str(e)}, status=400)

    return JsonResponse({"ok": True, "id": inst.pk})


@require_POST
@requiere(CAP_EDITAR)
def editar_institucion(request, pk):
    """Edita una institución (solo usuarios autenticados)."""
    if not request.user.is_authenticated:
        return JsonResponse({"ok": False, "error": "No autenticado"}, status=403)

    inst = get_object_or_404(Institucion, pk=pk)
    _aplicar_datos_institucion(inst, request.POST)

    try:
        inst.full_clean()
        inst.save()
        _aplicar_edificios_niveles(inst)
    except Exception as e:
        return JsonResponse({"ok": False, "error": str(e)}, status=400)

    return JsonResponse({"ok": True, "id": inst.pk})


def _aplicar_edificios_niveles(inst):
    """Aplica la relación M2M Institucion ↔ Edificio con sus niveles,
    tras que la institución ya está guardada (tiene PK)."""
    nuevos = inst._nuevos_edificios_niveles

    # Determinar edificios que se mantienen (con niveles actualizados)
    ids_mantenidos = {item["edificio_id"] for item in nuevos}

    # Eliminar relaciones de edificios ya no presentes
    InstitucionEdificio.objects.filter(institucion=inst).exclude(
        edificio_id__in=ids_mantenidos
    ).delete()

    # Actualizar o crear las relaciones
    for item in nuevos:
        InstitucionEdificio.objects.update_or_create(
            institucion=inst,
            edificio_id=item["edificio_id"],
            defaults={"nivel": item["nivel"]},
        )


def _aplicar_datos_institucion(inst, data, request_user=None):
    """Aplica los campos del form de institución a la instancia."""
    inst.nombre = data.get("nombre", "").strip()
    inst.siglas = data.get("siglas", "").strip()
    inst.estado = data.get("estado", inst.estado or "ACTIVO")
    inst.contacto_nombre = data.get("contacto_nombre", "").strip()
    inst.contacto_correo = data.get("contacto_correo", "").strip()
    inst.contacto_telefono = data.get("contacto_telefono", "").strip()

    # Niveles por edificio: el form envía `edificios_niveles` como JSON string
    # con el formato [{"edificio_id": 1, "nivel": ["PB","Nivel 1"]}, ...]
    edificios_niveles = []
    raw = (data.get("edificios_niveles", "") or "").strip()
    if raw:
        try:
            edificios_niveles = json.loads(raw)
        except (ValueError, TypeError):
            edificios_niveles = []

    # Compatibilidad con el campo simple `edificio` (un solo id) + `nivel`
    if not edificios_niveles:
        edificio_id = data.get("edificio", "").strip()
        nivel_raw = (data.get("nivel", "") or "").strip()
        nivel_lista = []
        if nivel_raw:
            try:
                nivel_lista = json.loads(nivel_raw)
            except (ValueError, TypeError):
                nivel_lista = []
        if edificio_id:
            try:
                edificios_niveles = [{
                    "edificio_id": int(edificio_id),
                    "nivel": nivel_lista,
                }]
            except (ValueError, TypeError):
                edificios_niveles = []

    # Reconstruir la relación M2M con sus niveles
    nuevas = []
    for item in edificios_niveles:
        try:
            ed_id = int(item.get("edificio_id"))
        except (TypeError, ValueError):
            continue
        if ed_id <= 0:
            continue
        nivel = item.get("nivel") or []
        if not isinstance(nivel, list):
            nivel = []
        nuevas.append({
            "edificio_id": ed_id,
            "nivel": [str(n) for n in nivel],
        })

    inst._nuevos_edificios_niveles = nuevas


@requiere(CAP_DIRECTORIO)
def enlaces_activos_institucion_json(request, pk):
    """Devuelve la tabla de enlaces activos de una institución (para el modal)."""
    inst = get_object_or_404(Institucion, pk=pk)
    enlaces = EnlaceAutorizado.objects.filter(
        institucion=inst, estado="ACTIVO"
    ).order_by("primer_apellido", "segundo_apellido", "nombres")
    filas = []
    for e in enlaces:
        filas.append({
            "id": e.pk,
            "nombre": e.nombre_completo,
            "correo": e.correo_principal or "",
            "telefono": e.telefono_principal or "",
            "estado": e.estado,
            "estado_display": e.get_estado_display(),
            "sincronizado": e.sincronizado,
        })
    ed_niveles = list(
        InstitucionEdificio.objects.filter(institucion=inst).select_related("edificio")
    )
    edificios_nombre = ", ".join(ie.edificio.nombre for ie in ed_niveles)
    return JsonResponse({
        "institucion": inst.nombre_completo,
        "siglas": inst.siglas or "",
        "edificio": edificios_nombre,
        "edificios_niveles": [
            {"edificio_id": ie.edificio_id, "nombre": ie.edificio.nombre, "niveles": ie.nivel or []}
            for ie in ed_niveles
        ],
        "telefono_institucion": inst.contacto_telefono or "",
        "enlaces": filas,
        "total": enlaces.count(),
    })


# ---------------------------------------------------------------------------
# Adjuntos de enlaces
# ---------------------------------------------------------------------------
@requiere(CAP_DIRECTORIO)
def adjuntos_json(request, pk):
    """Lista los adjuntos de un enlace."""
    enlace = get_object_or_404(EnlaceAutorizado, pk=pk)
    adjuntos = []
    for adj in enlace.adjuntos.all():
        adjuntos.append({
            "id": adj.pk,
            "nombre": adj.nombre or adj.archivo.name.split("/")[-1],
            "descripcion": adj.descripcion or "",
            "archivo_url": adj.archivo.url,
            "archivo_nombre": adj.archivo.name.split("/")[-1],
            "tamano": adj.tamano_legible,
            "icono": adj.icono,
            "subido_en": adj.subido_en.strftime("%d/%m/%Y %H:%M"),
        })
    return JsonResponse({"adjuntos": adjuntos})


@requiere(CAP_EDITAR)
@require_POST
def subir_adjunto(request, pk):
    """Sube un archivo adjunto a un enlace."""
    if not request.user.is_authenticated:
        return JsonResponse({"ok": False, "error": "No autenticado"}, status=403)

    enlace = get_object_or_404(EnlaceAutorizado, pk=pk)
    archivo = request.FILES.get("archivo")
    if not archivo:
        return JsonResponse({"ok": False, "error": "No se envió ningún archivo"}, status=400)

    nombre = request.POST.get("nombre", "").strip() or archivo.name
    descripcion = request.POST.get("descripcion", "").strip()

    adjunto = Adjunto.objects.create(
        enlace=enlace,
        archivo=archivo,
        nombre=nombre,
        descripcion=descripcion,
    )
    return JsonResponse({
        "ok": True,
        "adjunto": {
            "id": adjunto.pk,
            "nombre": adjunto.nombre,
            "descripcion": adjunto.descripcion,
            "archivo_url": adjunto.archivo.url,
            "archivo_nombre": adjunto.archivo.name.split("/")[-1],
            "tamano": adjunto.tamano_legible,
            "icono": adjunto.icono,
            "subido_en": adjunto.subido_en.strftime("%d/%m/%Y %H:%M"),
        },
    })


@requiere(CAP_EDITAR)
@require_POST
def eliminar_adjunto(request, pk, adjunto_id):
    """Elimina un adjunto."""
    if not request.user.is_authenticated:
        return JsonResponse({"ok": False, "error": "No autenticado"}, status=403)

    adjunto = get_object_or_404(Adjunto, pk=adjunto_id, enlace__pk=pk)
    adjuento_archivo = adjunto.archivo
    adjunto.delete()
    try:
        adjuento_archivo.delete(save=False)
    except Exception:
        pass
    return JsonResponse({"ok": True})


@requiere(CAP_DIRECTORIO)
def niveles_edificio(request):
    """Devuelve los niveles disponibles para un edificio dado."""
    edificio_id = request.GET.get("edificio_id")
    if not edificio_id:
        return JsonResponse({"niveles": ["PB"]})
    try:
        edificio = Edificio.objects.get(pk=int(edificio_id))
        return JsonResponse({"niveles": edificio.niveles_disponibles()})
    except (Edificio.DoesNotExist, ValueError):
        return JsonResponse({"niveles": ["PB"]})


# ---------------------------------------------------------------------------
# Importación de enlaces desde Excel
# ---------------------------------------------------------------------------

# Mapeo de encabezados Excel → campos del modelo
_MAPA_CAMPOS = {
    "nombres": "nombres",
    "primer apellido": "primer_apellido",
    "pr. apellido": "primer_apellido",
    "segundo apellido": "segundo_apellido",
    "seg. apellido": "segundo_apellido",
    "genero": "genero",
    "género": "genero",
    "correo principal": "correo_principal",
    "correo secundario": "correo_secundario",
    "telefono principal": "telefono_principal",
    "telefono 1": "telefono_principal",
    "teléfono 1": "telefono_principal",
    "telefono secundario": "telefono_secundario",
    "telefono 2": "telefono_secundario",
    "teléfono 2": "telefono_secundario",
    "extension": "extension_telefonica",
    "extensión": "extension_telefonica",
    "extension ccg": "extension_telefonica",
    "estado": "estado",
    "estatus": "estado",
    "nivel referencia": "nivel_referencia",
    "nivel ref.": "nivel_referencia",
    "referencia_nivel": "nivel_referencia",
    "usuario sig": "usuario_sig",
    "password sig": "password_sig",
    "nombre sig": "nombre_sig",
    "pin sig": "pin_sig",
    "institucion": "institucion_nombre",
    "institución": "institucion_nombre",
    "siglas inst.": "institucion_siglas",
    "siglas": "institucion_siglas",
    "abreviatura (institución) (mao_instituciones)": "institucion_siglas",
    "edificio": "edificio_nombre",
    "unidad funcional": "edificio_nombre",
    "mao_unidfuncionalautonoma": "edificio_nombre",
    "fecha alta": "fecha_alta",
    "oficio alta": "oficio_alta",
    "observ. alta": "observaciones_alta",
    "observaciones alta": "observaciones_alta",
    "fecha seguimiento": "fecha_seguimiento",
    "oficio seguimiento": "oficio_seguimiento",
    "observ. seguimiento": "observaciones_seguimiento",
    "observaciones seguimiento": "observaciones_seguimiento",
    "fecha baja": "fecha_baja",
    "oficio baja": "oficio_baja",
    "observ. baja": "observaciones_baja",
    "observaciones baja": "observaciones_baja",
    "comentarios": "comentarios",
}

_GENERO_MAP = {"m": "M", "masculino": "M", "f": "F", "femenino": "F", "o": "O", "otro": "O"}
_ESTADO_MAP = {"activo": "ACTIVO", "inactivo": "INACTIVO", "act": "ACTIVO", "inact": "INACTIVO"}


def _parsear_fecha(valor):
    """Intenta parsear una fecha desde distintos formatos comunes."""
    if not valor:
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if hasattr(valor, "date"):
        return valor.date()
    s = str(valor).strip()
    for fmt in (
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d-%m-%Y",
        "%d/%m/%y",
        "%d/%m/%Y %H:%M:%S",
        "%d-%m-%Y %H:%M:%S",
        "%m/%d/%Y",
        "%m/%d/%y",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _normalizar_genero(valor):
    if not valor:
        return ""
    return _GENERO_MAP.get(str(valor).strip().lower(), "")


def _normalizar_estado(valor):
    if not valor:
        return "ACTIVO"
    return _ESTADO_MAP.get(str(valor).strip().lower(), "ACTIVO")


@requiere(CAP_IMPORTAR)
def importar_enlaces(request):
    """
    Importa enlaces desde un archivo Excel subido.
    GET: muestra formulario
    POST: procesa archivo y muestra resultados
    """
    if not request.user.is_authenticated:
        from django.contrib.admin.sites import site as admin_site
        from django.contrib.auth.views import redirect_to_login
        return redirect_to_login(request.get_full_path())

    resultado = None

    if request.method == "POST":
        archivo = request.FILES.get("archivo")
        if not archivo:
            resultado = {"error": "No se eligió ningún archivo."}
        else:
            try:
                wb = load_workbook(archivo, data_only=True)
                ws = wb.active

                # --- Leer encabezados (fila 1) ---
                encabezados_raw = []
                for cell in ws[1]:
                    encabezados_raw.append(str(cell.value or "").strip())

                # Mapear columnas
                mapa_col = {}
                for idx, enc in enumerate(encabezados_raw):
                    clave = enc.lower().strip()
                    if clave in _MAPA_CAMPOS:
                        mapa_col[_MAPA_CAMPOS[clave]] = idx

                # Verificar campos obligatorios
                obligatorios = {"nombres", "primer_apellido", "institucion_nombre"}
                faltantes = obligatorios - set(mapa_col.keys())
                if faltantes:
                    resultado = {
                        "error": f"Faltan columnas obligatorias: {', '.join(faltantes)}. "
                                 f"Encabezados encontrados: {', '.join(encabezados_raw)}"
                    }
                else:
                    # --- Procesar filas ---
                    creados = 0
                    actualizados = 0
                    errores = []
                    instituciones_cache = {}

                    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                        if not any(row):
                            continue

                        def _val(campo):
                            # Nota: NO se aplica strip. En la importación los
                            # datos se guardan tal cual vienen del Excel.
                            idx = mapa_col.get(campo)
                            if idx is None or idx >= len(row):
                                return ""
                            v = row[idx]
                            return str(v) if v is not None else ""

                        def _val_crudo(campo):
                            idx = mapa_col.get(campo)
                            if idx is None or idx >= len(row):
                                return None
                            return row[idx]

                        nombres = _val("nombres")[:150]
                        pr_apellido = _val("primer_apellido")[:100]
                        if not nombres or not pr_apellido:
                            errores.append(f"Fila {row_idx}: nombres o primer apellido vacíos, omitida.")
                            continue

                        seg_apellido = _val("segundo_apellido")[:100]
                        inst_nombre = _val("institucion_nombre")
                        inst_siglas = _val("institucion_siglas")
                        edificio_nombre = _val("edificio_nombre")

                        # Buscar o crear institución + edificios (el nombre es el
                        # identificador; las siglas pueden repetirse, p.ej. SEFIN).
                        # El campo edificio puede traer varios separados por ';'.
                        inst_key = inst_nombre.upper() if inst_nombre else inst_siglas.upper()
                        if inst_key not in instituciones_cache:
                            institucion = None
                            if inst_nombre:
                                institucion = Institucion.objects.filter(
                                    nombre__iexact=inst_nombre
                                ).first()
                            if not institucion and inst_siglas:
                                institucion = Institucion.objects.filter(
                                    siglas__iexact=inst_siglas
                                ).first()

                            edificios_objs = []
                            if edificio_nombre:
                                for parte in [p.strip() for p in edificio_nombre.split(";") if p.strip()]:
                                    nombre_edificio_trunc = parte[:100]
                                    edificio_obj, _ = Edificio.objects.get_or_create(
                                        nombre__iexact=nombre_edificio_trunc,
                                        defaults={"nombre": nombre_edificio_trunc},
                                    )
                                    if edificio_obj not in edificios_objs:
                                        edificios_objs.append(edificio_obj)

                            if not institucion:
                                institucion = Institucion.objects.create(
                                    nombre=inst_nombre[:200],
                                    siglas=inst_siglas[:60],
                                )

                            # Vincular edificios a la institución (M2M through)
                            for edificio_obj in edificios_objs:
                                InstitucionEdificio.objects.get_or_create(
                                    institucion=institucion,
                                    edificio=edificio_obj,
                                )
                            instituciones_cache[inst_key] = institucion

                        institucion = instituciones_cache[inst_key]

                        # Buscar enlace existente por nombres + primer_apellido + institución
                        enlace = EnlaceAutorizado.objects.filter(
                            nombres__iexact=nombres,
                            primer_apellido__iexact=pr_apellido,
                            institucion=institucion,
                        ).first()

                        usuario_sig = _val("usuario_sig")[:50]

                        # No permitir usuarios SIG que ya existan en otro enlace
                        if usuario_sig:
                            existente_us_sig = EnlaceAutorizado.objects.filter(
                                usuario_sig__iexact=usuario_sig
                            ).first()
                            if existente_us_sig and (
                                not enlace or existente_us_sig.id != enlace.id
                            ):
                                errores.append(
                                    f"Fila {row_idx}: el usuario SIG '{usuario_sig}' ya está "
                                    f"asignado a {existente_us_sig.nombre_completo}. Fila omitida."
                                )
                                continue

                        datos = {
                            "nombres": nombres,
                            "primer_apellido": pr_apellido,
                            "segundo_apellido": seg_apellido,
                            "institucion": institucion,
                            "genero": _normalizar_genero(_val("genero")),
                            "correo_principal": _val("correo_principal") or None,
                            "correo_secundario": _val("correo_secundario"),
                            "telefono_principal": _val("telefono_principal")[:20],
                            "telefono_secundario": _val("telefono_secundario")[:20],
                            "extension_telefonica": _val("extension_telefonica")[:10],
                            "estado": _normalizar_estado(_val("estado")),
                            "nivel_referencia": _val("nivel_referencia")[:150],
                            "usuario_sig": usuario_sig,
                            # La importación marca sincronizado por defecto (aunque
                            # no venga en el Excel): los datos ya existen en el SIG.
                            "sincronizado": True,
                            "password_sig": _val("password_sig")[:128],
                            "nombre_sig": _val("nombre_sig")[:100],
                            "pin_sig": _val("pin_sig")[:10],
                            "fecha_alta": _parsear_fecha(_val_crudo("fecha_alta")),
                            "oficio_alta": _val("oficio_alta")[:100],
                            "observaciones_alta": _val("observaciones_alta"),
                            "fecha_seguimiento": _parsear_fecha(_val_crudo("fecha_seguimiento")),
                            "oficio_seguimiento": _val("oficio_seguimiento")[:100],
                            "observaciones_seguimiento": _val("observaciones_seguimiento"),
                            "fecha_baja": _parsear_fecha(_val_crudo("fecha_baja")),
                            "oficio_baja": _val("oficio_baja")[:100],
                            "observaciones_baja": _val("observaciones_baja"),
                            "comentarios": _val("comentarios"),
                        }

                        # La importación NO sincroniza con el SIG: se marca el
                        # flag para que el signal post_save lo ignore.
                        if enlace:
                            # Marcar COPY y no activar flag: usamos atributo temporal
                            enlace._sig_sync_desactivado = True
                            for k, v in datos.items():
                                setattr(enlace, k, v)
                            enlace.save()
                            actualizados += 1
                        else:
                            enlace = EnlaceAutorizado(**datos)
                            enlace._sig_sync_desactivado = True
                            enlace.save()
                            creados += 1

                    resultado = {
                        "ok": True,
                        "creados": creados,
                        "actualizados": actualizados,
                        "errores": errores,
                        "total_errores": len(errores),
                    }

            except Exception as e:
                resultado = {"error": f"Error al procesar el archivo: {str(e)}"}

    return render(
        request,
        "enlaces_ccg/importar_enlaces.html",
        {"resultado": resultado},
    )


@requiere(CAP_IMPORTAR)
def descargar_plantilla_importacion(request):
    """Genera y descarga una plantilla Excel vacía con los encabezados correctos."""
    COLUMNAS = [
        ("Nombres", True),
        ("Primer Apellido", True),
        ("Segundo Apellido", False),
        ("Genero", False),
        ("Correo Principal", False),
        ("Correo Secundario", False),
        ("Telefono 1", False),
        ("Telefono 2", False),
        ("Extension CCG", False),
        ("Estado", False),
        ("Nivel Referencia", False),
        ("Institución", True),
        ("Abreviatura (Institución)", False),
        ("Edificio / Unidad Funcional", False),
        ("Usuario SIG", False),
        ("Password SIG", False),
        ("Nombre SIG", False),
        ("PIN SIG", False),
        ("Fecha Alta (dd/mm/yyyy)", False),
        ("Oficio Alta", False),
        ("Observ. Alta", False),
        ("Fecha Seguimiento (dd/mm/yyyy)", False),
        ("Oficio Seguimiento", False),
        ("Observ. Seguimiento", False),
        ("Fecha Baja (dd/mm/yyyy)", False),
        ("Oficio Baja", False),
        ("Observ. Baja", False),
        ("Comentarios", False),
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Plantilla Enlaces"

    # Estilos
    title_font = Font(name="Calibri", bold=True, size=14, color="0D1B3E")
    subtitle_font = Font(name="Calibri", size=10, color="666666")
    header_fill = PatternFill(start_color="0D1B3E", end_color="0D1B3E", fill_type="solid")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    required_font = Font(name="Calibri", bold=True, color="CC0000", size=11)
    note_font = Font(name="Calibri", italic=True, size=9, color="888888")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    num_col = len(COLUMNAS)

    # Título
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_col)
    ws.cell(row=1, column=1, value="Plantilla de Importación — Enlaces CCG").font = title_font
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_col)
    ws.cell(row=2, column=1, value="Complete los datos y suba el archivo en la sección de importación.").font = subtitle_font
    ws.row_dimensions[2].height = 20

    # Encabezados (fila 3) con indicador * si es obligatorio
    for ci, (nombre, obligatorio) in enumerate(COLUMNAS, 1):
        celda = ws.cell(row=3, column=ci, value=f"* {nombre}" if obligatorio else nombre)
        celda.fill = header_fill
        celda.font = required_font if obligatorio else header_font
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = thin_border

    # Fila de notas (fila 4)
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=num_col)
    ws.cell(row=4, column=1,
            value="* = obligatorio  |  Genero: M / F / O  |  Estado: ACTIVO / INACTIVO  |  Fechas: dd/mm/yyyy").font = note_font

    # Ejemplo (fila 5)
    ejemplo = [
        "Juan Carlos", "Pérez", "López", "M",
        "jperez@gob.hn", "", "22001001", "", "101",
        "ACTIVO", "Nacional", "SEPLAN", "SEPLAN", "Torre 1",
        "jcperez", "****", "Juan Pérez", "1234",
        "01/01/2024", "OFICIO-001/2024", "Alta inicial",
        "", "", "",
        "", "", "",
        "Observación de ejemplo",
    ]
    for ci, val in enumerate(ejemplo, 1):
        celda = ws.cell(row=5, column=ci, value=val)
        celda.font = Font(name="Calibri", size=10, color="999999", italic=True)
        celda.border = thin_border

    # Ajustar anchos
    for ci in range(1, num_col + 1):
        ws.column_dimensions[ws.cell(row=3, column=ci).column_letter].width = min(
            len(str(COLUMNAS[ci - 1][0])) + 6, 30
        )

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    resp = HttpResponse(
        buf.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp["Content-Disposition"] = 'attachment; filename="plantilla_enlaces_ccg.xlsx"'
    return resp


# ---------------------------------------------------------------------------
# Importación de Instituciones desde Excel
# ---------------------------------------------------------------------------
_MAPA_INST = {
    "institucion": "nombre",
    "institución": "nombre",
    "abreviatura": "siglas",
    "abreviatura (institución)": "siglas",
    "abreviatura (institución) (mao_instituciones)": "siglas",
    "siglas": "siglas",
    "edificio": "edificio_nombre",
    "unidad funcional": "edificio_nombre",
    "unidad funcional / edificio": "edificio_nombre",
    "mao_unidfuncionalautonoma": "edificio_nombre",
    "mao_unidadfuncionalautonoma": "edificio_nombre",
    "estado": "estado",
    "estatus": "estado",
}


@requiere(CAP_IMPORTAR)
def importar_instituciones(request):
    """Importa instituciones desde un archivo Excel."""
    if not request.user.is_authenticated:
        from django.contrib.auth.views import redirect_to_login
        return redirect_to_login(request.get_full_path())

    resultado = None

    if request.method == "POST":
        archivo = request.FILES.get("archivo")
        if not archivo:
            resultado = {"error": "No se eligió ningún archivo."}
        else:
            try:
                wb = load_workbook(archivo, data_only=True)
                ws = wb.active

                encabezados_raw = []
                for cell in ws[1]:
                    encabezados_raw.append(str(cell.value or "").strip())

                mapa_col = {}
                for idx, enc in enumerate(encabezados_raw):
                    clave = enc.lower().strip()
                    if clave in _MAPA_INST:
                        mapa_col[_MAPA_INST[clave]] = idx

                if "nombre" not in mapa_col:
                    resultado = {
                        "error": f"Falta la columna obligatoria 'Institución'. "
                                 f"Encabezados encontrados: {', '.join(encabezados_raw)}. "
                                 f"Columnas detectadas: {list(mapa_col.keys())}"
                    }
                else:
                    creadas = 0
                    actualizadas = 0
                    errores = []
                    edificios_cache = {}

                    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                        if not any(row):
                            continue

                        def _val(campo):
                            idx = mapa_col.get(campo)
                            if idx is None or idx >= len(row):
                                return ""
                            v = row[idx]
                            return str(v).strip() if v is not None else ""

                        nombre = _val("nombre")
                        if not nombre:
                            errores.append(f"Fila {row_idx}: nombre vacío, omitida.")
                            continue

                        siglas = _val("siglas")[:60]
                        edificio_nombre = _val("edificio_nombre")
                        estado_raw = _val("estado").strip().lower()

                        # Normalizar estado: ACTIVO/INACTIVO (acepta variantes)
                        _s = estado_raw.replace(" ", "")
                        if _s.startswith("inact") or _s in ("baja", "0", "nocontrata"):
                            estado = "INACTIVO"
                        elif _s and _s not in (
                            "activo", "act", "si", "1", "true", "act",
                        ):
                            # Valor no vacío que no parece "activo" -> inactivo
                            estado = "INACTIVO"
                        else:
                            estado = "ACTIVO"

                        # Buscar edificios (pueden venir varios separados por ';')
                        edificios_objs = []
                        if edificio_nombre:
                            for parte in [p.strip() for p in edificio_nombre.split(";") if p.strip()]:
                                cache_key = parte.upper()
                                if cache_key not in edificios_cache:
                                    # Buscar por nombre exacto
                                    edificio_obj = Edificio.objects.filter(
                                        nombre__iexact=parte
                                    ).first()
                                    # Buscar por nombre que contenga el texto
                                    if not edificio_obj:
                                        edificio_obj = Edificio.objects.filter(
                                            nombre__icontains=parte
                                        ).first()
                                    # Buscar por siglas
                                    if not edificio_obj:
                                        edificio_obj = Edificio.objects.filter(
                                            siglas__iexact=parte
                                        ).first()
                                    # Crear si no existe
                                    if not edificio_obj:
                                        edificio_obj = Edificio.objects.create(
                                            nombre=parte[:100]
                                        )
                                    edificios_cache[cache_key] = edificio_obj
                                edificio_obj = edificios_cache[cache_key]
                                if edificio_obj not in edificios_objs:
                                    edificios_objs.append(edificio_obj)

                        # Buscar institución existente por NOMBRE exacto (el nombre es
                        # el identificador único; las siglas pueden repetirse, p.ej. SEFIN)
                        institucion = Institucion.objects.filter(nombre__iexact=nombre).first()

                        if institucion:
                            institucion.nombre = nombre[:200]
                            if siglas:
                                institucion.siglas = siglas
                            institucion.estado = estado
                            institucion.save()
                            actualizadas += 1
                        else:
                            institucion = Institucion.objects.create(
                                nombre=nombre[:200],
                                siglas=siglas,
                                estado=estado,
                            )
                            creadas += 1

                        # Vincular edificios (M2M through)
                        for edificio_obj in edificios_objs:
                            InstitucionEdificio.objects.get_or_create(
                                institucion=institucion,
                                edificio=edificio_obj,
                            )

                    resultado = {
                        "ok": True,
                        "creadas": creadas,
                        "actualizadas": actualizadas,
                        "errores": errores,
                        "total_errores": len(errores),
                    }

            except Exception as e:
                resultado = {"error": f"Error al procesar el archivo: {str(e)}"}

    return render(
        request,
        "enlaces_ccg/importar_instituciones.html",
        {"resultado": resultado},
    )


@requiere(CAP_IMPORTAR)
def descargar_plantilla_instituciones(request):
    """Genera y descarga una plantilla Excel para importar instituciones."""
    COLUMNAS = [
        ("Institución", True),
        ("Abreviatura", False),
        ("Unidad Funcional / Edificio", False),
        ("Estado", False),
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Plantilla Instituciones"

    title_font = Font(name="Calibri", bold=True, size=14, color="0D1B3E")
    subtitle_font = Font(name="Calibri", size=10, color="666666")
    header_fill = PatternFill(start_color="0D1B3E", end_color="0D1B3E", fill_type="solid")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    required_font = Font(name="Calibri", bold=True, color="CC0000", size=11)
    note_font = Font(name="Calibri", italic=True, size=9, color="888888")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    num_col = len(COLUMNAS)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_col)
    ws.cell(row=1, column=1, value="Plantilla de Importación — Instituciones CCG").font = title_font
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_col)
    ws.cell(row=2, column=1, value="Complete los datos y suba el archivo en la sección de importación.").font = subtitle_font
    ws.row_dimensions[2].height = 20

    for ci, (nombre, obligatorio) in enumerate(COLUMNAS, 1):
        celda = ws.cell(row=3, column=ci, value=f"* {nombre}" if obligatorio else nombre)
        celda.fill = header_fill
        celda.font = required_font if obligatorio else header_font
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = thin_border

    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=num_col)
    ws.cell(row=4, column=1,
            value="* = obligatorio  |  Estado: ACTIVO / INACTIVO (vacío = ACTIVO)").font = note_font

    ejemplo = [
        "Secretaría de Salud", "SESAL", "Torre 1", "ACTIVO",
    ]
    for ci, val in enumerate(ejemplo, 1):
        celda = ws.cell(row=5, column=ci, value=val)
        celda.font = Font(name="Calibri", size=10, color="999999", italic=True)
        celda.border = thin_border

    for ci in range(1, num_col + 1):
        ws.column_dimensions[ws.cell(row=3, column=ci).column_letter].width = min(
            len(str(COLUMNAS[ci - 1][0])) + 6, 35
        )

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    resp = HttpResponse(
        buf.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp["Content-Disposition"] = 'attachment; filename="plantilla_instituciones_ccg.xlsx"'
    return resp


# ---------------------------------------------------------------------------
# Documentos y carpetas
# ---------------------------------------------------------------------------
def _es_staff(user):
    return user.is_authenticated and getattr(user, "is_staff", False)


@requiere(CAP_DIRECTORIO)
def documentos(request):
    """Página de Documentos: lista carpetas con su nº de documentos.

    Visible para todos con acceso a directorio. La gestión (crear/editar/
    eliminar carpetas y subir/eliminar archivos) es solo para staff.
    """
    carpetas = list(
        DocumentoCarpeta.objects.select_related("seccion").annotate(
            num_docs=Count("documentos")
        ).order_by("seccion__nombre", "nombre")
    )
    agrupadas = []
    seccion_actual = ("__ini__", None)
    for carp in carpetas:
        clave = carp.seccion_id
        if clave != seccion_actual[0]:
            seccion_actual = (clave, carp.seccion.nombre if carp.seccion else None)
            agrupadas.append({"seccion": seccion_actual[1], "carpetas": []})
        agrupadas[-1]["carpetas"].append(carp)
    return render(request, "enlaces_ccg/documentos.html", {
        "carpetas": carpetas,
        "carpetas_por_seccion": agrupadas,
        "secciones": list(Seccion.objects.all()),
        "puede_gestionar_documentos": _es_staff(request.user),
        "puede_editar_carpeta": _es_staff(request.user),
    })


@requiere(CAP_DIRECTORIO)
@require_POST
def crear_carpeta(request):
    """Crea una nueva carpeta de documentos (solo staff)."""
    if not _es_staff(request.user):
        return JsonResponse({"ok": False, "error": "Sin permisos"}, status=403)

    nombre = request.POST.get("nombre", "").strip()
    slug = request.POST.get("slug", "").strip().lower()
    descripcion = request.POST.get("descripcion", "").strip()

    if not nombre or not slug:
        return JsonResponse({"ok": False, "error": "Nombre e identificador son obligatorios"}, status=400)
    if not slug.replace("-", "").isalnum():
        return JsonResponse({"ok": False, "error": "Identificador inválido (solo letras, números y guiones)"}, status=400)
    if DocumentoCarpeta.objects.filter(slug=slug).exists():
        return JsonResponse({"ok": False, "error": f"Ya existe una carpeta con el identificador «{slug}»"}, status=400)

    DocumentoCarpeta.objects.create(slug=slug, nombre=nombre, descripcion=descripcion)
    return JsonResponse({"ok": True})


@requiere(CAP_DIRECTORIO)
def detalle_carpeta(request, carpeta_slug):
    """Muestra los documentos de una carpeta con su sección única.

    `secciones` es el catálogo global (para el combobox del modal de editar).
    """
    carpeta = get_object_or_404(DocumentoCarpeta, slug=carpeta_slug)
    documentos = list(carpeta.documentos.select_related("carpeta"))
    return render(request, "enlaces_ccg/documentos.html", {
        "carpeta": carpeta,
        "secciones": list(Seccion.objects.all()),
        "documentos": documentos,
        "puede_gestionar_documentos": _es_staff(request.user),
        "puede_editar_carpeta": _es_staff(request.user),
    })


@requiere(CAP_DIRECTORIO)
@require_POST
def renombrar_carpeta(request, carpeta_slug):
    """Edita nombre, slug, descripción y sección única de una carpeta
    (seccion_id = id del catálogo global, o vacío para ninguna). Solo staff."""
    if not _es_staff(request.user):
        return JsonResponse({"ok": False, "error": "Sin permisos"}, status=403)

    carpeta = get_object_or_404(DocumentoCarpeta, slug=carpeta_slug)
    slug_nuevo = request.POST.get("slug", "").strip().lower()
    nombre = request.POST.get("nombre", "").strip()
    descripcion = request.POST.get("descripcion", "").strip()

    # Validar slug si se intenta cambiar
    if slug_nuevo and slug_nuevo != carpeta.slug:
        if not slug_nuevo.replace("-", "").isalnum():
            return JsonResponse({"ok": False, "error": "Identificador inválido (solo letras, números y guiones)"}, status=400)
        if DocumentoCarpeta.objects.filter(slug=slug_nuevo).exclude(pk=carpeta.pk).exists():
            return JsonResponse({"ok": False, "error": f"Ya existe una carpeta con el identificador «{slug_nuevo}»"}, status=400)
        carpeta.slug = slug_nuevo

    if nombre:
        carpeta.nombre = nombre
    carpeta.descripcion = descripcion

    # Sección única de la carpeta
    seccion_id = request.POST.get("seccion_id", "").strip()
    carpeta.seccion = None
    if seccion_id and seccion_id.isdigit() and Seccion.objects.filter(pk=int(seccion_id)).exists():
        carpeta.seccion_id = int(seccion_id)

    carpeta.save(update_fields=["nombre", "slug", "descripcion", "seccion", "actualizado_en"])
    return JsonResponse({"ok": True, "slug": carpeta.slug})


@requiere(CAP_DIRECTORIO)
@require_POST
def crear_seccion(request):
    """Crea una sección en el catálogo global (solo staff)."""
    if not _es_staff(request.user):
        return JsonResponse({"ok": False, "error": "Sin permisos"}, status=403)

    nombre = request.POST.get("nombre", "").strip()
    if not nombre:
        return JsonResponse({"ok": False, "error": "El nombre de la sección es obligatorio"}, status=400)
    if Seccion.objects.filter(nombre__iexact=nombre).exists():
        return JsonResponse({"ok": False, "error": f"Ya existe una sección «{nombre}»"}, status=400)
    seccion = Seccion.objects.create(nombre=nombre)
    return JsonResponse({"ok": True, "id": seccion.pk, "nombre": seccion.nombre})


@requiere(CAP_DIRECTORIO)
@require_POST
def editar_seccion(request, seccion_id):
    """Edita el nombre de una sección del catálogo (solo staff)."""
    if not _es_staff(request.user):
        return JsonResponse({"ok": False, "error": "Sin permisos"}, status=403)

    seccion = get_object_or_404(Seccion, pk=seccion_id)
    nombre = request.POST.get("nombre", "").strip()
    if not nombre:
        return JsonResponse({"ok": False, "error": "El nombre de la sección es obligatorio"}, status=400)
    if Seccion.objects.filter(nombre__iexact=nombre).exclude(pk=seccion.pk).exists():
        return JsonResponse({"ok": False, "error": f"Ya existe una sección «{nombre}» en el catálogo"}, status=400)
    seccion.nombre = nombre
    seccion.save(update_fields=["nombre"])
    return JsonResponse({"ok": True, "id": seccion.pk, "nombre": seccion.nombre})


@requiere(CAP_DIRECTORIO)
@require_POST
def eliminar_seccion(request, seccion_id):
    """Elimina una sección del catálogo solo si no está en uso (solo staff)."""
    if not _es_staff(request.user):
        return JsonResponse({"ok": False, "error": "Sin permisos"}, status=403)

    seccion = get_object_or_404(Seccion, pk=seccion_id)
    if seccion.carpetas.exists():
        return JsonResponse(
            {"ok": False, "error": f"No se puede eliminar: la sección «{seccion.nombre}» está en uso"},
            status=400,
        )
    seccion.delete()
    return JsonResponse({"ok": True})


@requiere(CAP_DIRECTORIO)
@require_POST
def eliminar_carpeta(request, carpeta_slug):
    """Elimina una carpeta y sus archivos del disco (solo staff)."""
    if not _es_staff(request.user):
        return JsonResponse({"ok": False, "error": "Sin permisos"}, status=403)

    carpeta = get_object_or_404(DocumentoCarpeta, slug=carpeta_slug)
    for doc in carpeta.documentos.all():
        try:
            doc.archivo.delete(save=False)
        except Exception:
            pass
    carpeta.delete()
    return JsonResponse({"ok": True})


@requiere(CAP_DIRECTORIO)
@require_POST
def subir_documento(request, carpeta_slug):
    """Sube un archivo a una carpeta (solo staff)."""
    if not _es_staff(request.user):
        return JsonResponse({"ok": False, "error": "Sin permisos"}, status=403)

    carpeta = get_object_or_404(DocumentoCarpeta, slug=carpeta_slug)
    archivo = request.FILES.get("archivo")
    if not archivo:
        return JsonResponse({"ok": False, "error": "No se envió ningún archivo"}, status=400)

    nombre = request.POST.get("nombre", "").strip() or archivo.name
    doc = Documento.objects.create(
        carpeta=carpeta,
        archivo=archivo,
        nombre=nombre,
    )
    return JsonResponse({
        "ok": True,
        "documento": {
            "id": doc.pk,
            "nombre": doc.nombre,
            "archivo_nombre": doc.nombre_archivo,
            "archivo_url": doc.archivo.url,
            "tamano": doc.tamano_legible,
            "icono": doc.icono,
            "subido_en": doc.subido_en.strftime("%d/%m/%Y %H:%M"),
        },
    })


@requiere(CAP_DIRECTORIO)
@require_POST
def eliminar_documento(request, carpeta_slug, documento_id):
    """Elimina un documento de una carpeta y su archivo del disco (solo staff)."""
    if not _es_staff(request.user):
        return JsonResponse({"ok": False, "error": "Sin permisos"}, status=403)

    doc = get_object_or_404(Documento, pk=documento_id, carpeta__slug=carpeta_slug)
    archivo = doc.archivo
    doc.delete()
    try:
        archivo.delete(save=False)
    except Exception:
        pass
    return JsonResponse({"ok": True})
