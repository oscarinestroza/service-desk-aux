"""Pruebas del parser/importador del módulo de tickets (Fase 2)."""

import tempfile
from datetime import datetime, time, timedelta
from io import BytesIO
from pathlib import Path

import openpyxl
from django.core.management import call_command
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone
from unittest import mock

from .models import (
    ConfiguracionTickets,
    Edificio,
    EnlaceAutorizado,
    Falla,
    Institucion,
    Nivel,
    Servicio,
    Ticket,
)
from .sig_sync.tickets import (
    MODO_SYNC_COMPLETO,
    MODO_SYNC_PARCIAL,
    _calcular_numero_display,
    _normalizar_nombre,
    _parsear_fecha,
    aplicar_tickets,
    parsear_excel_tickets,
)

ENCABEZADOS = [
    "ID Solicitud servicio", "foliosolicitudservicio", "solicitud_solicitante",
    "Responsable_atencion", "solicitud_descripcion", "falla_descripcion",
    "falla_clasificacion", "servicio", "subservicio", "unidad", "area",
    "grupo", "nivel", "solicitud_fecha", "TipoRecepcion",
    "Fecha_TipoRecepcion", "suspencion_fecha", "cerro_fecha", "solicitud_tipo",
    "tiempo_tipo", "FechaHora_Diagnostico", "Diagnostico",
    "FechaHora_Actividades", "Actividades", "FechaHora_Observaciones",
    "Observaciones", "FechaHora_ObservacionesUsuario", "ObservacionesUsuario",
    "Clasificacion_Falla", "Categoria_Falla",
]


def _excel_tmp(filas):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SIG_Seguimiento_Solicitudes"
    ws.append(ENCABEZADOS)
    for f in filas:
        ws.append(f)
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    wb.save(tmp.name)
    tmp.close()
    return Path(tmp.name)


def _fila(tid, numero, fecha="2026-09-01 10:00:00", cierro="", desc="Falla",
          servicio="Elevadores", tipo="Solicitud"):
    return [
        tid, numero, "Solicitante X", "Resp", desc, "", "", servicio,
        "Mantenimiento", "Torre A", "Admin", "Grupo 1", "3", fecha,
        "tel", "", "", cierro, tipo, "No atendida", "", "", "", "", "",
        "", "", "", "", "",
    ]


class ParsearFechasTests(TestCase):
    def test_fecha_iso_devuelve_datetime(self):
        self.assertEqual(
            _parsear_fecha("2026-09-01 10:00:00"),
            timezone.make_aware(datetime(2026, 9, 1, 10, 0, 0)),
        )

    def test_fecha_vacia_devuelve_none(self):
        self.assertIsNone(_parsear_fecha(""))
        self.assertIsNone(_parsear_fecha(None))


class ParsearExcelTicketsTests(TestCase):
    def test_dedupe_ultima_fila_por_ticket_id(self):
        ruta = _excel_tmp([
            _fila(1, "SS26-0010", desc="PRIMERA"),
            _fila(1, "SS26-0010", desc="ULTIMA"),
            _fila(2, "SS26-0011"),
        ])
        datos, errores = parsear_excel_tickets(ruta)
        por_id = {d["ticket_id"]: d for d in datos}
        self.assertEqual(errores, 0)
        self.assertEqual(len(datos), 2)
        self.assertEqual(por_id["1"]["descripcion"], "ULTIMA")

    def test_estatus_derivado_de_cerro_fecha(self):
        ruta = _excel_tmp([
            _fila(1, "SS26-0020", cierro="2026-09-02 12:00:00"),
            _fila(2, "SS26-0021", cierro=""),
        ])
        datos, _ = parsear_excel_tickets(ruta)
        abierto = {d["ticket_id"]: d for d in datos}
        self.assertIsNotNone(abierto["1"]["fecha_cierre"])
        self.assertIsNone(abierto["2"]["fecha_cierre"])

    def test_raw_data_mantiene_columnas_de_contexto(self):
        ruta = _excel_tmp([_fila(1, "SS26-0030", desc="Cabina 2")])
        datos, _ = parsear_excel_tickets(ruta)
        self.assertEqual("Cabina 2", datos[0]["raw_data"]["solicitud_descripcion"])
        self.assertEqual("Elevadores", datos[0]["raw_data"]["servicio"])

    def test_fila_sin_id_se_cuenta_como_error(self):
        ruta = _excel_tmp([
            _fila(1, "SS26-0040"),
            [None] * len(ENCABEZADOS),
        ])
        datos, errores = parsear_excel_tickets(ruta)
        self.assertEqual(len(datos), 1)
        self.assertEqual(errores, 1)


class CalcularNumeroDisplayTests(TestCase):
    def test_sufijo_n_para_duplicados(self):
        datos = [
            {"numero": "SS23-100", "ticket_id": "900", },
            {"numero": "SS23-100", "ticket_id": "901", },
            {"numero": "SS23-100", "ticket_id": "902", },
            {"numero": "SS23-200", "ticket_id": "903", },
            {"numero": "", "ticket_id": "904", },
        ]
        _calcular_numero_display(datos)
        disp = {d["ticket_id"]: d["numero_display"] for d in datos}
        self.assertEqual(disp["900"], "SS23-100")
        self.assertEqual(disp["901"], "SS23-100-2")
        self.assertEqual(disp["902"], "SS23-100-3")
        self.assertEqual(disp["903"], "SS23-200")
        self.assertEqual(disp["904"], "904")

    def test_sufijo_respeta_orden_por_ticket_id(self):
        datos = [
            {"numero": "SS23-300", "ticket_id": "9"},
            {"numero": "SS23-300", "ticket_id": "2"},
        ]
        _calcular_numero_display(datos)
        disp = {d["ticket_id"]: d["numero_display"] for d in datos}
        self.assertEqual(disp["2"], "SS23-300")
        self.assertEqual(disp["9"], "SS23-300-2")


class AplicarTicketsTests(TestCase):
    def _filas_basicas(self, filas):
        ruta = _excel_tmp(filas)
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        return datos

    def test_parcial_no_toca_ausentes(self):
        datos1 = self._filas_basicas([
            _fila(1, "SS26-0100"),
            _fila(2, "SS26-0101"),
        ])
        aplicar_tickets(datos1, MODO_SYNC_PARCIAL)

        datos2 = self._filas_basicas([_fila(1, "SS26-0100")])
        conteos = aplicar_tickets(datos2, MODO_SYNC_PARCIAL)

        self.assertEqual(Ticket.objects.count(), 2)
        self.assertEqual(conteos["archivados"], 0)
        self.assertFalse(Ticket.objects.get(ticket_id="2").archivado)

    def test_completo_archiva_ausentes_y_desarchiva_presentes(self):
        datos1 = self._filas_basicas([
            _fila(1, "SS26-0200"),
            _fila(2, "SS26-0201"),
        ])
        aplicar_tickets(datos1, MODO_SYNC_COMPLETO)
        self.assertEqual(Ticket.objects.count(), 2)
        self.assertEqual(Ticket.objects.filter(archivado=False).count(), 2)

        datos2 = self._filas_basicas([_fila(1, "SS26-0200")])
        conteos = aplicar_tickets(datos2, MODO_SYNC_COMPLETO)

        self.assertEqual(Ticket.objects.count(), 2)  # nunca borra
        self.assertEqual(conteos["archivados"], 1)
        self.assertTrue(Ticket.objects.get(ticket_id="2").archivado)
        self.assertFalse(Ticket.objects.get(ticket_id="1").archivado)

    def test_estatus_en_bd(self):
        datos = self._filas_basicas([
            _fila(1, "SS26-0300", cierro="2026-09-02 12:00:00"),
            _fila(2, "SS26-0301", cierro=""),
        ])
        aplicar_tickets(datos, MODO_SYNC_PARCIAL)
        self.assertEqual(Ticket.objects.get(ticket_id="1").estatus, Ticket.ESTATUS_CERRADO)
        self.assertEqual(Ticket.objects.get(ticket_id="2").estatus, Ticket.ESTATUS_ABIERTO)

    def test_numero_display_persistido(self):
        datos = self._filas_basicas([
            _fila(100, "SS26-0400"),
            _fila(101, "SS26-0400"),
        ])
        aplicar_tickets(datos, MODO_SYNC_PARCIAL)
        self.assertEqual(Ticket.objects.get(ticket_id="100").numero_display, "SS26-0400")
        self.assertEqual(Ticket.objects.get(ticket_id="101").numero_display, "SS26-0400-2")


class VistaTicketsTests(TestCase):
    def setUp(self):
        from django.contrib.auth import get_user_model

        self.usuario = get_user_model().objects.create_superuser(
            "prueba_tickets", "prueba@test.local", "Prueba123!"
        )
        self.client.force_login(self.usuario)
        self.datos = self._filas_basicas([
            _fila(1, "SS26-0500", cierro="2026-09-02 12:00:00"),
            _fila(2, "SS26-0501", cierro=""),
            _fila(3, "SS26-0500", cierro=""),
        ])
        aplicar_tickets(self.datos, MODO_SYNC_PARCIAL)

    def _filas_basicas(self, filas):
        ruta = _excel_tmp(filas)
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        return datos

    def test_lista_tickets_paginada(self):
        respuesta = self.client.get(reverse("enlaces_ccg:seguimiento_tickets"))
        self.assertEqual(respuesta.status_code, 200)
        html = respuesta.content.decode("utf-8", "replace")
        self.assertIn("SS26-0500", html)
        self.assertIn("SS26-0500-2", html)

    def test_filtro_estatus_abiertos(self):
        respuesta = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"), {"estatus": "ABIERTO"}
        )
        html = respuesta.content.decode("utf-8", "replace")
        self.assertIn(">SS26-0501<", html)
        self.assertIn(">SS26-0500-2<", html)
        self.assertNotIn(">SS26-0500<", html)

    def test_detalle_por_pk_y_por_numero(self):
        t = Ticket.objects.get(ticket_id="1")
        r1 = self.client.get(reverse("enlaces_ccg:ticket_detalle", args=[t.pk]))
        self.assertEqual(r1.status_code, 200)
        self.assertIn("Cerrado", r1.content.decode("utf-8", "replace"))

        r2 = self.client.get(reverse("enlaces_ccg:ticket_por_numero", args=["SS26-0500"]))
        self.assertEqual(r2.status_code, 200)

        r3 = self.client.get(reverse("enlaces_ccg:ticket_por_numero", args=["SS26-0500-2"]))
        self.assertEqual(r3.status_code, 200)
        html = r3.content.decode("utf-8", "replace")
        self.assertIn("SS26-0500-2", html)

    def test_por_numero_inexistente_404(self):
        respuesta = self.client.get(
            reverse("enlaces_ccg:ticket_por_numero", args=["NOEXISTE"])
        )
        self.assertEqual(respuesta.status_code, 404)

    def test_sincronizar_post_encola_descarga_rapida(self):
        with mock.patch(
            "enlaces_ccg.sig_sync.tickets.sincronizar_tickets_task"
        ) as tarea:
            respuesta = self.client.post(
                reverse("enlaces_ccg:ticket_sincronizar"), {"modo": "completo"}
            )
            self.assertEqual(respuesta.status_code, 302)
            self.assertEqual(respuesta.url, reverse("enlaces_ccg:seguimiento_tickets"))
            tarea.delay.assert_called_once_with(modo="parcial")

    def test_filtro_estatus_abiertos_lista(self):
        respuesta = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"), {"estatus": "ABIERTO"}
        )
        html = respuesta.content.decode("utf-8", "replace")
        self.assertIn(">SS26-0501<", html)
        self.assertIn(">SS26-0500-2<", html)

    def test_columna_responsable_y_barra_sync(self):
        respuesta = self.client.get(reverse("enlaces_ccg:seguimiento_tickets"))
        html = respuesta.content.decode("utf-8", "replace")
        self.assertIn("Responsable de atención", html)
        self.assertIn('id="barraSync"', html)
        self.assertIn('id="modalTicket"', html)
        self.assertIn("btn-ver-ticket", html)

    def test_ticket_json(self):
        t = Ticket.objects.get(ticket_id="1")
        respuesta = self.client.get(
            reverse("enlaces_ccg:ticket_json", args=[t.pk])
        )
        self.assertEqual(respuesta.status_code, 200)
        datos = respuesta.json()
        self.assertEqual(datos["numero"], "SS26-0500")
        self.assertEqual(datos["estatus"], "CERRADO")
        self.assertIn("raw_data", datos)
        self.assertEqual(
            datos["url_detalle"],
            reverse("enlaces_ccg:ticket_detalle", args=[t.pk]),
        )

    def test_sincronizar_requiere_post(self):
        respuesta = self.client.get(reverse("enlaces_ccg:ticket_sincronizar"))
        self.assertEqual(respuesta.status_code, 405)

    def test_sincronizacion_estado(self):
        with mock.patch(
            "enlaces_ccg.sig_sync.tickets._lock_activo", return_value=False
        ):
            respuesta = self.client.get(
                reverse("enlaces_ccg:ticket_sincronizacion_estado")
            )
        self.assertEqual(respuesta.status_code, 200)
        datos = respuesta.json()
        self.assertIn("ultima_sincronizacion", datos)
        self.assertIn("sincronizando", datos)
        self.assertIn("habilitado", datos)
        self.assertIs(datos["sincronizando"], False)


class FasesProcesoTests(TestCase):
    def setUp(self):
        from django.contrib.auth import get_user_model

        self.usuario = get_user_model().objects.create_superuser(
            "prueba_fases", "pf@test.local", "Prueba123!"
        )
        self.client.force_login(self.usuario)

    def _crear(self, tid, numero, cierro="", actividades="", servicio="Elevadores"):
        ruta = _excel_tmp([
            _fila(tid, numero, cierro=cierro, servicio=servicio),
        ])
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        aplicar_tickets(datos, MODO_SYNC_PARCIAL)
        t = Ticket.objects.get(ticket_id=str(tid))
        if actividades:
            rd = dict(t.raw_data)
            rd["Actividades"] = actividades
            t.raw_data = rd
            t.save(update_fields=["raw_data"])
        return Ticket.objects.get(pk=t.pk)

    def _estado(self, fases, clave):
        return next(f["estado"] for f in fases if f["clave"] == clave)

    def test_abierto_proceso_activo(self):
        t = self._crear(1, "SS26-0600")
        fases = t.fases_proceso()
        self.assertEqual(self._estado(fases, "creado"), "completado")
        self.assertEqual(self._estado(fases, "canalizado"), "completado")
        self.assertEqual(self._estado(fases, "proceso"), "activo")
        self.assertEqual(self._estado(fases, "finalizado"), "pendiente")
        self.assertEqual(self._estado(fases, "completado"), "pendiente")
        detalle_creado = next(f for f in fases if f["clave"] == "creado")["detalle"]
        self.assertIn("Ticket SS26-0600 creado el", detalle_creado)
        self.assertIn("a las", detalle_creado)
        self.assertEqual(
            next(f for f in fases if f["clave"] == "canalizado")["detalle"], "Resp"
        )

    def test_sin_responsable_canalizado_sin_detalle(self):
        ruta = _excel_tmp([
            _fila(1, "SS26-0600", servicio=""),
        ])
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        aplicar_tickets(datos, MODO_SYNC_PARCIAL)
        t = Ticket.objects.get(ticket_id="1")
        rd = dict(t.raw_data)
        rd["Responsable_atencion"] = ""
        t.raw_data = rd
        t.save(update_fields=["raw_data"])
        fases = Ticket.objects.get(pk=t.pk).fases_proceso()
        self.assertEqual(self._estado(fases, "canalizado"), "pendiente")
        self.assertEqual(
            next(f for f in fases if f["clave"] == "canalizado")["detalle"],
            "Sin responsable de atención",
        )

    def test_creado_sin_fecha_cae_al_numero(self):
        t = self._crear(1, "SS26-0607")
        t.fecha = None
        t.save(update_fields=["fecha"])
        fases = Ticket.objects.get(pk=t.pk).fases_proceso()
        self.assertEqual(
            next(f for f in fases if f["clave"] == "creado")["detalle"], "SS26-0607"
        )

    def test_cerrado_finalizado(self):
        t = self._crear(1, "SS26-0601", cierro="2026-09-02 12:00:00")
        fases = t.fases_proceso()
        self.assertEqual(self._estado(fases, "proceso"), "completado")
        self.assertEqual(self._estado(fases, "finalizado"), "completado")
        self.assertEqual(self._estado(fases, "completado"), "pendiente")

    def test_actividades_marca_completado(self):
        t = self._crear(1, "SS26-0602", cierro="2026-09-02 12:00:00",
                        actividades="Cambio de lámpara")
        fases = t.fases_proceso()
        self.assertEqual(self._estado(fases, "completado"), "completado")

    def test_sin_servicio_canalizado_pendiente(self):
        t = self._crear(1, "SS26-0603", servicio="")
        fases = t.fases_proceso()
        self.assertEqual(self._estado(fases, "canalizado"), "pendiente")

    def test_abierta_con_solo_actividades_marca_completado(self):
        t = self._crear(1, "SS26-0604", actividades="Atención inicial")
        fases = t.fases_proceso()
        self.assertEqual(self._estado(fases, "completado"), "completado")
        self.assertEqual(self._estado(fases, "finalizado"), "pendiente")

    def test_timeline_renderiza_en_detalle(self):
        t = self._crear(1, "SS26-0605")
        respuesta = self.client.get(reverse("enlaces_ccg:ticket_detalle", args=[t.pk]))
        self.assertEqual(respuesta.status_code, 200)
        html = respuesta.content.decode("utf-8", "replace")
        self.assertIn("Proceso de atención", html)
        self.assertIn("Ticket creado", html)
        self.assertIn("Canalizado con el servicio", html)
        self.assertIn("En proceso de atención", html)
        self.assertIn("ccg-termino-activo", html)
        self.assertIn("Ticket SS26-0605 creado el", html)
        self.assertIn("Resp", html)
        self.assertIn("Trabajo en curso", html)
        self.assertNotIn("Número visible", html)
        self.assertNotIn("Tipo de solicitud", html)

    def test_comentarios_servicio_segun_estado(self):
        t = self._crear(1, "SS26-0606")
        rd = dict(t.raw_data)
        rd["Observaciones"] = "Falla intermitente"
        t.raw_data = rd
        t.save(update_fields=["raw_data"])
        t = Ticket.objects.get(pk=t.pk)
        co = t.comentarios_servicio()
        self.assertEqual(co["observaciones"], "Falla intermitente")
        self.assertIn("proceso de atención", co["sugerencia"])
        self.assertIn("responsable de atención Resp", co["sugerencia"])

        rd = dict(t.raw_data)
        rd["Actividades"] = "Revisión"
        t.raw_data = rd
        t.save(update_fields=["raw_data"])
        co = Ticket.objects.get(pk=t.pk).comentarios_servicio()
        self.assertIn("actividades registradas", co["sugerencia"])

        t.estatus = Ticket.ESTATUS_CERRADO
        t.fecha_cierre = timezone.now()
        t.save(update_fields=["estatus", "fecha_cierre"])
        co = Ticket.objects.get(pk=t.pk).comentarios_servicio()
        self.assertIn("finalizado", co["sugerencia"])
        self.assertIn("Solicitante X", co["sugerencia"])

    def test_registrar_seguimiento_requiere_descripcion(self):
        t = self._crear(1, "SS26-0608")
        respuesta = self.client.post(
            reverse("enlaces_ccg:ticket_registrar", args=[t.pk]),
            {"descripcion": "  "},
        )
        self.assertEqual(respuesta.status_code, 302)
        self.assertEqual(t.registros.count(), 0)

    def test_registrar_seguimiento_guarda_y_muestra(self):
        t = self._crear(1, "SS26-0609")
        respuesta = self.client.post(
            reverse("enlaces_ccg:ticket_registrar", args=[t.pk]),
            {"descripcion": "Avancé con la atención"},
        )
        self.assertEqual(respuesta.status_code, 302)
        registro = t.registros.first()
        self.assertIsNotNone(registro)
        self.assertEqual(registro.usuario, self.usuario)
        self.assertEqual(registro.descripcion, "Avancé con la atención")

        detalle = self.client.get(
            reverse("enlaces_ccg:ticket_detalle", args=[t.pk])
        )
        html = detalle.content.decode("utf-8", "replace")
        self.assertIn("Avancé con la atención", html)
        self.assertIn("Registrar seguimiento", html)
        self.assertIn("Comentarios del servicio", html)


class ProximaSincronizacionTests(TestCase):
    def setUp(self):
        self.cfg = ConfiguracionTickets.cargar()
        self.cfg.intervalo_minutos = 5
        self.cfg.descarga_completa_horas = []
        self.cfg.habilitado = True

    def test_deshabilitado_devuelve_none(self):
        self.cfg.habilitado = False
        self.assertIsNone(self.cfg.proxima_sincronizacion())

    def test_sin_ultima_es_inmediata(self):
        self.cfg.ultima_sincronizacion = None
        ahora = timezone.now()
        proxima = self.cfg.proxima_sincronizacion(ahora=ahora)
        self.assertEqual(proxima, ahora)

    def test_parcial_usa_intervalo(self):
        ahora = timezone.now()
        self.cfg.ultima_sincronizacion = ahora - timedelta(minutes=2)
        proxima = self.cfg.proxima_sincronizacion(ahora=ahora)
        self.assertEqual(
            proxima, self.cfg.ultima_sincronizacion + timedelta(minutes=5)
        )

    def test_hora_completa_adelanta_la_proxima(self):
        ahora = timezone.localtime(timezone.now()).replace(
            minute=0, second=0, microsecond=0
        )
        self.cfg.ultima_sincronizacion = ahora
        self.cfg.intervalo_minutos = 120
        self.cfg.descarga_completa_horas = [
            (ahora + timedelta(hours=1)).strftime("%H:%M")
        ]
        proxima = self.cfg.proxima_sincronizacion(ahora=ahora)
        self.assertEqual(proxima, ahora + timedelta(hours=1))

    def test_hora_ya_pasada_va_al_dia_siguiente(self):
        ahora = timezone.localtime(timezone.now()).replace(
            minute=30, second=0, microsecond=0
        )
        self.cfg.ultima_sincronizacion = ahora
        self.cfg.intervalo_minutos = 1800
        self.cfg.descarga_completa_horas = ["00:00"]
        proxima = self.cfg.proxima_sincronizacion(ahora=ahora)
        manana = timezone.make_aware(
            datetime.combine((ahora + timedelta(days=1)).date(), time(0, 0)),
            timezone.get_current_timezone(),
        )
        self.assertEqual(proxima, manana)


def _fila_v(tid, numero, solicitante="Cesar Augusto Zavala",
            servicio="Elevadores", falla="Falla Uno",
            nivel="Torre 2", grupo="Nivel 8"):
    """Fila de Excel con datos de vínculos (Fase 3)."""
    fila = [""] * len(ENCABEZADOS)
    fila[0] = tid
    fila[1] = numero
    fila[2] = solicitante
    fila[3] = "Resp"
    fila[4] = "Falla"
    fila[5] = falla
    fila[7] = servicio
    fila[11] = grupo
    fila[12] = nivel
    fila[13] = "2026-09-01 10:00:00"
    return fila


class NormalizarNombreTests(TestCase):
    def test_colapsa_espacios_acentos_y_mayusculas(self):
        self.assertEqual(
            _normalizar_nombre("  Cesar   Augusto Zavala "),
            "CESAR AUGUSTO ZAVALA",
        )
        self.assertEqual(
            _normalizar_nombre("César  Augusto Zavála"),
            _normalizar_nombre("CESAR AUGUSTO ZAVALA"),
        )

    def test_vacio(self):
        self.assertEqual(_normalizar_nombre(None), "")


class VinculacionTests(TestCase):
    def setUp(self):
        self.edificio = Edificio.objects.create(nombre="TORRE 2", siglas="T2")
        self.institucion = Institucion.objects.create(
            nombre="Institución Uno", siglas="IUNO"
        )
        self.enlace = EnlaceAutorizado.objects.create(
            nombres="César Augusto",
            primer_apellido="Zavala",
            nombre_sig="Cesar Augusto Zavala",
            institucion=self.institucion,
            estado="ACTIVO",
        )

    def _importar(self, filas, modo=MODO_SYNC_PARCIAL):
        ruta = _excel_tmp(filas)
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        aplicar_tickets(datos, modo)
        return datos

    def test_resuelve_vinculos(self):
        self._importar([_fila_v(1, "SS26-0700")])
        t = Ticket.objects.get(ticket_id="1")
        self.assertEqual(t.torre, self.edificio)  # case-insensitive TORRE 2
        self.assertEqual(t.institucion, self.institucion)  # del enlace
        self.assertEqual(t.solicitante, self.enlace)
        self.assertEqual(t.nivel.nombre, "Nivel 8")
        self.assertEqual(t.servicio.nombre, "Elevadores")
        self.assertEqual(t.falla.descripcion, "Falla Uno")
        self.assertEqual(t.solicitante_nombre, "")

    def test_sin_empate_guarda_snapshot(self):
        self._importar([_fila_v(1, "SS26-0701", solicitante="Nadie Desconocido")])
        t = Ticket.objects.get(ticket_id="1")
        self.assertIsNone(t.solicitante)
        self.assertIsNone(t.institucion)
        self.assertEqual(t.solicitante_nombre, "Nadie Desconocido")

    def test_catalogos_get_or_create_sin_duplicar(self):
        self._importar([
            _fila_v(1, "SS26-0702", servicio="Elevadores"),
            _fila_v(2, "SS26-0703", servicio="elevadores",
                    falla="Falla Uno"),
        ])
        self.assertEqual(Servicio.objects.count(), 1)
        self.assertEqual(Falla.objects.count(), 1)
        self.assertEqual(Nivel.objects.count(), 1)
        self.assertEqual(Servicio.objects.first().nombre, "Elevadores")

    def test_backfill_command_idempotente(self):
        # Se importa ANTES de que existan el edificio y el enlace.
        self.enlace.delete()
        self.edificio.delete()
        ruta = _excel_tmp([_fila_v(1, "SS26-0704")])
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        Ticket.objects.create(
            ticket_id="1",
            numero_display="SS26-0704",
            raw_data=datos[0]["raw_data"],
        )
        self.assertIsNone(Ticket.objects.get(ticket_id="1").torre)

        Edificio.objects.create(nombre="TORRE 2", siglas="T2")
        EnlaceAutorizado.objects.create(
            nombres="Cesar", primer_apellido="Zavala",
            nombre_sig="Cesar Augusto Zavala", institucion=self.institucion,
        )
        call_command("vincular_tickets", verbosity=0)

        t = Ticket.objects.get(ticket_id="1")
        self.assertEqual(t.torre.nombre, "TORRE 2")
        self.assertEqual(t.solicitante.nombre_sig, "Cesar Augusto Zavala")
        self.assertEqual(t.nivel.nombre, "Nivel 8")

        # Segunda corrida: no debe fallar (idempotente).
        call_command("vincular_tickets", "--todos", verbosity=0)


class VinculacionVistasTests(TestCase):
    def setUp(self):
        from django.contrib.auth import get_user_model

        self.usuario = get_user_model().objects.create_superuser(
            "prueba_vinculos", "pv@test.local", "Prueba123!"
        )
        self.client.force_login(self.usuario)
        self.edificio = Edificio.objects.create(nombre="TORRE 2", siglas="T2")
        self.institucion = Institucion.objects.create(
            nombre="Institución Uno", siglas="IUNO"
        )
        self.enlace = EnlaceAutorizado.objects.create(
            nombres="Cesar", primer_apellido="Zavala",
            nombre_sig="Cesar Augusto Zavala", institucion=self.institucion,
        )
        ruta = _excel_tmp([_fila_v(1, "SS26-0710"), _fila_v(2, "SS26-0711")])
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        aplicar_tickets(datos, MODO_SYNC_PARCIAL)

    def test_lista_muestra_vinculos_con_enlaces(self):
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets")
        ).content.decode("utf-8", "replace")
        self.assertIn("Cesar Augusto Zavala", html)
        self.assertIn("Falla Uno", html)
        self.assertIn(
            reverse("enlaces_ccg:detalle_edificio", args=[self.edificio.pk]), html
        )

    def test_filtro_por_servicio(self):
        s = Servicio.objects.get(nombre="Elevadores")
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"), {"servicio": s.pk}
        ).content.decode("utf-8", "replace")
        self.assertIn("SS26-0710", html)

    def test_filtro_por_falla_sin_resultados(self):
        f = Falla.objects.create(descripcion="Otra falla")
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"), {"falla": f.pk}
        ).content.decode("utf-8", "replace")
        self.assertIn("No hay tickets", html)

    def test_filtro_por_nivel(self):
        n = Nivel.objects.get(nombre="Nivel 8")
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"), {"nivel": n.pk}
        ).content.decode("utf-8", "replace")
        self.assertIn("SS26-0710", html)

    def test_detalle_muestra_vinculos(self):
        t = Ticket.objects.get(ticket_id="1")
        html = self.client.get(
            reverse("enlaces_ccg:ticket_detalle", args=[t.pk])
        ).content.decode("utf-8", "replace")
        self.assertIn("Cesar Augusto Zavala", html)
        self.assertIn("Institución Uno", html)
        self.assertIn("Elevadores", html)
        self.assertIn("Falla Uno", html)
        self.assertIn("?nivel=", html)
        self.assertIn("?servicio=", html)
        self.assertIn("Detalle del SIG", html)

    def test_ficha_enlace_muestra_pestanas_y_tickets(self):
        html = self.client.get(
            reverse("enlaces_ccg:enlace_detalle", args=[self.enlace.pk])
        ).content.decode("utf-8", "replace")
        self.assertIn("Datos generales", html)
        self.assertIn("Contacto", html)
        self.assertIn("Ubicación", html)
        self.assertIn("Datos SIG", html)
        self.assertIn("Alta", html)
        self.assertIn("Seguimiento", html)
        self.assertIn("Baja", html)
        self.assertIn("Tickets", html)
        self.assertIn("SS26-0710", html)
        self.assertNotIn('id="tab-alta"', html)
        self.assertNotIn('id="tab-sig"', html)