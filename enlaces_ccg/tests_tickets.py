"""Pruebas del parser/importador del módulo de tickets (Fase 2)."""

import tempfile
import json
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
    PermisoGrupo,
    ResponsableAtencion,
    Servicio,
    Ticket,
    TicketCierre,
    TicketLog,
    VistaGuardada,
)
from .roles import CAP_DIRECTORIO, CAP_EDITAR
from .sig_sync.tickets import (
    MODO_SYNC_COMPLETO,
    MODO_SYNC_PARCIAL,
    _calcular_numero_display,
    _normalizar_nombre,
    _parsear_fecha,
    _registrar_resultado,
    aplicar_tickets,
    parsear_excel_tickets,
    sincronizar_tickets,
    tick_sync_tickets_task,
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


def _fila_con_cierre(tid, numero, cierro, diagnostico="Dx", actividades="Act",
                     observaciones="Obs", **kwargs):
    """Fila del Excel con los campos del informe de cierre rellenados."""
    fila = _fila(tid, numero, **kwargs)
    fila[21] = diagnostico
    fila[23] = actividades
    fila[25] = observaciones
    fila[17] = cierro
    return fila


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

    def test_cierre_local_no_se_pisoteca_si_sig_sigue_abierto(self):
        datos = self._filas_basicas([_fila(1, "SS26-0401", cierro="")])
        aplicar_tickets(datos, MODO_SYNC_PARCIAL)
        ticket = Ticket.objects.get(ticket_id="1")
        fecha_cierre = timezone.make_aware(datetime(2026, 9, 2, 15, 30, 0))
        TicketCierre.objects.create(
            ticket=ticket,
            fecha_inicio=timezone.make_aware(datetime(2026, 9, 1, 8, 0, 0)),
            fecha_cierre=fecha_cierre,
            diagnostico="Dx local",
            actividades="Act local",
            observaciones="Obs local",
        )
        ticket.estatus = Ticket.ESTATUS_CERRADO
        ticket.fecha_cierre = fecha_cierre
        ticket.save(update_fields=["estatus", "fecha_cierre"])

        # El SIG vuelve a reportar el ticket SIN fecha de cierre (sigue abierto).
        aplicar_tickets(self._filas_basicas([_fila(1, "SS26-0401", cierro="")]),
                        MODO_SYNC_PARCIAL)
        ticket.refresh_from_db()
        self.assertEqual(ticket.estatus, Ticket.ESTATUS_CERRADO)
        self.assertEqual(ticket.fecha_cierre, fecha_cierre)
        self.assertFalse(ticket.cierre_sincronizado)

    def test_cierre_local_se_libera_cuando_sig_confirma(self):
        datos = self._filas_basicas([_fila(1, "SS26-0402", cierro="")])
        aplicar_tickets(datos, MODO_SYNC_PARCIAL)
        ticket = Ticket.objects.get(ticket_id="1")
        fecha_cierre = timezone.make_aware(datetime(2026, 9, 2, 15, 30, 0))
        TicketCierre.objects.create(
            ticket=ticket,
            fecha_inicio=timezone.make_aware(datetime(2026, 9, 1, 8, 0, 0)),
            fecha_cierre=fecha_cierre,
            diagnostico="Dx local",
            actividades="Act local",
            observaciones="Obs local",
        )
        ticket.estatus = Ticket.ESTATUS_CERRADO
        ticket.fecha_cierre = fecha_cierre
        ticket.save(update_fields=["estatus", "fecha_cierre"])

        # El SIG ya reporta el mismo cierre (fecha + campos).
        aplicar_tickets(self._filas_basicas([
            _fila_con_cierre(1, "SS26-0402", cierro="2026-09-02 12:00:00",
                             diagnostico="Dx local", actividades="Act local",
                             observaciones="Obs local"),
        ]), MODO_SYNC_PARCIAL)
        ticket.refresh_from_db()
        self.assertEqual(ticket.estatus, Ticket.ESTATUS_CERRADO)
        self.assertTrue(ticket.cierre_sincronizado)


class CierreSincronizadoTests(TestCase):
    def setUp(self):
        self.ticket = Ticket.objects.create(ticket_id="SIG-1", numero="SS26-0900")

    def _con_cierre(self, **campos):
        TicketCierre.objects.get_or_create(
            ticket=self.ticket,
            defaults={
                "fecha_inicio": timezone.make_aware(datetime(2026, 9, 1, 8, 0, 0)),
                "fecha_cierre": timezone.make_aware(datetime(2026, 9, 2, 15, 30, 0)),
                **campos,
            },
        )
        self.ticket.estatus = Ticket.ESTATUS_CERRADO
        self.ticket.save(update_fields=["estatus"])

    def test_sin_cierre_local_es_sincronizado(self):
        self.assertTrue(self.ticket.cierre_sincronizado)

    def test_esta_cerrado_solo_por_estatus_sig(self):
        self.ticket.estatus = Ticket.ESTATUS_CERRADO
        self.assertTrue(self.ticket.esta_cerrado)
        self.ticket.estatus = Ticket.ESTATUS_ABIERTO
        self.ticket.save(update_fields=["estatus"])
        self.assertFalse(self.ticket.esta_cerrado)

    def test_esta_cerrado_con_cierre_local_pendiente_de_sig(self):
        TicketCierre.objects.create(
            ticket=self.ticket,
            fecha_inicio=timezone.make_aware(datetime(2026, 9, 1, 8, 0, 0)),
            fecha_cierre=timezone.make_aware(datetime(2026, 9, 2, 15, 30, 0)),
            diagnostico="Dx", actividades="Act", observaciones="Obs",
        )
        self.ticket.estatus = Ticket.ESTATUS_ABIERTO
        self.ticket.save(update_fields=["estatus"])
        self.assertTrue(self.ticket.esta_cerrado)
        self.assertFalse(self.ticket.cierre_sincronizado)

    def test_empate_completo_por_dia(self):
        self._con_cierre(
            diagnostico="Dx", actividades="Act", observaciones="Obs",
        )
        self.ticket.raw_data = {
            "cerro_fecha": "2026-09-02 12:00:00",
            "Diagnostico": "Dx",
            "Actividades": "Act",
            "Observaciones": "Obs",
        }
        self.ticket.save(update_fields=["raw_data"])
        self.assertTrue(self.ticket.cierre_sincronizado)

    def test_sig_sin_cerro_fecha_queda_pendiente(self):
        self._con_cierre(
            diagnostico="Dx", actividades="Act", observaciones="Obs",
        )
        self.ticket.raw_data = {
            "Diagnostico": "Dx",
            "Actividades": "Act",
            "Observaciones": "Obs",
        }
        self.ticket.save(update_fields=["raw_data"])
        self.assertFalse(self.ticket.cierre_sincronizado)

    def test_campos_diferentes_quedan_pendiente(self):
        self._con_cierre(
            diagnostico="Dx local", actividades="Act", observaciones="Obs",
        )
        self.ticket.raw_data = {
            "cerro_fecha": "2026-09-02 12:00:00",
            "Diagnostico": "Dx SIG",
            "Actividades": "Act",
            "Observaciones": "Obs",
        }
        self.ticket.save(update_fields=["raw_data"])
        self.assertFalse(self.ticket.cierre_sincronizado)

    def test_fecha_diferente_queda_pendiente(self):
        self._con_cierre(
            diagnostico="Dx", actividades="Act", observaciones="Obs",
        )
        self.ticket.raw_data = {
            "cerro_fecha": "2026-09-03 12:00:00",
            "Diagnostico": "Dx",
            "Actividades": "Act",
            "Observaciones": "Obs",
        }
        self.ticket.save(update_fields=["raw_data"])
        self.assertFalse(self.ticket.cierre_sincronizado)


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
        self.assertIn('id="syncDot"', html)
        self.assertIn('id="badgeUltima"', html)
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
        self.assertIn("ultimo_estado", datos)
        self.assertIn("ultimo_mensaje", datos)
        self.assertIs(datos["sincronizando"], False)

    def test_sincronizacion_estado_activo(self):
        with mock.patch(
            "enlaces_ccg.sig_sync.tickets._lock_activo", return_value=True
        ):
            respuesta = self.client.get(
                reverse("enlaces_ccg:ticket_sincronizacion_estado")
            )
        self.assertIs(respuesta.json()["sincronizando"], True)


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

    def test_registrar_seguimiento_guarda(self):
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


class RegistroResultadoTests(TestCase):
    """`_registrar_resultado`: el badge "Última" no se reinicia si falla."""

    def setUp(self):
        self.cfg = ConfiguracionTickets.cargar()

    def test_ok_actualiza_ultima_e_intento(self):
        self.cfg.ultima_sincronizacion = None
        self.cfg.save(update_fields=["ultima_sincronizacion"])
        _registrar_resultado("parcial", TicketLog.ESTADO_OK, "todo bien")
        self.cfg.refresh_from_db()
        self.assertIsNotNone(self.cfg.ultima_sincronizacion)
        self.assertIsNotNone(self.cfg.ultimo_intento)
        self.assertEqual(self.cfg.ultimo_estado, "OK")
        self.assertEqual(TicketLog.objects.count(), 1)

    def test_error_no_reinicia_ultima_pero_si_intento(self):
        anterior = timezone.now() - timedelta(minutes=10)
        self.cfg.ultima_sincronizacion = anterior
        self.cfg.save(update_fields=["ultima_sincronizacion"])
        with mock.patch(
            "enlaces_ccg.sig_sync.tickets._registrar_error_log"
        ) as log_mock:
            _registrar_resultado("parcial", TicketLog.ESTADO_ERROR, "boom")
        self.cfg.refresh_from_db()
        # "Última" (exitosa) no cambia; el intento sí se registra.
        self.assertEqual(self.cfg.ultima_sincronizacion, anterior)
        self.assertIsNotNone(self.cfg.ultimo_intento)
        self.assertEqual(self.cfg.ultimo_estado, "ERROR")
        self.assertEqual(TicketLog.objects.count(), 1)
        log_mock.assert_called_once()


class EnHorarioTests(TestCase):
    """`ConfiguracionTickets.en_horario` / `proximo_inicio_horario`."""

    def setUp(self):
        self.cfg = ConfiguracionTickets.cargar()
        self.cfg.hora_inicio = time(7, 0)
        self.cfg.hora_fin = time(19, 0)
        self.cfg.dias_semana = []

    def _local(self, *args):
        return timezone.make_aware(datetime(*args), timezone.get_current_timezone())

    def test_dentro(self):
        # 2026-09-18 es viernes; 12:00 local
        self.assertTrue(self.cfg.en_horario(self._local(2026, 9, 18, 12, 0)))

    def test_fuera_antes(self):
        self.assertFalse(self.cfg.en_horario(self._local(2026, 9, 18, 6, 59)))

    def test_fuera_despues(self):
        self.assertFalse(self.cfg.en_horario(self._local(2026, 9, 18, 19, 0)))

    def test_sin_limites(self):
        self.cfg.hora_inicio = None
        self.cfg.hora_fin = None
        self.assertTrue(self.cfg.en_horario(self._local(2026, 9, 18, 3, 0)))

    def test_dia_no_permitido(self):
        self.cfg.dias_semana = [0, 1, 2, 3, 4]  # lunes a viernes
        # 2026-09-19 es sábado
        self.assertFalse(self.cfg.en_horario(self._local(2026, 9, 19, 12, 0)))
        # 2026-09-18 es viernes
        self.assertTrue(self.cfg.en_horario(self._local(2026, 9, 18, 12, 0)))

    def test_proximo_inicio_hoy(self):
        ahora = self._local(2026, 9, 18, 6, 0)
        prox = self.cfg.proximo_inicio_horario(ahora)
        self.assertEqual(prox, self._local(2026, 9, 18, 7, 0))

    def test_proximo_inicio_manana(self):
        ahora = self._local(2026, 9, 18, 20, 0)
        prox = self.cfg.proximo_inicio_horario(ahora)
        self.assertEqual(prox, self._local(2026, 9, 19, 7, 0))


class TickHorarioTests(TestCase):
    """El tick solo sincroniza el parcial dentro del horario."""

    def setUp(self):
        self.cfg = ConfiguracionTickets.cargar()
        self.cfg.habilitado = True
        self.cfg.intervalo_minutos = 5
        self.cfg.descarga_completa_horas = []
        self.cfg.hora_inicio = time(7, 0)
        self.cfg.hora_fin = time(19, 0)
        self.cfg.dias_semana = []
        self.cfg.save()

    def _local(self, *args):
        return timezone.make_aware(datetime(*args), timezone.get_current_timezone())

    def test_parcial_fuera_de_horario(self):
        fijo = self._local(2026, 9, 18, 21, 0)  # 21:00 local, fuera de horario
        with mock.patch(
            "enlaces_ccg.sig_sync.tickets.timezone.now", return_value=fijo
        ):
            res = tick_sync_tickets_task()
        self.assertFalse(res["sincronizado"])
        self.assertEqual(res["motivo"], "fuera_horario")

    def test_completa_fuera_de_horario_si_corre(self):
        self.cfg.descarga_completa_horas = ["21:00"]
        self.cfg.save(update_fields=["descarga_completa_horas"])
        fijo = self._local(2026, 9, 18, 21, 0)  # coincide con la hora completa
        with mock.patch(
            "enlaces_ccg.sig_sync.tickets.timezone.now", return_value=fijo
        ), mock.patch(
            "enlaces_ccg.sig_sync.tickets._adquirir_lock", return_value=None
        ), mock.patch(
            "enlaces_ccg.sig_sync.tickets.sincronizar_tickets",
            return_value={"creados": 0, "actualizados": 0, "archivados": 0},
        ) as sync_mock:
            res = tick_sync_tickets_task()
        self.assertTrue(res["sincronizado"])
        sync_mock.assert_called_once_with(modo=MODO_SYNC_COMPLETO)


class SincronizacionFallasTests(TestCase):
    """La sincronización de tickets recalibra las fallas solo tras completas."""

    def _correr_sincronizacion(self, modo):
        with mock.patch(
            "enlaces_ccg.sig_sync.tickets.descargar_excel_tickets",
            return_value={"path": "x.xlsx"},
        ), mock.patch(
            "enlaces_ccg.sig_sync.tickets.parsear_excel_tickets",
            return_value=([], 0),
        ), mock.patch(
            "enlaces_ccg.sig_sync.tickets.aplicar_tickets",
            return_value={"creados": 0, "actualizados": 0, "archivados": 0},
        ), mock.patch(
            "enlaces_ccg.sig_sync.tickets._limpiar_descargas"
        ), mock.patch(
            "enlaces_ccg.sig_sync.tickets._registrar_resultado"
        ) as registrar_mock, mock.patch(
            "enlaces_ccg.sig_sync.tickets.sincronizar_datos_fallas",
            return_value=3,
        ) as fallas_mock:
            res = sincronizar_tickets(modo=modo)
            return res, registrar_mock, fallas_mock

    def test_completo_recalibra_fallas(self):
        res, registrar_mock, fallas_mock = self._correr_sincronizacion(
            MODO_SYNC_COMPLETO
        )
        fallas_mock.assert_called_once_with()
        self.assertEqual(res["fallas_actualizadas"], 3)
        mensaje = registrar_mock.call_args[0][2]
        self.assertIn("3 fallas actualizadas", mensaje)

    def test_parcial_no_recalibra_fallas(self):
        res, registrar_mock, fallas_mock = self._correr_sincronizacion(
            MODO_SYNC_PARCIAL
        )
        fallas_mock.assert_not_called()
        self.assertEqual(res["fallas_actualizadas"], 0)
        mensaje = registrar_mock.call_args[0][2]
        self.assertNotIn("fallas actualizadas", mensaje)


class ResponsablesUsuariosTests(TestCase):
    def setUp(self):
        from django.contrib.auth import get_user_model

        self.resp_uno = ResponsableAtencion.objects.create(
            nombre="Resp Uno", activo=True
        )
        self.resp_dos = ResponsableAtencion.objects.create(
            nombre="Resp Dos", activo=True
        )
        self.resp_tres = ResponsableAtencion.objects.create(
            nombre="Resp Tres", activo=True
        )
        self.usuario = get_user_model().objects.create_superuser(
            "operador.resp", "op@test.local", "Prueba123!"
        )
        self.usuario.responsables_atencion.add(self.resp_uno, self.resp_dos)
        self.client.force_login(self.usuario)

        def fila_resp(tid, numero, responsable):
            fila = _fila_v(tid, numero)
            fila[3] = responsable
            return fila

        self.datos = self._filas_basicas([
            fila_resp(1, "SS26-0901", "Resp Uno"),
            fila_resp(2, "SS26-0902", "Resp Dos"),
            fila_resp(3, "SS26-0903", "Resp Tres"),
        ])
        aplicar_tickets(self.datos, MODO_SYNC_PARCIAL)

    def _filas_basicas(self, filas):
        ruta = _excel_tmp(filas)
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        return datos

    def test_lista_prefiltrada_por_responsables_del_usuario(self):
        respuesta = self.client.get(reverse("enlaces_ccg:seguimiento_tickets"))
        html = respuesta.content.decode("utf-8", "replace")
        self.assertEqual(respuesta.status_code, 200)
        self.assertIn(">SS26-0901<", html)
        self.assertIn(">SS26-0902<", html)
        self.assertNotIn(">SS26-0903<", html)
        self.assertIn("Ver todos los tickets", html)

    def test_todos_por_parametro_muestra_todo(self):
        respuesta = self.client.get(reverse("enlaces_ccg:seguimiento_tickets"), {"todos": "1"})
        html = respuesta.content.decode("utf-8", "replace")
        self.assertIn(">SS26-0903<", html)
        self.assertNotIn("Ver todos los tickets", html)

    def test_responsable_explicito_ignora_prefiltro(self):
        respuesta = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"),
            {"responsable": str(self.resp_tres.pk)},
        )
        html = respuesta.content.decode("utf-8", "replace")
        self.assertIn(">SS26-0903<", html)
        self.assertNotIn("Ver todos los tickets", html)

    def test_catalogo_usuarios_asigna_responsables_al_crear(self):
        from django.contrib.auth import get_user_model

        respuesta = self.client.post(
            reverse("enlaces_ccg:catalogo_usuarios"),
            {
                "accion": "crear_usuario",
                "username": "usuario.limpieza",
                "password": "Clave123!",
                "first_name": "Limpieza",
                "last_name": "Operativo",
                "email": "",
                "is_staff": "1",
                "is_active": "1",
                "responsables": [str(self.resp_uno.pk), str(self.resp_dos.pk)],
            },
        )
        self.assertEqual(respuesta.status_code, 302)
        u = get_user_model().objects.get(username="usuario.limpieza")
        self.assertEqual(
            set(u.responsables_atencion.values_list("pk", flat=True)),
            {self.resp_uno.pk, self.resp_dos.pk},
        )

    def test_catalogo_usuarios_reasigna_responsables_al_editar(self):
        from django.contrib.auth import get_user_model

        u = get_user_model().objects.create_user(
            "usuario.editar", "e@test.local", "Clave123!", is_active=True
        )
        u.responsables_atencion.add(self.resp_uno)

        respuesta = self.client.post(
            reverse("enlaces_ccg:catalogo_usuarios"),
            {
                "accion": "editar_usuario",
                "pk": str(u.pk),
                "username": "usuario.editar",
                "password": "",
                "first_name": "",
                "last_name": "",
                "email": "",
                "is_staff": "0",
                "is_active": "1",
                "responsables": [str(self.resp_tres.pk)],
            },
        )
        self.assertEqual(respuesta.status_code, 302)
        u.refresh_from_db()
        self.assertEqual(
            list(u.responsables_atencion.values_list("pk", flat=True)),
            [self.resp_tres.pk],
        )

    def test_desactivar_usuario_desvincula_responsables(self):
        from django.contrib.auth import get_user_model

        u = get_user_model().objects.create_user(
            "usuario.baja", "b@test.local", "Clave123!", is_active=True
        )
        u.responsables_atencion.add(self.resp_uno, self.resp_dos)

        respuesta = self.client.post(
            reverse("enlaces_ccg:catalogo_usuarios"),
            {"accion": "toggle_usuario", "pk": str(u.pk)},
        )
        self.assertEqual(respuesta.status_code, 302)
        u.refresh_from_db()
        self.assertFalse(u.is_active)
        self.assertEqual(u.responsables_atencion.count(), 0)

    def test_editar_usuario_inactivo_desvincula_responsables(self):
        from django.contrib.auth import get_user_model

        u = get_user_model().objects.create_user(
            "usuario.inactivo", "i@test.local", "Clave123!", is_active=True
        )
        u.responsables_atencion.add(self.resp_uno)

        respuesta = self.client.post(
            reverse("enlaces_ccg:catalogo_usuarios"),
            {
                "accion": "editar_usuario",
                "pk": str(u.pk),
                "username": "usuario.inactivo",
                "password": "",
                "first_name": "",
                "last_name": "",
                "email": "",
                "is_staff": "0",
                "is_active": "0",
                "responsables": [str(self.resp_uno.pk), str(self.resp_dos.pk)],
            },
        )
        self.assertEqual(respuesta.status_code, 302)
        u.refresh_from_db()
        self.assertFalse(u.is_active)
        self.assertEqual(u.responsables_atencion.count(), 0)

    def test_guardar_usuario_inactivo_desvincula_por_signal(self):
        from django.contrib.auth import get_user_model

        u = get_user_model().objects.create_user(
            "usuario.signal", "s@test.local", "Clave123!", is_active=True
        )
        u.responsables_atencion.add(self.resp_uno)
        u.is_active = False
        u.save(update_fields=["is_active"])
        u.refresh_from_db()
        self.assertEqual(u.responsables_atencion.count(), 0)


class VistaGuardadaTests(TestCase):
    def setUp(self):
        from django.contrib.auth import get_user_model

        self.resp_uno = ResponsableAtencion.objects.create(
            nombre="Resp Uno", activo=True
        )
        self.resp_dos = ResponsableAtencion.objects.create(
            nombre="Resp Dos", activo=True
        )
        self.resp_tres = ResponsableAtencion.objects.create(
            nombre="Resp Tres", activo=True
        )
        self.usuario = get_user_model().objects.create_superuser(
            "operador.vistas", "v@test.local", "Prueba123!"
        )
        self.usuario.responsables_atencion.add(self.resp_uno, self.resp_dos)
        self.client.force_login(self.usuario)

        hoy = timezone.localdate()
        fecha_mes = hoy.replace(day=max(1, hoy.day - 1))
        fecha_pasado = hoy.replace(day=1) - timedelta(days=1)

        def fila(tid, numero, responsable, fecha, cierro=""):
            f = _fila_v(tid, numero)
            f[3] = responsable
            f[13] = fecha.strftime("%Y-%m-%d 10:00:00")
            if cierro:
                f[17] = fecha.strftime("%Y-%m-%d 10:00:00")
            return f

        self.datos = self._filas_basicas([
            fila(1, "SS26-1001", "Resp Uno", fecha_mes),                       # A: mío, mes, abierto
            fila(2, "SS26-1002", "Resp Uno", fecha_pasado),                   # B: mío, mes pasado
            fila(3, "SS26-1003", "Resp Dos", fecha_mes),                      # C: mío, mes, abierto
            fila(4, "SS26-1004", "Resp Tres", fecha_mes),                     # D: ajeno
            fila(5, "SS26-1005", "Resp Uno", fecha_mes, cierro=True),         # E: mío, cerrado
        ])
        aplicar_tickets(self.datos, MODO_SYNC_PARCIAL)

    def _filas_basicas(self, filas):
        ruta = _excel_tmp(filas)
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        return datos

    def _guardar_vista(self, nombre, parametros):
        return self.client.post(
            reverse("enlaces_ccg:vista_guardar"),
            {"nombre": nombre, "parametros_json": json.dumps(parametros)},
        )

    def test_guardar_y_aplicar_vista_mis_abiertos_del_mes(self):
        respuesta = self._guardar_vista(
            "Mis abiertos del mes",
            {
                "estatus": ["ABIERTO"],
                "responsable": ["@mios"],
                "desde": ["@mes_inicio"],
                "hasta": ["@mes_fin"],
            },
        )
        self.assertEqual(respuesta.status_code, 302)
        self.assertIn("vista=", respuesta.url)
        vista = VistaGuardada.objects.get(
            usuario=self.usuario, nombre="Mis abiertos del mes"
        )
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"), {"vista": vista.pk}
        ).content.decode("utf-8", "replace")
        self.assertIn(">SS26-1001<", html)
        self.assertIn(">SS26-1003<", html)
        self.assertNotIn(">SS26-1002<", html)
        self.assertNotIn(">SS26-1004<", html)
        self.assertNotIn(">SS26-1005<", html)
        self.assertIn("Mis abiertos del mes", html)

    def test_predeterminada_se_aplica_al_entrar_sin_filtros(self):
        VistaGuardada.objects.create(
            usuario=self.usuario,
            modulo="tickets",
            nombre="Solo mios del mes",
            parametros={
                "estatus": ["ABIERTO"],
                "responsable": ["@mios"],
                "desde": ["@mes_inicio"],
                "hasta": ["@mes_fin"],
            },
            es_predeterminada=True,
        )
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets")
        ).content.decode("utf-8", "replace")
        self.assertIn("Solo mios del mes", html)
        self.assertIn(">SS26-1001<", html)
        self.assertNotIn(">SS26-1004<", html)

        html2 = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"), {"todos": "1"}
        ).content.decode("utf-8", "replace")
        self.assertIn(">SS26-1004<", html2)

    def test_vista_cero_no_aplica_predeterminada(self):
        VistaGuardada.objects.create(
            usuario=self.usuario,
            modulo="tickets",
            nombre="Solo mios del mes",
            parametros={"estatus": ["ABIERTO"]},
            es_predeterminada=True,
        )
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"), {"vista": "0"}
        ).content.decode("utf-8", "replace")
        # Sin vista: vuelve el prefiltro por mis responsables (incluye mes pasado).
        self.assertIn(">SS26-1002<", html)
        self.assertNotIn(">SS26-1004<", html)

    def test_predeterminar_marca_y_desmarca(self):
        v1 = VistaGuardada.objects.create(
            usuario=self.usuario, modulo="tickets", nombre="V1", parametros={}
        )
        v2 = VistaGuardada.objects.create(
            usuario=self.usuario, modulo="tickets", nombre="V2", parametros={}
        )
        self.client.post(reverse("enlaces_ccg:vista_predeterminar", args=[v1.pk]))
        v1.refresh_from_db()
        v2.refresh_from_db()
        self.assertTrue(v1.es_predeterminada)
        self.assertFalse(v2.es_predeterminada)
        self.client.post(reverse("enlaces_ccg:vista_predeterminar", args=[v2.pk]))
        v1.refresh_from_db()
        v2.refresh_from_db()
        self.assertFalse(v1.es_predeterminada)
        self.assertTrue(v2.es_predeterminada)

    def test_mios_siguen_a_responsables_actuales(self):
        vista = VistaGuardada.objects.create(
            usuario=self.usuario,
            modulo="tickets",
            nombre="Mis tickets",
            parametros={"responsable": ["@mios"]},
        )
        url = reverse("enlaces_ccg:seguimiento_tickets") + f"?vista={vista.pk}"
        html = self.client.get(url).content.decode("utf-8", "replace")
        self.assertNotIn(">SS26-1004<", html)

        self.usuario.responsables_atencion.add(self.resp_tres)
        html2 = self.client.get(url).content.decode("utf-8", "replace")
        self.assertIn(">SS26-1004<", html2)

    def test_eliminar_vista_propia(self):
        vista = VistaGuardada.objects.create(
            usuario=self.usuario, modulo="tickets", nombre="V", parametros={}
        )
        respuesta = self.client.post(
            reverse("enlaces_ccg:vista_eliminar", args=[vista.pk])
        )
        self.assertEqual(respuesta.status_code, 302)
        self.assertFalse(VistaGuardada.objects.filter(pk=vista.pk).exists())

    def test_no_eliminar_vista_ajena(self):
        from django.contrib.auth import get_user_model

        otro = get_user_model().objects.create_user(
            "otro.vistas", "o@test.local", "Clave123!"
        )
        vista = VistaGuardada.objects.create(
            usuario=otro, modulo="tickets", nombre="Ajeno", parametros={}
        )
        respuesta = self.client.post(
            reverse("enlaces_ccg:vista_eliminar", args=[vista.pk])
        )
        self.assertEqual(respuesta.status_code, 404)
        self.assertTrue(VistaGuardada.objects.filter(pk=vista.pk).exists())

    def test_nombre_repetido_reemplaza(self):
        self._guardar_vista("Mi vista", {"estatus": ["ABIERTO"]})
        self._guardar_vista("Mi vista", {"estatus": ["CERRADO"]})
        self.assertEqual(
            VistaGuardada.objects.filter(
                usuario=self.usuario, nombre="Mi vista"
            ).count(),
            1,
        )
        vista = VistaGuardada.objects.get(
            usuario=self.usuario, nombre="Mi vista"
        )
        self.assertEqual(vista.parametros, {"estatus": ["CERRADO"]})

    def test_guardar_sin_nombre_rechaza(self):
        respuesta = self.client.post(
            reverse("enlaces_ccg:vista_guardar"),
            {"nombre": "", "parametros_json": "{}"},
        )
        self.assertEqual(respuesta.status_code, 400)

    def test_semilla_general_existe(self):
        self.assertTrue(
            VistaGuardada.objects.filter(
                modulo="tickets",
                es_general=True,
                nombre="Mis tickets del mes (abiertos)",
            ).exists()
        )

    def test_vista_general_disponible_para_todos(self):
        from django.contrib.auth import get_user_model

        general = VistaGuardada.objects.create(
            usuario=None,
            modulo="tickets",
            nombre="General abiertos del mes",
            parametros={
                "estatus": ["ABIERTO"],
                "responsable": ["@mios"],
                "desde": ["@mes_inicio"],
                "hasta": ["@mes_fin"],
            },
            es_general=True,
        )
        otro = get_user_model().objects.create_superuser(
            "operador.general", "g@test.local", "Prueba123!"
        )
        self.client.force_login(otro)

        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets")
        ).content.decode("utf-8", "replace")
        self.assertIn("Disponibles para todos", html)
        self.assertIn("General abiertos del mes", html)

        # El viewer no tiene responsables: @mios se omite, filtra mes + abiertos.
        html2 = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"),
            {"vista": general.pk},
        ).content.decode("utf-8", "replace")
        self.assertIn(">SS26-1001<", html2)  # A, abierto este mes
        self.assertIn(">SS26-1004<", html2)  # D, abierto este mes, ajeno
        self.assertNotIn(">SS26-1002<", html2)  # B, abierto pero mes pasado
        self.assertNotIn(">SS26-1005<", html2)  # E, cerrado

        # No puede eliminar ni marcar como predeterminada una vista general.
        r1 = self.client.post(
            reverse("enlaces_ccg:vista_predeterminar", args=[general.pk])
        )
        self.assertEqual(r1.status_code, 404)
        r2 = self.client.post(
            reverse("enlaces_ccg:vista_eliminar", args=[general.pk])
        )
        self.assertEqual(r2.status_code, 404)


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

    def test_filtro_responsable_multiple(self):
        f1 = _fila("A1", "SS26-A1")
        f1[3] = "Resp Uno"
        f2 = _fila("A2", "SS26-A2")
        f2[3] = "Resp Dos"
        f3 = _fila("A3", "SS26-A3")
        f3[3] = "Resp Tres"
        ruta = _excel_tmp([f1, f2, f3])
        datos, _ = parsear_excel_tickets(ruta)
        _calcular_numero_display(datos)
        aplicar_tickets(datos, MODO_SYNC_PARCIAL)
        r1 = ResponsableAtencion.objects.get(nombre="Resp Uno")
        r2 = ResponsableAtencion.objects.get(nombre="Resp Dos")
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets"),
            {"responsable": [r1.pk, r2.pk]},
        ).content.decode("utf-8", "replace")
        self.assertIn("SS26-A1", html)
        self.assertIn("SS26-A2", html)
        self.assertNotIn("SS26-A3", html)

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
        self.assertNotIn("Detalle del SIG", html)
        self.assertNotIn("Comentarios del servicio", html)
        self.assertNotIn("Registrar seguimiento", html)

    def test_cierre_anterior_a_creacion_se_rechaza(self):
        t = Ticket.objects.get(ticket_id="1")
        creacion = timezone.localtime(t.fecha)
        anterior = creacion - timedelta(hours=1)
        respuesta = self.client.post(
            reverse("enlaces_ccg:ticket_cierre_guardar", args=[t.pk]),
            {
                "incluir_fechas": "1",
                "fecha_inicio_fecha": creacion.strftime("%Y-%m-%d"),
                "fecha_inicio_hora": creacion.strftime("%H:%M"),
                "fecha_cierre_fecha": anterior.strftime("%Y-%m-%d"),
                "fecha_cierre_hora": anterior.strftime("%H:%M"),
            },
        )
        self.assertEqual(respuesta.status_code, 302)
        self.assertFalse(TicketCierre.objects.filter(ticket=t).exists())

    def test_cierre_sin_incluir_fechas_no_guarda_fechas(self):
        t = Ticket.objects.get(ticket_id="1")
        self.client.post(
            reverse("enlaces_ccg:ticket_cierre_guardar", args=[t.pk]),
            {"diagnostico": "Dx", "actividades": "Act"},
        )
        cierre = TicketCierre.objects.get(ticket=t)
        self.assertIsNone(cierre.fecha_inicio)
        self.assertIsNone(cierre.fecha_cierre)

    def test_cierre_sin_datos_se_rechaza(self):
        t = Ticket.objects.get(ticket_id="1")
        respuesta = self.client.post(
            reverse("enlaces_ccg:ticket_cierre_guardar", args=[t.pk]),
            {},
        )
        self.assertEqual(respuesta.status_code, 302)
        self.assertFalse(TicketCierre.objects.filter(ticket=t).exists())

    def test_seguimiento_no_cierra_ticket(self):
        t = Ticket.objects.get(ticket_id="1")
        self.assertEqual(t.estatus, Ticket.ESTATUS_ABIERTO)
        self.client.post(
            reverse("enlaces_ccg:ticket_cierre_guardar", args=[t.pk]),
            {"diagnostico": "Dx"},
        )
        t.refresh_from_db()
        self.assertEqual(t.estatus, Ticket.ESTATUS_ABIERTO)
        self.assertFalse(t.esta_cerrado)
        self.assertTrue(TicketCierre.objects.filter(ticket=t).exists())

    def test_cierre_con_fechas_cierra_ticket(self):
        t = Ticket.objects.get(ticket_id="1")
        creacion = timezone.localtime(t.fecha)
        self.client.post(
            reverse("enlaces_ccg:ticket_cierre_guardar", args=[t.pk]),
            {
                "incluir_fechas": "1",
                "fecha_inicio_fecha": creacion.strftime("%Y-%m-%d"),
                "fecha_inicio_hora": creacion.strftime("%H:%M"),
                "fecha_cierre_fecha": creacion.strftime("%Y-%m-%d"),
                "fecha_cierre_hora": (creacion + timedelta(hours=1)).strftime(
                    "%H:%M"
                ),
            },
        )
        t.refresh_from_db()
        self.assertEqual(t.estatus, Ticket.ESTATUS_CERRADO)
        self.assertTrue(t.esta_cerrado)

    def test_detalle_seguimiento_muestra_pendiente_sig(self):
        t = Ticket.objects.get(ticket_id="1")
        self.client.post(
            reverse("enlaces_ccg:ticket_cierre_guardar", args=[t.pk]),
            {"diagnostico": "Dx"},
        )
        html = self.client.get(
            reverse("enlaces_ccg:ticket_detalle", args=[t.pk])
        ).content.decode("utf-8", "replace")
        self.assertIn("Pendiente sincronizar SIG", html)
        self.assertIn(
            "Trabajo en curso · Seguimiento pendiente de sincronizar en el SIG",
            html,
        )
        self.assertIn("Abierto", html)

    def test_check_fechas_desmarcado_si_cerrado(self):
        t = Ticket.objects.get(ticket_id="1")
        creacion = timezone.localtime(t.fecha)
        self.client.post(
            reverse("enlaces_ccg:ticket_cierre_guardar", args=[t.pk]),
            {
                "incluir_fechas": "1",
                "fecha_inicio_fecha": creacion.strftime("%Y-%m-%d"),
                "fecha_inicio_hora": creacion.strftime("%H:%M"),
                "fecha_cierre_fecha": creacion.strftime("%Y-%m-%d"),
                "fecha_cierre_hora": (creacion + timedelta(hours=1)).strftime(
                    "%H:%M"
                ),
            },
        )
        t.refresh_from_db()
        self.assertTrue(t.esta_cerrado)
        html = self.client.get(
            reverse("enlaces_ccg:ticket_cerrar", args=[t.pk])
        ).content.decode("utf-8", "replace")
        self.assertNotIn('name="incluir_fechas" value="1" checked', html)

    def test_seguimiento_en_cerrado_mantiene_fechas_y_estado(self):
        t = Ticket.objects.get(ticket_id="1")
        creacion = timezone.localtime(t.fecha)
        self.client.post(
            reverse("enlaces_ccg:ticket_cierre_guardar", args=[t.pk]),
            {
                "incluir_fechas": "1",
                "fecha_inicio_fecha": creacion.strftime("%Y-%m-%d"),
                "fecha_inicio_hora": creacion.strftime("%H:%M"),
                "fecha_cierre_fecha": creacion.strftime("%Y-%m-%d"),
                "fecha_cierre_hora": (creacion + timedelta(hours=1)).strftime(
                    "%H:%M"
                ),
            },
        )
        self.client.post(
            reverse("enlaces_ccg:ticket_cierre_guardar", args=[t.pk]),
            {"diagnostico": "Nuevo seguimiento"},
        )
        t.refresh_from_db()
        cierre = TicketCierre.objects.get(ticket=t)
        self.assertTrue(t.esta_cerrado)
        self.assertIsNotNone(cierre.fecha_cierre)
        self.assertEqual(cierre.diagnostico, "Nuevo seguimiento")

    def test_cierre_precarga_informe_desde_sig(self):
        t = Ticket.objects.create(
            ticket_id="B1",
            numero="SS26-B1",
            raw_data={
                "Diagnostico": "Dx SIG",
                "Actividades": "Act SIG",
                "Observaciones": "Obs SIG",
                "ObservacionesUsuario": "ObsU SIG",
                "falla_descripcion": "",
                "servicio": "",
                "nivel": "",
                "grupo": "",
                "solicitud_solicitante": "",
                "solicitud_descripcion": "",
            },
        )
        html = self.client.get(
            reverse("enlaces_ccg:ticket_cerrar", args=[t.pk])
        ).content.decode("utf-8", "replace")
        self.assertIn("Dx SIG", html)
        self.assertIn("Act SIG", html)
        self.assertIn("Obs SIG", html)
        self.assertIn("ObsU SIG", html)

    def test_pagina_cierre_renderiza_hora(self):
        t = Ticket.objects.get(ticket_id="1")
        html = self.client.get(
            reverse("enlaces_ccg:ticket_cerrar", args=[t.pk])
        ).content.decode("utf-8", "replace")
        self.assertIn('name="fecha_inicio_hora"', html)
        self.assertIn('name="fecha_cierre_hora"', html)
        self.assertIn('type="time"', html)
        self.assertIn('name="incluir_fechas"', html)
        self.assertIn("Atención del Ticket", html)

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


class VistaMisMesTests(TestCase):
    """Toggle opt-in de la vista predeterminada «Mis Tickets del Mes»."""

    def setUp(self):
        from django.contrib.auth import get_user_model

        self.usuario = get_user_model().objects.create_superuser(
            "operador.mismes", "mm@test.local", "Prueba123!"
        )
        self.client.force_login(self.usuario)

    def _toggle(self):
        return self.client.post(reverse("enlaces_ccg:vista_mis_mes"))

    def test_activar_crea_vista_predeterminada(self):
        respuesta = self._toggle()
        self.assertEqual(respuesta.status_code, 302)
        vista = VistaGuardada.objects.get(
            usuario=self.usuario, modulo="tickets", nombre="Mis Tickets del Mes"
        )
        self.assertTrue(vista.es_predeterminada)
        self.assertEqual(vista.parametros["responsable"], ["@mios"])
        self.assertEqual(vista.parametros["desde"], ["@mes_inicio"])
        self.assertEqual(vista.parametros["hasta"], ["@mes_fin"])

    def test_activar_reemplaza_otra_predeterminada(self):
        otra = VistaGuardada.objects.create(
            usuario=self.usuario,
            modulo="tickets",
            nombre="Otra",
            parametros={},
            es_predeterminada=True,
        )
        self._toggle()
        otra.refresh_from_db()
        self.assertFalse(otra.es_predeterminada)

    def test_desactivar_elimina_vista(self):
        self._toggle()
        self._toggle()
        self.assertFalse(
            VistaGuardada.objects.filter(
                usuario=self.usuario, nombre="Mis Tickets del Mes"
            ).exists()
        )

    def test_boton_visible_en_lista(self):
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets")
        ).content.decode("utf-8", "replace")
        self.assertIn("Mis Tickets del Mes", html)

    def test_predeterminada_filtra_mis_tickets_del_mes(self):
        resp = ResponsableAtencion.objects.create(nombre="Resp Mio", activo=True)
        self.usuario.responsables_atencion.add(resp)
        Ticket.objects.create(
            ticket_id="M1",
            numero="SS26-M1",
            numero_display="SS26-M1",
            fecha=timezone.now(),
            responsable_atencion=resp,
            estatus=Ticket.ESTATUS_ABIERTO,
            raw_data={"solicitud_solicitante": ""},
        )
        Ticket.objects.create(
            ticket_id="M2",
            numero="SS26-M2",
            numero_display="SS26-M2",
            fecha=timezone.now() - timedelta(days=40),
            responsable_atencion=resp,
            estatus=Ticket.ESTATUS_ABIERTO,
            raw_data={"solicitud_solicitante": ""},
        )
        self._toggle()
        html = self.client.get(
            reverse("enlaces_ccg:seguimiento_tickets")
        ).content.decode("utf-8", "replace")
        self.assertIn("SS26-M1", html)
        self.assertNotIn("SS26-M2", html)


class CatalogoFallasPaginacionTests(TestCase):
    def setUp(self):
        from django.contrib.auth import get_user_model

        self.usuario = get_user_model().objects.create_superuser(
            "admin.fallas", "af@test.local", "Prueba123!"
        )
        self.client.force_login(self.usuario)

    def test_paginacion_muestra_botones_arriba_y_abajo(self):
        for i in range(60):
            Falla.objects.create(descripcion=f"Falla {i:03d}")
        html = self.client.get(
            reverse("enlaces_ccg:catalogo_fallas")
        ).content.decode("utf-8", "replace")
        self.assertIn("Mostrando 1–50 de 60 fallas", html)
        self.assertIn("Página 1 de 2", html)
        self.assertIn("page=2", html)
        self.assertIn("card-footer", html)

    def test_pagina_dos_cambia_contenido(self):
        for i in range(60):
            Falla.objects.create(descripcion=f"Falla {i:03d}")
        html = self.client.get(
            reverse("enlaces_ccg:catalogo_fallas"), {"page": "2"}
        ).content.decode("utf-8", "replace")
        self.assertIn("Mostrando 51–60 de 60 fallas", html)


class PerfilInstitucionPermisosTests(TestCase):
    """En la ficha de institución solo quien puede editar ve los botones."""

    def setUp(self):
        self.inst = Institucion.objects.create(nombre="Institución Test", siglas="IT")
        self.enlace = EnlaceAutorizado.objects.create(
            nombres="Ana", primer_apellido="Perez", institucion=self.inst
        )

    def _usuario(self, username, capacidades):
        from django.contrib.auth import get_user_model
        from django.contrib.auth.models import Group

        usuario = get_user_model().objects.create_user(
            username, f"{username}@test.local", "Clave123!"
        )
        grupo = Group.objects.create(name=f"G-{username}")
        PermisoGrupo.objects.create(grupo=grupo, capacidades=capacidades)
        usuario.groups.add(grupo)
        return usuario

    def test_sin_permiso_no_ve_botones_de_edicion(self):
        usuario = self._usuario("solo.directorio", [CAP_DIRECTORIO])
        self.client.force_login(usuario)
        html = self.client.get(
            reverse("enlaces_ccg:perfil_institucion", args=[self.inst.pk])
        ).content.decode("utf-8", "replace")
        self.assertNotIn('data-crear-enlace class="btn', html)
        self.assertNotIn("btn-outline-primary btn-sm btn-editar-enlace", html)

    def test_con_permiso_ve_botones_de_edicion(self):
        usuario = self._usuario("editor", [CAP_DIRECTORIO, CAP_EDITAR])
        self.client.force_login(usuario)
        html = self.client.get(
            reverse("enlaces_ccg:perfil_institucion", args=[self.inst.pk])
        ).content.decode("utf-8", "replace")
        self.assertIn('data-crear-enlace class="btn', html)
        self.assertIn("btn-outline-primary btn-sm btn-editar-enlace", html)

    def test_sin_permiso_no_puede_crear_ni_editar(self):
        usuario = self._usuario("sin.editar", [CAP_DIRECTORIO])
        self.client.force_login(usuario)
        r_crear = self.client.post(
            reverse("enlaces_ccg:crear_enlace"),
            {"nombres": "X", "primer_apellido": "Y"},
        )
        self.assertEqual(r_crear.status_code, 403)
        r_editar = self.client.post(
            reverse("enlaces_ccg:editar_enlace", args=[self.enlace.pk]),
            {"nombres": "Z", "primer_apellido": "W"},
        )
        self.assertEqual(r_editar.status_code, 403)
        self.enlace.refresh_from_db()
        self.assertEqual(self.enlace.nombres, "Ana")