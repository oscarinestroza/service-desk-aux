"""Pruebas del parser/importador del módulo de tickets (Fase 2)."""

import tempfile
from datetime import datetime
from io import BytesIO
from pathlib import Path

import openpyxl
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone
from unittest import mock

from .models import Ticket
from .sig_sync.tickets import (
    MODO_SYNC_COMPLETO,
    MODO_SYNC_PARCIAL,
    _calcular_numero_display,
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

    def test_sincronizar_post_encola_y_redirige(self):
        with mock.patch(
            "enlaces_ccg.sig_sync.tickets.sincronizar_tickets_task"
        ) as tarea:
            respuesta = self.client.post(
                reverse("enlaces_ccg:ticket_sincronizar"), {"modo": "completo"}
            )
            self.assertEqual(respuesta.status_code, 302)
            self.assertEqual(respuesta.url, reverse("enlaces_ccg:seguimiento_tickets"))
            tarea.delay.assert_called_once_with(modo="completo")

    def test_sincronizar_requiere_post(self):
        respuesta = self.client.get(reverse("enlaces_ccg:ticket_sincronizar"))
        self.assertEqual(respuesta.status_code, 405)