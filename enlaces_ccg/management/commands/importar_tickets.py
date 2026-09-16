"""Management command para importar un Excel de tickets ya descargado.

Uso:
    python manage.py importar_tickets <ruta.xlsx> [--modo parcial|completo] [--hoja SIG_Seguimiento_Solicitudes]

El modo completo además archiva los tickets que ya no aparecen en el archivo.
"""

from django.core.management.base import BaseCommand, CommandError

from enlaces_ccg.sig_sync.tickets import (
    MODO_SYNC_COMPLETO,
    MODO_SYNC_PARCIAL,
    MODO_SYNC_VALORES,
    _calcular_numero_display,
    _registrar_resultado,
    aplicar_tickets,
    parsear_excel_tickets,
)


class Command(BaseCommand):
    help = "Importa un Excel de tickets del SIG a la base de datos."

    def add_arguments(self, parser):
        parser.add_argument("ruta", help="Ruta al archivo .xlsx del SIG.")
        parser.add_argument(
            "--modo",
            choices=MODO_SYNC_VALORES,
            default=MODO_SYNC_PARCIAL,
            help="parcial (default, upsert sin tocar ausentes) o completo.",
        )
        parser.add_argument(
            "--hoja", default=None, help="Nombre de la hoja (default: configurada)."
        )

    def handle(self, *args, **options):
        from enlaces_ccg.models import Ticket, TicketLog

        ruta = options["ruta"]
        modo = options["modo"]
        hoja = options["hoja"]

        self.stdout.write(f"Leyendo {ruta} ...")
        try:
            datos, errores = parsear_excel_tickets(ruta, hoja=hoja)
        except FileNotFoundError:
            raise CommandError(f"No existe el archivo: {ruta}")
        except Exception as e:  # noqa: BLE001
            raise CommandError(f"Error al leer el Excel: {e}") from e

        _calcular_numero_display(datos)
        self.stdout.write(
            f"{len(datos)} tickets únicos ({errores} filas con error). Aplicando {modo} ..."
        )
        conteos = aplicar_tickets(datos, modo)
        conteos["errores"] = errores
        mensaje = (
            f"{modo}: {len(datos)} filas, {conteos['creados']} creados, "
            f"{conteos['actualizados']} actualizados, {conteos['archivados']} archivados."
            + (f" {errores} filas con error." if errores else "")
        )
        _registrar_resultado(modo, TicketLog.ESTADO_OK, mensaje, conteos)

        self.stdout.write(self.style.SUCCESS(
            f"OK {modo}: creados {conteos['creados']}, "
            f"actualizados {conteos['actualizados']}, "
            f"archivados {conteos['archivados']}"
        ))
        self.stdout.write(
            f"Total en BD: {Ticket.objects.count()} tickets."
        )