"""Management command para probar la descarga del Excel de tickets del SIG.

Uso:
    python manage.py descargar_tickets                    # modo rapida (default)
    python manage.py descargar_tickets --modo completa
    python manage.py descargar_tickets --modo completa --desde 2020-01-01
    python manage.py descargar_tickets --report-url /admin/Solicitud/SeguimientoAtencion
"""

from django.core.management.base import BaseCommand, CommandError

from enlaces_ccg.sig_sync.tickets import (
    MODO_COMPLETA,
    MODO_RAPIDA,
    MODO_VALORES,
    descargar_excel_tickets,
)


class Command(BaseCommand):
    help = "Descarga el reporte de tickets desde el SIG (Excel) a media/."

    def add_arguments(self, parser):
        parser.add_argument(
            "--modo",
            choices=MODO_VALORES,
            default=MODO_RAPIDA,
            help="rapida (default, sin filtros, 3 min) o completa (filtros, 10 min).",
        )
        parser.add_argument(
            "--desde",
            default=None,
            help="Fecha 'desde' para la descarga completa (dd/mm/aaaa o AAAA-MM-DD).",
        )
        parser.add_argument(
            "--report-url",
            default=None,
            help="Ruta del reporte (default: ConfiguracionTickets.url_reporte).",
        )

    def handle(self, *args, **options):
        modo = options["modo"]
        desde = options["desde"]
        report_url = options["report_url"]
        try:
            resultado = descargar_excel_tickets(
                modo=modo, desde=desde, report_url=report_url
            )
        except Exception as e:  # noqa: BLE001
            raise CommandError(
                f"Descarga de tickets ({modo}) falló: {e}"
            ) from e
        self.stdout.write(self.style.SUCCESS(
            f"OK [{modo}] {resultado['nombre']} "
            f"({resultado['tamano_bytes']} bytes, {resultado['segundos']} s)"
        ))
        self.stdout.write(self.style.SUCCESS(f"Archivo: {resultado['path']}"))