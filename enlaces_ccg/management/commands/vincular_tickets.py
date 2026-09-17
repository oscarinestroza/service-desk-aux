"""Management command para (re)vincular los tickets con el directorio (Fase 3).

Recorre `Ticket.raw_data` y resuelve torre, institución, solicitante,
servicio, falla y nivel usando los catálogos y el directorio actuales.
Es idempotente y re-ejecutable (por ejemplo, cuando se agregan enlaces).

Uso:
    python manage.py vincular_tickets [--batch 2000] [--todos]

Por defecto solo revisa tickets sin vínculos (torre/servicio/falla/solicitante
nulos) y los que no han empatado solicitante. Con --todos reprocesa el
histórico completo.
"""

from django.core.management.base import BaseCommand
from django.db import transaction
from django.db.models import Q

from enlaces_ccg.sig_sync.tickets import (
    cargar_caches_vinculos,
    _resolver_vinculos,
)

CAMPOS_VINCULO = (
    "torre", "institucion", "solicitante", "servicio", "falla",
    "nivel", "solicitante_nombre",
)
CAMPOS_FK = ("torre", "institucion", "solicitante", "servicio", "falla", "nivel")


class Command(BaseCommand):
    help = "Resuelve los vínculos de los tickets con el directorio (Fase 3)."

    def add_arguments(self, parser):
        parser.add_argument(
            "--batch", type=int, default=2000,
            help="Tamaño del bloque para bulk_update (default 2000).",
        )
        parser.add_argument(
            "--todos", action="store_true",
            help="Reprocesa todos los tickets (no solo los pendientes).",
        )

    def handle(self, *args, **options):
        from enlaces_ccg.models import Ticket

        batch = max(1, options["batch"])
        caches = cargar_caches_vinculos()

        qs = Ticket.objects.all()
        if not options["todos"]:
            qs = qs.filter(
                Q(torre__isnull=True)
                | Q(servicio__isnull=True)
                | Q(falla__isnull=True)
                | Q(nivel__isnull=True)
                | Q(solicitante__isnull=True)
            )
        qs = qs.only("pk", "raw_data", *CAMPOS_VINCULO).order_by("pk")

        total = qs.count()
        self.stdout.write(f"Revisando {total} tickets ...")

        procesados = actualizados = 0
        pendientes = []

        def _flush():
            nonlocal actualizados
            if not pendientes:
                return
            with transaction.atomic():
                Ticket.objects.bulk_update(pendientes, CAMPOS_VINCULO)
            actualizados += len(pendientes)
            pendientes.clear()

        for t in qs.iterator(chunk_size=batch):
            procesados += 1
            vinculos = _resolver_vinculos(t.raw_data, caches)
            cambia = False
            for campo, valor in vinculos.items():
                if campo in CAMPOS_FK:
                    actual = getattr(t, f"{campo}_id")
                    nuevo = valor.pk if valor else None
                else:
                    actual = getattr(t, campo)
                    nuevo = valor
                if actual != nuevo:
                    cambia = True
                    break
            if not cambia:
                continue
            for campo, valor in vinculos.items():
                setattr(t, campo, valor)
            pendientes.append(t)
            if len(pendientes) >= batch:
                _flush()

        _flush()
        self.stdout.write(self.style.SUCCESS(
            f"OK: {procesados} tickets revisados, {actualizados} actualizados."
        ))
