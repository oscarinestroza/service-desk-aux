"""
Management command: sync_pendientes

Sincroniza enlaces con el SIG que estén pendientes o re-sincroniza todos.

Uso:
    python manage.py sync_pendientes              # Sync pendientes (sin usuario_sig)
    python manage.py sync_pendientes --all        # Re-sincronizar todos
    python manage.py sync_pendientes --id 42      # Sync enlace específico
    python manage.py sync_pendientes --dry-run    # Solo mostrar qué se haría
"""

import logging
from django.core.management.base import BaseCommand, CommandError
from django.conf import settings

from enlaces_ccg.models import EnlaceAutorizado, SyncLog

logger = logging.getLogger("sig_sync")


class Command(BaseCommand):
    help = "Sincroniza enlaces autorizados con el SIG"

    def add_arguments(self, parser):
        parser.add_argument(
            "--all",
            action="store_true",
            help="Re-sincronizar todos los enlaces (no solo pendientes)",
        )
        parser.add_argument(
            "--id",
            type=int,
            help="Sincistrar un enlace específico por ID",
        )
        parser.add_argument(
            "--dry-run",
            action="store_true",
            help="Mostrar qué se haría sin ejecutar",
        )
        parser.add_argument(
            "--sequential",
            action="store_true",
            help="Ejecutar secuencialmente (no en threads)",
        )

    def handle(self, *args, **options):
        if not getattr(settings, "SIG_USER", ""):
            raise CommandError(
                "SIG_USER no está configurado en settings. "
                "Configure las variables de entorno SIG_USER, SIG_PASSWORD, SIG_URL."
            )

        dry_run = options["dry_run"]
        sequential = options["sequential"]
        sync_all = options["all"]
        enlace_id = options["id"]

        if enlace_id:
            # Sync específico
            try:
                enlace = EnlaceAutorizado.objects.get(pk=enlace_id)
            except EnlaceAutorizado.DoesNotExist:
                raise CommandError(f"Enlace con ID {enlace_id} no existe")

            from enlaces_ccg.sig_sync.tasks import sincronizar_enlace, deshabilitar_enlace
            if enlace.usuario_sig:
                deshabilitar_enlace(enlace.pk)
            else:
                sincronizar_enlace(enlace.pk)
            self.stdout.write(f"Enlace {enlace_id} encolado para sync")
            return

        # Determinar enlaces a sincronizar
        qs = EnlaceAutorizado.objects.select_related(
            "institucion"
        ).prefetch_related("institucion__edificio")

        if sync_all:
            # Todos los que tienen usuario_sig o que necesitan crear
            qs = qs.filter(usuario_sig__gt="")
            action = "re-sincronizar"
        else:
            # Pendientes: sin usuario_sig pero con datos suficientes
            qs = qs.filter(usuario_sig="")
            action = "sincronizar nuevos"

        total = qs.count()
        self.stdout.write(f"Enlaces a {action}: {total}")

        if dry_run:
            for enlace in qs[:20]:
                accion = "crear" if not enlace.usuario_sig else "deshabilitar"
                self.stdout.write(
                    f"  [{accion}] {enlace.usuario_sig or 'SIN_USUARIO'} — "
                    f"{enlace.nombre_completo} ({enlace.institucion})"
                )
            if total > 20:
                self.stdout.write(f"  ... y {total - 20} más")
            return

        # Ejecutar sync
        from enlaces_ccg.sig_sync.tasks import sincronizar_enlace, deshabilitar_enlace

        success = 0
        errors = 0

        for enlace in qs:
            try:
                if enlace.usuario_sig:
                    deshabilitar_enlace(enlace.pk)
                else:
                    sincronizar_enlace(enlace.pk)
                success += 1
                self.stdout.write(
                    f"  ✓ {enlace.nombre_completo} — encolado"
                )
            except Exception as e:
                errors += 1
                self.stderr.write(
                    self.style.ERROR(f"  ✗ {enlace.nombre_completo}: {e}")
                )

        self.stdout.write(
            self.style.SUCCESS(
                f"\nCompletado: {success} exitosos, {errors} errores"
            )
        )
