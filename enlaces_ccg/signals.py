"""
Django signals para sincronización automática con el SIG.

Se dispara un sync en background cuando:
  1. Se crea un EnlaceAutorizado con usuario_sig → crear usuario en SIG
  2. El estado cambia de ACTIVO a INACTIVO (con usuario_sig) → deshabilitar
  3. El estado cambia de INACTIVO a ACTIVO (con usuario_sig) → re-activar
  4. Cambia el correo_principal (con usuario_sig) → actualizar correo en SIG
"""

import logging
from django.db.models.signals import post_save, pre_save
from django.dispatch import receiver

logger = logging.getLogger("sig_sync")

# Cache temporal para guardar estado y correo anteriores del enlace
_anterior = {}


@receiver(pre_save, sender="enlaces_ccg.EnlaceAutorizado")
def capturar_estado_anterior(sender, instance, **kwargs):
    """Captura el estado y correo anteriores ANTES del save para comparar."""
    if instance.pk:
        try:
            from enlaces_ccg.models import EnlaceAutorizado
            old = EnlaceAutorizado.objects.filter(pk=instance.pk).values(
                "estado", "correo_principal"
            ).first()
            _anterior[instance.pk] = old or {}
        except Exception:
            _anterior[instance.pk] = {}


def _normalizar_mail(valor):
    """Normaliza un correo para comparar (None y "" se tratan igual)."""
    if valor is None:
        return ""
    return str(valor).strip().lower()


@receiver(post_save, sender="enlaces_ccg.EnlaceAutorizado")
def sync_sig_on_enlace_save(sender, instance, created, **kwargs):
    """Dispara sincronización SIG al guardar un EnlaceAutorizado."""
    # Evitar loops infinitos
    if getattr(instance, "_sig_syncing", False):
        return

    usuario_sig = getattr(instance, "usuario_sig", "")

    if created and usuario_sig:
        # Nuevo enlace con usuario SIG → crear en SIG
        logger.info(
            "Signal: enlace %s creado con usuario_sig=%s → queue sync crear",
            instance.pk, usuario_sig,
        )
        from .sig_sync.tasks import sincronizar_enlace
        sincronizar_enlace(instance.pk)
        return

    if not created and usuario_sig:
        previo = _anterior.pop(instance.pk, {})
        old_estado = previo.get("estado")
        old_correo = previo.get("correo_principal")
        new_estado = instance.estado

        # 1) Cambio de estado
        if old_estado and old_estado != new_estado:
            if new_estado == "INACTIVO":
                logger.info(
                    "Signal: enlace %s cambió %s → %s → queue sync deshabilitar",
                    instance.pk, old_estado, new_estado,
                )
                from .sig_sync.tasks import deshabilitar_enlace
                deshabilitar_enlace(instance.pk)
            elif new_estado == "ACTIVO":
                logger.info(
                    "Signal: enlace %s cambió %s → %s → queue sync reactivar",
                    instance.pk, old_estado, new_estado,
                )
                from .sig_sync.tasks import reactivar_enlace
                reactivar_enlace(instance.pk)

        # 2) Cambio de correo principal
        if old_correo is not None and _normalizar_mail(old_correo) != _normalizar_mail(instance.correo_principal):
            logger.info(
                "Signal: enlace %s cambió correo %r → %r → queue sync actualizar correo",
                instance.pk, old_correo, instance.correo_principal,
            )
            from .sig_sync.tasks import actualizar_correo_enlace
            actualizar_correo_enlace(instance.pk)
