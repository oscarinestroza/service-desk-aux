from django.db import migrations

CAP_ATENDER = "atender"

GRUPOS_CON_ATENDER = ("Administrador", "Encargado de Servicio")


def sembrar(apps, schema_editor):
    Group = apps.get_model("auth", "Group")
    PermisoGrupo = apps.get_model("enlaces_ccg", "PermisoGrupo")
    for nombre in GRUPOS_CON_ATENDER:
        grupo = Group.objects.filter(name=nombre).first()
        if grupo is None:
            continue
        permiso = PermisoGrupo.objects.filter(grupo=grupo).first()
        if permiso is None:
            continue
        capacidades = list(permiso.capacidades or [])
        if CAP_ATENDER not in capacidades:
            capacidades.append(CAP_ATENDER)
            permiso.capacidades = capacidades
            permiso.save()


def quitar(apps, schema_editor):
    Group = apps.get_model("auth", "Group")
    PermisoGrupo = apps.get_model("enlaces_ccg", "PermisoGrupo")
    for nombre in GRUPOS_CON_ATENDER:
        grupo = Group.objects.filter(name=nombre).first()
        if grupo is None:
            continue
        permiso = PermisoGrupo.objects.filter(grupo=grupo).first()
        if permiso is None:
            continue
        capacidades = [c for c in (permiso.capacidades or []) if c != CAP_ATENDER]
        permiso.capacidades = capacidades
        permiso.save()


class Migration(migrations.Migration):

    dependencies = [
        ("enlaces_ccg", "0043_configuraciontickets_dias_semana_and_more"),
    ]

    operations = [
        migrations.RunPython(sembrar, quitar),
    ]