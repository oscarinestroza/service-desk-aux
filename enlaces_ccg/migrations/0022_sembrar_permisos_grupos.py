from django.db import migrations

ROL_ADMIN = "Administrador"
ROL_OPERADOR = "Operador MAO"
ROL_ENCARGADO = "Encargado de Servicio"

CAP_DIRECTORIO = "directorio"
CAP_EDITAR = "editar"
CAP_IMPORTAR = "importar"
CAP_TICKETS = "tickets"
CAP_REVISIONES = "revisiones"
CAP_ADMIN = "admin"

CAPACIDADES_POR_ROL = {
    ROL_ADMIN: [CAP_DIRECTORIO, CAP_EDITAR, CAP_IMPORTAR, CAP_TICKETS, CAP_REVISIONES, CAP_ADMIN],
    ROL_OPERADOR: [CAP_DIRECTORIO, CAP_EDITAR, CAP_TICKETS, CAP_REVISIONES],
    ROL_ENCARGADO: [CAP_DIRECTORIO, CAP_TICKETS],
}


def sembrar(apps, schema_editor):
    Group = apps.get_model("auth", "Group")
    PermisoGrupo = apps.get_model("enlaces_ccg", "PermisoGrupo")
    for nombre, caps in CAPACIDADES_POR_ROL.items():
        grupo = Group.objects.filter(name=nombre).first()
        if grupo is None:
            continue
        PermisoGrupo.objects.update_or_create(
            grupo=grupo,
            defaults={"capacidades": caps},
        )


def quitar(apps, schema_editor):
    PermisoGrupo = apps.get_model("enlaces_ccg", "PermisoGrupo")
    PermisoGrupo.objects.all().delete()


class Migration(migrations.Migration):

    dependencies = [
        ("enlaces_ccg", "0021_permisogrupo"),
    ]

    operations = [
        migrations.RunPython(sembrar, quitar),
    ]