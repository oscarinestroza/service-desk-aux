# Migración manual: cada carpeta usa UNA sola sección del catálogo global.
#
# - Se añade DocumentoCarpeta.seccion (FK -> Seccion), tomando como sección la
#   primera sección asignada a la carpeta (orden mínimo).
# - Se elimina CarpetaSeccion (asignación múltiple) y Documento.seccion (la
#   sección pasa a ser propiedad de la carpeta).

import django.db.models.deletion
from django.db import migrations, models


def asignar_seccion_a_carpeta(apps, schema_editor):
    DocumentoCarpeta = apps.get_model("enlaces_ccg", "DocumentoCarpeta")
    CarpetaSeccion = apps.get_model("enlaces_ccg", "CarpetaSeccion")
    for carpeta in DocumentoCarpeta.objects.all():
        asignada = CarpetaSeccion.objects.filter(carpeta=carpeta).order_by("orden", "seccion__nombre").first()
        if asignada is not None:
            carpeta.seccion = asignada.seccion
            carpeta.save(update_fields=["seccion"])


class Migration(migrations.Migration):

    dependencies = [
        ("enlaces_ccg", "0029_secciones_catalogo_global"),
    ]

    operations = [
        migrations.AddField(
            model_name="documentocarpeta",
            name="seccion",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name="carpetas",
                to="enlaces_ccg.seccion",
                verbose_name="Sección",
                help_text="Cada carpeta usa una sola sección del catálogo (una sección puede usarse en varias carpetas).",
            ),
        ),
        migrations.RunPython(asignar_seccion_a_carpeta, migrations.RunPython.noop),
        migrations.RemoveField(
            model_name="documento",
            name="seccion",
        ),
        migrations.RemoveConstraint(
            model_name="carpetaseccion",
            name="seccion_unica_por_carpeta",
        ),
        migrations.DeleteModel(
            name="CarpetaSeccion",
        ),
    ]