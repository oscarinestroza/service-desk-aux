# Migración manual: catálogo global de secciones reutilizables entre carpetas.
#
# Transforma el modelo per-carpeta `SeccionDocumento` en un catálogo global
# `Seccion` (nombre único reutilizable en varias carpetas) más `CarpetaSeccion`
# (asignación carpeta <-> sección con orden).
#
# Documento.seccion pasa a referenciar `Seccion`. Para no perder la asignación
# existente se respalda en una columna temporal antes de recrear la FK.

import django.db.models.deletion
from django.db import migrations, models


def migrar_a_catalogo(apps, schema_editor):
    """Crea Seccion/CarpetaSeccion a partir de SeccionDocumento y respalda la
    asignación de Documento.seccion en una columna temporal (nuevos ids)."""
    Seccion = apps.get_model("enlaces_ccg", "Seccion")
    CarpetaSeccion = apps.get_model("enlaces_ccg", "CarpetaSeccion")
    SeccionDocumento = apps.get_model("enlaces_ccg", "SeccionDocumento")

    mapa_antiguo_nuevo = {}
    carpeta_ids = list(SeccionDocumento.objects.order_by("carpeta_id").values_list("carpeta_id", flat=True).distinct())
    for carpeta_id in carpeta_ids:
        for orden, cs in enumerate(
            SeccionDocumento.objects.filter(carpeta_id=carpeta_id).order_by("orden", "nombre")
        ):
            seccion, _ = Seccion.objects.get_or_create(nombre=cs.nombre)
            mapa_antiguo_nuevo[cs.pk] = seccion.pk
            CarpetaSeccion.objects.get_or_create(
                carpeta_id=carpeta_id, seccion_id=seccion.pk, defaults={"orden": orden}
            )

    tabla = schema_editor.quote_name("enlaces_ccg_documento")
    with schema_editor.connection.cursor() as cursor:
        cursor.execute("ALTER TABLE {} ADD COLUMN _seccion_tmp INTEGER NULL".format(tabla))
        cursor.execute("SELECT id, seccion_id FROM {} WHERE seccion_id IS NOT NULL".format(tabla))
        for documento_id, seccion_antigua in cursor.fetchall():
            nuevo = mapa_antiguo_nuevo.get(seccion_antigua)
            if nuevo is not None:
                cursor.execute(
                    "UPDATE {} SET _seccion_tmp = %s WHERE id = %s".format(tabla),
                    [nuevo, documento_id],
                )


def restaurar_seccion(apps, schema_editor):
    """Copia la asignación respaldada a la nueva FK y elimina la columna temporal."""
    tabla = schema_editor.quote_name("enlaces_ccg_documento")
    with schema_editor.connection.cursor() as cursor:
        cursor.execute(
            "UPDATE {} SET seccion_id = _seccion_tmp WHERE _seccion_tmp IS NOT NULL".format(tabla)
        )
        cursor.execute("ALTER TABLE {} DROP COLUMN _seccion_tmp".format(tabla))


class Migration(migrations.Migration):

    dependencies = [
        ("enlaces_ccg", "0028_secciondocumento_documento_seccion_and_more"),
    ]

    operations = [
        migrations.CreateModel(
            name="Seccion",
            fields=[
                (
                    "id",
                    models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID"),
                ),
                (
                    "nombre",
                    models.CharField(
                        help_text="Nombre de la sección (catálogo compartido entre carpetas).",
                        max_length=200,
                        unique=True,
                        verbose_name="Nombre",
                    ),
                ),
            ],
            options={
                "verbose_name": "Sección de documentos",
                "verbose_name_plural": "Secciones de documentos",
                "ordering": ["nombre"],
            },
        ),
        migrations.CreateModel(
            name="CarpetaSeccion",
            fields=[
                (
                    "id",
                    models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID"),
                ),
                ("orden", models.PositiveIntegerField(default=0)),
                (
                    "carpeta",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="secciones_asignadas",
                        to="enlaces_ccg.documentocarpeta",
                        verbose_name="Carpeta",
                    ),
                ),
                (
                    "seccion",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="carpetas_asignadas",
                        to="enlaces_ccg.seccion",
                        verbose_name="Sección",
                    ),
                ),
            ],
            options={
                "verbose_name": "Sección asignada",
                "verbose_name_plural": "Secciones asignadas",
                "ordering": ["orden", "seccion__nombre"],
            },
        ),
        migrations.AddConstraint(
            model_name="carpetaseccion",
            constraint=models.UniqueConstraint(fields=("carpeta", "seccion"), name="seccion_unica_por_carpeta"),
        ),
        migrations.RunPython(migrar_a_catalogo, migrations.RunPython.noop),
        migrations.RemoveField(
            model_name="documento",
            name="seccion",
        ),
        migrations.AddField(
            model_name="documento",
            name="seccion",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.SET_NULL,
                related_name="documentos",
                to="enlaces_ccg.seccion",
                verbose_name="Sección",
            ),
        ),
        migrations.RunPython(restaurar_seccion, migrations.RunPython.noop),
        migrations.RemoveConstraint(
            model_name="secciondocumento",
            name="unica_seccion_por_carpeta",
        ),
        migrations.RemoveField(
            model_name="secciondocumento",
            name="carpeta",
        ),
        migrations.DeleteModel(
            name="SeccionDocumento",
        ),
    ]