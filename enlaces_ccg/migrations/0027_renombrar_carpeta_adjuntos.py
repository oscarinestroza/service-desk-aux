from django.db import migrations


def renombrar_carpeta_nuevos_enlaces(apps, schema_editor):
    """Renombra la carpeta por defecto a 'Adjuntos: Nuevos Enlaces MAO'."""
    DocumentoCarpeta = apps.get_model("enlaces_ccg", "DocumentoCarpeta")
    DocumentoCarpeta.objects.filter(slug="adjuntos-nuevos-enlaces").update(
        nombre="Adjuntos: Nuevos Enlaces MAO",
    )


def reversa(apps, schema_editor):
    DocumentoCarpeta = apps.get_model("enlaces_ccg", "DocumentoCarpeta")
    DocumentoCarpeta.objects.filter(slug="adjuntos-nuevos-enlaces").update(
        nombre="Adjuntos para Nuevos enlaces",
    )


class Migration(migrations.Migration):

    dependencies = [
        ('enlaces_ccg', '0026_documentocarpeta_documento'),
    ]

    operations = [
        migrations.RunPython(renombrar_carpeta_nuevos_enlaces, reversa),
    ]