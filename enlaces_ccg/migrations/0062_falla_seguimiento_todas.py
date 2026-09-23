from django.db import migrations


def activar_seguimiento_todas(apps, schema_editor):
    Falla = apps.get_model("enlaces_ccg", "Falla")
    Falla.objects.update(seguimiento=True)


def desactivar_seguimiento_todas(apps, schema_editor):
    Falla = apps.get_model("enlaces_ccg", "Falla")
    Falla.objects.update(seguimiento=False)


class Migration(migrations.Migration):

    dependencies = [
        ("enlaces_ccg", "0061_falla_seguimiento"),
    ]

    operations = [
        migrations.RunPython(
            activar_seguimiento_todas, desactivar_seguimiento_todas
        ),
    ]