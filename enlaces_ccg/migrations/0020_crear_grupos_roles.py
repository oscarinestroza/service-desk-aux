from django.db import migrations

ROLES = ("Administrador", "Operador MAO", "Encargado de Servicio")


def crear_grupos(apps, schema_editor):
    Group = apps.get_model("auth", "Group")
    Permission = apps.get_model("auth", "Permission")
    for nombre in ROLES:
        grupo, creado = Group.objects.get_or_create(name=nombre)
        # Administrador recibe permisos de Django (staff usa is_staff del usuario)
        # Los permisos finos por vista se controlan con capacidades en roles.py.
        # Asignamos todos los permisos solo al Administrador.
        if nombre == "Administrador":
            for perm in Permission.objects.all():
                grupo.permissions.add(perm)


def eliminar_grupos(apps, schema_editor):
    Group = apps.get_model("auth", "Group")
    Group.objects.filter(name__in=ROLES).delete()


class Migration(migrations.Migration):

    dependencies = [
        ("enlaces_ccg", "0019_renombrar_niveles_a_nivel"),
    ]

    operations = [
        migrations.RunPython(crear_grupos, eliminar_grupos),
    ]
