import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("auth", "0012_alter_user_first_name_max_length"),
        ("enlaces_ccg", "0020_crear_grupos_roles"),
    ]

    operations = [
        migrations.CreateModel(
            name="PermisoGrupo",
            fields=[
                (
                    "id",
                    models.BigAutoField(
                        auto_created=True,
                        primary_key=True,
                        serialize=False,
                        verbose_name="ID",
                    ),
                ),
                (
                    "capacidades",
                    models.JSONField(
                        default=list,
                        blank=True,
                        verbose_name="Capacidades",
                        help_text="Lista de claves de capacidad (directorio, editar, importar, tickets, revisiones, admin).",
                    ),
                ),
                (
                    "grupo",
                    models.OneToOneField(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="permiso_grupo",
                        to="auth.group",
                        verbose_name="Grupo",
                    ),
                ),
            ],
            options={
                "verbose_name": "Permisos de grupo",
                "verbose_name_plural": "Permisos de grupos",
            },
        ),
    ]