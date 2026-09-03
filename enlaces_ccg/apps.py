from django.apps import AppConfig


class EnlacesCcgConfig(AppConfig):
    name = 'enlaces_ccg'

    def ready(self):
        import enlaces_ccg.signals  # noqa: F401
