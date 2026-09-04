import os

from celery import Celery

# Configurar el módulo de settings por defecto para 'celery' cli.
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "enlacesMAO.settings")

app = Celery("enlacesMAO")

# Namespace CELERY_ en settings.py (ej. CELERY_BROKER_URL). 'task' es el
# namespace por defecto de Celery; se preserva con task_default_* si hiciera falta.
app.config_from_object("django.conf:settings", namespace="CELERY")

# Auto-descubre tareas en los apps registrados (task decorador @shared_task).
app.autodiscover_tasks()

# Las tareas SIG viven en un subpaquete (enlaces_ccg/sig_sync/tasks.py) que el
# autodiscover no alcanza. Se importan aquí para que se registren al arrancar.
import enlaces_ccg.sig_sync.tasks  # noqa: E402,F401


@app.task(bind=True)
def debug_task(self):
    """Tarea de depuración útil para verificar que el worker responde."""
    print(f"Request: {self.request!r}")