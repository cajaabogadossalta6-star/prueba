import os
import time
from celery import Celery

# Celery app configured to use Redis by default
CELERY_BROKER_URL = os.environ.get("CELERY_BROKER_URL", "redis://redis:6379/0")
CELERY_RESULT_BACKEND = os.environ.get("CELERY_RESULT_BACKEND", CELERY_BROKER_URL)

celery_app = Celery(
    "backend.tasks",
    broker=CELERY_BROKER_URL,
    backend=CELERY_RESULT_BACKEND,
)

@celery_app.task(bind=True)
def proceso_descarga(self, pasos: int = 10, pausa: float = 1.0):
    """Simula un proceso de descarga y reporta progreso.

    Args:
        pasos: Número de iteraciones a ejecutar.
        pausa: Segundos a esperar entre pasos.
    Returns:
        Información final de la tarea.
    """
    for i in range(pasos):
        time.sleep(pausa)
        self.update_state(state="PROGRESS", meta={"current": i + 1, "total": pasos})
    return {"current": pasos, "total": pasos, "status": "Completado"}
