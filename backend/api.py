from fastapi import FastAPI, HTTPException
from celery.result import AsyncResult

from backend.tasks import proceso_descarga, celery_app

app = FastAPI()

@app.post("/descargas")
def iniciar_descarga(pasos: int = 10, pausa: float = 1.0):
    """Inicia la descarga asíncrona y devuelve el ID de la tarea."""
    task = proceso_descarga.delay(pasos, pausa)
    return {"id": task.id}

@app.get("/tareas/{task_id}")
def obtener_estado(task_id: str):
    """Devuelve el estado y porcentaje de avance de una tarea."""
    resultado = AsyncResult(task_id, app=celery_app)
    if resultado.state == "PENDING":
        return {"state": resultado.state, "progress": 0}
    if resultado.state == "PROGRESS":
        info = resultado.info or {}
        total = info.get("total", 1)
        current = info.get("current", 0)
        porcentaje = int(current / total * 100)
        return {"state": resultado.state, "progress": porcentaje}
    if resultado.state == "SUCCESS":
        return {"state": resultado.state, "progress": 100, "result": resultado.result}
    if resultado.state == "FAILURE":
        raise HTTPException(status_code=500, detail=str(resultado.info))
    return {"state": resultado.state}
