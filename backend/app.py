"""Aplicación FastAPI para gestionar descargas y licencias."""

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
import io
import requests

from .core.descargas import (
    iniciar_descarga,
    obtener_pdf,
    obtener_excel,
)

app = FastAPI(title="API de Descargas")

FIREBASE_URL = "https://recibos-anses-default-rtdb.firebaseio.com"


@app.post("/descargas", summary="Inicia una descarga", tags=["Descargas"])
def post_descargas():
    """Inicia una nueva tarea de descarga.

    **Ejemplo**

    ```bash
    curl -X POST http://localhost:8000/descargas
    ```
    """
    descarga_id = iniciar_descarga()
    return {"id": descarga_id}


@app.get("/descargas/{descarga_id}/pdf", summary="Obtiene el PDF de una descarga", tags=["Descargas"])
def get_descarga_pdf(descarga_id: str):
    """Devuelve el PDF asociado a una descarga.

    **Ejemplo**

    ```bash
    curl -X GET http://localhost:8000/descargas/<id>/pdf -o reporte.pdf
    ```
    """
    pdf = obtener_pdf(descarga_id)
    if not pdf:
        raise HTTPException(status_code=404, detail="Descarga no encontrada")
    return StreamingResponse(io.BytesIO(pdf), media_type="application/pdf", headers={"Content-Disposition": f"attachment; filename={descarga_id}.pdf"})


@app.get("/descargas/{descarga_id}/excel", summary="Obtiene el Excel de una descarga", tags=["Descargas"])
def get_descarga_excel(descarga_id: str):
    """Devuelve el Excel asociado a una descarga.

    **Ejemplo**

    ```bash
    curl -X GET http://localhost:8000/descargas/<id>/excel -o reporte.csv
    ```
    """
    excel = obtener_excel(descarga_id)
    if not excel:
        raise HTTPException(status_code=404, detail="Descarga no encontrada")
    return StreamingResponse(io.BytesIO(excel), media_type="text/csv", headers={"Content-Disposition": f"attachment; filename={descarga_id}.csv"})


@app.get("/licencias/{licencia_id}", summary="Valida una licencia", tags=["Licencias"])
def get_licencia(licencia_id: str):
    """Verifica si una licencia está activa.

    **Ejemplo**

    ```bash
    curl -X GET http://localhost:8000/licencias/<id>
    ```
    """
    try:
        respuesta = requests.get(f"{FIREBASE_URL}/licenses/{licencia_id}.json", timeout=10)
    except requests.RequestException:
        raise HTTPException(status_code=503, detail="Servicio de licencias no disponible")

    if respuesta.status_code != 200:
        raise HTTPException(status_code=respuesta.status_code, detail="Error al consultar la licencia")

    datos = respuesta.json()
    activa = bool(datos and datos.get("active"))
    return {"id": licencia_id, "activa": activa}
