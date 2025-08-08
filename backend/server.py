import os
from fastapi import Depends, FastAPI, HTTPException
from fastapi.responses import FileResponse, RedirectResponse
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer

from .storage import (
    generate_presigned_url,
    get_file_path,
    start_cleanup_thread,
)

app = FastAPI()
security = HTTPBearer()
API_TOKEN = os.getenv("API_TOKEN", "changeme")


@app.on_event("startup")
async def startup() -> None:
    # Start background cleanup of old files
    start_cleanup_thread()


def authenticate(credentials: HTTPAuthorizationCredentials = Depends(security)) -> None:
    if credentials.credentials != API_TOKEN:
        raise HTTPException(status_code=401, detail="Unauthorized")


@app.get("/descargas/{file_id}/pdf", dependencies=[Depends(authenticate)])
async def descargar_pdf(file_id: str):
    url = generate_presigned_url(file_id, "pdf")
    if url:
        return RedirectResponse(url)
    file_path = get_file_path(file_id, "pdf")
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    return FileResponse(file_path)


@app.get("/descargas/{file_id}/excel", dependencies=[Depends(authenticate)])
async def descargar_excel(file_id: str):
    url = generate_presigned_url(file_id, "xlsx")
    if url:
        return RedirectResponse(url)
    file_path = get_file_path(file_id, "xlsx")
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    return FileResponse(file_path)
