"""Lógica de gestión de descargas."""

from uuid import uuid4
from typing import Dict, Optional

from .reportes import generar_pdf, generar_excel

# Almacén en memoria de resultados de descargas
_descargas: Dict[str, Dict[str, bytes]] = {}


def iniciar_descarga() -> str:
    """Inicia una descarga y genera los reportes asociados.

    Returns:
        Identificador único de la descarga.
    """
    descarga_id = str(uuid4())
    _descargas[descarga_id] = {
        "pdf": generar_pdf(descarga_id),
        "excel": generar_excel(descarga_id),
    }
    return descarga_id


def obtener_pdf(descarga_id: str) -> Optional[bytes]:
    """Obtiene el PDF generado para una descarga."""
    descarga = _descargas.get(descarga_id)
    if descarga:
        return descarga.get("pdf")
    return None


def obtener_excel(descarga_id: str) -> Optional[bytes]:
    """Obtiene el Excel generado para una descarga."""
    descarga = _descargas.get(descarga_id)
    if descarga:
        return descarga.get("excel")
    return None
