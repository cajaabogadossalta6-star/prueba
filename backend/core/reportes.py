"""Funciones de generación de reportes."""

import io


def generar_pdf(descarga_id: str) -> bytes:
    """Genera un PDF ficticio para la descarga.

    Args:
        descarga_id: Identificador de la descarga.

    Returns:
        Contenido en bytes del PDF generado.
    """
    contenido = f"Reporte PDF para {descarga_id}\n".encode("utf-8")
    # Un PDF real requeriría una librería adicional, esto es un marcador de posición.
    return contenido


def generar_excel(descarga_id: str) -> bytes:
    """Genera un archivo Excel ficticio para la descarga.

    Args:
        descarga_id: Identificador de la descarga.

    Returns:
        Contenido en bytes del Excel generado.
    """
    # Generamos un contenido CSV simple como ejemplo.
    contenido = "columna1,columna2\nvalor1,valor2\n".encode("utf-8")
    return contenido
