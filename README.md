# API de Descargas

Esta API expone endpoints para iniciar descargas y obtener reportes en formato PDF o Excel.

## Endpoints

### `POST /descargas`
Inicia una nueva descarga.
```bash
curl -X POST http://localhost:8000/descargas
```

### `GET /descargas/{id}/pdf`
Obtiene el PDF generado para la descarga especificada.
```bash
curl -L -o reporte.pdf http://localhost:8000/descargas/<id>/pdf
```

### `GET /descargas/{id}/excel`
Obtiene el archivo Excel generado para la descarga especificada.
```bash
curl -L -o reporte.csv http://localhost:8000/descargas/<id>/excel
```

### `GET /licencias/{id}`
Valida si una licencia está activa.
```bash
curl http://localhost:8000/licencias/<id>
```

La documentación interactiva de la API está disponible en `/docs` al ejecutar la aplicación con `uvicorn backend.app:app --reload`.
