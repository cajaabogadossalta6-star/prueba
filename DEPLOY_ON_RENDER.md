# Guía para desplegar en Render

Sigue estos pasos para ejecutar la aplicación automáticamente en [Render](https://render.com):

1. **Haz un fork o sube este repositorio a tu cuenta de Git.**
2. Inicia sesión en Render y haz clic en **New ➝ Blueprint**.
3. Proporciona la URL del repositorio. Render detectará el archivo `render.yaml` y creará los servicios automáticamente:
   - API FastAPI.
   - Worker de Celery.
   - Base de datos Redis.
   - Sitio estático con el frontend compilado.
4. Haz clic en **Apply**. Se iniciará la construcción y despliegue de cada servicio.
5. Cuando finalice el despliegue, tendrás disponible la URL pública del frontend y de la API.

No es necesario configurar comandos manuales; todo está definido en `render.yaml`.
