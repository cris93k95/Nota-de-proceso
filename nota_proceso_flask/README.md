# Nota de Proceso (Flask)

## Ejecutar

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abrir: http://127.0.0.1:5000

## Deploy en Render (recomendado)

### Opción rápida (desde Render UI)

1. Sube esta carpeta a un repositorio en GitHub.
2. En Render, crea un **Web Service** conectado al repo.
3. Configura:
	- Build Command: `pip install -r requirements.txt`
	- Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT`
	- Root Directory: `nota_proceso_flask`
	- Variable `DATABASE_URL`: conexión PostgreSQL (si usas `render.yaml`, se configura sola)
	- Variable `SECRET_KEY`: clave secreta Flask (obligatoria en producción)
	- Variable `GOOGLE_CLIENT_ID`: OAuth Client ID de Google
	- Variable `GOOGLE_CLIENT_SECRET`: OAuth Client Secret de Google
	- Variable `ADMIN_EMAIL`: correo Google del administrador
	- Variable `COLLAB_EMAIL`: correo Google del docente colaborador
	- Opcional: `ADMIN_NAME` y `COLLAB_NAME`
4. Deploy.

### Opción con archivo de configuración

Este proyecto ya incluye `render.yaml` con:

- Web service Flask
- Base PostgreSQL administrada (`nota-proceso-db`)
- Enlace automático `DATABASE_URL`

Con eso tendrás persistencia real (no se pierde al reiniciar el servicio).

## Uso en iPad

- Una vez desplegada, abre la URL pública (`https://...onrender.com`) en Safari.
- Puedes agregarla a pantalla de inicio para uso más cómodo.

## Nota importante sobre datos

- Si existe `DATABASE_URL`, la app guarda estado en PostgreSQL (tabla `app_state`).
- Si no existe `DATABASE_URL`, usa `data.json` local como respaldo.
- Recomendado: usar **Exportar JSON** regularmente como backup adicional.

### Migración segura (usuarios)

- Al iniciar con la nueva versión, se crean tablas `app_users` y `user_state`.
- El contenido histórico de `app_state` se copia automáticamente al usuario `ADMIN_EMAIL` la primera vez.
- No se elimina `app_state`, por lo que los datos previos no se pierden.

## Incluye

- Cursos, estudiantes, periodos y clases
- Marcas C/I/S por clase
- Cálculo de nota 1.0 a 7.0 con exigencia configurable
- PDF individual y grupal desde backend Python
- Importar Excel (primera columna = nombre estudiante)
- Exportar/Importar JSON
