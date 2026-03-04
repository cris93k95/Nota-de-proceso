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
4. Deploy.

### Opción con archivo de configuración

Este proyecto ya incluye `render.yaml`, por lo que Render puede tomar la configuración automáticamente.

## Uso en iPad

- Una vez desplegada, abre la URL pública (`https://...onrender.com`) en Safari.
- Puedes agregarla a pantalla de inicio para uso más cómodo.

## Nota importante sobre datos

- El estado se guarda en `data.json` dentro del servidor.
- En hosting gratuito, el almacenamiento puede ser efímero.
- Recomendado: usar **Exportar JSON** regularmente como respaldo.
- Siguiente mejora sugerida: migrar almacenamiento a base de datos persistente (por ejemplo, PostgreSQL).

## Incluye

- Cursos, estudiantes, periodos y clases
- Marcas C/I/S por clase
- Cálculo de nota 1.0 a 7.0 con exigencia configurable
- PDF individual y grupal desde backend Python
- Importar Excel (primera columna = nombre estudiante)
- Exportar/Importar JSON
