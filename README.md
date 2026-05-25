# Conexia SEROS Bot

## Arquitectura
- **backend/** → FastAPI + Playwright → deploy en **PikaPods** (puerto 8000)
- **frontend/** → HTML/JS estático → deploy en **Vercel**

## Deploy paso a paso

### 1. Backend en PikaPods
1. Subir carpeta `backend/` a un repo GitHub (ej: `conexia-seros-backend`)
2. PikaPods → Add Pod → From Dockerfile → URL del repo → Puerto: `8000`
3. Variable de entorno: `VERCEL_URL=https://tu-app.vercel.app`
4. Guardar la URL que te da PikaPods (ej: `https://abc123.pikapods.com`)

### 2. Frontend en Vercel
1. Subir carpeta `frontend/` a otro repo GitHub (ej: `conexia-seros-frontend`)
2. Vercel → New Project → importar ese repo → Deploy
3. En `index.html` reemplazar `TU-APP.pikapods.com` con la URL real del backend

### 3. Uso
1. Abrir la URL de Vercel
2. Subir `credenciales_seros.xlsx`
3. Clic en **Iniciar extracción**
4. Activar notificaciones con el botón 🔔
5. Al terminar: notificación del navegador + sonido + botones de descarga

## Seguridad
- El Excel de credenciales **nunca** se sube a GitHub
- Está en `.gitignore`
- Se sube directamente desde la interfaz web a PikaPods
