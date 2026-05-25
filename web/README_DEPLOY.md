# Deploy en Vercel — Finanzas Familia

App Next.js que procesa audio/texto/foto/PDF con **Gemini** y guarda el Excel en **Vercel Blob**.

## Estado del proyecto

`npm run build` compila sin errores. Las rutas API están implementadas:
- `GET  /api/summary` → resumen anual de ingresos/egresos
- `POST /api/process` → procesa con Gemini y escribe al Excel
- `POST /api/upload`  → reemplaza la planilla con una versión editada
- `GET  /api/excel`   → descarga la planilla

## Por qué Blob

Vercel es serverless: no puede guardar archivos en disco entre requests. El `.xlsx` vive en Vercel Blob; la app lo descarga, le agrega filas y lo vuelve a subir.

---

## Pasos para hacer el deploy

### 1. Crear el proyecto en Vercel

- Conectá el repo de GitHub (o usá `vercel` CLI).
- En **Project Settings → General → Root Directory** escribí: `web`
- Framework Preset: **Next.js** (Vercel lo detecta solo)

### 2. Activar Vercel Blob

- Ir a **Storage → Create Database → Blob**
- Activar **Private** access
- Conectar el store al proyecto → Vercel agrega `BLOB_READ_WRITE_TOKEN` automáticamente

### 3. Agregar variables de entorno

En **Project Settings → Environment Variables**:

| Variable | Valor |
|---|---|
| `GEMINI_API_KEY` | Tu clave de https://aistudio.google.com/apikey |
| `GEMINI_MODEL` | `gemini-2.5-flash` |
| `BLOB_READ_WRITE_TOKEN` | Lo agrega Vercel al activar Blob |

### 4. Subir la planilla base a Blob (una sola vez)

Desde la raíz del proyecto, con el token en el entorno local:

```powershell
# PowerShell:
$env:BLOB_READ_WRITE_TOKEN = "vercel_blob_rw_..."
cd web
npm run seed
```

Esto sube `Finanzas_Familia_2026.xlsx` al Blob store. Solo se hace una vez.

### 5. Deploy

```bash
vercel   # o push al repo conectado a Vercel
```

---

## Probar en local

Creá `web/.env.local` con las tres variables:

```
GEMINI_API_KEY=AIza...
GEMINI_MODEL=gemini-2.5-flash
BLOB_READ_WRITE_TOKEN=vercel_blob_rw_...
```

Luego:

```bash
cd web
npm run dev   # http://localhost:3000
```

Sin `BLOB_READ_WRITE_TOKEN`, la app lee/escribe el `.xlsx` directamente del disco
(desde la raíz del proyecto), sin necesidad de Blob.

---

## Notas

- Audio del navegador = webm; si Gemini lo rechaza, subí un mp3/wav.
- Cerrar la planilla en Excel antes de procesar (Windows la bloquea al escribir).
- Cada escritura hace backup automático en `backups/`.
- Si usás Vercel Hobby y los procesos con audio tardan más de 10 s, considerá
  subir a Pro (el código ya tiene `maxDuration = 60` en `/api/process`).
- Concurrencia: si dos personas cargan exactamente a la vez puede pisarse una
  escritura (para una familia es muy improbable).
