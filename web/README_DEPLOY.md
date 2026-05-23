# Deploy en Vercel — Finanzas Familia (web)

App Next.js que procesa audio/texto/foto/PDF con **Gemini** y guarda el Excel en **Vercel Blob**.

## Por qué Blob
Vercel es serverless: no puede guardar archivos en su disco. El `.xlsx` vive en Vercel Blob;
la app lo baja, le agrega filas y lo vuelve a subir. La descarga lo trae desde Blob.

## Pasos
1. Generar la planilla base (desde la raíz del proyecto):
   ```
   py generar_finanzas.py
   ```
2. Instalar dependencias de la web:
   ```
   cd web
   npm install
   ```
3. En Vercel: creá el proyecto apuntando a la carpeta `web/`. Activá **Storage -> Blob**
   (te da `BLOB_READ_WRITE_TOKEN`). Cargá las variables de entorno:
   - `GEMINI_API_KEY` (de https://aistudio.google.com/apikey)
   - `GEMINI_MODEL` (ej. `gemini-2.5-flash`)
   - `BLOB_READ_WRITE_TOKEN`
4. Subir la planilla base a Blob (una sola vez), con el token en el entorno:
   ```
   # PowerShell:  $env:BLOB_READ_WRITE_TOKEN="vercel_blob_rw_..."; npm run seed
   ```
5. Deploy: `vercel` (o conectá el repo de GitHub).

## Probar local
```
cd web
npm run dev      # http://localhost:3000  (necesita las 3 variables en web/.env.local)
```

## Notas honestas
- No pude testear el deploy en vivo desde acá; sí verifiqué que `npm run build` compile.
- exceljs preserva fórmulas/hojas al reescribir; Excel recalcula al abrir.
- Audio del navegador = webm; si Gemini lo rechaza, subí un mp3/wav.
- Concurrencia: si dos personas cargan exactamente a la vez puede pisarse una escritura
  (para una familia es muy improbable).
