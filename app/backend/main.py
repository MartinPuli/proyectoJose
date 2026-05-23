"""API web: subir audio/texto/documento, procesar con Gemini y descargar la planilla."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from core import excel_io
import gemini_agent

app = FastAPI(title="Finanzas Familia")
FRONT = Path(__file__).resolve().parents[1] / "frontend" / "index.html"

@app.get("/", response_class=HTMLResponse)
def home():
    return FRONT.read_text(encoding="utf-8")

@app.post("/procesar")
async def procesar(texto: str = Form(""), archivo: UploadFile = File(None)):
    try:
        data = mime = name = None
        if archivo is not None and archivo.filename:
            data = await archivo.read()
            mime = archivo.content_type
            name = archivo.filename
        res = gemini_agent.procesar(texto=texto or "", file_bytes=data, mime=mime, filename=name)
        return JSONResponse(res)
    except Exception as e:
        return JSONResponse({"error": str(e)}, status_code=400)

@app.get("/excel")
def descargar_excel():
    return FileResponse(
        excel_io.EXCEL_PATH,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=excel_io.EXCEL_PATH.name,
    )

@app.get("/resumen")
def resumen():
    from core import finanzas_core
    return finanzas_core.resumen_mensual(0)
