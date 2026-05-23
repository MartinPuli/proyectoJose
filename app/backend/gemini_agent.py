"""Agente Gemini con function calling automático sobre las tools del núcleo.

Gemini entiende audio, imágenes y PDF de forma nativa, así que la misma llamada
sirve para texto, comprobantes (foto/PDF) o audios.
"""
import sys
from pathlib import Path
from datetime import date
sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from google import genai
from google.genai import types
from core import finanzas_core as core
import config

SYSTEM = """Sos el asistente de finanzas de una familia que vive de alquileres de locales.
Hoy es {fecha}. Cuando el usuario te manda un audio, texto o comprobante, identificá las
operaciones y cargalas usando las herramientas disponibles:
- Cobros de alquiler -> agregar_cobro (identificá al inquilino por nombre o local; usá listar_inquilinos si hace falta).
- Gastos o ingresos de la familia -> agregar_movimiento.
- Dato del IPC del INDEC -> actualizar_ipc (mes YYYY-MM).
Reglas: si no dicen el año, asumí el actual; moneda por defecto ARS; para egresos elegí
una categoría razonable. Después de cargar, respondé en español qué registraste."""

def procesar(texto: str = "", file_bytes: bytes = None, mime: str = "", filename: str = "") -> dict:
    if not config.GEMINI_API_KEY:
        raise RuntimeError("Falta GEMINI_API_KEY en .env (sacala de aistudio.google.com).")
    core.get_and_clear_log()
    client = genai.Client(api_key=config.GEMINI_API_KEY)
    parts = []
    if file_bytes:
        parts.append(types.Part.from_bytes(data=file_bytes, mime_type=mime or "application/octet-stream"))
    parts.append(types.Part.from_text(text=texto or "Procesá el comprobante adjunto y cargá lo que corresponda."))
    resp = client.models.generate_content(
        model=config.GEMINI_MODEL,
        contents=[types.Content(role="user", parts=parts)],
        config=types.GenerateContentConfig(
            system_instruction=SYSTEM.format(fecha=date.today()),
            tools=core.TOOLS,
        ),
    )
    return {"resumen": (resp.text or "").strip(), "operaciones": core.get_and_clear_log()}
