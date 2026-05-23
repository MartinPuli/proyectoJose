# Puesta en marcha

## Claves
- **GEMINI_API_KEY** (app web): sacala gratis en https://aistudio.google.com/apikey
- No hace falta nada más: Gemini procesa audio, imágenes y PDF de forma nativa.

Copiá `.env.example` a `.env` y completá `GEMINI_API_KEY`.

## App web
```
.\run.ps1
```
Abrí http://localhost:8000

## Servidor MCP (para Claude)
```
pip install -r mcp/requirements.txt
```
Configuralo en Claude Desktop (ver `mcp/README_MCP.md`).

## Notas
- Cerrá la planilla en Excel antes de procesar (Windows la bloquea al escribir).
- Cada escritura hace backup automático en `backups/`.
- Audio del navegador = webm; si Gemini lo rechaza, subí un mp3/wav.
- App pública en internet o MCP en claude.ai (nube) = requieren hosting; pedímelo.
