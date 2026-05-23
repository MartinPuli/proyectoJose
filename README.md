# Finanzas Familia

Economía doméstica de una familia que vive de alquileres (29 inquilinos), con reajuste
por inflación INDEC y carga de datos por **audio, texto o documentos** usando **Gemini**.
Las mismas operaciones se exponen como **MCP** para usarlas desde **Claude**.

## Arquitectura
```
core/            Tools y acceso al Excel (única fuente de verdad)
mcp/             Servidor MCP -> usar desde Claude escribiéndole
app/             Web (FastAPI + Gemini) -> audio/texto/foto/PDF
generar_finanzas.py / validar.py   Generan y validan la planilla
armar_app.py     Scaffolder (regenera core/mcp/app)
scripts/modificar_proyecto.py      CLI de mantenimiento
```

## Tools disponibles
`agregar_cobro` · `agregar_movimiento` · `actualizar_ipc` · `listar_inquilinos` · `resumen_mensual`

## Arrancar la app
```
.\run.ps1
```
http://localhost:8000 · Ver `SETUP.md`.

## Usar desde Claude (MCP)
Ver `mcp/README_MCP.md`.
