# Servidor MCP — Finanzas Familia

Expone las tools del núcleo (`core/finanzas_core.py`) para usarlas **desde Claude**
escribiéndole en lenguaje natural.

## Instalar
```
pip install -r mcp/requirements.txt
```

## Conectar en Claude Desktop
Editá `claude_desktop_config.json` (Configuración → Developer → Edit Config) y agregá:
```json
{
  "mcpServers": {
    "finanzas-familia": {
      "command": "py",
      "args": ["C:\\Users\\marti\\documents\\proyectojose\\mcp\\server.py"]
    }
  }
}
```
Reiniciá Claude Desktop. Vas a ver las tools disponibles y podés decir:
"Cargá un egreso de 48.500 de luz pagado con débito".

## Usarlo en la nube (claude.ai)
Para claude.ai web hay que **hostear** el server como MCP remoto (HTTP/SSE) en internet
con autenticación. Es un paso aparte; pedímelo cuando quieras y te guío con el deploy.
