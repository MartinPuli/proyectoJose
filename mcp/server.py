"""Servidor MCP de Finanzas Familia. Expone las tools del núcleo a Claude.

Conectar en Claude Desktop (claude_desktop_config.json):

  "mcpServers": {
    "finanzas-familia": {
      "command": "py",
      "args": ["C:\\Users\\marti\\documents\\proyectojose\\mcp\\server.py"]
    }
  }

Luego, desde Claude, podés escribir: "Cargá el cobro del Local 12, 350 mil".
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from mcp.server.fastmcp import FastMCP
from core import finanzas_core as core

mcp = FastMCP("finanzas-familia")

@mcp.tool()
def agregar_cobro(monto: float, inquilino: str = "", id_inquilino: int = 0, fecha: str = "",
                  moneda: str = "ARS", medio_pago: str = "", estado: str = "Cobrado",
                  periodo: str = "", notas: str = "") -> dict:
    """Registra el cobro del alquiler de un inquilino (hoja Cobros)."""
    return core.agregar_cobro(monto, inquilino, id_inquilino, fecha, moneda,
                              medio_pago, estado, periodo, notas)

@mcp.tool()
def agregar_movimiento(monto: float, tipo: str = "Egreso", miembro: str = "Familia",
                       categoria: str = "", descripcion: str = "", fecha: str = "",
                       moneda: str = "ARS", medio_pago: str = "", notas: str = "") -> dict:
    """Registra un ingreso o egreso de la familia (hoja Movimientos)."""
    return core.agregar_movimiento(monto, tipo, miembro, categoria, descripcion,
                                   fecha, moneda, medio_pago, notas)

@mcp.tool()
def actualizar_ipc(mes: str, indice: float) -> dict:
    """Carga el índice IPC del INDEC para un mes (YYYY-MM)."""
    return core.actualizar_ipc(mes, indice)

@mcp.tool()
def listar_inquilinos() -> list:
    """Lista los inquilinos (id, nombre, local)."""
    return core.listar_inquilinos()

@mcp.tool()
def resumen_mensual(mes: int = 0) -> dict:
    """Resumen de ingresos, egresos y resultado neto (mes 1-12, o 0 = año)."""
    return core.resumen_mensual(mes)

if __name__ == "__main__":
    mcp.run()
