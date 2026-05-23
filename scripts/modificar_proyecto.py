"""CLI de mantenimiento.

  py scripts/modificar_proyecto.py regenerar   # regenera la planilla
  py scripts/modificar_proyecto.py validar      # valida fórmulas
  py scripts/modificar_proyecto.py scaffold      # regenera app/mcp/core
  py scripts/modificar_proyecto.py app           # levanta la app web (Gemini)
  py scripts/modificar_proyecto.py mcp           # corre el servidor MCP (para Claude)
  py scripts/modificar_proyecto.py backup        # copia la planilla a backups/
"""
import subprocess, sys, shutil
from datetime import datetime
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]

def _run(args, cwd=None):
    print(">", " ".join(args)); return subprocess.call(args, cwd=str(cwd or ROOT))

def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "help"
    if cmd == "regenerar": _run(["py", "generar_finanzas.py"])
    elif cmd == "validar": _run(["py", "validar.py"])
    elif cmd == "scaffold": _run(["py", "armar_app.py"])
    elif cmd == "app": _run(["uvicorn", "main:app", "--reload", "--port", "8000"], cwd=ROOT / "app" / "backend")
    elif cmd == "mcp": _run(["py", "mcp/server.py"])
    elif cmd == "backup":
        x = ROOT / "Finanzas_Familia_2026.xlsx"; d = ROOT / "backups"; d.mkdir(exist_ok=True)
        out = d / f"{x.stem}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"; shutil.copy2(x, out); print("Backup:", out)
    else: print(__doc__)

if __name__ == "__main__":
    main()
