"""Prueba el núcleo Python (las tools) sobre una copia de la planilla."""
import os, shutil, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TEST_XLSX = ROOT / "_test_core.xlsx"
shutil.copy(ROOT / "Finanzas_Familia_2026.xlsx", TEST_XLSX)
os.environ["EXCEL_PATH"] = str(TEST_XLSX)
sys.path.insert(0, str(ROOT))

from core import finanzas_core as c

fallos = []

def check(nombre, cond, extra=""):
    print(("  OK  " if cond else " FAIL ") + nombre + ("  " + str(extra) if extra else ""))
    if not cond:
        fallos.append(nombre)

inq = c.listar_inquilinos()
check("listar_inquilinos -> 29", len(inq) == 29, len(inq))

r1 = c.agregar_cobro(350000, inquilino="Local 12", fecha="2026-05-10", medio_pago="Transferencia")
check("agregar_cobro resuelve id por local", r1["id_inquilino"] == 12, r1)

r2 = c.agregar_cobro(1000, id_inquilino=1, moneda="USD", fecha="2026-05-11")
check("agregar_cobro USD por id", r2["id_inquilino"] == 1 and r2["hoja"] == "Cobros", r2)

r3 = c.agregar_movimiento(48500, tipo="Egreso", categoria="Servicios (luz/gas/agua)",
                          descripcion="Luz", fecha="2026-05-09", medio_pago="Débito", miembro="Yo")
check("agregar_movimiento egreso", r3["hoja"] == "Movimientos", r3)

r4 = c.agregar_movimiento(900000, tipo="Ingreso", categoria="Sueldo", fecha="2026-05-01", miembro="Hermana")
check("agregar_movimiento ingreso", r4["tipo"] == "Ingreso", r4)

try:
    c.actualizar_ipc("2026-04", 7600.0)
    ipc_ok = True
except Exception as e:
    ipc_ok = False
    print("   IPC error:", e)
check("actualizar_ipc abril 2026", ipc_ok)

res = c.resumen_mensual(5)
check("resumen alquileres mayo", res["ingresos_alquileres"] == 350000 + 1000 * 1150, res["ingresos_alquileres"])
check("resumen otros ingresos", res["otros_ingresos"] == 900000, res["otros_ingresos"])
check("resumen egresos", res["egresos"] == 48500, res["egresos"])
check("resumen neto", res["resultado_neto"] == (350000 + 1150000 + 900000 - 48500), res["resultado_neto"])

print("\nRESULTADO:", "TODO OK" if not fallos else f"{len(fallos)} FALLOS: {fallos}")
sys.exit(1 if fallos else 0)
