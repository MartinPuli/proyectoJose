// Prueba el motor exceljs (lib/excel.js) que usa la app de Vercel.
import { fileURLToPath } from "url";
import { dirname, resolve } from "path";
import { readFileSync, writeFileSync } from "fs";
import {
  loadWB, wbToBuffer, agregarCobro, agregarMovimiento,
  actualizarIpc, listarInquilinos, resumenMensual,
} from "./lib/excel.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, "..");
const BASE = resolve(ROOT, "Finanzas_Familia_2026.xlsx");
const OUT = resolve(ROOT, "_test_excel.xlsx");

const fallos = [];
const check = (nombre, cond, extra = "") => {
  console.log((cond ? "  OK  " : " FAIL ") + nombre + (extra !== "" ? "  " + JSON.stringify(extra) : ""));
  if (!cond) fallos.push(nombre);
};

const wb = await loadWB(readFileSync(BASE));

const inq = listarInquilinos(wb);
check("listarInquilinos -> 29", inq.length === 29, inq.length);

const r1 = agregarCobro(wb, { monto: 350000, inquilino: "Local 12", fecha: "2026-05-10", medio_pago: "Transferencia" });
check("agregarCobro resuelve id por local", r1.id_inquilino === 12, r1);

const r2 = agregarCobro(wb, { monto: 1000, id_inquilino: 1, moneda: "USD", fecha: "2026-05-11" });
check("agregarCobro USD por id (fila 6)", r2.fila === 6, r2);

const r3 = agregarMovimiento(wb, { monto: 48500, tipo: "Egreso", categoria: "Servicios (luz/gas/agua)", descripcion: "Luz", fecha: "2026-05-09", medio_pago: "Débito", miembro: "Yo" });
check("agregarMovimiento egreso", r3.hoja === "Movimientos", r3);

const r4 = agregarMovimiento(wb, { monto: 900000, tipo: "Ingreso", categoria: "Sueldo", fecha: "2026-05-01", miembro: "Hermana" });
check("agregarMovimiento ingreso", r4.tipo === "Ingreso", r4);

let ipcOk = true;
try { actualizarIpc(wb, { mes: "2026-04", indice: 7600 }); } catch (e) { ipcOk = false; console.log("   IPC error:", e.message); }
check("actualizarIpc abril 2026", ipcOk);

const res = resumenMensual(wb, 5);
check("resumen alquileres mayo", res.ingresos_alquileres === 350000 + 1000 * 1150, res.ingresos_alquileres);
check("resumen otros ingresos", res.otros_ingresos === 900000, res.otros_ingresos);
check("resumen egresos", res.egresos === 48500, res.egresos);
check("resumen neto", res.resultado_neto === 350000 + 1150000 + 900000 - 48500, res.resultado_neto);

// round-trip: guardar y volver a leer
writeFileSync(OUT, await wbToBuffer(wb));
const wb2 = await loadWB(readFileSync(OUT));
const cob = wb2.getWorksheet("Cobros");
check("round-trip: cobro monto en fila 5", cob.getRow(5).getCell(6).value === 350000, cob.getRow(5).getCell(6).value);
check("round-trip: id inquilino en fila 5", cob.getRow(5).getCell(2).value === 12, cob.getRow(5).getCell(2).value);
const mov = wb2.getWorksheet("Movimientos");
check("round-trip: movimiento tipo en fila 5", mov.getRow(5).getCell(3).value === "Egreso", mov.getRow(5).getCell(3).value);
const fI = mov.getRow(5).getCell(9).value;
check("round-trip: fórmula preservada (Mov!I5)", fI && typeof fI === "object" && "formula" in fI, fI);

console.log("\nRESULTADO:", fallos.length ? `${fallos.length} FALLOS: ${fallos}` : "TODO OK");
process.exit(fallos.length ? 1 : 0);
