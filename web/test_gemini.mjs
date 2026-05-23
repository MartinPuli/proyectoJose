// Prueba REAL de la integración con Gemini (function calling) sobre el Excel local.
// Correr:  node --env-file=.env.local test_gemini.mjs
import { fileURLToPath } from "url";
import { dirname, resolve } from "path";
import { readFileSync } from "fs";
import { loadWB } from "./lib/excel.js";
import { procesar } from "./lib/gemini.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(__dirname, "..");
const wb = await loadWB(readFileSync(resolve(ROOT, "Finanzas_Familia_2026.xlsx")));

const texto =
  "Cobré el alquiler del Local 12, 350 mil por transferencia. " +
  "Y pagué la factura de luz, 48.500 con débito. " +
  "Ah, y el IPC de abril dio 7600.";

console.log("Mensaje de prueba:\n  " + texto + "\n");
const res = await procesar({ texto, wb });
console.log("Resumen de Gemini:\n  " + res.resumen + "\n");
console.log("Operaciones que escribió:");
for (const o of res.operaciones) console.log("  -", JSON.stringify(o));
console.log("\nTotal operaciones:", res.operaciones.length);
process.exit(res.operaciones.length > 0 ? 0 : 2);
