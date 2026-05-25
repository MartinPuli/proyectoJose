/**
 * Rellena las categorías y presupuestos sugeridos en la hoja Presupuesto.
 * - Reemplaza la fórmula `Parametros!I12+` por el nombre de categoría hardcodeado
 * - Llena la columna "Presup. mensual" con valores sugeridos (modificables a mano)
 *
 * Ejecutar desde web/: node scripts/rellenar_presupuesto.mjs
 */
import ExcelJS from "exceljs";
import { readFileSync, writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const XLSX = resolve(__dirname, "../../Finanzas_Familia_2026.xlsx");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(readFileSync(XLSX));

// Categorías + presupuestos mensuales sugeridos (ARS). El usuario puede ajustar.
const PRESUPUESTO = [
  { cat: "Vivienda/Expensas",            mensual: 100000 },
  { cat: "Servicios (luz/gas/agua)",     mensual: 60000  },
  { cat: "Internet/Teléfono",            mensual: 30000  },
  { cat: "Supermercado/Comida",          mensual: 350000 },
  { cat: "Salud/Obra social",            mensual: 80000  },
  { cat: "Transporte/Auto",              mensual: 100000 },
  { cat: "Educación",                    mensual: 50000  },
  { cat: "Impuestos (ABL/AFIP/Rentas)",  mensual: 80000  },
  { cat: "Mantenimiento locales",        mensual: 50000  },
  { cat: "Seguros",                      mensual: 40000  },
  { cat: "Honorarios/Gestión",           mensual: 30000  },
  { cat: "Ocio/Viajes",                  mensual: 80000  },
  { cat: "Indumentaria",                 mensual: 40000  },
  { cat: "Ahorro/Inversión",             mensual: 100000 },
  { cat: "Otros",                        mensual: 30000  },
];

const ws = wb.getWorksheet("Presupuesto");
if (!ws) throw new Error("No hay Presupuesto");

// Filas 5-19: 15 categorías
for (let i = 0; i < PRESUPUESTO.length; i++) {
  const r = 5 + i;
  const { cat, mensual } = PRESUPUESTO[i];
  // Col A: nombre de categoría (texto plano, no fórmula)
  ws.getRow(r).getCell(1).value = cat;
  // Col B: presupuesto mensual sugerido
  ws.getRow(r).getCell(2).value = mensual;
}

// Cambiar la nota (fila 2) para aclarar que los montos son sugeridos
const noteRow = ws.getRow(2);
noteRow.getCell(1).value = "Cargá / ajustá el presupuesto mensual por categoría. El gasto real se calcula automáticamente desde Movimientos.";

wb.calcProperties = { fullCalcOnLoad: true };
const buf = await wb.xlsx.writeBuffer();
writeFileSync(XLSX, Buffer.from(buf));
console.log("✅ Presupuesto rellenado");

// Verificación
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(readFileSync(XLSX));
const ws2 = wb2.getWorksheet("Presupuesto");
console.log("\nCategorías y montos mensuales:");
for (let i = 0; i < 15; i++) {
  const r = 5 + i;
  const a = ws2.getRow(r).getCell(1).value;
  const b = ws2.getRow(r).getCell(2).value;
  console.log(`  ${a}: $${b?.toLocaleString?.("es-AR") ?? b}`);
}
