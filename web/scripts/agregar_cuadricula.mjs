/**
 * Agrega una cuadrícula suave a todo el rango usado de cada hoja.
 * Conserva títulos, headers y datos — solo ajusta los bordes.
 *
 * Ejecutar desde web/: node scripts/agregar_cuadricula.mjs
 */
import ExcelJS from "exceljs";
import { readFileSync, writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const XLSX = resolve(__dirname, "../../Finanzas_Familia_2026.xlsx");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(readFileSync(XLSX));

// Color un poco más visible que antes, pero todavía suave
const GRID_COLOR = "FFD1D5DB"; // gris medio claro
const grid = { style: "thin", color: { argb: GRID_COLOR } };
const gridAll = { top: grid, left: grid, bottom: grid, right: grid };

// Para cada hoja, aplicar bordes a todo el rango usado
const PER_SHEET = {
  "Portada":            { skipBorders: true },                  // carátula sin grilla
  "Parametros":         { startRow: 4, cols: 2 },
  "Inquilinos":         { startRow: 4, cols: 15 },
  "Cobros":             { startRow: 4, cols: 12 },
  "Tablero Inquilinos": { startRow: 4, cols: 18 },
  "Movimientos":        { startRow: 4, cols: 11 },
  "Resumen Mensual":    { startRow: 4, cols: 9 },
  "Presupuesto":        { startRow: 4, cols: 6 },
  "Indicadores":        { startRow: 5, cols: 8, twoPanels: true },
};

for (const [name, cfg] of Object.entries(PER_SHEET)) {
  const ws = wb.getWorksheet(name);
  if (!ws) continue;
  if (cfg.skipBorders) { console.log("· saltada:", name); continue; }

  const lastRow = ws.rowCount;
  for (let r = cfg.startRow; r <= lastRow; r++) {
    for (let c = 1; c <= cfg.cols; c++) {
      // En Indicadores hay dos paneles separados (cols 2-4 y cols 6-8)
      if (cfg.twoPanels && (c === 1 || c === 5)) continue;
      const cell = ws.getRow(r).getCell(c);
      cell.border = gridAll;
    }
  }
  console.log("✓ cuadrícula:", name);
}

// Para Inquilinos, Cobros, Movimientos, etc. también quiero que la fila de header
// tenga el mismo borde inferior suave (no el medium que tenía antes)
for (const [name, cfg] of Object.entries(PER_SHEET)) {
  if (cfg.skipBorders || cfg.twoPanels) continue;
  const ws = wb.getWorksheet(name);
  if (!ws) continue;
  const hdrRow = cfg.startRow - 1;
  if (hdrRow < 1) continue;
  for (let c = 1; c <= cfg.cols; c++) {
    const cell = ws.getRow(hdrRow).getCell(c);
    cell.border = gridAll;
  }
}

wb.calcProperties = { fullCalcOnLoad: true };
const buf = await wb.xlsx.writeBuffer();
writeFileSync(XLSX, Buffer.from(buf));
console.log("\n✅ Cuadrícula aplicada");
