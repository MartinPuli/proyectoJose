/**
 * Agrega un bloque "EGRESOS POR MIEMBRO — MENSUAL" al final de Parametros.
 * Lee de Movimientos via SUMIFS, con los 4 miembros estándar (Yo, Hermana, Mamá, Familia).
 *
 * Ejecutar desde web/: node scripts/egresos_por_miembro.mjs
 */
import ExcelJS from "exceljs";
import { readFileSync, writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const XLSX = resolve(__dirname, "../../Finanzas_Familia_2026.xlsx");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(readFileSync(XLSX));

const ws = wb.getWorksheet("Parametros");
if (!ws) throw new Error("No hay Parametros");

// ─── Paleta (igual a las demás hojas) ──────────────────────────────
const COL = {
  title:    "FF374151",
  header:   "FFF3F4F6",
  text:     "FF1F2937",
  textW:    "FFFFFFFF",
  noteText: "FF6B7280",
  grid:     "FFD1D5DB",
};
const FONT = "Calibri";
const fill = (a) => ({ type: "pattern", pattern: "solid", fgColor: { argb: a } });
const thin = { style: "thin", color: { argb: COL.grid } };
const borderAll = { top: thin, left: thin, bottom: thin, right: thin };

// ─── Limpiar primero la zona donde voy a escribir (filas 28+) ─────
for (let r = 28; r <= 50; r++) {
  const row = ws.getRow(r);
  for (let c = 1; c <= 6; c++) {
    row.getCell(c).value = null;
    row.getCell(c).fill = { type: "pattern", pattern: "none" };
    row.getCell(c).border = {};
  }
  try { ws.unMergeCells(r, 1, r, 6); } catch {}
}

// ─── Layout ────────────────────────────────────────────────────────
const MIEMBROS = ["Yo", "Hermana", "Mamá", "Familia"];
const MESES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
const T_TITLE = 28;
const T_HEAD  = 29;
const T_FIRST = 30;
const T_LAST  = T_FIRST + 11; // 41
const T_TOTAL = T_LAST + 1;   // 42

// Asegurar anchos
const widths = [9, 14, 14, 14, 14, 14];
widths.forEach((w, i) => {
  const col = ws.getColumn(i + 1);
  if (!col.width || col.width < w) col.width = w;
});

// Título de sección
ws.mergeCells(T_TITLE, 1, T_TITLE, 6);
const t = ws.getRow(T_TITLE); t.height = 26;
t.getCell(1).value = "EGRESOS POR MIEMBRO — MENSUAL (ARS)";
t.getCell(1).font = { name: FONT, bold: true, size: 12, color: { argb: COL.textW } };
t.getCell(1).fill = fill(COL.title);
t.getCell(1).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

// Headers
const hRow = ws.getRow(T_HEAD); hRow.height = 24;
const headers = ["Mes", ...MIEMBROS, "Total"];
headers.forEach((h, i) => {
  const c = hRow.getCell(i + 1);
  c.value = h;
  c.font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
  c.fill = fill(COL.header);
  c.alignment = { vertical: "middle", horizontal: "center" };
  c.border = borderAll;
});

// 12 filas de meses
for (let i = 0; i < 12; i++) {
  const r = T_FIRST + i;
  const mes = i + 1; // 1..12
  const row = ws.getRow(r); row.height = 22;

  // Col 1: nombre del mes
  row.getCell(1).value = MESES[i];
  row.getCell(1).font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
  row.getCell(1).fill = fill("FFFAFAFA");
  row.getCell(1).alignment = { vertical: "middle", horizontal: "center" };
  row.getCell(1).border = borderAll;

  // Cols 2-5: SUMIFS por miembro
  MIEMBROS.forEach((m, idx) => {
    const c = idx + 2;
    const cell = row.getCell(c);
    cell.value = {
      formula: `SUMIFS(Movimientos!$I:$I,Movimientos!$C:$C,"Egreso",Movimientos!$D:$D,"${m}",Movimientos!$B:$B,${mes})`,
      result: 0,
    };
    cell.numFmt = "#,##0";
    cell.font = { name: FONT, size: 10, color: { argb: COL.text } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = borderAll;
  });

  // Col 6: total del mes
  const totalCell = row.getCell(6);
  totalCell.value = { formula: `SUM(B${r}:E${r})`, result: 0 };
  totalCell.numFmt = "#,##0";
  totalCell.font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
  totalCell.fill = fill("FFFAFAFA");
  totalCell.alignment = { vertical: "middle", horizontal: "center" };
  totalCell.border = borderAll;
}

// Fila TOTAL (año)
const tot = ws.getRow(T_TOTAL); tot.height = 24;
tot.getCell(1).value = "TOTAL";
tot.getCell(1).font = { name: FONT, bold: true, size: 11, color: { argb: COL.textW } };
tot.getCell(1).fill = fill(COL.title);
tot.getCell(1).alignment = { vertical: "middle", horizontal: "center" };
tot.getCell(1).border = borderAll;

for (let c = 2; c <= 6; c++) {
  const cell = tot.getCell(c);
  const colLetter = String.fromCharCode(64 + c);
  cell.value = { formula: `SUM(${colLetter}${T_FIRST}:${colLetter}${T_LAST})`, result: 0 };
  cell.numFmt = "#,##0";
  cell.font = { name: FONT, bold: true, size: 11, color: { argb: COL.text } };
  cell.fill = fill(COL.header);
  cell.alignment = { vertical: "middle", horizontal: "center" };
  cell.border = { top: { style: "medium", color: { argb: COL.title } }, left: thin, right: thin, bottom: thin };
}

// ─── Guardar ───────────────────────────────────────────────────────
wb.calcProperties = { fullCalcOnLoad: true };
const buf = await wb.xlsx.writeBuffer();
writeFileSync(XLSX, Buffer.from(buf));
console.log("✅ Bloque agregado en Parametros, filas 28-42");

// Verificación
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(readFileSync(XLSX));
const ws2 = wb2.getWorksheet("Parametros");
console.log("R28 título:", ws2.getRow(28).getCell(1).value);
console.log("R29 headers:", [1,2,3,4,5,6].map(c => ws2.getRow(29).getCell(c).value).join(" | "));
console.log("R30 ene:", [1,2,3,4,5,6].map(c => {
  const v = ws2.getRow(30).getCell(c).value;
  return (typeof v === "object" && v?.formula) ? "F" : v;
}).join(" | "));
console.log("R42 total:", [1,2,3,4,5,6].map(c => {
  const v = ws2.getRow(42).getCell(c).value;
  return (typeof v === "object" && v?.formula) ? "F:"+v.formula : v;
}).join(" | "));
