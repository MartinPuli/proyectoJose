/**
 * - Borra el bloque mensual grandote de Parametros (filas 28-42)
 * - Reemplaza la sección "Egresos por miembro (año)" en Indicadores por una compacta:
 *     Miembro | Mes actual | Año
 *   con los nombres reales: José, Clara, Laura, Familia
 *
 * Ejecutar desde web/: node scripts/renombrar_miembros.mjs
 */
import ExcelJS from "exceljs";
import { readFileSync, writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const XLSX = resolve(__dirname, "../../Finanzas_Familia_2026.xlsx");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(readFileSync(XLSX));

// ─── Paleta ────────────────────────────────────────────────────────
const COL = {
  title:  "FF374151",
  header: "FFF3F4F6",
  text:   "FF1F2937",
  textW:  "FFFFFFFF",
  grid:   "FFD1D5DB",
};
const FONT = "Calibri";
const fill = (a) => ({ type: "pattern", pattern: "solid", fgColor: { argb: a } });
const thin = { style: "thin", color: { argb: COL.grid } };
const borderAll = { top: thin, left: thin, bottom: thin, right: thin };

// ─── 1. Parametros: borrar bloque mensual grandote ─────────────────
{
  const ws = wb.getWorksheet("Parametros");
  for (let r = 28; r <= 50; r++) {
    const row = ws.getRow(r);
    for (let c = 1; c <= 6; c++) {
      row.getCell(c).value = null;
      row.getCell(c).fill = { type: "pattern", pattern: "none" };
      row.getCell(c).border = {};
      row.getCell(c).font = undefined;
    }
    try { ws.unMergeCells(r, 1, r, 6); } catch {}
  }
  console.log("✓ Parametros: bloque mensual viejo borrado");
}

// ─── 2. Indicadores: rehacer sección Egresos por miembro ───────────
{
  const ws = wb.getWorksheet("Indicadores");

  // Limpiar las celdas F-H (cols 6-8) de filas 4-10 (rango actual de la sección)
  for (let r = 4; r <= 10; r++) {
    const row = ws.getRow(r);
    for (let c = 6; c <= 8; c++) {
      row.getCell(c).value = null;
      row.getCell(c).fill = { type: "pattern", pattern: "none" };
      row.getCell(c).border = {};
      row.getCell(c).font = undefined;
    }
    try { ws.unMergeCells(r, 6, r, 8); } catch {}
  }

  // Asegurar anchos de cols 6-8
  if (!ws.getColumn(6).width || ws.getColumn(6).width < 14) ws.getColumn(6).width = 14;
  if (!ws.getColumn(7).width || ws.getColumn(7).width < 14) ws.getColumn(7).width = 14;
  if (!ws.getColumn(8).width || ws.getColumn(8).width < 14) ws.getColumn(8).width = 14;

  // Título de sección (R4, cols F-H merged)
  ws.mergeCells(4, 6, 4, 8);
  const t = ws.getRow(4);
  t.height = 22;
  t.getCell(6).value = "Egresos por miembro";
  t.getCell(6).font = { name: FONT, bold: true, size: 11, color: { argb: COL.textW } };
  t.getCell(6).fill = fill(COL.title);
  t.getCell(6).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

  // Headers (R5)
  const h = ws.getRow(5);
  ["Miembro", "Mes actual", "Año"].forEach((txt, i) => {
    const c = h.getCell(6 + i);
    c.value = txt;
    c.font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
    c.fill = fill(COL.header);
    c.alignment = { vertical: "middle", horizontal: "center" };
    c.border = borderAll;
  });
  h.height = 24;

  // Filas de miembros (R6-R9)
  const MIEMBROS = ["José", "Clara", "Laura", "Familia"];
  MIEMBROS.forEach((m, i) => {
    const r = 6 + i;
    const row = ws.getRow(r);
    row.height = 22;

    // F: Nombre
    row.getCell(6).value = m;
    row.getCell(6).font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
    row.getCell(6).fill = fill("FFFAFAFA");
    row.getCell(6).alignment = { vertical: "middle", horizontal: "left", indent: 1 };
    row.getCell(6).border = borderAll;

    // G: Mes actual
    row.getCell(7).value = {
      formula: `SUMIFS(Movimientos!$I:$I,Movimientos!$C:$C,"Egreso",Movimientos!$D:$D,"${m}",Movimientos!$B:$B,MONTH(TODAY()))`,
      result: 0,
    };
    row.getCell(7).numFmt = "#,##0";
    row.getCell(7).font = { name: FONT, size: 10, color: { argb: COL.text } };
    row.getCell(7).alignment = { vertical: "middle", horizontal: "center" };
    row.getCell(7).border = borderAll;

    // H: Año
    row.getCell(8).value = {
      formula: `SUMIFS(Movimientos!$I:$I,Movimientos!$C:$C,"Egreso",Movimientos!$D:$D,"${m}")`,
      result: 0,
    };
    row.getCell(8).numFmt = "#,##0";
    row.getCell(8).font = { name: FONT, size: 10, color: { argb: COL.text } };
    row.getCell(8).alignment = { vertical: "middle", horizontal: "center" };
    row.getCell(8).border = borderAll;
  });

  // Fila TOTAL (R10)
  const tr = ws.getRow(10);
  tr.height = 24;
  tr.getCell(6).value = "TOTAL";
  tr.getCell(6).font = { name: FONT, bold: true, size: 10, color: { argb: COL.textW } };
  tr.getCell(6).fill = fill(COL.title);
  tr.getCell(6).alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  tr.getCell(6).border = borderAll;

  tr.getCell(7).value = { formula: "SUM(G6:G9)", result: 0 };
  tr.getCell(7).numFmt = "#,##0";
  tr.getCell(7).font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
  tr.getCell(7).fill = fill(COL.header);
  tr.getCell(7).alignment = { vertical: "middle", horizontal: "center" };
  tr.getCell(7).border = { top: { style: "medium", color: { argb: COL.title } }, left: thin, right: thin, bottom: thin };

  tr.getCell(8).value = { formula: "SUM(H6:H9)", result: 0 };
  tr.getCell(8).numFmt = "#,##0";
  tr.getCell(8).font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
  tr.getCell(8).fill = fill(COL.header);
  tr.getCell(8).alignment = { vertical: "middle", horizontal: "center" };
  tr.getCell(8).border = { top: { style: "medium", color: { argb: COL.title } }, left: thin, right: thin, bottom: thin };

  console.log("✓ Indicadores: sección 'Egresos por miembro' reescrita con José/Clara/Laura/Familia (Mes actual + Año)");
}

// ─── Guardar ───────────────────────────────────────────────────────
wb.calcProperties = { fullCalcOnLoad: true };
const buf = await wb.xlsx.writeBuffer();
writeFileSync(XLSX, Buffer.from(buf));
console.log("\n✅ Guardado");

// ─── Verificación ──────────────────────────────────────────────────
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(readFileSync(XLSX));
const ind = wb2.getWorksheet("Indicadores");
console.log("\nIndicadores R4-R10 (cols F-H):");
for (let r = 4; r <= 10; r++) {
  const cells = [];
  for (let c = 6; c <= 8; c++) {
    const v = ind.getRow(r).getCell(c).value;
    if (v === null || v === undefined) continue;
    const s = (typeof v === "object" && v?.formula) ? "F:" + v.formula.slice(0, 50) : v;
    cells.push(`[${String.fromCharCode(64+c)}]${s}`);
  }
  if (cells.length) console.log(`R${r}: ${cells.join(" | ")}`);
}
