/**
 * - Borra hoja "Tablero Inquilinos"
 * - Limpia todos los datos de "Movimientos" (preservando headers y formato)
 * - Reconstruye "Presupuesto" limpio (preservando fórmulas y categorías de Parametros)
 *
 * Ejecutar desde web/: node scripts/tareas_varias.mjs
 */
import ExcelJS from "exceljs";
import { readFileSync, writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const XLSX = resolve(__dirname, "../../Finanzas_Familia_2026.xlsx");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(readFileSync(XLSX));

// ─── Paleta (igual al script de formato anterior) ─────────────────
const COL = {
  title:    "FF374151",
  header:   "FFF3F4F6",
  text:     "FF1F2937",
  textWhite:"FFFFFFFF",
  noteText: "FF6B7280",
  grid:     "FFD1D5DB",
};
const FONT = "Calibri";
const fill = (a) => ({ type: "pattern", pattern: "solid", fgColor: { argb: a } });
const thin = { style: "thin", color: { argb: COL.grid } };
const borderAll = { top: thin, left: thin, bottom: thin, right: thin };

// ─── 1. Borrar Tablero Inquilinos ──────────────────────────────────
{
  const ws = wb.getWorksheet("Tablero Inquilinos");
  if (ws) {
    wb.removeWorksheet(ws.id);
    console.log("✗ borrada: Tablero Inquilinos");
  }
}

// ─── 2. Limpiar datos de Movimientos (fila 5+) ─────────────────────
{
  const ws = wb.getWorksheet("Movimientos");
  if (ws) {
    const lastRow = ws.rowCount;
    let cleared = 0;
    for (let r = 5; r <= lastRow; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= 11; c++) {
        row.getCell(c).value = null;
      }
      cleared++;
    }
    console.log(`✓ Movimientos limpiada: ${cleared} filas borradas (estructura preservada)`);
  }
}

// ─── 3. Reconstruir Presupuesto ────────────────────────────────────
{
  const old = wb.getWorksheet("Presupuesto");
  if (old) wb.removeWorksheet(old.id);

  const ws = wb.addWorksheet("Presupuesto");
  ws.columns = [
    { width: 32 },  // A: Categoría
    { width: 16 },  // B: Presup. mensual
    { width: 16 },  // C: Presup. anual
    { width: 18 },  // D: Gasto real (año)
    { width: 14 },  // E: Desvío
    { width: 12 },  // F: % usado
  ];

  // Título (fila 1)
  ws.mergeCells("A1:F1");
  const t = ws.getRow(1); t.height = 28;
  t.getCell(1).value = "PRESUPUESTO ANUAL POR CATEGORÍA (ARS)";
  t.getCell(1).font = { name: FONT, bold: true, size: 13, color: { argb: COL.textWhite } };
  t.getCell(1).fill = fill(COL.title);
  t.getCell(1).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

  // Nota (fila 2)
  ws.mergeCells("A2:F2");
  const n = ws.getRow(2); n.height = 18;
  n.getCell(1).value = "Cargá el presupuesto mensual por categoría. El gasto real se calcula desde Movimientos.";
  n.getCell(1).font = { name: FONT, italic: true, size: 9, color: { argb: COL.noteText } };
  n.getCell(1).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

  // Headers (fila 4)
  ws.getRow(3).height = 8;
  const headers = ["Categoría", "Presup. mensual", "Presup. anual", "Gasto real (año)", "Desvío", "% usado"];
  const hRow = ws.getRow(4); hRow.height = 26;
  headers.forEach((h, i) => {
    const c = hRow.getCell(i + 1);
    c.value = h;
    c.font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
    c.fill = fill(COL.header);
    c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    c.border = borderAll;
  });

  // 15 filas de categorías (5 a 19) — leen desde Parametros!I12:I26
  for (let i = 0; i < 15; i++) {
    const r = 5 + i;
    const row = ws.getRow(r); row.height = 22;

    // A: Categoría (jala de Parametros)
    row.getCell(1).value = { formula: `Parametros!I${12 + i}`, result: "" };
    // B: Presup. mensual — manual (vacío)
    row.getCell(2).value = null;
    // C: Presup. anual = B*12
    row.getCell(3).value = { formula: `IF(B${r}="","",B${r}*12)`, result: "" };
    // D: Gasto real = SUMIFS de Movimientos
    row.getCell(4).value = {
      formula: `SUMIFS(Movimientos!$I:$I,Movimientos!$C:$C,"Egreso",Movimientos!$E:$E,$A${r})`,
      result: 0
    };
    // E: Desvío = C - D
    row.getCell(5).value = { formula: `IF(C${r}="","",C${r}-D${r})`, result: "" };
    // F: % usado = D/C
    row.getCell(6).value = { formula: `IF(OR(C${r}="",C${r}=0),"",D${r}/C${r})`, result: "" };

    row.getCell(2).numFmt = "#,##0";
    row.getCell(3).numFmt = "#,##0";
    row.getCell(4).numFmt = "#,##0";
    row.getCell(5).numFmt = "#,##0";
    row.getCell(6).numFmt = "0.0%";

    // Estilo común
    for (let c = 1; c <= 6; c++) {
      const cell = row.getCell(c);
      cell.font = { name: FONT, size: 10, color: { argb: COL.text } };
      cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center" };
      cell.border = borderAll;
    }
  }

  // Fila TOTAL (20)
  const r = 20;
  const totalRow = ws.getRow(r); totalRow.height = 26;
  totalRow.getCell(1).value = "TOTAL";
  totalRow.getCell(2).value = { formula: `SUM(B5:B19)`, result: 0 };
  totalRow.getCell(3).value = { formula: `SUM(C5:C19)`, result: 0 };
  totalRow.getCell(4).value = { formula: `SUM(D5:D19)`, result: 0 };
  totalRow.getCell(5).value = { formula: `SUM(E5:E19)`, result: 0 };
  totalRow.getCell(6).value = { formula: `IF(OR(C${r}="",C${r}=0),"",D${r}/C${r})`, result: "" };

  for (let c = 1; c <= 6; c++) {
    const cell = totalRow.getCell(c);
    cell.font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
    cell.fill = fill(COL.header);
    cell.alignment = { vertical: "middle", horizontal: c === 1 ? "left" : "center" };
    cell.border = { top: { style: "medium", color: { argb: COL.title } }, left: thin, right: thin, bottom: thin };
  }
  totalRow.getCell(2).numFmt = "#,##0";
  totalRow.getCell(3).numFmt = "#,##0";
  totalRow.getCell(4).numFmt = "#,##0";
  totalRow.getCell(5).numFmt = "#,##0";
  totalRow.getCell(6).numFmt = "0.0%";

  ws.views = [{ state: "frozen", ySplit: 4 }];
  console.log("✓ Presupuesto reconstruido");
}

// ─── Guardar ───────────────────────────────────────────────────────
wb.calcProperties = { fullCalcOnLoad: true };
const buf = await wb.xlsx.writeBuffer();
writeFileSync(XLSX, Buffer.from(buf));
console.log("\n✅ Guardado");

// ─── Verificación ──────────────────────────────────────────────────
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(readFileSync(XLSX));
console.log("Hojas finales:", wb2.worksheets.map(w => w.name).join(", "));
const wsM = wb2.getWorksheet("Movimientos");
let movDatos = 0;
for (let r = 5; r <= wsM.rowCount; r++) {
  if (wsM.getRow(r).getCell(7).value) movDatos++;
}
console.log(`Movimientos: filas con datos = ${movDatos} (debería ser 0)`);
