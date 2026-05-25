/**
 * - Borra las hojas Contratos, Morosidad, Fondo Emergencia
 * - Aplica un formato uniforme y minimalista a todas las hojas restantes
 *
 * Estética: gris oscuro #374151 para títulos, gris muy clarito #F3F4F6 para headers,
 * blanco para datos, bordes finos #E5E7EB. Sin emojis ni alternancia de colores.
 *
 * Ejecutar desde web/: node scripts/unificar_formato.mjs
 */
import ExcelJS from "exceljs";
import { readFileSync, writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const XLSX = resolve(__dirname, "../../Finanzas_Familia_2026.xlsx");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(readFileSync(XLSX));

// ─── 1. Borrar las hojas que el usuario no quiere ──────────────────
for (const n of ["Contratos", "Morosidad", "Fondo Emergencia"]) {
  const s = wb.getWorksheet(n);
  if (s) {
    wb.removeWorksheet(s.id);
    console.log("✗ borrada:", n);
  }
}

// ─── 2. Paleta unificada ───────────────────────────────────────────
const COL = {
  title:    "FF374151", // gris oscuro
  header:   "FFF3F4F6", // gris muy clarito
  text:     "FF1F2937", // texto oscuro
  textWhite:"FFFFFFFF",
  noteText: "FF6B7280", // gris medio
  border:   "FFE5E7EB",
  panelLbl: "FFFAFAFA", // fondo de labels en paneles
};

const FONT = "Calibri";
const fill = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const thinBorder = { style: "thin", color: { argb: COL.border } };
const borderAll = { top: thinBorder, left: thinBorder, bottom: thinBorder, right: thinBorder };

// ─── 3. Estilos base ───────────────────────────────────────────────
function styleTitleRow(ws, rowNum, cols) {
  const row = ws.getRow(rowNum);
  row.height = 28;
  // Fusionar si hay más de 1 col
  if (cols > 1) {
    try { ws.unMergeCells(rowNum, 1, rowNum, cols); } catch {}
    ws.mergeCells(rowNum, 1, rowNum, cols);
  }
  const c = row.getCell(1);
  // Quitar emojis del título si los tiene
  if (typeof c.value === "string") {
    c.value = c.value.replace(/[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}]/gu, "").trim();
  }
  c.font = { name: FONT, bold: true, size: 13, color: { argb: COL.textWhite } };
  c.fill = fill(COL.title);
  c.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  c.border = {};
}

function styleNoteRow(ws, rowNum, cols) {
  const row = ws.getRow(rowNum);
  row.height = 18;
  if (cols > 1) {
    try { ws.unMergeCells(rowNum, 1, rowNum, cols); } catch {}
    ws.mergeCells(rowNum, 1, rowNum, cols);
  }
  const c = row.getCell(1);
  c.font = { name: FONT, italic: true, size: 9, color: { argb: COL.noteText } };
  c.fill = { type: "pattern", pattern: "none" };
  c.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
}

function styleHeaderRow(ws, rowNum, cols) {
  const row = ws.getRow(rowNum);
  row.height = 26;
  for (let c = 1; c <= cols; c++) {
    const cell = row.getCell(c);
    cell.font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
    cell.fill = fill(COL.header);
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = {
      top: thinBorder, bottom: { style: "medium", color: { argb: COL.title } },
      left: thinBorder, right: thinBorder,
    };
  }
}

function styleDataRow(ws, rowNum, cols, leftAlignCols = []) {
  const row = ws.getRow(rowNum);
  if (!row.height || row.height < 18) row.height = 22;
  for (let c = 1; c <= cols; c++) {
    const cell = row.getCell(c);
    cell.font = { name: FONT, size: 10, color: { argb: COL.text } };
    cell.fill = { type: "pattern", pattern: "none" };
    const ha = leftAlignCols.includes(c) ? "left" : "center";
    cell.alignment = { vertical: "middle", horizontal: ha, wrapText: false };
    cell.border = borderAll;
  }
}

// ─── 4. Aplicar por hoja ──────────────────────────────────────────
const TABLES = {
  // hoja → { hdr: fila de headers, cols: nº de columnas a estilar, leftCols: cuáles van alineadas a la izq }
  "Inquilinos":         { hdr: 4, cols: 15, leftCols: [2, 3, 15] },
  "Cobros":             { hdr: 4, cols: 12, leftCols: [3, 5, 10, 11, 12] },
  "Tablero Inquilinos": { hdr: 4, cols: 18, leftCols: [2] },
  "Movimientos":        { hdr: 4, cols: 11, leftCols: [4, 5, 6, 10, 11] },
  "Resumen Mensual":    { hdr: 4, cols: 9,  leftCols: [1] },
  "Presupuesto":        { hdr: 4, cols: 6,  leftCols: [1] },
};

for (const [name, conf] of Object.entries(TABLES)) {
  const ws = wb.getWorksheet(name);
  if (!ws) { console.log("? hoja no encontrada:", name); continue; }
  ws.properties.tabColor = undefined;
  styleTitleRow(ws, 1, conf.cols);
  styleNoteRow(ws, 2, conf.cols);
  styleHeaderRow(ws, conf.hdr, conf.cols);
  const lastRow = ws.rowCount;
  for (let r = conf.hdr + 1; r <= lastRow; r++) {
    styleDataRow(ws, r, conf.cols, conf.leftCols);
  }
  // Freeze panel
  ws.views = [{ state: "frozen", ySplit: conf.hdr, xSplit: 0 }];
  console.log("✓ formato aplicado:", name);
}

// ─── 5. Parametros (panel key-value) ──────────────────────────────
{
  const ws = wb.getWorksheet("Parametros");
  if (ws) {
    ws.properties.tabColor = undefined;
    // Detectar última fila con dato
    let lastRow = 4;
    for (let r = 4; r <= 30; r++) {
      if (ws.getRow(r).getCell(1).value) lastRow = r;
    }
    styleTitleRow(ws, 1, 4);
    styleNoteRow(ws, 2, 4);
    // Filas 4+: label (col A) + valor (col B)
    for (let r = 4; r <= lastRow; r++) {
      const row = ws.getRow(r);
      row.height = 22;
      // label
      const cl = row.getCell(1);
      if (cl.value) {
        cl.font = { name: FONT, bold: true, size: 10, color: { argb: COL.text } };
        cl.fill = fill(COL.panelLbl);
        cl.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
        cl.border = borderAll;
      }
      // value
      const cv = row.getCell(2);
      cv.font = { name: FONT, size: 11, color: { argb: COL.text } };
      cv.fill = { type: "pattern", pattern: "none" };
      cv.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      cv.border = borderAll;
    }
    ws.getColumn(1).width = 32;
    ws.getColumn(2).width = 20;
    console.log("✓ formato aplicado: Parametros");
  }
}

// ─── 6. Indicadores (dual table: paneles izq y der) ───────────────
{
  const ws = wb.getWorksheet("Indicadores");
  if (ws) {
    ws.properties.tabColor = undefined;
    styleTitleRow(ws, 1, 8);
    styleNoteRow(ws, 2, 8);
    // Row 4: sub-section labels (col B y col F)
    const r4 = ws.getRow(4);
    r4.height = 22;
    for (const c of [2, 6]) {
      const cell = r4.getCell(c);
      if (cell.value) {
        cell.font = { name: FONT, bold: true, size: 11, color: { argb: COL.textWhite } };
        cell.fill = fill(COL.title);
        cell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
      }
    }
    try { ws.mergeCells(4, 2, 4, 4); } catch {}
    try { ws.mergeCells(4, 6, 4, 8); } catch {}

    // Row 5: column headers
    styleHeaderRow(ws, 5, 8);
    // Pero solo cols 2-4 y 6-8 son reales, las otras vacías
    const r5 = ws.getRow(5);
    for (const c of [1, 5]) {
      r5.getCell(c).fill = { type: "pattern", pattern: "none" };
      r5.getCell(c).border = {};
    }

    // Data rows (6+)
    for (let r = 6; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      row.height = 22;
      for (const c of [2, 3, 4, 6, 7, 8]) {
        const cell = row.getCell(c);
        cell.font = { name: FONT, size: 10, color: { argb: COL.text } };
        cell.fill = { type: "pattern", pattern: "none" };
        cell.alignment = { vertical: "middle", horizontal: c === 2 || c === 4 || c === 6 ? "left" : "center", wrapText: false };
        cell.border = borderAll;
      }
    }
    ws.views = [{ state: "frozen", ySplit: 5 }];
    console.log("✓ formato aplicado: Indicadores");
  }
}

// ─── 7. Portada (cover) ────────────────────────────────────────────
{
  const ws = wb.getWorksheet("Portada");
  if (ws) {
    ws.properties.tabColor = undefined;
    // Solo limpiar colores chillones pero respetar la estructura
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= 3; c++) {
        const cell = row.getCell(c);
        const v = cell.value;
        if (v === null || v === undefined || v === "") {
          cell.fill = { type: "pattern", pattern: "none" };
          continue;
        }
        const isTitle = r === 2;
        const isSubtitle = r === 3;
        cell.font = {
          name: FONT,
          bold: isTitle,
          size: isTitle ? 18 : (isSubtitle ? 11 : 10),
          color: { argb: isSubtitle ? COL.noteText : COL.text },
        };
        cell.fill = { type: "pattern", pattern: "none" };
        cell.alignment = { vertical: "middle", horizontal: "left", indent: 1, wrapText: true };
      }
    }
    console.log("✓ formato aplicado: Portada");
  }
}

// ─── 8. Guardar ────────────────────────────────────────────────────
wb.calcProperties = { fullCalcOnLoad: true };
const buf = await wb.xlsx.writeBuffer();
writeFileSync(XLSX, Buffer.from(buf));
console.log("\n✅ Excel guardado");

// ─── 9. Verificación ──────────────────────────────────────────────
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(readFileSync(XLSX));
console.log("\nHojas finales:", wb2.worksheets.map(w => w.name).join(", "));
