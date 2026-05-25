/**
 * Convierte las fórmulas `=Presupuesto!A5+` en Indicadores (cols F, filas 15-29)
 * a texto plano para que las categorías siempre se muestren sin depender de
 * la recalculación de Excel.
 *
 * Ejecutar desde web/: node scripts/fix_categorias_indicadores.mjs
 */
import ExcelJS from "exceljs";
import { readFileSync, writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const XLSX = resolve(__dirname, "../../Finanzas_Familia_2026.xlsx");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(readFileSync(XLSX));

const ws = wb.getWorksheet("Indicadores");
const pres = wb.getWorksheet("Presupuesto");

// Leer los 15 nombres de categoría desde Presupuesto!A5:A19 (texto plano)
const cats = [];
for (let i = 0; i < 15; i++) {
  const v = pres.getRow(5 + i).getCell(1).value;
  cats.push(v ?? "");
}
console.log("Categorías leídas de Presupuesto:", cats.length);

// Reemplazar col F (filas 15-29) por texto plano
for (let i = 0; i < 15; i++) {
  const r = 15 + i;
  const cell = ws.getRow(r).getCell(6);
  cell.value = cats[i];
  // Alineación a la izquierda (es texto, no número)
  cell.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
}

wb.calcProperties = { fullCalcOnLoad: true };
const buf = await wb.xlsx.writeBuffer();
writeFileSync(XLSX, Buffer.from(buf));
console.log("✅ Categorías en Indicadores convertidas a texto plano");

// Verificación
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(readFileSync(XLSX));
const ws2 = wb2.getWorksheet("Indicadores");
for (let r = 15; r <= 29; r++) {
  console.log(`R${r}: F = "${ws2.getRow(r).getCell(6).value}"`);
}
