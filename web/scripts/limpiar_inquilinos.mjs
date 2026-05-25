/**
 * Limpia la hoja Inquilinos:
 * - Quita columnas: Rubro, IPC base, Mes últ. ajuste, Coef., Depósito, Contacto
 * - Frecuencia ajuste → "Trimestral" para todos
 * - IPC ajuste → manual (coef., default 1.0)
 * - Alquiler ajustado / (ARS) se recalculan con fórmulas simples
 *
 * Ejecutar desde web/: node scripts/limpiar_inquilinos.mjs
 */
import ExcelJS from "exceljs";
import { readFileSync, writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const XLSX = resolve(__dirname, "../../Finanzas_Familia_2026.xlsx");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(readFileSync(XLSX));
const old = wb.getWorksheet("Inquilinos");
if (!old) throw new Error("No hay hoja Inquilinos");

// ─── 1. Leer datos existentes (filas 5–33, 29 inquilinos) ──────────
const datos = [];
for (let i = 1; i <= 29; i++) {
  const r = 4 + i;
  const row = old.getRow(r);
  // Lectura: si el valor es una fórmula, tomar el `result` cacheado
  const get = (c) => {
    const v = row.getCell(c).value;
    if (v && typeof v === "object" && "formula" in v) return v.result;
    return v;
  };
  datos.push({
    id:        get(1),
    nombre:    get(2),
    local:     get(3),
    // rubro: get(4) — DESCARTADO
    inicio:    get(5),
    fin:       get(6),
    mesBase:   get(7),
    alqBase:   get(8),
    moneda:    get(9),
    // frecuencia: get(10) — IGNORADO (forzamos "Trimestral")
    // ipcBase: get(11) — DESCARTADO
    // mesUltAjuste: get(12) — DESCARTADO
    // ipcAjuste: get(13) — DESCARTADO (era fórmula, queda manual)
    // coef: get(14) — DESCARTADO
    // alqAjustado: get(15) — RECALCULADO con fórmula
    // alqAjustadoARS: get(16) — RECALCULADO con fórmula
    diaVenc:   get(17),
    // deposito: get(18) — DESCARTADO
    estado:    get(19),
    // contacto: get(20) — DESCARTADO
    notas:     get(21),
  });
}

// Eliminar la hoja vieja
const oldId = old.id;
wb.removeWorksheet(oldId);

// ─── 2. Crear hoja nueva con estructura limpia ─────────────────────
const ws = wb.addWorksheet("Inquilinos", { properties: { tabColor: { argb: "FF3498DB" } } });

ws.columns = [
  { width: 5 },   // A: ID
  { width: 22 },  // B: Inquilino
  { width: 20 },  // C: Local / Propiedad
  { width: 14 },  // D: Inicio contrato
  { width: 14 },  // E: Fin contrato
  { width: 14 },  // F: Mes base (ajuste)
  { width: 14 },  // G: Alquiler base
  { width: 8 },   // H: Moneda
  { width: 13 },  // I: Frecuencia ajuste
  { width: 11 },  // J: IPC ajuste
  { width: 16 },  // K: Alquiler ajustado
  { width: 18 },  // L: Alquiler ajustado (ARS)
  { width: 10 },  // M: Día venc.
  { width: 13 },  // N: Estado
  { width: 30 },  // O: Notas
];

// Título
ws.mergeCells("A1:O1");
const t = ws.getRow(1); t.height = 28;
t.getCell(1).value = "LISTADO DE INQUILINOS";
t.getCell(1).font = { bold: true, color: { argb: "FFFFFFFF" }, size: 13 };
t.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2874A6" } };
t.getCell(1).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

// Nota
ws.mergeCells("A2:O2");
const n = ws.getRow(2); n.height = 16;
n.getCell(1).value = "IPC ajuste: coeficiente manual (1.0 = sin ajuste; 1.45 = +45%). Alquiler ajustado se calcula solo.";
n.getCell(1).font = { italic: true, color: { argb: "FF777777" }, size: 9 };

// Headers (fila 4 para mantener HDR=4 igual que antes)
ws.getRow(3).height = 8; // separador
const HDR = 4;
const headers = [
  "ID", "Inquilino", "Local /\nPropiedad", "Inicio\ncontrato", "Fin\ncontrato",
  "Mes base\n(ajuste)", "Alquiler\nbase", "Moneda", "Frecuencia\najuste",
  "IPC ajuste\n(coef.)", "Alquiler\najustado", "Alquiler\najustado (ARS)",
  "Día\nvenc.", "Estado", "Notas",
];
const hRow = ws.getRow(HDR); hRow.height = 36;
headers.forEach((h, i) => {
  const c = hRow.getCell(i + 1);
  c.value = h;
  c.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
  c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2874A6" } };
  c.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
});

// ─── 3. Cargar los 29 inquilinos con sus datos ─────────────────────
for (let i = 0; i < datos.length; i++) {
  const d = datos[i];
  const r = HDR + 1 + i;
  const row = ws.getRow(r);
  const altBg = i % 2 === 0 ? "FFFFFFFF" : "FFF5F9FC";

  row.getCell(1).value = d.id ?? (i + 1);
  row.getCell(2).value = d.nombre ?? "";
  row.getCell(3).value = d.local ?? "";
  row.getCell(4).value = d.inicio instanceof Date ? d.inicio : null;
  row.getCell(5).value = d.fin instanceof Date ? d.fin : null;
  row.getCell(6).value = d.mesBase instanceof Date ? d.mesBase : null;
  row.getCell(7).value = typeof d.alqBase === "number" ? d.alqBase : 0;
  row.getCell(8).value = d.moneda ?? "ARS";
  row.getCell(9).value = "Trimestral"; // forzado para todos
  row.getCell(10).value = 1; // IPC ajuste — coef. manual, default 1.0
  // Alquiler ajustado = base * IPC ajuste (USD también ajusta ahora — más coherente)
  row.getCell(11).value = {
    formula: `IF(G${r}="","",G${r}*J${r})`,
    result: (typeof d.alqBase === "number" ? d.alqBase : 0) * 1
  };
  // Alquiler ajustado (ARS) — convierte USD usando rango nombrado TC (Parametros!B4)
  row.getCell(12).value = {
    formula: `IF(K${r}="","",IF(H${r}="USD",K${r}*TC,K${r}))`,
    result: ""
  };
  row.getCell(13).value = typeof d.diaVenc === "number" ? d.diaVenc : (d.diaVenc ?? null);
  row.getCell(14).value = d.estado ?? "Activo";
  row.getCell(15).value = d.notas ?? "";

  // Formatos
  row.getCell(4).numFmt = "dd/mm/yyyy";
  row.getCell(5).numFmt = "dd/mm/yyyy";
  row.getCell(6).numFmt = "dd/mm/yyyy";
  row.getCell(7).numFmt = "#,##0";
  row.getCell(10).numFmt = "0.0000";
  row.getCell(11).numFmt = "#,##0";
  row.getCell(12).numFmt = "#,##0";
  row.getCell(13).numFmt = "0";

  // Dropdown para Moneda
  row.getCell(8).dataValidation = {
    type: "list", allowBlank: false, formulae: ['"ARS,USD"'],
  };
  // Dropdown para Frecuencia
  row.getCell(9).dataValidation = {
    type: "list", allowBlank: false, formulae: ['"Mensual,Trimestral,Semestral,Anual"'],
  };
  // Dropdown para Estado
  row.getCell(14).dataValidation = {
    type: "list", allowBlank: true, formulae: ['"Activo,Pausado,Vencido"'],
  };

  // Estilos
  for (let c = 1; c <= 15; c++) {
    row.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: altBg } };
    if (!row.getCell(c).alignment) {
      row.getCell(c).alignment = { vertical: "middle", horizontal: "center" };
    }
    row.getCell(c).border = { bottom: { style: "hair", color: { argb: "FFDDDDDD" } } };
  }
  row.getCell(2).alignment = { vertical: "middle", horizontal: "left" };
  row.getCell(3).alignment = { vertical: "middle", horizontal: "left" };
  row.getCell(15).alignment = { vertical: "middle", horizontal: "left" };
  row.height = 22;
  row.commit();
}

// Mover la hoja Inquilinos a la posición 3 (antes era esa posición)
ws.orderNo = 2; // Vercel ignora esto pero ayuda en UI

// Freeze header + columnas izquierdas
ws.views = [{ state: "frozen", ySplit: HDR, xSplit: 3 }];

// ─── 4. Guardar ────────────────────────────────────────────────────
wb.calcProperties = { fullCalcOnLoad: true };
const buf = await wb.xlsx.writeBuffer();
writeFileSync(XLSX, Buffer.from(buf));
console.log("✅ Hoja Inquilinos limpiada");

// ─── 5. Verificación ───────────────────────────────────────────────
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(readFileSync(XLSX));
const ws2 = wb2.getWorksheet("Inquilinos");
console.log("Columnas nuevas:");
for (let c = 1; c <= 15; c++) {
  console.log(`  ${c}: ${ws2.getRow(HDR).getCell(c).value}`);
}
console.log(`\nFila 5 (primer inquilino):`);
const ej = ws2.getRow(5);
console.log(`  ID=${ej.getCell(1).value}, nombre=${ej.getCell(2).value}, frec=${ej.getCell(9).value}, IPC ajuste=${ej.getCell(10).value}`);

// Check que todas tengan Trimestral
let countTrim = 0;
for (let r = 5; r <= 33; r++) {
  if (ws2.getRow(r).getCell(9).value === "Trimestral") countTrim++;
}
console.log(`Inquilinos con Trimestral: ${countTrim}/29`);
