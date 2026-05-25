/**
 * Agrega las hojas: Contratos, Morosidad, Fondo Emergencia
 * Ejecutar desde web/: node scripts/agregar_hojas.mjs
 */
import ExcelJS from "exceljs";
import { readFileSync, writeFileSync } from "fs";
import { resolve, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const XLSX = resolve(__dirname, "../../Finanzas_Familia_2026.xlsx");

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(readFileSync(XLSX));

// Borrar si ya existen (idempotente)
for (const n of ["Contratos", "Morosidad", "Fondo Emergencia"]) {
  const s = wb.getWorksheet(n); if (s) wb.removeWorksheet(s.id);
}

// ─── utilidades ────────────────────────────────────────────────
const bg   = (argb) => ({ type: "pattern", pattern: "solid", fgColor: { argb } });
const cen  = { horizontal: "center", vertical: "middle" };
const ft   = (bold, argb, sz = 10) => ({ bold, color: { argb }, size: sz });
const hair = { style: "hair", color: { argb: "FFDDDDDD" } };

function titleRow(ws, text, cols, fillArgb) {
  const rn = ws.rowCount + 1;
  ws.mergeCells(rn, 1, rn, cols);
  const r = ws.getRow(rn);
  r.height = 28;
  const c = r.getCell(1);
  c.value = text; c.font = ft(true, "FFFFFFFF", 13);
  c.fill = bg(fillArgb); c.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
}

function noteRow(ws, text, cols) {
  const rn = ws.rowCount + 1;
  ws.mergeCells(rn, 1, rn, cols);
  const r = ws.getRow(rn); r.height = 16;
  const c = r.getCell(1);
  c.value = text; c.font = { italic: true, color: { argb: "FF777777" }, size: 9 };
}

function headersRow(ws, headers, fillArgb) {
  const r = ws.getRow(ws.rowCount + 1); r.height = 32;
  headers.forEach((h, i) => {
    const c = r.getCell(i + 1);
    c.value = h; c.font = ft(true, "FFFFFFFF");
    c.fill = bg(fillArgb); c.alignment = { ...cen, wrapText: true };
  });
  return ws.rowCount; // devuelve número de fila del header
}

function styleDataRow(row, ncols, altRow) {
  const color = altRow ? "FFF4F7FA" : "FFFFFFFF";
  for (let c = 1; c <= ncols; c++) {
    row.getCell(c).fill = bg(color);
    if (!row.getCell(c).alignment) row.getCell(c).alignment = { ...cen };
    row.getCell(c).border = { bottom: hair };
  }
  row.height = 20;
}

// ════════════════════════════════════════════════════════════════
// HOJA 1 — CONTRATOS Y VENCIMIENTOS
// ════════════════════════════════════════════════════════════════
const wC = wb.addWorksheet("Contratos", { properties: { tabColor: { argb: "FF1B2A4A" } } });
wC.columns = [
  { width: 5 }, { width: 22 }, { width: 18 }, { width: 15 }, { width: 8 },
  { width: 13 }, { width: 13 }, { width: 11 }, { width: 17 }, { width: 14 },
  { width: 15 }, { width: 12 }, { width: 14 }, { width: 13 }, { width: 13 }, { width: 32 },
];

titleRow(wC, "Contratos y Vencimientos", 16, "FF1B2A4A");
noteRow(wC, "Completá las columnas en amarillo. Las demás se calculan automáticamente.", 16);
const HC = headersRow(wC, [
  "ID", "Inquilino", "Local", "Alquiler\nmensual", "Mon.",
  "Inicio", "Fin contrato", "Días p/\nvencer", "Estado", "Ocupación",
  "Depósito\ngarantía", "Dep.\ndevuelto", "Cláusula\najuste",
  "Últ. ajuste", "Próx. ajuste", "Notas"
], "FF1B2A4A");

const INPUT_BG = bg("FFFFFDE7"); // amarillo suave para celdas de entrada

for (let i = 1; i <= 29; i++) {
  const r = HC + i;
  const row = wC.getRow(r);

  row.getCell(1).value = i;
  row.getCell(2).value = { formula: `IFERROR(VLOOKUP(A${r},Inquilinos!$A:$B,2,0),"")`, result: "" };
  row.getCell(3).value = { formula: `IFERROR(VLOOKUP(A${r},Inquilinos!$A:$C,3,0),"")`, result: "" };
  row.getCell(4).value = 0; row.getCell(4).numFmt = "#,##0";
  row.getCell(5).value = "ARS";
  // cols 6 y 7: fecha inicio / fin → entrada manual
  row.getCell(6).numFmt = "dd/mm/yyyy";
  row.getCell(7).numFmt = "dd/mm/yyyy";
  row.getCell(8).value = { formula: `IF(G${r}="","",G${r}-TODAY())`, result: "" };
  row.getCell(8).numFmt = "0";
  row.getCell(9).value = {
    formula: `IF(J${r}="Desocupado","🏚 Desocupado",IF(G${r}="","Sin fecha",IF(H${r}<0,"🔴 Vencido",IF(H${r}<=60,"⚠ Por vencer","✅ Vigente"))))`,
    result: "Sin fecha"
  };
  // Col 10: Ocupación — dropdown Ocupado / Desocupado
  row.getCell(10).value = "Ocupado";
  row.getCell(10).dataValidation = {
    type: "list", allowBlank: false, formulae: ['"Ocupado,Desocupado"'],
    showErrorMessage: true, errorTitle: "Valor inválido",
    error: "Elegí Ocupado o Desocupado de la lista."
  };
  row.getCell(11).value = 0; row.getCell(11).numFmt = "#,##0"; // Depósito
  row.getCell(12).value = "No";   // Dep. devuelto
  row.getCell(13).value = "IPC";  // Cláusula ajuste
  row.getCell(14).numFmt = "dd/mm/yyyy"; // Último ajuste
  row.getCell(15).numFmt = "dd/mm/yyyy"; // Próximo ajuste

  styleDataRow(row, 16, i % 2 === 0);
  // Celdas de entrada en amarillo (incluyendo Ocupación col 10)
  for (const c of [4, 5, 6, 7, 10, 11, 12, 13, 14, 15, 16]) {
    row.getCell(c).fill = i % 2 === 0 ? bg("FFFFFAD5") : bg("FFFFFDE7");
  }
  row.getCell(2).alignment = { vertical: "middle", horizontal: "left" };
  row.getCell(3).alignment = { vertical: "middle", horizontal: "left" };
  row.getCell(16).alignment = { vertical: "middle", horizontal: "left" };
  row.commit();
}

wC.views = [{ state: "frozen", ySplit: HC, xSplit: 3, activeCell: "D" + (HC + 1) }];

// ════════════════════════════════════════════════════════════════
// HOJA 2 — MOROSIDAD
// ════════════════════════════════════════════════════════════════
const wM = wb.addWorksheet("Morosidad", { properties: { tabColor: { argb: "FFC0392B" } } });
wM.columns = [
  { width: 5 }, { width: 22 }, { width: 18 }, { width: 16 },
  { width: 15 }, { width: 14 }, { width: 18 }, { width: 22 }, { width: 35 },
];

titleRow(wM, "Morosidad — Estado de pagos por inquilino", 9, "FFC0392B");
noteRow(wM, "Último pago, días y cobros se calculan desde Cobros. La columna Notas es de llenado manual.", 9);
const HM = headersRow(wM, [
  "ID", "Inquilino", "Local", "Alquiler\ncontratado",
  "Último\npago", "Días sin\npagar", "Cobrado\neste año",
  "Estado", "Notas / Gestión"
], "FFC0392B");

for (let i = 1; i <= 29; i++) {
  const r = HM + i;
  const row = wM.getRow(r);

  row.getCell(1).value = i;
  row.getCell(2).value = { formula: `IFERROR(VLOOKUP(A${r},Inquilinos!$A:$B,2,0),"")`, result: "" };
  row.getCell(3).value = { formula: `IFERROR(VLOOKUP(A${r},Inquilinos!$A:$C,3,0),"")`, result: "" };
  // Alquiler desde Contratos (col D = alquiler mensual)
  row.getCell(4).value = { formula: `IFERROR(VLOOKUP(A${r},Contratos!$A:$D,4,0),0)`, result: 0 };
  row.getCell(4).numFmt = "#,##0";
  // Último pago: MAXIFS — devuelve 0 si no hay cobros (muestra vacío con ;;)
  row.getCell(5).value = {
    formula: `MAXIFS(Cobros!$A$5:$A$500,Cobros!$B$5:$B$500,A${r})`,
    result: 0
  };
  row.getCell(5).numFmt = "dd/mm/yyyy;;"; // 0 → vacío
  // Días sin pagar
  row.getCell(6).value = {
    formula: `IF(E${r}=0,"—",TEXT(TODAY()-E${r},"0"))`,
    result: "—"
  };
  // Cobrado este año (SUMIFS para filtrar también por año)
  row.getCell(7).value = {
    formula: `SUMIFS(Cobros!$F$5:$F$500,Cobros!$B$5:$B$500,A${r},Cobros!$A$5:$A$500,">="&DATE(YEAR(TODAY()),1,1))`,
    result: 0
  };
  row.getCell(7).numFmt = "#,##0";
  // Estado
  row.getCell(8).value = {
    formula: `IF(IFERROR(VLOOKUP(A${r},Contratos!$A:$J,10,0),"")="Desocupado","🏚 Desocupado",IF(E${r}=0,"❓ Sin pagos registrados",IF(TODAY()-E${r}>90,"🔴 Morosidad grave",IF(TODAY()-E${r}>60,"🟠 Moroso",IF(TODAY()-E${r}>30,"⚡ En riesgo","✅ Al día")))))`,
    result: "❓ Sin pagos registrados"
  };
  row.getCell(9).value = "";

  styleDataRow(row, 9, i % 2 === 0);
  row.getCell(2).alignment = { vertical: "middle", horizontal: "left" };
  row.getCell(3).alignment = { vertical: "middle", horizontal: "left" };
  row.getCell(9).alignment = { vertical: "middle", horizontal: "left" };
  row.commit();
}

wM.views = [{ state: "frozen", ySplit: HM, xSplit: 3, activeCell: "D" + (HM + 1) }];

// ════════════════════════════════════════════════════════════════
// HOJA 3 — FONDO DE EMERGENCIA
// ════════════════════════════════════════════════════════════════
const wF = wb.addWorksheet("Fondo Emergencia", { properties: { tabColor: { argb: "FF145A32" } } });
wF.columns = [
  { width: 32 }, { width: 20 }, { width: 20 }, { width: 26 }, { width: 28 },
];

titleRow(wF, "Fondo de Emergencia", 5, "FF145A32");

// Panel de resumen (filas 2-8)
const panelRows = [
  {
    label: "🎯  Objetivo del fondo (ARS)",
    val: 0, fmt: "#,##0",
    note: "← Escribí acá el monto que querés acumular",
    isInput: true,
  },
  {
    label: "💰  Saldo actual",
    val: { formula: `SUMIF(D13:D1000,"Depósito",B13:B1000)-SUMIF(D13:D1000,"Retiro",B13:B1000)`, result: 0 },
    fmt: "#,##0", note: "", isInput: false,
  },
  {
    label: "📊  % alcanzado",
    val: { formula: `IF(B2=0,0,B3/B2)`, result: 0 },
    fmt: "0.0%", note: "", isInput: false,
  },
  {
    label: "📅  Gasto mensual promedio (del año)",
    val: { formula: `IFERROR(SUMIFS(Movimientos!$G$5:$G$500,Movimientos!$C$5:$C$500,"Egreso")/MAX(1,MONTH(TODAY())),0)`, result: 0 },
    fmt: "#,##0", note: "← Calculado de Movimientos automáticamente", isInput: false,
  },
  {
    label: "🛡️  Meses de cobertura",
    val: { formula: `IF(OR(B2=0,B5=0),"—",TEXT(B3/B5,"0.0")&" meses")`, result: "—" },
    fmt: "@", note: "← Cuántos meses del gasto mensual cubre el fondo", isInput: false,
  },
  {
    label: "💡  Contribución mensual para llegar al objetivo",
    val: { formula: `IF(B2=0,0,MAX(0,B2-B3)/12)`, result: 0 },
    fmt: "#,##0", note: "← Cuánto depositar por mes para alcanzarlo en 12 meses", isInput: false,
  },
];

panelRows.forEach(({ label, val, fmt, note, isInput }, i) => {
  const rn = i + 2;
  const row = wF.getRow(rn);
  row.height = 24;
  const labelBg  = isInput ? "FFFFFDE7" : "FFF0FFF4";
  const valBg    = isInput ? "FFFFFDE7" : "FFE8F8F0";

  row.getCell(1).value = label;
  row.getCell(1).font = ft(true, "FF1A1A1A");
  row.getCell(1).fill = bg(labelBg);
  row.getCell(1).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

  row.getCell(2).value = val;
  row.getCell(2).font = { bold: true, size: 13, color: { argb: isInput ? "FF1B4F72" : "FF145A32" } };
  row.getCell(2).numFmt = fmt;
  row.getCell(2).fill = bg(valBg);
  row.getCell(2).alignment = cen;

  wF.mergeCells(rn, 3, rn, 5);
  row.getCell(3).value = note;
  row.getCell(3).font = { italic: true, color: { argb: "FF888888" }, size: 9 };
  row.getCell(3).fill = bg(labelBg);
  row.getCell(3).alignment = { vertical: "middle" };
});

// Separador
wF.getRow(8).height = 10;
wF.getRow(9).height = 10;

// Título de la tabla de movimientos
wF.mergeCells("A10:E10");
const logT = wF.getRow(10); logT.height = 24;
logT.getCell(1).value = "Movimientos del fondo";
logT.getCell(1).font = ft(true, "FFFFFFFF");
logT.getCell(1).fill = bg("FF1D6A54");
logT.getCell(1).alignment = { vertical: "middle", horizontal: "left", indent: 1 };

// Nota de uso
wF.mergeCells("A11:E11");
const logN = wF.getRow(11); logN.height = 16;
logN.getCell(1).value = "💡  Tipo: escribí \"Depósito\" para agregar fondos al fondo, \"Retiro\" cuando lo uses.";
logN.getCell(1).font = { italic: true, color: { argb: "FF555555" }, size: 9 };

// Cabecera de la tabla (fila 12)
const FH = 12;
const fhRow = wF.getRow(FH); fhRow.height = 26;
["Fecha", "Monto", "Saldo acumulado", "Tipo (Depósito / Retiro)", "Descripción"].forEach((h, i) => {
  const c = fhRow.getCell(i + 1);
  c.value = h; c.font = ft(true, "FFFFFFFF");
  c.fill = bg("FF1D6A54"); c.alignment = { ...cen, wrapText: true };
});

// 50 filas de movimientos (fila 13 en adelante)
for (let i = 0; i < 50; i++) {
  const r = FH + 1 + i;
  const row = wF.getRow(r);
  const altBg = i % 2 === 0 ? "FFFFFFFF" : "FFF2FBF5";

  // Saldo acumulado
  if (i === 0) {
    row.getCell(3).value = {
      formula: `IF(B${r}="","",IF(D${r}="Depósito",ABS(B${r}),-ABS(B${r})))`,
      result: ""
    };
  } else {
    const prev = r - 1;
    row.getCell(3).value = {
      formula: `IF(B${r}="","",IF(C${prev}="",IF(D${r}="Depósito",ABS(B${r}),-ABS(B${r})),C${prev}+IF(D${r}="Depósito",ABS(B${r}),-ABS(B${r}))))`,
      result: ""
    };
  }

  row.getCell(1).numFmt = "dd/mm/yyyy";
  row.getCell(2).numFmt = "#,##0";
  row.getCell(3).numFmt = "#,##0";

  for (let c = 1; c <= 5; c++) {
    row.getCell(c).fill = bg(altBg);
    row.getCell(c).alignment = c === 5 ? { vertical: "middle", horizontal: "left" } : cen;
    row.getCell(c).border = { bottom: hair };
  }
  row.height = 18;
  row.commit();
}

// ─── guardar ───────────────────────────────────────────────────
wb.calcProperties = { fullCalcOnLoad: true };
const buf = await wb.xlsx.writeBuffer();
writeFileSync(XLSX, Buffer.from(buf));
console.log("✅ Hojas creadas: Contratos, Morosidad, Fondo Emergencia");

// ─── verificación ──────────────────────────────────────────────
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(readFileSync(XLSX));
const hojas = wb2.worksheets.map(ws => ws.name);
console.log("Hojas en el archivo:", hojas.join(", "));

for (const nombre of ["Contratos", "Morosidad", "Fondo Emergencia"]) {
  const ws = wb2.getWorksheet(nombre);
  const filas = ws.rowCount;
  const titulo = ws.getRow(1).getCell(1).value;
  console.log(`  ${nombre}: ${filas} filas — título: "${titulo}"`);
}
