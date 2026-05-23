import ExcelJS from "exceljs";

const HDR = 4;

export async function loadWB(buf) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}

export async function wbToBuffer(wb) {
  if (wb.calcProperties) wb.calcProperties.fullCalcOnLoad = true;
  const ab = await wb.xlsx.writeBuffer();
  return Buffer.from(ab);
}

function nextRow(ws, keyCol, start, end) {
  let r = start;
  while (r <= end) {
    const v = ws.getRow(r).getCell(keyCol).value;
    if (v === null || v === undefined || v === "") break;
    r++;
  }
  return r;
}

function parseFecha(s) {
  if (!s) return new Date();
  const d = new Date(String(s).length <= 10 ? String(s) + "T00:00:00" : String(s));
  return isNaN(d.getTime()) ? new Date() : d;
}

// Tipo de cambio USD→ARS (Parametros!B4, nombre definido "TC").
function getTC(wb) {
  try {
    const par = wb.getWorksheet("Parametros");
    return Number(par.getCell("B4").value) || 1;
  } catch (e) { return 1; }
}
const esUSD = (moneda) => String(moneda || "").toUpperCase() === "USD";
const aARS = (monto, moneda, tc) => Math.round(Number(monto) * (esUSD(moneda) ? tc : 1));
const aUSD = (monto, moneda, tc) => Math.round(esUSD(moneda) ? Number(monto) : Number(monto) / tc);

export function listarInquilinos(wb) {
  const ws = wb.getWorksheet("Inquilinos");
  const out = [];
  for (let r = HDR + 1; r <= HDR + 29; r++) {
    const row = ws.getRow(r);
    const id = row.getCell(1).value;
    if (id) out.push({ id, nombre: row.getCell(2).value, local: row.getCell(3).value });
  }
  return out;
}

function resolverId(wb, id, texto) {
  if (id) return id;
  const t = (texto || "").toString().trim().toLowerCase();
  if (!t) return null;
  for (const i of listarInquilinos(wb)) {
    const n = (i.nombre || "").toString().toLowerCase();
    const l = (i.local || "").toString().toLowerCase();
    if (t === String(i.id) || (n && (t.includes(n) || n.includes(t))) || (l && (t.includes(l) || l.includes(t)))) return i.id;
  }
  return null;
}

export function agregarCobro(wb, a) {
  const ws = wb.getWorksheet("Cobros");
  const r = nextRow(ws, 1, HDR + 1, HDR + 400);
  const row = ws.getRow(r);
  const tc = getTC(wb);
  const f = parseFecha(a.fecha);
  const monto = a.monto != null ? Number(a.monto) : null;
  const moneda = (a.moneda || "ARS").toUpperCase();
  const idInq = resolverId(wb, a.id_inquilino, a.inquilino) || null;
  const c1 = row.getCell(1); c1.value = f; c1.numFmt = "dd/mm/yyyy";
  row.getCell(2).value = idInq;
  // Inquilino (col 3): VLOOKUP por ID; cacheamos el nombre resuelto.
  const nombreInq = idInq != null ? ((listarInquilinos(wb).find((i) => String(i.id) === String(idInq)) || {}).nombre || null) : null;
  row.getCell(3).value = idInq != null
    ? { formula: `IF($B${r}="","",IFERROR(VLOOKUP($B${r},Inquilinos!$A:$B,2,0),"ID inexistente"))`, result: nombreInq || "ID inexistente" }
    : null;
  // Mes (col 4): =MONTH(fecha) con valor cacheado.
  row.getCell(4).value = { formula: `IF($A${r}="","",MONTH($A${r}))`, result: f.getMonth() + 1 };
  row.getCell(5).value = a.periodo || null;
  row.getCell(6).value = monto;
  row.getCell(7).value = moneda;
  // Monto (ARS) col 8 y Monto (USD) col 9: conversión por TC, fórmula + valor cacheado.
  row.getCell(8).value = monto != null
    ? { formula: `IF($F${r}="","",IF($G${r}="USD",$F${r}*TC,$F${r}))`, result: aARS(monto, moneda, tc) }
    : null;
  row.getCell(9).value = monto != null
    ? { formula: `IF($F${r}="","",IF($G${r}="USD",$F${r},$F${r}/TC))`, result: aUSD(monto, moneda, tc) }
    : null;
  row.getCell(10).value = a.medio_pago || null;
  row.getCell(11).value = a.estado || "Cobrado";
  row.getCell(12).value = a.notas || null;
  row.commit && row.commit();
  return { hoja: "Cobros", fila: r, id_inquilino: idInq, monto, moneda };
}

export function agregarMovimiento(wb, a) {
  const ws = wb.getWorksheet("Movimientos");
  const r = nextRow(ws, 1, HDR + 1, HDR + 400);
  const row = ws.getRow(r);
  const tc = getTC(wb);
  const f = parseFecha(a.fecha);
  const monto = a.monto != null ? Number(a.monto) : null;
  const moneda = (a.moneda || "ARS").toUpperCase();
  const c1 = row.getCell(1); c1.value = f; c1.numFmt = "dd/mm/yyyy";
  // Mes (col 2): fórmula =MONTH(fecha) con el valor ya cacheado para que se vea al instante.
  row.getCell(2).value = { formula: `IF($A${r}="","",MONTH($A${r}))`, result: f.getMonth() + 1 };
  row.getCell(3).value = a.tipo || "Egreso";
  row.getCell(4).value = a.miembro || "Familia";
  row.getCell(5).value = a.categoria || null;
  row.getCell(6).value = a.descripcion || null;
  row.getCell(7).value = monto;
  row.getCell(8).value = moneda;
  // Monto (ARS) (col 9): convierte USD→ARS con TC; fórmula + valor cacheado.
  row.getCell(9).value = monto != null
    ? { formula: `IF($G${r}="","",IF($H${r}="USD",$G${r}*TC,$G${r}))`, result: aARS(monto, moneda, tc) }
    : null;
  row.getCell(10).value = a.medio_pago || null;
  row.getCell(11).value = a.notas || null;
  row.commit && row.commit();
  return { hoja: "Movimientos", fila: r, tipo: a.tipo || "Egreso", monto, categoria: a.categoria, medio_pago: a.medio_pago || null };
}

export function actualizarIpc(wb, a) {
  const ws = wb.getWorksheet("Inflacion INDEC");
  const mm = (a.mes || "").slice(0, 7);
  for (let r = HDR + 1; r <= HDR + 60; r++) {
    const b = ws.getRow(r).getCell(2).value;
    if (b instanceof Date) {
      const key = b.getFullYear() + "-" + String(b.getMonth() + 1).padStart(2, "0");
      if (key === mm) {
        ws.getRow(r).getCell(3).value = Number(a.indice);
        return { hoja: "Inflacion INDEC", fila: r, mes: mm, indice: a.indice };
      }
    }
  }
  throw new Error("El mes " + mm + " no está en la tabla de Inflacion INDEC (2022-2026).");
}

export function resumenMensual(wb, mes) {
  const par = wb.getWorksheet("Parametros");
  const tc = Number(par.getCell("B4").value) || 1;
  const m = Number(mes) || 0;
  const toARS = (monto, moneda) => Number(monto) * (String(moneda).toUpperCase() === "USD" ? tc : 1);
  let alq = 0;
  const cob = wb.getWorksheet("Cobros");
  for (let r = HDR + 1; r <= HDR + 400; r++) {
    const row = cob.getRow(r);
    const f = row.getCell(1).value, monto = row.getCell(6).value, mon = row.getCell(7).value;
    if (monto && f instanceof Date && (m === 0 || f.getMonth() + 1 === m)) alq += toARS(monto, mon);
  }
  let ing = 0, egr = 0;
  const mov = wb.getWorksheet("Movimientos");
  for (let r = HDR + 1; r <= HDR + 400; r++) {
    const row = mov.getRow(r);
    const f = row.getCell(1).value, tipo = row.getCell(3).value, monto = row.getCell(7).value, mon = row.getCell(8).value;
    if (monto && f instanceof Date && (m === 0 || f.getMonth() + 1 === m)) {
      const ars = toARS(monto, mon);
      if (String(tipo).toLowerCase() === "ingreso") ing += ars; else egr += ars;
    }
  }
  const totalIng = alq + ing;
  return {
    mes: m || "año completo",
    ingresos_alquileres: Math.round(alq), otros_ingresos: Math.round(ing),
    total_ingresos: Math.round(totalIng), egresos: Math.round(egr),
    resultado_neto: Math.round(totalIng - egr), moneda: "ARS",
  };
}
