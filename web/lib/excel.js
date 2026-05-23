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
  const c1 = row.getCell(1); c1.value = parseFecha(a.fecha); c1.numFmt = "dd/mm/yyyy";
  row.getCell(2).value = resolverId(wb, a.id_inquilino, a.inquilino) || null;
  row.getCell(5).value = a.periodo || null;
  row.getCell(6).value = a.monto != null ? Number(a.monto) : null;
  row.getCell(7).value = (a.moneda || "ARS").toUpperCase();
  row.getCell(10).value = a.medio_pago || null;
  row.getCell(11).value = a.estado || "Cobrado";
  row.getCell(12).value = a.notas || null;
  row.commit && row.commit();
  return { hoja: "Cobros", fila: r, id_inquilino: row.getCell(2).value, monto: a.monto, moneda: a.moneda || "ARS" };
}

export function agregarMovimiento(wb, a) {
  const ws = wb.getWorksheet("Movimientos");
  const r = nextRow(ws, 1, HDR + 1, HDR + 400);
  const row = ws.getRow(r);
  const c1 = row.getCell(1); c1.value = parseFecha(a.fecha); c1.numFmt = "dd/mm/yyyy";
  row.getCell(3).value = a.tipo || "Egreso";
  row.getCell(4).value = a.miembro || "Familia";
  row.getCell(5).value = a.categoria || null;
  row.getCell(6).value = a.descripcion || null;
  row.getCell(7).value = a.monto != null ? Number(a.monto) : null;
  row.getCell(8).value = (a.moneda || "ARS").toUpperCase();
  row.getCell(10).value = a.medio_pago || null;
  row.getCell(11).value = a.notas || null;
  row.commit && row.commit();
  return { hoja: "Movimientos", fila: r, tipo: a.tipo || "Egreso", monto: a.monto, categoria: a.categoria };
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
