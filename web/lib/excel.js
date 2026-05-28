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

export function listarCategorias(wb) {
  const ws = wb.getWorksheet("Presupuesto");
  if (!ws) return [];
  const out = [];
  for (let r = HDR + 1; r <= HDR + 15; r++) {
    const v = ws.getRow(r).getCell(1).value;
    if (v && String(v).trim().toUpperCase() !== "TOTAL") out.push(String(v).trim());
  }
  return out;
}

// Mapea texto libre (lo que sugiere el LLM) a una de las categorías exactas del
// Presupuesto. Si nada matchea => "Otros". Sin esto el SUMIFS del Presupuesto
// devuelve 0 cuando la categoría guardada no coincide exacto (ej. "seguro auto").
function _toks(s) {
  return s.replace(/[\/()\-,.]/g, " ").split(/\s+/).filter((t) => t.length >= 3);
}
export function normalizarCategoria(wb, categoria) {
  if (!categoria) return "Otros";
  const cats = listarCategorias(wb);
  if (!cats.length) return categoria;
  const c = String(categoria).trim();
  for (const x of cats) if (c === x) return x;
  const cl = c.toLowerCase();
  for (const x of cats) if (cl === x.toLowerCase()) return x;
  for (const x of cats) { const xl = x.toLowerCase(); if (cl.includes(xl) || xl.includes(cl)) return x; }
  const inTok = _toks(cl);
  for (const tok of inTok) {
    for (const x of cats) {
      for (const xt of _toks(x.toLowerCase())) {
        if (tok === xt || tok.startsWith(xt) || xt.startsWith(tok)) return x;
      }
    }
  }
  return "Otros";
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
  const tipo = a.tipo || "Egreso";
  row.getCell(3).value = tipo;
  row.getCell(4).value = a.miembro || "Familia";
  // Normalizar a una categoría del Presupuesto solo para egresos (los ingresos no se contabilizan ahí).
  const cat = String(tipo).toLowerCase() === "egreso" ? normalizarCategoria(wb, a.categoria) : (a.categoria || null);
  row.getCell(5).value = cat;
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
  return { hoja: "Movimientos", fila: r, tipo, monto, categoria: cat, medio_pago: a.medio_pago || null };
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

const MES_NOMBRES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];

// Datos crudos para que la IA arme conclusiones de fin de mes: mes actual + 3 meses
// previos por categoría, cobros pendientes y top egresos. NO juzga, solo agrupa.
export function analisisMes(wb, mesObjetivo) {
  const par = wb.getWorksheet("Parametros");
  const tc = Number(par && par.getCell("B4").value) || 1;
  const toARS = (monto, moneda) => Number(monto) * (String(moneda || "").toUpperCase() === "USD" ? tc : 1);
  const hoy = new Date();
  const mes = Number(mesObjetivo) || (hoy.getMonth() + 1);
  const anio = hoy.getFullYear();

  const cob = wb.getWorksheet("Cobros");
  let alquileres = 0;
  const inquilinosCobrados = new Set();
  for (let r = HDR + 1; r <= HDR + 400; r++) {
    const row = cob.getRow(r);
    const f = row.getCell(1).value, id = row.getCell(2).value;
    const monto = row.getCell(6).value, mon = row.getCell(7).value;
    if (monto && f instanceof Date && f.getMonth() + 1 === mes && f.getFullYear() === anio) {
      alquileres += toARS(monto, mon);
      if (id) inquilinosCobrados.add(String(id));
    }
  }

  const mov = wb.getWorksheet("Movimientos");
  const egresosMes = {};
  const ingresosOtros = { total: 0 };
  const detalleEgresos = [];
  for (let r = HDR + 1; r <= HDR + 400; r++) {
    const row = mov.getRow(r);
    const f = row.getCell(1).value;
    if (!(f instanceof Date)) continue;
    const tipo = String(row.getCell(3).value || "").toLowerCase();
    const miembro = row.getCell(4).value || "";
    const cat = row.getCell(5).value || "Otros";
    const desc = row.getCell(6).value || "";
    const monto = row.getCell(7).value, mon = row.getCell(8).value;
    if (!monto) continue;
    const ars = toARS(monto, mon);
    if (f.getMonth() + 1 === mes && f.getFullYear() === anio) {
      if (tipo === "egreso") {
        egresosMes[cat] = (egresosMes[cat] || 0) + ars;
        detalleEgresos.push({ fecha: f.toISOString().slice(0, 10), categoria: String(cat), descripcion: String(desc), monto: Math.round(ars), miembro: String(miembro) });
      } else if (tipo === "ingreso") {
        ingresosOtros.total += ars;
      }
    }
  }

  const previosPorCat = {};
  const previosTotales = [];
  for (let offset = 1; offset <= 3; offset++) {
    let m = mes - offset, a = anio;
    while (m <= 0) { m += 12; a -= 1; }
    const acumCat = {};
    let totalEgr = 0;
    for (let r = HDR + 1; r <= HDR + 400; r++) {
      const row = mov.getRow(r);
      const f = row.getCell(1).value;
      if (!(f instanceof Date)) continue;
      if (f.getMonth() + 1 !== m || f.getFullYear() !== a) continue;
      const tipo = String(row.getCell(3).value || "").toLowerCase();
      const cat = row.getCell(5).value || "Otros";
      const monto = row.getCell(7).value, mon = row.getCell(8).value;
      if (!monto || tipo !== "egreso") continue;
      const ars = toARS(monto, mon);
      acumCat[cat] = (acumCat[cat] || 0) + ars;
      totalEgr += ars;
    }
    previosTotales.push({ mes: m, anio: a, nombre: MES_NOMBRES[m - 1], egresos_total: Math.round(totalEgr) });
    for (const [k, v] of Object.entries(acumCat)) previosPorCat[k] = (previosPorCat[k] || []).concat(v);
  }
  const promedioPrevios = {};
  for (const k of Object.keys(previosPorCat)) {
    const arr = previosPorCat[k];
    promedioPrevios[k] = Math.round(arr.reduce((s, x) => s + x, 0) / 3);
  }

  const inq = wb.getWorksheet("Inquilinos");
  const pendientes = [];
  for (let r = HDR + 1; r <= HDR + 33; r++) {
    const row = inq.getRow(r);
    const id = row.getCell(1).value;
    if (!id) continue;
    const estado = String(row.getCell(14).value || "").trim();
    if (estado.toLowerCase() !== "activo") continue;
    if (inquilinosCobrados.has(String(id))) continue;
    pendientes.push({
      id, nombre: String(row.getCell(2).value || ""), local: String(row.getCell(3).value || ""),
      alquiler_ars: Math.round(Number(row.getCell(12).value) || 0),
    });
  }

  const totalEgr = Object.values(egresosMes).reduce((s, x) => s + x, 0);
  return {
    mes: { numero: mes, nombre: MES_NOMBRES[mes - 1], anio },
    moneda: "ARS",
    ingresos: {
      alquileres: Math.round(alquileres), otros: Math.round(ingresosOtros.total),
      total: Math.round(alquileres + ingresosOtros.total),
    },
    egresos: {
      total: Math.round(totalEgr),
      por_categoria: Object.fromEntries(Object.entries(egresosMes).map(([k, v]) => [k, Math.round(v)])),
    },
    promedio_meses_previos: {
      meses: previosTotales,
      por_categoria: promedioPrevios,
    },
    resultado_neto: Math.round(alquileres + ingresosOtros.total - totalEgr),
    cobros_pendientes: pendientes,
    top_egresos: detalleEgresos.sort((a, b) => b.monto - a.monto).slice(0, 10),
  };
}
