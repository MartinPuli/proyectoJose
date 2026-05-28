export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;
import { NextResponse } from "next/server";
import { GoogleGenAI } from "@google/genai";
import { readWorkbookBuffer } from "../../../lib/blob.js";
import { loadWB, analisisMes } from "../../../lib/excel.js";

function clean(s) { return String(s || "").replace(/^﻿/, "").trim(); }

const fmtARS = (n) => "$" + Number(n || 0).toLocaleString("es-AR");

function armarPrompt(d) {
  const lineas = [];
  lineas.push(`Mes a analizar: ${d.mes.nombre} ${d.mes.anio}.`);
  lineas.push(`Ingresos del mes: alquileres ${fmtARS(d.ingresos.alquileres)}, otros ${fmtARS(d.ingresos.otros)}, total ${fmtARS(d.ingresos.total)}.`);
  lineas.push(`Egresos del mes: total ${fmtARS(d.egresos.total)}.`);
  lineas.push(`Resultado neto: ${fmtARS(d.resultado_neto)}.`);
  lineas.push("");
  lineas.push("Egresos por categoría este mes:");
  for (const [k, v] of Object.entries(d.egresos.por_categoria).sort((a, b) => b[1] - a[1])) {
    const prev = d.promedio_meses_previos.por_categoria[k];
    if (prev) {
      const dif = v - prev;
      const pct = prev > 0 ? Math.round((dif / prev) * 100) : null;
      const signo = dif > 0 ? "+" : "";
      lineas.push(`  · ${k}: ${fmtARS(v)} (promedio últimos 3 meses ${fmtARS(prev)}, diferencia ${signo}${fmtARS(dif)}${pct !== null ? ` = ${signo}${pct}%` : ""})`);
    } else {
      lineas.push(`  · ${k}: ${fmtARS(v)} (sin historial previo)`);
    }
  }
  lineas.push("");
  lineas.push("Top egresos individuales del mes:");
  for (const e of d.top_egresos) {
    lineas.push(`  · ${e.fecha} — ${e.categoria} — ${e.descripcion || "(sin descripción)"} — ${fmtARS(e.monto)} (${e.miembro})`);
  }
  if (d.cobros_pendientes.length) {
    lineas.push("");
    lineas.push(`Inquilinos activos sin cobro registrado este mes (${d.cobros_pendientes.length}):`);
    for (const i of d.cobros_pendientes) lineas.push(`  · ${i.nombre} — Local ${i.local} — alquiler esperado ${fmtARS(i.alquiler_ars)}`);
  }
  lineas.push("");
  lineas.push("Meses previos comparados (totales de egreso):");
  for (const m of d.promedio_meses_previos.meses) lineas.push(`  · ${m.nombre} ${m.anio}: ${fmtARS(m.egresos_total)}`);
  return lineas.join("\n");
}

const SYSTEM = `Sos un analista financiero familiar. Recibís datos de un mes (Argentina, ARS) y devolvés un análisis útil y honesto.

Formato OBLIGATORIO de tu respuesta (sin encabezados extra, sin saludo):
**Resumen ejecutivo**
2 o 3 oraciones: cómo cerró el mes en una pincelada.

**Lo bueno**
- Lista 1 a 3 cosas positivas (categorías donde se gastó menos que el promedio, ingresos altos, etc.).

**Lo que destaca**
- Lista 1 a 4 alertas o cambios significativos respecto al promedio de los últimos 3 meses. Mostrá el porcentaje cuando sea informativo. Si una categoría subió >20% o bajó >20% es relevante.

**Cobros pendientes**
Solo si hay inquilinos sin cobrar este mes. Mencioná cuántos son y el total esperado. Si no hay, omití esta sección.

**Sugerencias para el mes que viene**
- 2 o 3 acciones concretas y aplicables. No genéricas tipo "ahorrá más" — sí "fijate que en Servicios estás 35% arriba, ¿hubo un consumo puntual o se mantiene?".

REGLAS:
- Hablá en argentino, voseo, directo, sin adornos.
- Usá los números que te paso, no inventes.
- Si los datos son escasos (mes recién empezado, sin historial previo), decílo y no fuerces conclusiones.
- Inflación argentina alta: aumentos del 5-10% no son alarma, recién por encima del 20% merecen mención.
- Nada de emojis ni claims sobre cosas que no están en los datos.`;

export async function GET(req) {
  try {
    const url = new URL(req.url);
    const raw = url.searchParams.get("mes");
    let mes = new Date().getMonth() + 1;
    if (raw !== null) {
      const n = Number(raw);
      if (Number.isFinite(n) && n >= 1 && n <= 12) mes = n;
    }
    const wb = await loadWB(await readWorkbookBuffer());
    const datos = analisisMes(wb, mes);

    const apiKey = clean(process.env.GEMINI_API_KEY);
    if (!apiKey) throw new Error("Falta GEMINI_API_KEY.");
    const ai = new GoogleGenAI({ apiKey });
    const model = clean(process.env.GEMINI_MODEL) || "gemini-2.5-flash";

    const promptUsuario = armarPrompt(datos);
    const resp = await ai.models.generateContent({
      model,
      contents: [{ role: "user", parts: [{ text: promptUsuario }] }],
      config: { systemInstruction: SYSTEM, temperature: 0.4 },
    });
    const texto = (resp.text || "").trim();
    return NextResponse.json({ texto, datos });
  } catch (e) {
    return NextResponse.json({ error: String(e.message || e) }, { status: 400 });
  }
}
