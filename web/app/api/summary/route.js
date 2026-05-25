export const runtime = "nodejs";
export const dynamic = "force-dynamic";
import { NextResponse } from "next/server";
import { readWorkbookBuffer } from "../../../lib/blob.js";
import { loadWB, resumenMensual } from "../../../lib/excel.js";

const MES_NOMBRES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];

export async function GET(req) {
  try {
    const wb = await loadWB(await readWorkbookBuffer());
    const url = new URL(req.url);
    const raw = url.searchParams.get("mes");
    // Por defecto: mes actual (1-12). Si llega ?mes=0, devuelve año completo.
    let mes = new Date().getMonth() + 1;
    if (raw !== null) {
      const n = Number(raw);
      if (Number.isFinite(n) && n >= 0 && n <= 12) mes = n;
    }
    const res = resumenMensual(wb, mes);
    res.mes_nombre = mes === 0 ? "año completo" : MES_NOMBRES[mes - 1];
    res.mes_numero = mes;
    return NextResponse.json(res);
  } catch (e) {
    return NextResponse.json({ error: String(e.message || e) }, { status: 400 });
  }
}
