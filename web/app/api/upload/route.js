export const runtime = "nodejs";
export const dynamic = "force-dynamic";
import { NextResponse } from "next/server";
import { writeWorkbookBuffer } from "../../../lib/blob.js";
import { loadWB } from "../../../lib/excel.js";

const HOJAS_REQUERIDAS = ["Cobros", "Movimientos", "Inquilinos", "Inflacion INDEC", "Parametros"];

export async function POST(req) {
  try {
    const form = await req.formData();
    const file = form.get("archivo");
    if (!file || !file.size) {
      return NextResponse.json({ error: "Subí un archivo .xlsx." }, { status: 400 });
    }
    if (!String(file.name || "").toLowerCase().endsWith(".xlsx")) {
      return NextResponse.json({ error: "El archivo tiene que ser un .xlsx." }, { status: 400 });
    }
    const buf = Buffer.from(await file.arrayBuffer());
    // Validar que tenga el formato esperado antes de reemplazar
    let wb;
    try {
      wb = await loadWB(buf);
    } catch (e) {
      return NextResponse.json({ error: "No pude leer el Excel: " + (e.message || e) }, { status: 400 });
    }
    const faltan = HOJAS_REQUERIDAS.filter((s) => !wb.getWorksheet(s));
    if (faltan.length) {
      return NextResponse.json(
        { error: "Falta(n) la(s) hoja(s): " + faltan.join(", ") + ". ¿Es la planilla correcta?" },
        { status: 400 }
      );
    }
    await writeWorkbookBuffer(buf);
    return NextResponse.json({ ok: true, mensaje: "Planilla actualizada con tu versión." });
  } catch (e) {
    return NextResponse.json({ error: String(e.message || e) }, { status: 400 });
  }
}
