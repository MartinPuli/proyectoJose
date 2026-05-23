export const runtime = "nodejs";
export const dynamic = "force-dynamic";
import { NextResponse } from "next/server";
import { readWorkbookBuffer, writeWorkbookBuffer } from "../../../lib/blob.js";
import { loadWB, wbToBuffer } from "../../../lib/excel.js";
import { procesar } from "../../../lib/gemini.js";

export async function POST(req) {
  try {
    const form = await req.formData();
    const texto = form.get("texto") || "";
    const file = form.get("archivo");
    let historial = [];
    try { historial = JSON.parse(form.get("historial") || "[]"); } catch (e) { historial = []; }
    let fileBase64 = null, mime = null;
    if (file && typeof file.arrayBuffer === "function" && file.size) {
      const buf = Buffer.from(await file.arrayBuffer());
      fileBase64 = buf.toString("base64");
      mime = file.type || null;
    }
    const wb = await loadWB(await readWorkbookBuffer());
    const res = await procesar({ texto, fileBase64, mime, wb, historial });
    if (res.operaciones.length) await writeWorkbookBuffer(await wbToBuffer(wb));
    return NextResponse.json(res);
  } catch (e) {
    return NextResponse.json({ error: String(e.message || e) }, { status: 400 });
  }
}
