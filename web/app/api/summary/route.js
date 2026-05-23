export const runtime = "nodejs";
export const dynamic = "force-dynamic";
import { NextResponse } from "next/server";
import { readWorkbookBuffer } from "../../../lib/blob.js";
import { loadWB, resumenMensual } from "../../../lib/excel.js";

export async function GET() {
  try {
    const wb = await loadWB(await readWorkbookBuffer());
    return NextResponse.json(resumenMensual(wb, 0));
  } catch (e) {
    return NextResponse.json({ error: String(e.message || e) }, { status: 400 });
  }
}
