export const runtime = "nodejs";
export const dynamic = "force-dynamic";
import { readWorkbookBuffer } from "../../../lib/blob.js";

export async function GET() {
  try {
    const buf = await readWorkbookBuffer();
    return new Response(buf, {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="Finanzas_Familia_2026.xlsx"',
      },
    });
  } catch (e) {
    return new Response(String(e.message || e), { status: 400 });
  }
}
