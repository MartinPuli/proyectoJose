"""Genera la app web Next.js (deployable en Vercel) en la carpeta web/.

El Excel vive en Vercel Blob (el filesystem de Vercel es efímero). La app agrega
filas con exceljs y procesa audio/texto/imagen/PDF con Gemini (function calling).
"""
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent / "web"
FORCE = "--force" in sys.argv  # por defecto NO pisa archivos existentes
FILES = {}

FILES["package.json"] = r'''{
  "name": "finanzas-familia-web",
  "version": "1.0.0",
  "private": true,
  "scripts": {
    "dev": "next dev",
    "build": "next build",
    "start": "next start",
    "seed": "node scripts/seed-blob.mjs"
  },
  "dependencies": {
    "@google/genai": "^1.0.0",
    "@vercel/blob": "^0.27.0",
    "exceljs": "^4.4.0",
    "next": "14.2.35",
    "react": "^18.3.1",
    "react-dom": "^18.3.1"
  }
}
'''

FILES["next.config.js"] = r'''/** @type {import('next').NextConfig} */
const nextConfig = {
  experimental: { serverComponentsExternalPackages: ["exceljs"] },
};
module.exports = nextConfig;
'''

FILES[".env.example"] = r'''# Variables de entorno (Vercel -> Project Settings -> Environment Variables)
GEMINI_API_KEY=AIza...                 # https://aistudio.google.com/apikey
GEMINI_MODEL=gemini-2.5-flash
BLOB_READ_WRITE_TOKEN=vercel_blob_rw_... # Vercel -> Storage -> Blob -> .env
'''

FILES[".gitignore"] = r'''node_modules
.next
.env
.env.local
'''

FILES["app/layout.js"] = r'''export const metadata = { title: "Finanzas Familia", description: "Carga por audio/texto/documento" };
export default function RootLayout({ children }) {
  return (
    <html lang="es">
      <body style={{ margin: 0, background: "#f4f6fb", color: "#1b1b1b", fontFamily: "system-ui, Segoe UI, Roboto, sans-serif" }}>
        {children}
      </body>
    </html>
  );
}
'''

FILES["app/page.js"] = r'''"use client";
import { useState, useRef } from "react";

export default function Home() {
  const [texto, setTexto] = useState("");
  const [estado, setEstado] = useState("");
  const [salida, setSalida] = useState(null);
  const [grabando, setGrabando] = useState(false);
  const fileRef = useRef(null);
  const recRef = useRef(null);
  const chunksRef = useRef([]);
  const audioRef = useRef(null);

  async function toggleRec() {
    if (grabando) { recRef.current && recRef.current.stop(); return; }
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      const rec = new MediaRecorder(stream);
      chunksRef.current = [];
      rec.ondataavailable = (e) => chunksRef.current.push(e.data);
      rec.onstop = () => {
        audioRef.current = new Blob(chunksRef.current, { type: "audio/webm" });
        setEstado("🎧 Audio grabado, listo para procesar.");
        setGrabando(false);
        stream.getTracks().forEach((t) => t.stop());
      };
      rec.start(); recRef.current = rec; setGrabando(true); setEstado("Grabando...");
    } catch (e) { setEstado("No pude acceder al micrófono: " + e.message); }
  }

  async function enviar() {
    const fd = new FormData();
    fd.append("texto", texto || "");
    const f = fileRef.current && fileRef.current.files[0];
    if (f) fd.append("archivo", f);
    else if (audioRef.current) fd.append("archivo", audioRef.current, "audio.webm");
    setEstado("Procesando con Gemini...");
    try {
      const r = await fetch("/api/process", { method: "POST", body: fd });
      const data = await r.json();
      setSalida(data);
      setEstado("Listo.");
      audioRef.current = null; setTexto(""); if (fileRef.current) fileRef.current.value = "";
    } catch (e) { setEstado("Error: " + e.message); }
  }

  const card = { background: "#fff", borderRadius: 14, padding: 20, boxShadow: "0 2px 10px rgba(0,0,0,.06)", marginBottom: 18 };
  const btn = { border: 0, borderRadius: 10, padding: "11px 16px", fontSize: 15, cursor: "pointer" };

  return (
    <main style={{ maxWidth: 720, margin: "0 auto", padding: "0 16px" }}>
      <header style={{ background: "#1F3864", color: "#fff", padding: "20px 24px", margin: "0 -16px 24px", borderRadius: "0 0 14px 14px" }}>
        <h1 style={{ margin: 0, fontSize: 20 }}>💰 Finanzas Familia — carga rápida</h1>
        <p style={{ margin: "4px 0 0", opacity: .85, fontSize: 13 }}>Mandá un audio, texto o foto/PDF. Gemini lo entiende y lo carga en la planilla.</p>
      </header>
      <div style={card}>
        <textarea value={texto} onChange={(e) => setTexto(e.target.value)}
          placeholder="Ej: Cobré el alquiler del Local 12, 350 mil por transferencia. / Pagué la luz 48.500 con débito."
          style={{ width: "100%", minHeight: 90, border: "1px solid #cfd6e4", borderRadius: 10, padding: 12, fontSize: 15 }} />
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 12 }}>
          <button onClick={toggleRec} style={{ ...btn, background: grabando ? "#7a1f15" : "#c0392b", color: "#fff" }}>
            {grabando ? "⏹️ Detener" : "🎤 Grabar audio"}
          </button>
          <input type="file" ref={fileRef} accept="audio/*,image/*,.pdf" />
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 12 }}>
          <button onClick={enviar} style={{ ...btn, background: "#1F3864", color: "#fff" }}>Procesar y cargar</button>
          <a href="/api/excel"><button type="button" style={{ ...btn, background: "#eef1f8", color: "#1F3864" }}>⬇️ Descargar Excel actual</button></a>
        </div>
        <p style={{ color: "#667", fontSize: 13 }}>{estado}</p>
      </div>
      {salida && (
        <div style={card}>
          <strong>Resultado</strong>
          {salida.error ? (
            <p style={{ color: "#c0392b" }}>⚠️ {salida.error}</p>
          ) : (
            <div>
              <p>{salida.resumen}</p>
              {(salida.operaciones || []).map((o, i) => (
                <div key={i} style={{ borderLeft: "3px solid #2e5090", padding: "6px 10px", margin: "8px 0", background: "#f8fafc", borderRadius: 6, fontSize: 13 }}>
                  <b>{o.hoja}</b> (fila {o.fila}): {JSON.stringify(o)}
                </div>
              ))}
              {!(salida.operaciones || []).length && <div style={{ color: "#667", fontSize: 13 }}>No se cargó ninguna fila.</div>}
            </div>
          )}
        </div>
      )}
    </main>
  );
}
'''

FILES["lib/excel.js"] = r'''import ExcelJS from "exceljs";

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
'''

FILES["lib/blob.js"] = r'''import { put, list } from "@vercel/blob";

const PATHNAME = "finanzas/Finanzas_Familia_2026.xlsx";

export async function readWorkbookBuffer() {
  const { blobs } = await list({ prefix: PATHNAME });
  if (!blobs.length) {
    throw new Error("No hay planilla en Blob. Subila una vez con: npm run seed");
  }
  const res = await fetch(blobs[0].url, { cache: "no-store" });
  return Buffer.from(await res.arrayBuffer());
}

export async function writeWorkbookBuffer(buf) {
  await put(PATHNAME, buf, {
    access: "public",
    addRandomSuffix: false,
    allowOverwrite: true,
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
'''

FILES["lib/gemini.js"] = r'''import { GoogleGenAI } from "@google/genai";
import { agregarCobro, agregarMovimiento, actualizarIpc, listarInquilinos, resumenMensual } from "./excel.js";

const SYSTEM = `Sos el asistente de finanzas de una familia que vive de alquileres de locales.
Cuando el usuario manda un audio, texto o comprobante, identificá las operaciones y cargalas
con las herramientas: agregar_cobro (alquileres; identificá al inquilino por nombre o local),
agregar_movimiento (gastos/ingresos de la familia), actualizar_ipc (índice INDEC, mes YYYY-MM).
Si no dicen el año, asumí el actual. Moneda por defecto ARS. Después de cargar, respondé en
español qué registraste.`;

const declaraciones = [
  { name: "agregar_cobro", description: "Registra el cobro del alquiler de un inquilino.",
    parameters: { type: "OBJECT", properties: {
      monto: { type: "NUMBER" }, inquilino: { type: "STRING", description: "nombre o local" },
      id_inquilino: { type: "INTEGER" }, fecha: { type: "STRING", description: "YYYY-MM-DD" },
      moneda: { type: "STRING" }, medio_pago: { type: "STRING" }, estado: { type: "STRING" },
      periodo: { type: "STRING" }, notas: { type: "STRING" } }, required: ["monto"] } },
  { name: "agregar_movimiento", description: "Registra un ingreso o egreso de la familia.",
    parameters: { type: "OBJECT", properties: {
      monto: { type: "NUMBER" }, tipo: { type: "STRING", description: "Ingreso o Egreso" },
      miembro: { type: "STRING", description: "Yo, Hermana, Mamá o Familia" }, categoria: { type: "STRING" },
      descripcion: { type: "STRING" }, fecha: { type: "STRING" }, moneda: { type: "STRING" },
      medio_pago: { type: "STRING" }, notas: { type: "STRING" } }, required: ["monto"] } },
  { name: "actualizar_ipc", description: "Carga el índice IPC del INDEC para un mes.",
    parameters: { type: "OBJECT", properties: {
      mes: { type: "STRING", description: "YYYY-MM" }, indice: { type: "NUMBER" } }, required: ["mes", "indice"] } },
  { name: "listar_inquilinos", description: "Lista los inquilinos (id, nombre, local).",
    parameters: { type: "OBJECT", properties: {} } },
  { name: "resumen_mensual", description: "Resumen de ingresos/egresos (mes 1-12, 0=año).",
    parameters: { type: "OBJECT", properties: { mes: { type: "INTEGER" } } } },
];

function exec(wb, name, args) {
  args = args || {};
  switch (name) {
    case "agregar_cobro": return agregarCobro(wb, args);
    case "agregar_movimiento": return agregarMovimiento(wb, args);
    case "actualizar_ipc": return actualizarIpc(wb, args);
    case "listar_inquilinos": return listarInquilinos(wb);
    case "resumen_mensual": return resumenMensual(wb, args.mes || 0);
    default: return { error: "tool desconocida: " + name };
  }
}

export async function procesar({ texto, fileBase64, mime, wb }) {
  if (!process.env.GEMINI_API_KEY) throw new Error("Falta GEMINI_API_KEY.");
  const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });
  const chat = ai.chats.create({
    model: process.env.GEMINI_MODEL || "gemini-2.5-flash",
    config: {
      systemInstruction: SYSTEM,
      tools: [{ functionDeclarations: declaraciones }],
    },
  });
  const parts = [];
  if (fileBase64) parts.push({ inlineData: { data: fileBase64, mimeType: mime || "application/octet-stream" } });
  parts.push({ text: texto || "Procesá el comprobante adjunto y cargá lo que corresponda." });

  let response = await chat.sendMessage({ message: parts });
  const operaciones = [];
  for (let i = 0; i < 6; i++) {
    const calls = response.functionCalls || [];
    if (!calls.length) break;
    const responseParts = [];
    for (const c of calls) {
      let out;
      try { out = exec(wb, c.name, c.args); } catch (e) { out = { error: String(e.message || e) }; }
      if (out && out.hoja) operaciones.push(out);
      responseParts.push({ functionResponse: { name: c.name, response: { result: out } } });
    }
    response = await chat.sendMessage({ message: responseParts });
  }
  return { resumen: (response.text || "").trim(), operaciones };
}
'''

FILES["app/api/process/route.js"] = r'''export const runtime = "nodejs";
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
    let fileBase64 = null, mime = null;
    if (file && typeof file.arrayBuffer === "function" && file.size) {
      const buf = Buffer.from(await file.arrayBuffer());
      fileBase64 = buf.toString("base64");
      mime = file.type || null;
    }
    const wb = await loadWB(await readWorkbookBuffer());
    const res = await procesar({ texto, fileBase64, mime, wb });
    if (res.operaciones.length) await writeWorkbookBuffer(await wbToBuffer(wb));
    return NextResponse.json(res);
  } catch (e) {
    return NextResponse.json({ error: String(e.message || e) }, { status: 400 });
  }
}
'''

FILES["app/api/excel/route.js"] = r'''export const runtime = "nodejs";
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
'''

FILES["app/api/summary/route.js"] = r'''export const runtime = "nodejs";
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
'''

FILES["scripts/seed-blob.mjs"] = r'''// Sube la planilla base a Vercel Blob (correr una vez).
// Uso (local):  BLOB_READ_WRITE_TOKEN=... node scripts/seed-blob.mjs [ruta-al-xlsx]
import { put } from "@vercel/blob";
import { readFileSync } from "fs";
import { resolve } from "path";

const src = process.argv[2] || resolve(process.cwd(), "..", "Finanzas_Familia_2026.xlsx");
const buf = readFileSync(src);
const r = await put("finanzas/Finanzas_Familia_2026.xlsx", buf, {
  access: "public", addRandomSuffix: false, allowOverwrite: true,
  contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
});
console.log("Planilla subida a Blob:", r.url);
'''

FILES["README_DEPLOY.md"] = r'''# Deploy en Vercel — Finanzas Familia (web)

App Next.js que procesa audio/texto/foto/PDF con **Gemini** y guarda el Excel en **Vercel Blob**.

## Por qué Blob
Vercel es serverless: no puede guardar archivos en su disco. El `.xlsx` vive en Vercel Blob;
la app lo baja, le agrega filas y lo vuelve a subir. La descarga lo trae desde Blob.

## Pasos
1. Generar la planilla base (desde la raíz del proyecto):
   ```
   py generar_finanzas.py
   ```
2. Instalar dependencias de la web:
   ```
   cd web
   npm install
   ```
3. En Vercel: creá el proyecto apuntando a la carpeta `web/`. Activá **Storage -> Blob**
   (te da `BLOB_READ_WRITE_TOKEN`). Cargá las variables de entorno:
   - `GEMINI_API_KEY` (de https://aistudio.google.com/apikey)
   - `GEMINI_MODEL` (ej. `gemini-2.5-flash`)
   - `BLOB_READ_WRITE_TOKEN`
4. Subir la planilla base a Blob (una sola vez), con el token en el entorno:
   ```
   # PowerShell:  $env:BLOB_READ_WRITE_TOKEN="vercel_blob_rw_..."; npm run seed
   ```
5. Deploy: `vercel` (o conectá el repo de GitHub).

## Probar local
```
cd web
npm run dev      # http://localhost:3000  (necesita las 3 variables en web/.env.local)
```

## Notas honestas
- No pude testear el deploy en vivo desde acá; sí verifiqué que `npm run build` compile.
- exceljs preserva fórmulas/hojas al reescribir; Excel recalcula al abrir.
- Audio del navegador = webm; si Gemini lo rechaza, subí un mp3/wav.
- Concurrencia: si dos personas cargan exactamente a la vez puede pisarse una escritura
  (para una familia es muy improbable).
'''

count = 0
for rel, content in FILES.items():
    p = BASE / rel
    if p.exists() and not FORCE:
        print("skip (ya existe):", rel)
        continue
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(content, encoding="utf-8")
    count += 1
print(f"OK: {count} archivos escritos en web/ (usá --force para sobrescribir)")
for rel in FILES:
    print("  - web/" + rel)
