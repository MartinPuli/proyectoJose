import { GoogleGenAI } from "@google/genai";
import { agregarCobro, agregarMovimiento, listarInquilinos, listarCategorias, resumenMensual } from "./excel.js";

function buildSystem(wb) {
  const hoy = new Date().toISOString().slice(0, 10);
  const CATEGORIAS = listarCategorias(wb).join(", ") || "Otros";
  return `Sos el asistente de finanzas de una familia que vive de alquileres de locales.
Hoy es ${hoy}. Cuando el usuario manda un audio, texto o comprobante, identificá las operaciones
y cargalas con las herramientas: agregar_cobro (alquileres; identificá al inquilino por nombre o
local), agregar_movimiento (gastos/ingresos de la familia).

REGLAS IMPORTANTES:
- Para registrar SIEMPRE llamá a la herramienta correspondiente. NUNCA digas que registraste,
  anoté o cargué algo si no llamaste a la herramienta en este mismo turno. No inventes registros.
- Si falta un dato OBLIGATORIO (sobre todo el MONTO, o a qué inquilino corresponde un cobro),
  NO adivines: registrá lo que sí puedas y hacé UNA pregunta corta y concreta SOLO por lo que falta.
- Si el mensaje tiene varias operaciones, registrá todas las que tengan datos completos.
- Fecha: si no dicen el año, usá ${hoy.slice(0, 4)}. Si no dicen fecha, usá hoy (${hoy}).
- MONEDA: "dólares", "dolar", "USD", "U$S", "verdes" => USD. "pesos", "mango", "$", o sin aclarar => ARS.
  "luca"/"lucas" = miles; "palo"/"palos" = millones. Default ARS.
- Para egresos elegí EXACTAMENTE una de estas categorías: ${CATEGORIAS}.
- MIEMBRO: completá 'miembro' con EXACTAMENTE uno de estos valores: "José", "Clara", "Laura", "Familia".
  Mapeo:
    · "soy José" / "yo" / "José" / "pagué yo" => José
    · "soy Clara" / "Clara" / "mi hermana" / "la hermana" => Clara
    · "soy Laura" / "Laura" / "mi mamá" / "mamá" / "mami" => Laura
    · gastos compartidos o cuando no se aclara => Familia (default)
- Completá 'descripcion' con un texto corto de qué fue el gasto/ingreso.
- MEDIO DE PAGO: si dicen cómo se pagó o cobró, completá SIEMPRE 'medio_pago' con EXACTAMENTE
  una de estas opciones: Efectivo, Transferencia, Cheque, Débito, Mercado Pago, Otro.
  Mapeo: "efectivo"/"cash"/"plata" => Efectivo; "transferencia"/"transfer"/"CBU"/"alias" => Transferencia;
  "cheque"/"echeq" => Cheque; "débito"/"tarjeta de débito" => Débito; "MP"/"mercadopago"/"mercado pago" => Mercado Pago;
  "crédito"/"tarjeta"/"tarjeta de crédito" o cualquier otro => Otro. Aplica tanto a gastos como a cobros.
Respondé SIEMPRE en español, breve. Si registraste algo, decí qué. Si te falta info, terminá con la pregunta.`;
}

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
      miembro: { type: "STRING", description: "EXACTAMENTE uno de: José, Clara, Laura, Familia" }, categoria: { type: "STRING" },
      descripcion: { type: "STRING" }, fecha: { type: "STRING" }, moneda: { type: "STRING" },
      medio_pago: { type: "STRING" }, notas: { type: "STRING" } }, required: ["monto"] } },
  { name: "listar_inquilinos", description: "Lista los inquilinos (id, nombre, local).",
    parameters: { type: "OBJECT", properties: {} } },
  { name: "listar_categorias", description: "Lista las categorías válidas para 'agregar_movimiento' (las del Presupuesto). Si no usás una de estas, el egreso no aparece en el Presupuesto.",
    parameters: { type: "OBJECT", properties: {} } },
  { name: "resumen_mensual", description: "Resumen de ingresos/egresos (mes 1-12, 0=año).",
    parameters: { type: "OBJECT", properties: { mes: { type: "INTEGER" } } } },
];

function exec(wb, name, args) {
  args = args || {};
  switch (name) {
    case "agregar_cobro": return agregarCobro(wb, args);
    case "agregar_movimiento": return agregarMovimiento(wb, args);
    case "listar_inquilinos": return listarInquilinos(wb);
    case "listar_categorias": return listarCategorias(wb);
    case "resumen_mensual": return resumenMensual(wb, args.mes || 0);
    default: return { error: "tool desconocida: " + name };
  }
}

// historial: [{ rol: "user"|"model", texto: string }] de turnos previos, para que
// pueda responder repreguntas manteniendo el contexto de la conversación.
function aHistorialGemini(historial) {
  if (!Array.isArray(historial)) return [];
  return historial
    .filter((h) => h && h.texto)
    .map((h) => ({ role: h.rol === "model" ? "model" : "user", parts: [{ text: String(h.texto) }] }));
}

// Saca BOM/espacios invisibles que a veces aparecen al cargar env vars
// desde PowerShell (UTF-16 LE con BOM); ya pasó con GEMINI_API_KEY antes.
function clean(s) { return String(s || "").replace(/^﻿/, "").trim(); }

export async function procesar({ texto, fileBase64, mime, wb, historial }) {
  const apiKey = clean(process.env.GEMINI_API_KEY);
  if (!apiKey) throw new Error("Falta GEMINI_API_KEY.");
  const ai = new GoogleGenAI({ apiKey });
  const chat = ai.chats.create({
    model: clean(process.env.GEMINI_MODEL) || "gemini-2.5-flash",
    config: {
      systemInstruction: buildSystem(wb),
      tools: [{ functionDeclarations: declaraciones }],
    },
    history: aHistorialGemini(historial),
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

  const resumen = (response.text || "").trim();
  // Necesita respuesta del usuario si no registró nada, o si terminó preguntando algo.
  const pregunta = operaciones.length === 0 || /\?\s*$/.test(resumen);
  return { resumen, operaciones, pregunta };
}
