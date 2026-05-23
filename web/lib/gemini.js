import { GoogleGenerativeAI } from "@google/generative-ai";
import { agregarCobro, agregarMovimiento, actualizarIpc, listarInquilinos, resumenMensual } from "./excel.js";

const CATEGORIAS =
  "Vivienda/Expensas, Servicios (luz/gas/agua), Internet/Teléfono, Supermercado/Comida, " +
  "Salud/Obra social, Transporte/Auto, Educación, Impuestos (ABL/AFIP/Rentas), " +
  "Mantenimiento locales, Seguros, Honorarios/Gestión, Ocio/Viajes, Indumentaria, " +
  "Ahorro/Inversión, Otros";

function buildSystem() {
  const hoy = new Date().toISOString().slice(0, 10);
  return `Sos el asistente de finanzas de una familia que vive de alquileres de locales.
Hoy es ${hoy}. Cuando el usuario manda un audio, texto o comprobante, identificá las operaciones
y cargalas con las herramientas: agregar_cobro (alquileres; identificá al inquilino por nombre o
local), agregar_movimiento (gastos/ingresos de la familia), actualizar_ipc (índice INDEC, mes YYYY-MM).

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
- Si mencionan a un miembro (yo, mi hermana, mi mamá) completá 'miembro'.
- Completá 'descripcion' con un texto corto de qué fue el gasto/ingreso.
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

// historial: [{ rol: "user"|"model", texto: string }] de turnos previos, para que
// pueda responder repreguntas manteniendo el contexto de la conversación.
function aHistorialGemini(historial) {
  if (!Array.isArray(historial)) return [];
  return historial
    .filter((h) => h && h.texto)
    .map((h) => ({ role: h.rol === "model" ? "model" : "user", parts: [{ text: String(h.texto) }] }));
}

export async function procesar({ texto, fileBase64, mime, wb, historial }) {
  if (!process.env.GEMINI_API_KEY) throw new Error("Falta GEMINI_API_KEY.");
  const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
  const model = genAI.getGenerativeModel({
    model: process.env.GEMINI_MODEL || "gemini-2.5-flash",
    systemInstruction: buildSystem(),
    tools: [{ functionDeclarations: declaraciones }],
  });
  const parts = [];
  if (fileBase64) parts.push({ inlineData: { data: fileBase64, mimeType: mime || "application/octet-stream" } });
  parts.push({ text: texto || "Procesá el comprobante adjunto y cargá lo que corresponda." });

  const chat = model.startChat({ history: aHistorialGemini(historial) });
  let result = await chat.sendMessage(parts);
  const operaciones = [];
  for (let i = 0; i < 6; i++) {
    const calls = (result.response.functionCalls && result.response.functionCalls()) || [];
    if (!calls.length) break;
    const responses = [];
    for (const c of calls) {
      let out;
      try { out = exec(wb, c.name, c.args); } catch (e) { out = { error: String(e.message || e) }; }
      if (out && out.hoja) operaciones.push(out);
      responses.push({ functionResponse: { name: c.name, response: { result: out } } });
    }
    result = await chat.sendMessage(responses);
  }
  let resumen = "";
  try { resumen = result.response.text(); } catch (e) { resumen = ""; }
  resumen = (resumen || "").trim();
  // Necesita respuesta del usuario si no registró nada, o si terminó preguntando algo.
  const pregunta = operaciones.length === 0 || /\?\s*$/.test(resumen);
  return { resumen, operaciones, pregunta };
}
