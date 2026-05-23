// Sube la planilla base a Vercel Blob (correr una vez).
// Uso (local):  BLOB_READ_WRITE_TOKEN=... node scripts/seed-blob.mjs [ruta-al-xlsx]
import { put } from "@vercel/blob";
import { readFileSync } from "fs";
import { resolve } from "path";

const src = process.argv[2] || resolve(process.cwd(), "..", "Finanzas_Familia_2026.xlsx");
const buf = readFileSync(src);
const r = await put("finanzas/Finanzas_Familia_2026.xlsx", buf, {
  access: "private", addRandomSuffix: false, allowOverwrite: true,
  contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
});
console.log("Planilla subida a Blob:", r.url);
