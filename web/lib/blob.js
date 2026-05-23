// Almacenamiento de la planilla.
// - Local (sin BLOB_READ_WRITE_TOKEN): lee/escribe el archivo Finanzas_Familia_2026.xlsx
//   que está en la raíz del proyecto. No necesitás Vercel para usarlo en tu compu.
// - En Vercel (con BLOB_READ_WRITE_TOKEN): usa Vercel Blob.
import { readFile, writeFile } from "fs/promises";
import { resolve } from "path";

const PATHNAME = "finanzas/Finanzas_Familia_2026.xlsx";
const LOCAL_PATH = process.env.EXCEL_PATH || resolve(process.cwd(), "..", "Finanzas_Familia_2026.xlsx");
const useBlob = () => Boolean(process.env.BLOB_READ_WRITE_TOKEN);

export async function readWorkbookBuffer() {
  if (!useBlob()) {
    return await readFile(LOCAL_PATH);
  }
  const { list } = await import("@vercel/blob");
  const { blobs } = await list({ prefix: PATHNAME });
  if (!blobs.length) {
    throw new Error("No hay planilla en Blob. Subila una vez con: npm run seed");
  }
  const res = await fetch(blobs[0].url, { cache: "no-store" });
  return Buffer.from(await res.arrayBuffer());
}

export async function writeWorkbookBuffer(buf) {
  if (!useBlob()) {
    await writeFile(LOCAL_PATH, buf);
    return;
  }
  const { put } = await import("@vercel/blob");
  await put(PATHNAME, buf, {
    access: "public",
    addRandomSuffix: false,
    allowOverwrite: true,
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
