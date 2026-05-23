// Almacenamiento de la planilla.
// - Local (sin BLOB_READ_WRITE_TOKEN): lee/escribe el archivo Finanzas_Familia_2026.xlsx
//   que está en la raíz del proyecto. No necesitás Vercel para usarlo en tu compu.
// - En Vercel (con BLOB_READ_WRITE_TOKEN): usa Vercel Blob en modo PRIVADO
//   (los datos de la familia no quedan en una URL pública).
import { readFile, writeFile } from "fs/promises";
import { resolve } from "path";

const PATHNAME = "finanzas/Finanzas_Familia_2026.xlsx";
const LOCAL_PATH = process.env.EXCEL_PATH || resolve(process.cwd(), "..", "Finanzas_Familia_2026.xlsx");
const useBlob = () => Boolean(process.env.BLOB_READ_WRITE_TOKEN);

export async function readWorkbookBuffer() {
  if (!useBlob()) {
    return await readFile(LOCAL_PATH);
  }
  const { get } = await import("@vercel/blob");
  const res = await get(PATHNAME, { access: "private" });
  if (!res) {
    throw new Error("No hay planilla en Blob. Subila una vez con: npm run seed");
  }
  // res.stream es un ReadableStream web → lo pasamos a Buffer.
  return Buffer.from(await new Response(res.stream).arrayBuffer());
}

export async function writeWorkbookBuffer(buf) {
  if (!useBlob()) {
    await writeFile(LOCAL_PATH, buf);
    return;
  }
  const { put } = await import("@vercel/blob");
  await put(PATHNAME, buf, {
    access: "private",
    addRandomSuffix: false,
    allowOverwrite: true,
    contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
