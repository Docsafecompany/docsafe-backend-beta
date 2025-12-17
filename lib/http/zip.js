// lib/http/zip.js
import path from "path";

export const outName = (single, base, name) => (single ? name : `${base}_${name}`);

export const baseName = (filename = "document") => filename.replace(/\.[^.]+$/, "");

export function sendZip(res, zip, zipName) {
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);
  res.send(zip.toBuffer());
}

// (optionnel) utile si tu veux plus tard
export const getExt = (fn = "") => (fn.includes(".") ? fn.split(".").pop().toLowerCase() : "");

export const getBaseFromPath = (filename = "") => path.parse(filename).name;
