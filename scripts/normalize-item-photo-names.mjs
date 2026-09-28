import process from "node:process";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import sharePointHandler from "../api/xrf-sharepoint.js";

const apiBase = process.env.MXKPCS_API_BASE || "http://127.0.0.1:3000";
const direct = process.argv.includes("--direct");
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

if (direct) {
  const rawEnv = await readFile(path.join(projectRoot, ".env.local"), "utf8");
  const allowedEnvironment = new Set(["MXKPCS_CLIENT_SECRET", "MXKPCS_TENANT", "MXKPCS_CLIENT_ID", "MXKPCS_SITE_ID"]);
  for (const line of rawEnv.split(/\r?\n/)) {
    const match = /^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/.exec(line);
    if (!match || !allowedEnvironment.has(match[1]) || process.env[match[1]]) continue;
    let value = match[2];
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) value = value.slice(1, -1);
    process.env[match[1]] = value;
  }
  process.env.VERCEL_ENV = "development";
}

async function invokeDirect(method, body) {
  let output = "";
  const response = {
    statusCode: 200,
    setHeader() {},
    end(value = "") { output = Buffer.isBuffer(value) ? value.toString("utf8") : String(value); },
  };
  await sharePointHandler({ method, url: "/api/xrf-sharepoint", headers: {}, body }, response);
  const payload = output ? JSON.parse(output) : null;
  if (response.statusCode < 200 || response.statusCode >= 300 || !payload?.ok) {
    throw new Error(payload?.error || `HTTP ${response.statusCode}`);
  }
  return payload.result ?? payload;
}

async function request(method, body) {
  if (direct) return invokeDirect(method, body);
  const response = await fetch(`${apiBase}/api/xrf-sharepoint`, {
    method,
    headers: { Accept: "application/json", ...(body ? { "Content-Type": "application/json" } : {}) },
    body: body ? JSON.stringify(body) : undefined,
  });
  const payload = await response.json().catch(() => null);
  if (!response.ok || !payload?.ok) throw new Error(payload?.error || `HTTP ${response.status}`);
  return payload.result ?? payload;
}

const bootstrap = await request("GET");
const stamp = new Date().toISOString().replace(/[-:]/g, "").replace("T", "_").slice(0, 15);
const backupDir = path.join(projectRoot, "db_test_backup");
const snapshotPath = path.join(backupDir, `XRF_DB_PRE_PHOTO_RENAME_${stamp}.json`);
await mkdir(backupDir, { recursive: true });
await writeFile(snapshotPath, JSON.stringify(bootstrap), { flag: "wx" });
console.log(`SNAPSHOT ${snapshotPath}`);
const targets = (bootstrap?.lists?.XRF_Items || [])
  .map(row => ({
    itemSpItemId: row.id,
    itemId: String(row?.fields?.Title || "").trim(),
    partNumber: String(row?.fields?.field_1 || "").trim(),
    driveItemId: String(row?.fields?.Photo_File_ID || "").trim(),
    fileName: String(row?.fields?.Photo_File_Name || "").trim(),
  }))
  .filter(row => row.itemSpItemId && row.itemId && row.driveItemId && row.fileName);

let renamed = 0;
let moved = 0;
let deletedEmptyFolders = 0;
let changed = 0;
for (const target of targets) {
  const result = await request("POST", { action: "normalizeItemPhotoName", payload: target });
  if (result.renamed) renamed += 1;
  if (result.moved) moved += 1;
  if (result.deletedEmptyFolder) deletedEmptyFolders += 1;
  if (result.renamed || result.moved) changed += 1;
  console.log(`${result.moved ? "MOVED" : result.renamed ? "RENAMED" : "UNCHANGED"} ${target.partNumber} -> ${result.storedName}`);
}

console.log(JSON.stringify({ checked: targets.length, renamed, moved, deletedEmptyFolders, unchanged: targets.length - changed }, null, 2));
