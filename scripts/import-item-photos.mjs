import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import sharePointHandler from "../api/xrf-sharepoint.js";

const apply = process.argv.includes("--apply");
const direct = process.argv.includes("--direct");
const summaryOnly = process.argv.includes("--summary");
const apiBase = process.env.MXKPCS_API_BASE || "http://localhost:3000";
const photoDir = path.resolve(process.env.MXKPCS_PHOTO_DIR || "item_photos_import");
const allowed = /\.(jpe?g|png|webp)$/i;

if (direct) {
  const rawEnv = await readFile(path.resolve(".env.local"), "utf8");
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

async function invokeDirect(method, url, body) {
  let output = "";
  const responseHeaders = {};
  const response = {
    statusCode: 200,
    setHeader(name, value) { responseHeaders[name] = value; },
    end(value = "") { output = Buffer.isBuffer(value) ? value.toString("utf8") : String(value); },
  };
  await sharePointHandler({ method, url, headers: {}, body }, response);
  const payload = output ? JSON.parse(output) : null;
  if (response.statusCode < 200 || response.statusCode >= 300 || !payload?.ok) throw new Error(payload?.error || `HTTP ${response.statusCode}`);
  return payload.result ?? payload;
}

async function api(method, body) {
  if (direct) return invokeDirect(method, "/api/xrf-sharepoint", body);
  const response = await fetch(`${apiBase}/api/xrf-sharepoint`, {
    method,
    headers: { Accept: "application/json", ...(body ? { "Content-Type": "application/json" } : {}) },
    body: body ? JSON.stringify(body) : undefined,
  });
  const payload = await response.json().catch(() => null);
  if (!response.ok || !payload?.ok) throw new Error(payload?.error || `HTTP ${response.status}`);
  return payload.result ?? payload;
}

async function columnsFor(listName) {
  if (direct) return (await invokeDirect("GET", `/api/xrf-sharepoint?columnsFor=${encodeURIComponent(listName)}`))?.columns || [];
  const response = await fetch(`${apiBase}/api/xrf-sharepoint?columnsFor=${encodeURIComponent(listName)}`, { headers: { Accept: "application/json" } });
  const payload = await response.json().catch(() => null);
  if (!response.ok || !payload?.ok) throw new Error(payload?.error || `HTTP ${response.status}`);
  return payload.columns || [];
}

async function upload(itemId, itemSpItemId, fileName, bytes) {
  const session = await api("POST", {
    action: "beginItemPhotoUpload",
    payload: { itemId, itemSpItemId, fileName, size: bytes.length },
  });
  let offset = 0;
  let driveItemId = "";
  while (offset < bytes.length) {
    const end = Math.min(offset + session.chunkBytes, bytes.length);
    const chunk = await api("POST", {
      action: "uploadPrecisionFileChunk",
      payload: {
        uploadUrl: session.uploadUrl,
        offset,
        totalSize: bytes.length,
        base64: bytes.subarray(offset, end).toString("base64"),
      },
    });
    offset = end;
    if (chunk.complete) driveItemId = chunk.driveItemId || "";
  }
  if (!driveItemId) throw new Error("업로드 완료 파일 ID를 받지 못했습니다.");
  return api("POST", {
    action: "finishItemPhotoUpload",
    payload: { itemId, itemSpItemId, storedName: session.storedName, driveItemId },
  });
}

const itemColumnNames = new Set((await columnsFor("XRF_Items")).map(column => column.name));
const missingColumns = ["Photo_File_Name", "Photo_File_ID", "Photo_File_URL"].filter(name => !itemColumnNames.has(name));
if (missingColumns.length) throw new Error(`XRF_Items에 필요한 사진 참조 열이 없습니다: ${missingColumns.join(", ")}`);

const bootstrap = await api("GET");
const rows = bootstrap?.lists?.XRF_Items || [];
const byKey = new Map();
for (const row of rows) {
  const itemId = String(row?.fields?.Title || "").trim();
  const partNumber = String(row?.fields?.field_1 || "").trim();
  if (itemId) byKey.set(itemId.toLocaleLowerCase("en-US"), { row, itemId, partNumber });
  if (partNumber) byKey.set(partNumber.toLocaleLowerCase("en-US"), { row, itemId, partNumber });
}

const names = (await readdir(photoDir, { withFileTypes: true }))
  .filter(entry => entry.isFile() && allowed.test(entry.name))
  .map(entry => entry.name)
  .sort((a, b) => a.localeCompare(b, "ko", { numeric: true }));
const matches = [];
const unmatched = [];
const duplicates = [];
const usedItems = new Map();
for (const fileName of names) {
  const key = path.basename(fileName, path.extname(fileName)).toLocaleLowerCase("en-US");
  const item = byKey.get(key);
  if (!item) {
    unmatched.push(fileName);
    continue;
  }
  if (usedItems.has(item.itemId)) {
    duplicates.push([usedItems.get(item.itemId), fileName, item.itemId]);
    continue;
  }
  usedItems.set(item.itemId, fileName);
  matches.push({ fileName, itemSpItemId: item.row?.id || null, ...item });
}

console.log(JSON.stringify({ mode: apply ? "APPLY" : "DRY_RUN", apiBase, photoDir, files: names.length,
  matched: summaryOnly ? matches.length : matches.map(row => ({ file: row.fileName, partNumber: row.partNumber, itemId: row.itemId })),
  unmatched, duplicates }, null, 2));

if (unmatched.length || duplicates.length) {
  throw new Error("매핑되지 않은 사진 또는 품목별 중복 사진이 있어 업로드를 중단했습니다.");
}
if (!apply) {
  console.log("DRY RUN 완료: 실제 업로드는 npm run photos:import 로 실행합니다.");
}

for (const row of apply ? matches : []) {
  const bytes = await readFile(path.join(photoDir, row.fileName));
  if (!bytes.length || bytes.length > 10 * 1024 * 1024) throw new Error(`${row.fileName}: 10MB 이하의 비어 있지 않은 사진이어야 합니다.`);
  const stored = await upload(row.itemId, row.itemSpItemId, row.fileName, bytes);
  console.log(`UPLOADED ${row.partNumber} -> ${stored.name} (${stored.id})`);
}
if (apply) console.log(`품목 사진 ${matches.length}개 업로드 완료`);
