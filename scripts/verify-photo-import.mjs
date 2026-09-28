import { readFile } from "node:fs/promises";

const snapshotPath = process.argv[2];
const apiBase = process.env.MXKPCS_API_BASE || "http://127.0.0.1:3000";
if (!snapshotPath) throw new Error("Usage: node scripts/verify-photo-import.mjs <pre-import snapshot>");

const before = JSON.parse(await readFile(snapshotPath, "utf8")).lists;
const response = await fetch(`${apiBase}/api/xrf-sharepoint`, { headers: { Accept: "application/json" } });
const payload = await response.json();
if (!response.ok || !payload?.ok) throw new Error(payload?.error || `HTTP ${response.status}`);
const after = payload.lists;
const listNames = [
  "XRF_Items",
  "XRF_Requests",
  "XRF_Measurements",
  "XRF_ElementResults",
  "XRF_PrecisionAnalyses",
  "XRF_PrecisionElementResults",
];
const photoFields = new Set(["Photo_File_Name", "Photo_File_ID", "Photo_File_URL"]);
const sharePointSystemFields = new Set([
  "@odata.etag",
  "Modified",
  "EditorLookupId",
  "_UIVersionString",
  "AppEditorLookupId",
]);
const changes = [];

for (const listName of listNames) {
  const beforeRows = new Map((before[listName] || []).map(row => [String(row.id), row]));
  const afterRows = new Map((after[listName] || []).map(row => [String(row.id), row]));
  for (const [id, beforeRow] of beforeRows) {
    const afterRow = afterRows.get(id);
    if (!afterRow) {
      changes.push(`${listName}/${id}/ROW_DELETED`);
      continue;
    }
    const keys = new Set([...Object.keys(beforeRow.fields || {}), ...Object.keys(afterRow.fields || {})]);
    for (const key of keys) {
      if (listName === "XRF_Items" && (photoFields.has(key) || sharePointSystemFields.has(key))) continue;
      const beforeValue = beforeRow.fields?.[key] ?? null;
      const afterValue = afterRow.fields?.[key] ?? null;
      if (JSON.stringify(beforeValue) !== JSON.stringify(afterValue)) changes.push(`${listName}/${id}/${key}`);
    }
  }
  for (const id of afterRows.keys()) {
    if (!beforeRows.has(id)) changes.push(`${listName}/${id}/ROW_ADDED`);
  }
}

console.log(JSON.stringify({ nonPhotoFieldChanges: changes.length, changes }, null, 2));
if (changes.length) process.exitCode = 1;
