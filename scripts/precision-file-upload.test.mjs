import assert from "node:assert/strict";
import handler from "../api/xrf-sharepoint.js";

const originalFetch = globalThis.fetch;
const originalSecret = process.env.MXKPCS_CLIENT_SECRET;
process.env.MXKPCS_CLIENT_SECRET = "test-only-secret";

const files = [];
const uploadedBytes = new Map();
let folderCreated = false;
let patchedPdfName = null;
let patchedConfirmation = null;
let patchedMeasurementFile = null;
let patchedPrecisionReport = null;
let sessionName = "";
let sessionBytes = Buffer.alloc(0);
let precisionHeaderAvailable = true;

function json(value, status = 200) {
  return new Response(JSON.stringify(value), { status, headers: { "Content-Type": "application/json" } });
}

globalThis.fetch = async (input, options = {}) => {
  const url = String(input);
  const method = String(options.method || "GET").toUpperCase();
  if (url.includes("/oauth2/v2.0/token")) return json({ access_token: "fake-token" });
  if (url.includes("/lists/dd0bf876-6066-4877-be6d-eb10eb5ce35c/columns") && method === "GET") {
    return json({ value: ["field_6", "Source_File_ID", "Source_File_URL"].map(name => ({ name })) });
  }
  if (url.includes("/lists/3c387146-a8c5-4b94-9bf3-251452b19a1c/columns") && method === "GET") {
    return json({ value: ["field_5", "field_6", "Report_File_ID", "Report_File_URL"].map(name => ({ name })) });
  }
  if (url.includes("/lists/dd0bf876-6066-4877-be6d-eb10eb5ce35c/items/58/fields") && method === "PATCH") {
    patchedMeasurementFile = JSON.parse(options.body);
    return json({ id: "58" });
  }
  if (url.includes("/lists/3c387146-a8c5-4b94-9bf3-251452b19a1c/items/5/fields") && method === "PATCH") {
    const fields = JSON.parse(options.body);
    patchedPdfName = fields.field_5;
    patchedConfirmation = fields.field_6;
    if (fields.Report_File_ID) patchedPrecisionReport = fields;
    return json({ id: "5" });
  }
  if (url.includes("/lists/dd0bf876-6066-4877-be6d-eb10eb5ce35c/items") && method === "GET") {
    return json({ value: [{ id: "58", fields: { Title: "C-18-01_20260716_01" } }] });
  }
  if (url.includes("/lists/3c387146-a8c5-4b94-9bf3-251452b19a1c/items") && method === "GET") {
    return json({ value: precisionHeaderAvailable
      ? [{ id: "5", fields: { Title: "PREC_C-18-01_20260716_01", field_1: "C-18-01_20260716_01" } }]
      : [] });
  }
  if (url.includes("/lists/0462cdd6-f3ce-40bf-a814-471796a6147b/drive")) {
    return json({ id: "report-drive", driveType: "documentLibrary" });
  }
  if (url.includes("/drives/report-drive/root:/PREC_C-18-01_20260716_01") && method === "GET") {
    return folderCreated ? json({ id: "folder-1", folder: {} }) : json({ error: "not found" }, 404);
  }
  if (url.includes("/drives/report-drive/root/children") && method === "POST") {
    folderCreated = true;
    return json({ id: "folder-1", folder: {} }, 201);
  }
  if (url.includes("/drives/report-drive/items/folder-1/children") && method === "GET") {
    return json({ value: files });
  }
  if (url.endsWith("/createUploadSession") && method === "POST") {
    sessionName = decodeURIComponent(url.split("/items/folder-1:/")[1].split(":/createUploadSession")[0]);
    sessionBytes = Buffer.alloc(0);
    return json({ uploadUrl: "https://kochind.sharepoint.com/_api/v2.0/uploadSession?test=1" });
  }
  if (url.includes("/_api/v2.0/uploadSession") && method === "PUT") {
    const bytes = Buffer.from(options.body);
    const [, startText, endText, totalText] = /^bytes (\d+)-(\d+)\/(\d+)$/.exec(options.headers["Content-Range"]);
    assert.equal(Number(startText), sessionBytes.length);
    assert.equal(Number(endText) - Number(startText) + 1, bytes.length);
    sessionBytes = Buffer.concat([sessionBytes, bytes]);
    if (sessionBytes.length < Number(totalText)) return json({ nextExpectedRanges: [`${sessionBytes.length}-`] }, 202);
    uploadedBytes.set(sessionName, sessionBytes);
    const item = {
      id: `file-${files.length + 1}`, name: sessionName, size: sessionBytes.length,
      file: {}, webUrl: `https://kochind.sharepoint.com/${encodeURIComponent(sessionName)}`,
      lastModifiedDateTime: new Date(2026, 8, 16, 12, files.length).toISOString(),
    };
    files.push(item);
    return json(item, 201);
  }
  throw new Error(`Unexpected mock Graph request: ${method} ${url}`);
};

async function invoke(method, body, url = "/api/xrf-sharepoint") {
  let output = "";
  const response = {
    statusCode: 200,
    setHeader() {},
    end(value = "") { output = value; },
  };
  await handler({ method, body, url }, response);
  return { status: response.statusCode, body: JSON.parse(output) };
}

async function upload(kind, fileName, bytes) {
  const measurementId = "C-18-01_20260716_01";
  const start = await invoke("POST", { action: "beginPrecisionFileUpload", payload: {
    measurementId, kind, fileName, size: bytes.length,
  } });
  assert.equal(start.status, 200);
  assert.equal(start.body.ok, true);
  const { uploadUrl, storedName } = start.body.result;
  let offset = 0;
  let driveItemId = null;
  while (offset < bytes.length) {
    const end = Math.min(offset + start.body.result.chunkBytes, bytes.length);
    const part = await invoke("POST", { action: "uploadPrecisionFileChunk", payload: {
      uploadUrl, offset, totalSize: bytes.length, base64: bytes.subarray(offset, end).toString("base64"),
    } });
    assert.equal(part.body.ok, true);
    offset = end;
    if (part.body.result.complete) driveItemId = part.body.result.driveItemId;
    else assert.deepEqual(part.body.result.nextExpectedRanges, [`${offset}-`]);
  }
  assert.ok(driveItemId);
  const finish = await invoke("POST", { action: "finishPrecisionFileUpload", payload: {
    measurementId, storedName, driveItemId,
  } });
  assert.equal(finish.body.result.name, fileName);
  assert.deepEqual(uploadedBytes.get(storedName), bytes);
}

try {
  const originalConsoleError = console.error;
  console.error = () => {};
  let prematureReport;
  let unsafeChunk;
  try {
    prematureReport = await invoke("POST", { action: "beginPrecisionFileUpload", payload: {
      measurementId: "C-18-01_20260716_01", kind: "report", fileName: "early.pdf", size: 12,
    } });
    unsafeChunk = await invoke("POST", { action: "uploadPrecisionFileChunk", payload: {
      uploadUrl: "http://localhost/internal", offset: 0, totalSize: 4, base64: "dGVzdA==",
    } });
  } finally {
    console.error = originalConsoleError;
  }
  assert.equal(prematureReport.body.ok, false);
  assert.match(prematureReport.body.error, /XRF 원본 파일을 먼저 업로드/);
  assert.equal(unsafeChunk.body.ok, false);
  precisionHeaderAvailable = false;
  await upload("source", "measurement-only.xls", Buffer.from("measurement source"));
  precisionHeaderAvailable = true;
  await upload("source", "original.xls", Buffer.alloc(8 * 320 * 1024 + 17, 0x61));
  await upload("report", "analysis.pdf", Buffer.from("%PDF-1.4\ntest"));
  const listed = await invoke("GET", null, "/api/xrf-sharepoint?precisionFilesFor=C-18-01_20260716_01");
  assert.equal(listed.body.files.source.name, "original.xls");
  assert.equal(listed.body.files.report.name, "analysis.pdf");
  assert.equal(patchedPdfName, "analysis.pdf");
  assert.equal(patchedConfirmation, false);
  assert.equal(patchedMeasurementFile.field_6, "original.xls");
  assert.equal(patchedMeasurementFile.Source_File_ID, "file-2");
  assert.equal(patchedPrecisionReport.Report_File_ID, "file-3");
  console.log("precision-file-upload: PASS");
} finally {
  globalThis.fetch = originalFetch;
  if (originalSecret === undefined) delete process.env.MXKPCS_CLIENT_SECRET;
  else process.env.MXKPCS_CLIENT_SECRET = originalSecret;
}
