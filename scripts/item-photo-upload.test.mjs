import assert from "node:assert/strict";
import sharePointHandler from "../api/xrf-sharepoint.js";

const originalFetch = globalThis.fetch;
const originalSecret = process.env.MXKPCS_CLIENT_SECRET;
const originalVercelEnv = process.env.VERCEL_ENV;

async function invoke(action, payload) {
  let output = "";
  const res = {
    statusCode: 200,
    setHeader() {},
    end(value = "") { output = value; },
  };
  await sharePointHandler({ method: "POST", url: "/api/xrf-sharepoint", headers: {}, body: { action, payload } }, res);
  return { status: res.statusCode, body: output ? JSON.parse(output) : null };
}

try {
  process.env.MXKPCS_CLIENT_SECRET = "test-secret";
  process.env.VERCEL_ENV = "development";
  let storedName = "";
  let patchedFields = null;
  let renamedFileName = "";

  globalThis.fetch = async (url, options = {}) => {
    const value = String(url);
    const method = String(options.method || "GET").toUpperCase();
    if (value.includes("/oauth2/v2.0/token")) {
      return new Response(JSON.stringify({ access_token: "fake-token" }), { status: 200 });
    }
    if (value.includes("/lists/651d2a7e-d6b9-43ce-9c44-2991682eb058/items/65?") && method === "GET") {
      return new Response(JSON.stringify({ id: "65", fields: { Title: "ITM_A_01", field_1: "A-01" } }), { status: 200 });
    }
    if (value.includes("/lists/ef7471b3-dea2-447f-a54a-254cddc7ed1a/drive") && method === "GET") {
      return new Response(JSON.stringify({ id: "photo-drive", driveType: "documentLibrary" }), { status: 200 });
    }
    if (value.includes("/drives/photo-drive/root:/ITEM_ITM_A_01") && method === "GET") {
      return new Response(JSON.stringify({ id: "photo-folder", name: "ITEM_ITM_A_01", folder: {} }), { status: 200 });
    }
    if (value.includes("/createUploadSession") && method === "POST") {
      storedName = JSON.parse(options.body).item.name;
      return new Response(JSON.stringify({ uploadUrl: "https://tenant.sharepoint.com/upload-session/photo-1" }), { status: 200 });
    }
    if (value === "https://tenant.sharepoint.com/upload-session/photo-1" && method === "PUT") {
      return new Response(JSON.stringify({ id: "photo-file-1" }), { status: 201 });
    }
    if (value.includes("/drives/photo-drive/items/photo-folder/children") && method === "GET") {
      return new Response(JSON.stringify({ value: [{
        id: "photo-file-1", name: storedName, size: 4,
        lastModifiedDateTime: "2026-09-22T01:00:00Z",
        webUrl: "https://tenant.sharepoint.com/photo/A-01.jpg",
        file: { mimeType: "image/jpeg" },
      }] }), { status: 200 });
    }
    if (value.includes("/drives/photo-drive/items/photo-file-1") && method === "PATCH") {
      renamedFileName = JSON.parse(options.body).name;
      storedName = renamedFileName;
      return new Response(JSON.stringify({
        id: "photo-file-1", name: storedName, size: 4,
        lastModifiedDateTime: "2026-09-22T01:00:00Z",
        webUrl: "https://tenant.sharepoint.com/photo/A-01.jpg",
        file: { mimeType: "image/jpeg" },
      }), { status: 200 });
    }
    if (value.includes("/lists/651d2a7e-d6b9-43ce-9c44-2991682eb058/columns") && method === "GET") {
      return new Response(JSON.stringify({ value: ["Photo_File_Name", "Photo_File_ID", "Photo_File_URL"].map(name => ({ name })) }), { status: 200 });
    }
    if (value.includes("/lists/651d2a7e-d6b9-43ce-9c44-2991682eb058/items/65/fields") && method === "PATCH") {
      patchedFields = JSON.parse(options.body);
      return new Response(JSON.stringify(patchedFields), { status: 200 });
    }
    throw new Error(`Unexpected fetch: ${method} ${value}`);
  };

  const begin = await invoke("beginItemPhotoUpload", { itemId: "ITM_A_01", itemSpItemId: "65", fileName: "A-01.jpg", size: 4 });
  assert.equal(begin.status, 200);
  assert.equal(begin.body.result.storedName, "A-01.jpg");
  const chunk = await invoke("uploadPrecisionFileChunk", {
    uploadUrl: begin.body.result.uploadUrl,
    offset: 0,
    totalSize: 4,
    base64: Buffer.from([1, 2, 3, 4]).toString("base64"),
  });
  assert.equal(chunk.body.result.driveItemId, "photo-file-1");
  const finish = await invoke("finishItemPhotoUpload", {
    itemId: "ITM_A_01",
    itemSpItemId: "65",
    storedName: begin.body.result.storedName,
    driveItemId: chunk.body.result.driveItemId,
  });
  assert.equal(finish.status, 200);
  assert.equal(finish.body.result.name, "A-01.jpg");
  assert.deepEqual(patchedFields, {
    Photo_File_Name: "A-01.jpg",
    Photo_File_ID: "photo-file-1",
    Photo_File_URL: "https://tenant.sharepoint.com/photo/A-01.jpg",
  });
  storedName = "photo_11111111-1111-4111-8111-111111111111_A-01.jpg";
  const normalize = await invoke("normalizeItemPhotoName", {
    itemId: "ITM_A_01", itemSpItemId: "65", driveItemId: "photo-file-1", fileName: "A-01.jpg",
  });
  assert.equal(normalize.status, 200);
  assert.equal(normalize.body.result.renamed, true);
  assert.equal(normalize.body.result.name, "A-01.jpg");
  assert.equal(renamedFileName, "A-01.jpg");
  console.log("item-photo-upload: PASS");
} finally {
  globalThis.fetch = originalFetch;
  if (originalSecret === undefined) delete process.env.MXKPCS_CLIENT_SECRET;
  else process.env.MXKPCS_CLIENT_SECRET = originalSecret;
  if (originalVercelEnv === undefined) delete process.env.VERCEL_ENV;
  else process.env.VERCEL_ENV = originalVercelEnv;
}
