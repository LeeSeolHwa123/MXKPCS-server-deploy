import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import sharePointHandler from "../api/xrf-sharepoint.js";

const originalFetch = globalThis.fetch;
const names = ["VERCEL_ENV", "VERCEL_TARGET_ENV", "VERCEL", "NODE_ENV", "MXKPCS_CLIENT_SECRET"];
const originalEnvironment = Object.fromEntries(names.map(name => [name, process.env[name]]));

async function invoke(method, body = undefined) {
  let output = "";
  const headers = {};
  const response = {
    statusCode: 200,
    setHeader(name, value) { headers[name] = value; },
    end(value = "") { output = value; },
  };
  await sharePointHandler({ method, body, url: "/api/test", headers: {} }, response);
  return { status: response.statusCode, headers, body: output ? JSON.parse(output) : null };
}

try {
  process.env.VERCEL_ENV = "production";
  delete process.env.VERCEL_TARGET_ENV;
  process.env.VERCEL = "1";
  process.env.NODE_ENV = "production";
  process.env.MXKPCS_CLIENT_SECRET = "test-only-secret";

  globalThis.fetch = async (url, options = {}) => {
    const method = String(options.method || "GET").toUpperCase();
    if (String(url).includes("/oauth2/v2.0/token")) {
      return new Response(JSON.stringify({ access_token: "fake-token" }), { status: 200 });
    }
    if (String(url).includes("/lists/ef7471b3-dea2-447f-a54a-254cddc7ed1a/drive")) {
      return new Response(JSON.stringify({ id: "photo-drive", driveType: "documentLibrary" }), { status: 200 });
    }
    if (String(url).includes("/drives/photo-drive/root/children")) {
      return new Response(JSON.stringify({ value: [] }), { status: 200 });
    }
    if (method === "GET") return new Response(JSON.stringify({ value: [] }), { status: 200 });
    if (method === "POST" && String(url).includes("/items")) {
      return new Response(JSON.stringify({ id: "probe-row", fields: {} }), { status: 201 });
    }
    if (method === "DELETE") return new Response(null, { status: 204 });
    throw new Error(`Unexpected mock Graph request: ${method} ${url}`);
  };

  const options = await invoke("OPTIONS");
  assert.equal(options.status, 204);
  assert.equal(options.headers.Allow, "GET, POST, DELETE, OPTIONS");

  const bootstrap = await invoke("GET");
  assert.equal(bootstrap.status, 200);
  assert.equal(bootstrap.body.ok, true);
  assert.equal(bootstrap.body.readOnly, false);
  assert.equal(bootstrap.body.authRequired, false);
  assert.equal(bootstrap.body.writeAccessEmail, null);

  const writeProbe = await invoke("POST", { action: "writeProbe", payload: {} });
  assert.equal(writeProbe.status, 200);
  assert.equal(writeProbe.body.ok, true);

  const appSource = await readFile(new URL("../app.jsx", import.meta.url), "utf8");
  assert.doesNotMatch(appSource, /WRITE_ACCESS_API|writeAccessForm|쓰기 권한 로그인|Write access/);
  assert.match(appSource, /precision-handoff-auth-title/);

  console.log("public-write-access: PASS");
} finally {
  globalThis.fetch = originalFetch;
  for (const [name, value] of Object.entries(originalEnvironment)) {
    if (value === undefined) delete process.env[name];
    else process.env[name] = value;
  }
}
