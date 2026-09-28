import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import sharePointHandler from "../api/xrf-sharepoint.js";
import accessHandler from "../api/write-access.js";

const originalFetch = globalThis.fetch;
const names = ["VERCEL_ENV", "VERCEL_TARGET_ENV", "VERCEL", "NODE_ENV", "MXKPCS_CLIENT_SECRET", "MXKPCS_WRITE_ACCESS_HASHES"];
const originalEnvironment = Object.fromEntries(names.map(name => [name, process.env[name]]));

async function invoke(handler, method, body = undefined, cookie = "") {
  let output = "";
  const headers = {};
  const response = {
    statusCode: 200,
    setHeader(name, value) { headers[name] = value; },
    end(value = "") { output = value; },
  };
  await handler({ method, body, url: "/api/test", headers: cookie ? { cookie } : {} }, response);
  return { status: response.statusCode, headers, body: output ? JSON.parse(output) : null };
}

try {
  process.env.VERCEL_ENV = "production";
  delete process.env.VERCEL_TARGET_ENV;
  delete process.env.VERCEL;
  delete process.env.NODE_ENV;
  process.env.MXKPCS_CLIENT_SECRET = "test-only-secret";
  const code = "A".repeat(43);
  const codeHash = createHash("sha256").update(code).digest("hex");
  const email = "newly-approved@example.com";
  process.env.MXKPCS_WRITE_ACCESS_HASHES = `old-account@example.com:${codeHash}`;

  let graphCalls = 0;
  globalThis.fetch = async () => { graphCalls += 1; throw new Error("Graph must not be called"); };

  const anonymousOptions = await invoke(sharePointHandler, "OPTIONS");
  assert.equal(anonymousOptions.status, 204);
  assert.equal(anonymousOptions.headers.Allow, "GET, POST, OPTIONS");
  for (const method of ["POST", "DELETE", "PATCH"]) {
    const denied = await invoke(sharePointHandler, method, { action: "writeProbe" });
    assert.equal(denied.status, 403);
  }
  assert.equal(graphCalls, 0);

  const wrong = await invoke(accessHandler, "POST", { action: "login", email, code: "B".repeat(43) });
  assert.equal(wrong.status, 401);
  const login = await invoke(accessHandler, "POST", { action: "login", email, code });
  assert.equal(login.status, 200);
  assert.equal(login.body.canWrite, true);
  assert.equal(login.body.email, email);
  const cookie = login.headers["Set-Cookie"].split(";")[0];
  assert.ok(cookie.startsWith("mxk_write_access="));
  const another = await invoke(accessHandler, "POST", { action: "login", email: "another@example.com", code });
  assert.equal(another.status, 200);
  assert.equal(another.body.email, "another@example.com");
  const limitedEmail = "rate-limit@example.com";
  for (let attempt = 0; attempt < 5; attempt += 1) {
    assert.equal((await invoke(accessHandler, "POST", { action: "login", email: limitedEmail, code: "B".repeat(43) })).status, 401);
  }
  const limited = await invoke(accessHandler, "POST", { action: "login", email: limitedEmail, code });
  assert.equal(limited.status, 429);
  assert.ok(Number(limited.headers["Retry-After"]) > 0);
  process.env.MXKPCS_WRITE_ACCESS_HASHES = codeHash;
  assert.equal((await invoke(accessHandler, "GET", undefined, cookie)).body.canWrite, true);
  assert.equal((await invoke(accessHandler, "GET", undefined, `${cookie}tampered`)).body.canWrite, false);

  for (const [method, action] of [["DELETE", "deleteRecord"], ["POST", "deleteRecord"], ["POST", "writeProbe"]]) {
    const denied = await invoke(sharePointHandler, method, { action }, cookie);
    assert.equal(denied.status, 403);
  }
  assert.equal(graphCalls, 0);

  globalThis.fetch = async (url, options = {}) => {
    graphCalls += 1;
    if (String(url).includes("/oauth2/v2.0/token")) {
      return new Response(JSON.stringify({ access_token: "fake-token" }), { status: 200 });
    }
    assert.equal(options.method, "GET");
    if (String(url).includes("/drives/photo-drive/root/children")) {
      return new Response(JSON.stringify({ value: [] }), { status: 200 });
    }
    assert.match(String(url), /graph\.microsoft\.com\/v1\.0\/sites\//);
    if (String(url).includes("/lists/ef7471b3-dea2-447f-a54a-254cddc7ed1a/drive")) {
      return new Response(JSON.stringify({ id: "photo-drive", driveType: "documentLibrary" }), { status: 200 });
    }
    return new Response(JSON.stringify({ value: [] }), { status: 200 });
  };
  const anonymousGet = await invoke(sharePointHandler, "GET");
  assert.equal(anonymousGet.status, 200);
  assert.equal(anonymousGet.body.readOnly, true);
  const approvedGet = await invoke(sharePointHandler, "GET", undefined, cookie);
  assert.equal(approvedGet.status, 200);
  assert.equal(approvedGet.body.readOnly, false);
  assert.equal(approvedGet.body.writeAccessEmail, email);
  assert.equal(Object.keys(approvedGet.body.lists).length, 6);
  // Item photo bytes are fetched lazily and the second bootstrap reuses the cached OAuth token.
  assert.equal(graphCalls, 13);

  process.env.MXKPCS_WRITE_ACCESS_HASHES = "f".repeat(64);
  assert.equal((await invoke(accessHandler, "GET", undefined, cookie)).body.canWrite, false);
  assert.equal((await invoke(sharePointHandler, "POST", { action: "createRequest" }, cookie)).status, 403);
  process.env.MXKPCS_WRITE_ACCESS_HASHES = codeHash;
  const logout = await invoke(accessHandler, "POST", { action: "logout" }, cookie);
  assert.equal(logout.status, 200);
  assert.match(logout.headers["Set-Cookie"], /Max-Age=0/);

  process.env.VERCEL_ENV = "development";
  assert.equal((await invoke(sharePointHandler, "OPTIONS")).headers.Allow, "GET, POST, DELETE, OPTIONS");

  console.log("public-share-access: PASS");
} finally {
  globalThis.fetch = originalFetch;
  for (const [name, value] of Object.entries(originalEnvironment)) {
    if (value === undefined) delete process.env[name];
    else process.env[name] = value;
  }
}
