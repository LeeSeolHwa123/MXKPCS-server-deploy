import assert from "node:assert/strict";
import handler from "../api/xrf-sharepoint.js";

const originalFetch = globalThis.fetch;
const originalEnvironment = Object.fromEntries(
  ["VERCEL_ENV", "VERCEL_TARGET_ENV", "VERCEL", "NODE_ENV", "MXKPCS_CLIENT_SECRET"]
    .map(name => [name, process.env[name]])
);

async function invoke(method, body = undefined) {
  let output = "";
  const headers = {};
  const response = {
    statusCode: 200,
    setHeader(name, value) { headers[name] = value; },
    end(value = "") { output = value; },
  };
  await handler({ method, body, url: "/api/xrf-sharepoint" }, response);
  return { status: response.statusCode, headers, body: output ? JSON.parse(output) : null };
}

try {
  process.env.VERCEL_ENV = "production";
  delete process.env.VERCEL_TARGET_ENV;
  delete process.env.VERCEL;
  delete process.env.NODE_ENV;
  delete process.env.MXKPCS_CLIENT_SECRET;
  let graphCalls = 0;
  globalThis.fetch = async () => { graphCalls += 1; throw new Error("Graph must not be called"); };

  const options = await invoke("OPTIONS");
  assert.equal(options.status, 204);
  assert.equal(options.headers.Allow, "GET, OPTIONS");
  for (const method of ["POST", "DELETE", "PATCH"]) {
    const denied = await invoke(method, { action: "writeProbe" });
    assert.equal(denied.status, 403);
    assert.equal(denied.body.ok, false);
  }
  assert.equal(graphCalls, 0);

  process.env.VERCEL_ENV = "preview";
  assert.equal((await invoke("POST", { action: "writeProbe" })).status, 403);
  assert.equal(graphCalls, 0);

  delete process.env.VERCEL_ENV;
  process.env.VERCEL = "1";
  process.env.NODE_ENV = "production";
  assert.equal((await invoke("POST", { action: "writeProbe" })).status, 403);
  assert.equal(graphCalls, 0);
  delete process.env.VERCEL;
  delete process.env.NODE_ENV;

  process.env.VERCEL_ENV = "production";
  process.env.MXKPCS_CLIENT_SECRET = "test-only-secret";
  globalThis.fetch = async (url, options = {}) => {
    graphCalls += 1;
    if (String(url).includes("/oauth2/v2.0/token")) {
      return new Response(JSON.stringify({ access_token: "fake-token" }), { status: 200 });
    }
    assert.equal(options.method, "GET");
    assert.match(String(url), /graph\.microsoft\.com\/v1\.0\/sites\//);
    return new Response(JSON.stringify({ value: [] }), { status: 200 });
  };
  const get = await invoke("GET");
  assert.equal(get.status, 200);
  assert.equal(get.body.ok, true);
  assert.equal(get.body.readOnly, true);
  assert.equal(Object.keys(get.body.lists).length, 6);
  assert.equal(graphCalls, 7);

  process.env.VERCEL_ENV = "development";
  const localOptions = await invoke("OPTIONS");
  assert.equal(localOptions.headers.Allow, "GET, POST, DELETE, OPTIONS");

  console.log("public-share-readonly: PASS");
} finally {
  globalThis.fetch = originalFetch;
  for (const [name, value] of Object.entries(originalEnvironment)) {
    if (value === undefined) delete process.env[name];
    else process.env[name] = value;
  }
}
