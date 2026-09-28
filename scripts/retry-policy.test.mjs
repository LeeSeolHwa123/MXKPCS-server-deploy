import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const apiUrl = new URL("../api/xrf-sharepoint.js", import.meta.url);
const source = await readFile(apiUrl, "utf8");
const instrumented = source
  .replace(
    "function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }",
    "function sleep() { return Promise.resolve(); }",
  )
  + "\nfunction resetGraphTokenCache(){ graphTokenCache = { key: '', token: '', expiresAt: 0 }; }\n"
  + "export { fetchWithRetry, acquireGraphToken, resetGraphTokenCache };\n";
const policy = await import(`data:text/javascript;base64,${Buffer.from(instrumented).toString("base64")}`);

const originalFetch = globalThis.fetch;
const originalSecret = process.env.MXKPCS_CLIENT_SECRET;

function response(status, body = "") {
  return new Response(body, {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function fakeFetch(steps) {
  let calls = 0;
  const fetch = async () => {
    const step = steps[calls];
    calls += 1;
    if (!step) throw new Error(`Unexpected fake fetch call ${calls}`);
    if (step instanceof Error) throw step;
    return step;
  };
  return { fetch, calls: () => calls };
}

async function expectSuccess(name, { method = "GET", steps, fetchPolicy, expectedCalls }) {
  const fake = fakeFetch(steps);
  globalThis.fetch = fake.fetch;
  const result = await policy.fetchWithRetry("https://graph.test/resource", { method }, fetchPolicy);
  assert.equal(result.status, 200, `${name}: response status`);
  assert.equal(fake.calls(), expectedCalls, `${name}: fetch calls`);
  console.log(`${name}: PASS`);
}

async function expectError(name, { method = "GET", steps, fetchPolicy, expectedCalls, expectedStatus }) {
  const fake = fakeFetch(steps);
  globalThis.fetch = fake.fetch;
  await assert.rejects(
    () => policy.fetchWithRetry("https://graph.test/resource", { method }, fetchPolicy),
    error => expectedStatus == null || error?.status === expectedStatus,
    `${name}: expected error`,
  );
  assert.equal(fake.calls(), expectedCalls, `${name}: fetch calls`);
  console.log(`${name}: PASS`);
}

async function expectTokenSuccess(name, steps, expectedCalls) {
  policy.resetGraphTokenCache();
  const fake = fakeFetch(steps);
  globalThis.fetch = fake.fetch;
  process.env.MXKPCS_CLIENT_SECRET = "retry-policy-test-only";
  const token = await policy.acquireGraphToken();
  assert.equal(token, "test-token", `${name}: token response`);
  assert.equal(fake.calls(), expectedCalls, `${name}: fetch calls`);
  console.log(`${name}: PASS`);
}

try {
  await expectSuccess("R01", { steps: [response(200)], expectedCalls: 1 });
  await expectSuccess("R02", { steps: [response(500, "server error"), response(200)], expectedCalls: 2 });
  await expectSuccess("R03", { steps: [new Error("network error"), response(200)], expectedCalls: 2 });
  await expectError("R04", { steps: [response(400, "bad request")], expectedCalls: 1, expectedStatus: 400 });
  await expectError("R05", { method: "POST", steps: [response(500, "server error")], expectedCalls: 1, expectedStatus: 500 });
  await expectError("R06", { method: "POST", steps: [new Error("network error")], expectedCalls: 1 });
  await expectError("R07", { method: "POST", steps: [response(400, "bad request")], expectedCalls: 1, expectedStatus: 400 });
  await expectSuccess("R08", { method: "POST", steps: [response(429, "throttled"), response(200)], expectedCalls: 2 });
  await expectError("R09", { method: "PATCH", steps: [response(500, "server error")], expectedCalls: 1, expectedStatus: 500 });
  await expectSuccess("R10", { method: "PATCH", steps: [response(429, "throttled"), response(200)], expectedCalls: 2 });
  await expectTokenSuccess("R11", [response(500, "server error"), response(200, '{"access_token":"test-token"}')], 2);
  await expectTokenSuccess("R12", [new Error("network error"), response(200, '{"access_token":"test-token"}')], 2);
  console.log("retry-policy: PASS");
} finally {
  globalThis.fetch = originalFetch;
  if (originalSecret === undefined) delete process.env.MXKPCS_CLIENT_SECRET;
  else process.env.MXKPCS_CLIENT_SECRET = originalSecret;
}
