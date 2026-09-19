import { checkWriteCode, clearWriteSessionCookie, verifyWriteSession, writeSessionCookie } from "./write-access-auth.js";

function sendJson(res, status, body) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store, max-age=0");
  res.end(JSON.stringify(body));
}

async function parseBody(req) {
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.body === "string") return JSON.parse(req.body || "{}");
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > 4096) throw new Error("요청이 너무 큽니다.");
    chunks.push(chunk);
  }
  return JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
}

export default async function handler(req, res) {
  if (req.method === "GET") {
    const session = verifyWriteSession(req);
    return sendJson(res, 200, { ok: true, canWrite: !!session, email: session?.email || null });
  }
  if (req.method !== "POST") {
    res.setHeader("Allow", "GET, POST");
    return sendJson(res, 405, { ok: false, error: "Method Not Allowed" });
  }
  const origin = req.headers?.origin;
  const host = req.headers?.host;
  if (origin && host) {
    const expected = `${req.headers?.["x-forwarded-proto"] === "https" || process.env.VERCEL === "1" ? "https" : "http"}://${host}`;
    if (origin !== expected) return sendJson(res, 403, { ok: false, error: "요청 출처를 확인할 수 없습니다." });
  }
  try {
    const body = await parseBody(req);
    if (body.action === "logout") {
      res.setHeader("Set-Cookie", clearWriteSessionCookie(req));
      return sendJson(res, 200, { ok: true, canWrite: false, email: null });
    }
    if (body.action !== "login") return sendJson(res, 400, { ok: false, error: "지원하지 않는 요청입니다." });
    const identity = checkWriteCode(body.email, body.code);
    if (!identity) return sendJson(res, 401, { ok: false, error: "계정 또는 접근 코드가 올바르지 않습니다." });
    res.setHeader("Set-Cookie", writeSessionCookie(identity, req));
    return sendJson(res, 200, { ok: true, canWrite: true, email: identity.email });
  } catch {
    return sendJson(res, 400, { ok: false, error: "인증 요청을 처리할 수 없습니다." });
  }
}
