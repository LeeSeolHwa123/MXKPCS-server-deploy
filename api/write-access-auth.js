import { createHash, createHmac, timingSafeEqual } from "node:crypto";

const COOKIE_NAME = "mxk_write_access";
const SESSION_SECONDS = 12 * 60 * 60;

function approvedHashes() {
  const entries = new Map();
  for (const entry of String(process.env.MXKPCS_WRITE_ACCESS_HASHES || "").split(/[\r\n,;]+/)) {
    const separator = entry.indexOf(":");
    if (separator < 1) continue;
    const email = entry.slice(0, separator).trim().toLowerCase();
    const hash = entry.slice(separator + 1).trim().toLowerCase();
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) && /^[a-f0-9]{64}$/.test(hash)) {
      entries.set(email, hash);
    }
  }
  return entries;
}

function sessionKey() {
  const secret = process.env.MXKPCS_CLIENT_SECRET;
  if (!secret) return null;
  return createHmac("sha256", secret).update("mxkpcs/write-access/session/v1").digest();
}

function signature(payload, key) {
  return createHmac("sha256", key).update(payload).digest("base64url");
}

function equalText(a, b) {
  const left = Buffer.from(String(a));
  const right = Buffer.from(String(b));
  return left.length === right.length && timingSafeEqual(left, right);
}

function cookieValue(req) {
  const raw = String(req.headers?.cookie || "");
  const pair = raw.split(";").map(part => part.trim()).find(part => part.startsWith(`${COOKIE_NAME}=`));
  return pair ? pair.slice(COOKIE_NAME.length + 1) : "";
}

export function verifyWriteSession(req) {
  const key = sessionKey();
  const value = cookieValue(req);
  if (!key || !value) return null;
  const [encoded, mac, extra] = value.split(".");
  if (!encoded || !mac || extra || !equalText(mac, signature(encoded, key))) return null;
  try {
    const session = JSON.parse(Buffer.from(encoded, "base64url").toString("utf8"));
    if (session.v !== 1 || !Number.isInteger(session.exp) || session.exp <= Date.now()) return null;
    const email = String(session.email || "").toLowerCase();
    if (session.accessHash !== approvedHashes().get(email)) return null;
    return { email };
  } catch {
    return null;
  }
}

export function checkWriteCode(emailInput, codeInput) {
  const email = String(emailInput || "").trim().toLowerCase();
  const code = String(codeInput || "").trim();
  const expected = approvedHashes().get(email);
  if (!expected || !/^[A-Za-z0-9_-]{40,100}$/.test(code) || !sessionKey()) return null;
  const actual = createHash("sha256").update(code).digest("hex");
  return equalText(actual, expected) ? email : null;
}

export function writeSessionCookie(email, req) {
  const key = sessionKey();
  if (!key) throw new Error("쓰기 권한 인증이 설정되지 않았습니다.");
  const encoded = Buffer.from(JSON.stringify({ v: 1, email, accessHash: approvedHashes().get(email), exp: Date.now() + SESSION_SECONDS * 1000 })).toString("base64url");
  const secure = req.headers?.["x-forwarded-proto"] === "https" || process.env.VERCEL === "1";
  return `${COOKIE_NAME}=${encoded}.${signature(encoded, key)}; HttpOnly; SameSite=Strict; Path=/api; Max-Age=${SESSION_SECONDS}${secure ? "; Secure" : ""}`;
}

export function clearWriteSessionCookie(req) {
  const secure = req.headers?.["x-forwarded-proto"] === "https" || process.env.VERCEL === "1";
  return `${COOKIE_NAME}=; HttpOnly; SameSite=Strict; Path=/api; Max-Age=0${secure ? "; Secure" : ""}`;
}
