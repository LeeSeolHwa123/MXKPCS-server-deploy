// Self-host entry point for Windows Server.
// Serves the Vite build output (dist/) and forwards /api/xrf-sharepoint to the
// same handler used on Vercel — no framework, so the handler needs no changes.
import http from "node:http";
import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import apiHandler from "./api/xrf-sharepoint.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DIST_DIR = path.join(__dirname, "dist");

function loadEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return;
  for (const line of fs.readFileSync(filePath, "utf8").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    let value = trimmed.slice(eq + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    if (!(key in process.env)) process.env[key] = value;
  }
}

loadEnvFile(process.env.ENV_FILE || path.join(__dirname, "deployment", "MXKPCS_Server_Env"));

const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || "0.0.0.0";

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
  ".ico": "image/x-icon",
  ".woff": "font/woff",
  ".woff2": "font/woff2",
  ".txt": "text/plain; charset=utf-8",
  ".map": "application/json; charset=utf-8",
};

async function serveStatic(req, res, pathname) {
  const safePath = path.normalize(pathname).replace(/^(\.\.[/\\])+/, "");
  const hasExtension = path.extname(pathname) !== "";
  let targetPath = hasExtension && pathname !== "/"
    ? path.join(DIST_DIR, safePath)
    : path.join(DIST_DIR, "index.html");

  if (!targetPath.startsWith(DIST_DIR)) {
    res.statusCode = 403;
    return res.end("Forbidden");
  }

  try {
    const stat = await fsp.stat(targetPath);
    if (stat.isDirectory()) targetPath = path.join(targetPath, "index.html");
    const ext = path.extname(targetPath).toLowerCase();
    res.statusCode = 200;
    res.setHeader("Content-Type", MIME_TYPES[ext] || "application/octet-stream");
    res.setHeader(
      "Cache-Control",
      ext === ".html" ? "no-cache" : "public, max-age=31536000, immutable"
    );
    fs.createReadStream(targetPath).pipe(res);
  } catch {
    res.statusCode = 404;
    res.end("Not found");
  }
}

const server = http.createServer(async (req, res) => {
  try {
    const pathname = new URL(req.url || "/", "http://localhost").pathname;

    if (pathname === "/api/xrf-sharepoint") {
      await apiHandler(req, res);
      return;
    }

    if (req.method !== "GET" && req.method !== "HEAD") {
      res.statusCode = 405;
      res.end("Method Not Allowed");
      return;
    }

    await serveStatic(req, res, pathname);
  } catch (error) {
    console.error("[server] request failed:", error);
    if (!res.headersSent) {
      res.statusCode = 500;
      res.setHeader("Content-Type", "application/json; charset=utf-8");
    }
    res.end(JSON.stringify({ ok: false, error: "Internal Server Error" }));
  }
});

server.listen(PORT, HOST, () => {
  console.log(`[server] MXKPCS listening on http://${HOST}:${PORT}`);
});

process.on("SIGINT", () => server.close(() => process.exit(0)));
process.on("SIGTERM", () => server.close(() => process.exit(0)));
