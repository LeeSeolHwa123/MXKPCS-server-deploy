import { mkdir, readFile, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import path from "node:path";
import process from "node:process";
import { promisify } from "node:util";
import sharePointHandler from "../api/xrf-sharepoint.js";

const projectRoot = path.resolve(import.meta.dirname, "..");
const execFileAsync = promisify(execFile);
const deployment = String(
  process.env.MXKPCS_VERCEL_DEPLOYMENT
  || "https://mxkpcs-pb05lfser-yah2685-2386.vercel.app",
).trim();
const allowedEnvironment = new Set([
  "MXKPCS_CLIENT_SECRET",
  "MXKPCS_TENANT",
  "MXKPCS_CLIENT_ID",
  "MXKPCS_SITE_ID",
  "MXKPCS_VERCEL_DEPLOYMENT",
  "VERCEL_OIDC_TOKEN",
]);

async function loadLocalEnvironment() {
  const envPath = path.join(projectRoot, ".env.local");
  const raw = await readFile(envPath, "utf8").catch(error => {
    if (error?.code === "ENOENT") return "";
    throw error;
  });
  for (const line of raw.split(/\r?\n/)) {
    const match = /^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/.exec(line);
    if (!match || !allowedEnvironment.has(match[1]) || process.env[match[1]]) continue;
    let value = match[2];
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    process.env[match[1]] = value;
  }
}

async function invokeGet(url) {
  if (!process.env.MXKPCS_CLIENT_SECRET) {
    const vercelCli = path.join(path.dirname(process.execPath), "node_modules", "vercel", "dist", "vc.js");
    const { stdout } = await execFileAsync(process.execPath, [
      vercelCli,
      "curl",
      url,
      "--deployment",
      deployment,
      "--yes",
    ], {
      cwd: projectRoot,
      env: process.env,
      windowsHide: true,
      maxBuffer: 50 * 1024 * 1024,
      encoding: "utf8",
    });
    const text = stdout.trim();
    const payload = JSON.parse(text.slice(text.indexOf("{"), text.lastIndexOf("}") + 1));
    if (!payload?.ok) throw new Error(payload?.error || `GET ${url} failed through Vercel`);
    return payload;
  }

  let output = "";
  const response = {
    statusCode: 200,
    setHeader() {},
    end(value = "") {
      output = Buffer.isBuffer(value) ? value.toString("utf8") : String(value);
    },
  };
  await sharePointHandler({ method: "GET", url, headers: {} }, response);
  const payload = output ? JSON.parse(output) : null;
  if (response.statusCode < 200 || response.statusCode >= 300 || !payload?.ok) {
    throw new Error(payload?.error || `GET ${url} failed with HTTP ${response.statusCode}`);
  }
  return payload;
}

async function mapLimit(values, limit, task) {
  const results = new Array(values.length);
  let cursor = 0;
  async function worker() {
    while (cursor < values.length) {
      const index = cursor++;
      results[index] = await task(values[index]);
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, values.length) }, worker));
  return results;
}

await loadLocalEnvironment();

const dateKey = new Intl.DateTimeFormat("en-CA", {
  timeZone: "Asia/Seoul",
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
}).format(new Date());
const outputDir = path.join(projectRoot, "outputs", `sharepoint_db_${dateKey}`);
const outputPath = path.join(outputDir, `MXKPCS_SharePoint_DB_${dateKey}.json`);
const listNames = [
  "XRF_Items",
  "XRF_Requests",
  "XRF_Measurements",
  "XRF_ElementResults",
  "XRF_PrecisionAnalyses",
  "XRF_PrecisionElementResults",
  "XRF_PrecisionExceptions",
  "XRF_ReportLibrary",
  "XRF_ItemPhotoLibrary",
];

const bootstrap = await invokeGet("/api/xrf-sharepoint");
const schemaEntries = await mapLimit(listNames, 3, async listName => {
  const [infoPayload, columnsPayload] = await Promise.all([
    invokeGet(`/api/xrf-sharepoint?listInfoFor=${encodeURIComponent(listName)}`),
    invokeGet(`/api/xrf-sharepoint?columnsFor=${encodeURIComponent(listName)}`),
  ]);
  return [listName, {
    list: infoPayload.list || null,
    columns: columnsPayload.columns || [],
  }];
});

const snapshot = {
  formatVersion: 1,
  extractedAt: bootstrap.generatedAt || new Date().toISOString(),
  siteId: bootstrap.siteId || null,
  source: "SharePoint via api/xrf-sharepoint.js",
  scope: {
    listRows: Object.keys(bootstrap.lists || {}),
    schema: listNames,
    documentBinariesIncluded: false,
    credentialsIncluded: false,
  },
  counts: bootstrap.counts || {},
  schemas: Object.fromEntries(schemaEntries),
  lists: bootstrap.lists || {},
};

await mkdir(outputDir, { recursive: true });
await writeFile(outputPath, `${JSON.stringify(snapshot, null, 2)}\n`, "utf8");

console.log(JSON.stringify({
  outputPath,
  extractedAt: snapshot.extractedAt,
  counts: snapshot.counts,
  schemaCount: Object.keys(snapshot.schemas).length,
  credentialsIncluded: false,
  documentBinariesIncluded: false,
}, null, 2));
