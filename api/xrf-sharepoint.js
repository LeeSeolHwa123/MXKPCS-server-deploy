// Node.js SharePoint API Handler
// SharePoint XRF DB READ + WRITE gateway
// Client secret must exist only in server-side environment variables.
import { randomUUID } from "node:crypto";
import { verifyWriteSession } from "./write-access-auth.js";

const GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
const DEFAULT_TENANT = "kochind.com";
const DEFAULT_CLIENT_ID = "bdaa0272-4761-4a5a-834f-09c15aa79e64";
const DEFAULT_SITE_ID = "kochind.sharepoint.com,094129d8-9492-4d1d-9223-c716e28317a9,c79c6d5a-5dab-47fb-bc5e-43a44c51c84a";

const LIST_IDS = {
  XRF_Items: "651d2a7e-d6b9-43ce-9c44-2991682eb058",
  XRF_Requests: "187a06a8-5a0a-40d0-9164-8cb45372ef32",
  XRF_Measurements: "dd0bf876-6066-4877-be6d-eb10eb5ce35c",
  XRF_ElementResults: "b0a877ee-5334-4237-9b8f-e94235b118a4",
  XRF_PrecisionAnalyses: "3c387146-a8c5-4b94-9bf3-251452b19a1c",
  XRF_PrecisionElementResults: "3fc979f1-4d6c-4d8a-9c75-3cab099d9971",
  XRF_ReportLibrary: "0462cdd6-f3ce-40bf-a814-471796a6147b",
  XRF_ItemPhotoLibrary: "ef7471b3-dea2-447f-a54a-254cddc7ed1a",
};

const WRITABLE_LISTS = new Set([
  "XRF_Items",
  "XRF_Requests",
  "XRF_Measurements",
  "XRF_ElementResults",
  "XRF_PrecisionAnalyses",
  "XRF_PrecisionElementResults",
]);

const XRF_SOURCE_ELEMENTS = ["Pb", "Hg", "Cr", "Cd", "Cl", "Br"];
const MAX_RETRIES = 4;
const RETRY_BASE_MS = 350;
const PRECISION_FILE_CHUNK_BYTES = 8 * 320 * 1024; // 2.5 MiB; base64 JSON stays below Vercel's 4.5 MB request limit.
const PRECISION_FILE_MAX_BYTES = 250 * 1024 * 1024;
const ITEM_PHOTO_MAX_BYTES = 10 * 1024 * 1024;
const FILE_REFERENCE_COLUMNS = {
  XRF_Items: [
    { name: "Photo_File_Name", displayName: "Item Photo File Name", text: { allowMultipleLines: false, maxLength: 255 } },
    { name: "Photo_File_ID", displayName: "Item Photo File ID", text: { allowMultipleLines: false, maxLength: 255 } },
    { name: "Photo_File_URL", displayName: "Item Photo File URL", text: { allowMultipleLines: true, linesForEditing: 3 } },
  ],
  XRF_Requests: [
    { name: "Requester_Name", displayName: "Requester Name", text: { allowMultipleLines: false, maxLength: 255 } },
  ],
  XRF_Measurements: [
    { name: "Source_File_ID", displayName: "XRF Source File ID", text: { allowMultipleLines: false, maxLength: 255 } },
    { name: "Source_File_URL", displayName: "XRF Source File URL", text: { allowMultipleLines: true, linesForEditing: 3 } },
    { name: "Requester_Name", displayName: "Requester Name", text: { allowMultipleLines: false, maxLength: 255 } },
  ],
  XRF_PrecisionAnalyses: [
    { name: "Report_File_ID", displayName: "Precision Report File ID", text: { allowMultipleLines: false, maxLength: 255 } },
    { name: "Report_File_URL", displayName: "Precision Report File URL", text: { allowMultipleLines: true, linesForEditing: 3 } },
  ],
};
let graphTokenCache = { key: "", token: "", expiresAt: 0 };

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
function cleanText(v) { return v == null ? "" : String(v); }
function nonEmpty(v) { const s = cleanText(v).trim(); return s ? s : null; }
function dateOnly(v) { const s = cleanText(v).trim(); return s ? s.slice(0, 10) : null; }
function yyyymmdd(v) { return cleanText(v).slice(0, 10).replace(/-/g, ""); }
function unique(arr) { return [...new Set(arr.filter(Boolean))]; }
function bool(v) { return v === true || v === 1 || v === "1" || String(v).toLowerCase() === "true"; }
function createFields(obj) {
  return Object.fromEntries(Object.entries(obj || {}).filter(([, value]) => value !== null && value !== undefined));
}

function sendJson(res, status, body) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.setHeader("Cache-Control", "no-store, max-age=0");
  res.end(JSON.stringify(body));
}

async function fetchWithRetry(url, options = {}, policy = {}) {
  const method = String(options?.method || "GET").toUpperCase();
  const isRead = method === "GET" || method === "HEAD" || method === "OPTIONS";
  const retry429 = policy.retry429 ?? true;
  const retry5xx = policy.retry5xx ?? isRead;
  const retryNetwork = policy.retryNetwork ?? isRead;
  let lastError = null;

  for (let attempt = 0; attempt <= MAX_RETRIES; attempt += 1) {
    let response;
    try {
      response = await fetch(url, options);
    } catch (error) {
      lastError = error;
      if (!retryNetwork || attempt === MAX_RETRIES) throw error;
      await sleep(RETRY_BASE_MS * Math.pow(2, attempt));
      continue;
    }

    if (response.ok) return response;

    const text = await response.text();
    const retryable = (response.status === 429 && retry429)
      || (response.status >= 500 && retry5xx);
    if (!retryable || attempt === MAX_RETRIES) {
      const error = new Error(`Graph HTTP ${response.status}: ${text.slice(0, 1800)}`);
      error.status = response.status;
      throw error;
    }

    const retryAfter = Number(response.headers.get("retry-after"));
    await sleep(Number.isFinite(retryAfter) && retryAfter > 0
      ? retryAfter * 1000
      : RETRY_BASE_MS * Math.pow(2, attempt));
  }
  throw lastError || new Error("Graph request failed");
}

async function acquireGraphToken() {
  const tenant = process.env.MXKPCS_TENANT || DEFAULT_TENANT;
  const clientId = process.env.MXKPCS_CLIENT_ID || DEFAULT_CLIENT_ID;
  const clientSecret = process.env.MXKPCS_CLIENT_SECRET;
  if (!clientSecret) throw new Error("서버 환경변수 MXKPCS_CLIENT_SECRET가 설정되어 있지 않습니다.");
  const cacheKey = `${tenant}|${clientId}|${clientSecret.length}`;
  if (graphTokenCache.key === cacheKey && graphTokenCache.token && Date.now() < graphTokenCache.expiresAt) {
    return graphTokenCache.token;
  }

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const response = await fetchWithRetry(
    `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body },
    { retry429: true, retry5xx: true, retryNetwork: true }
  );
  const payload = await response.json();
  if (!payload?.access_token) throw new Error("Microsoft Graph access token을 받지 못했습니다.");
  graphTokenCache = {
    key: cacheKey,
    token: payload.access_token,
    expiresAt: Date.now() + Math.max(60, Number(payload.expires_in || 3600) - 120) * 1000,
  };
  return payload.access_token;
}

function graphHeaders(accessToken, extra = {}) {
  return { Authorization: `Bearer ${accessToken}`, Accept: "application/json", ...extra };
}

function precisionFileFolderName(measurementId) {
  const id = cleanText(measurementId).trim();
  if (!/^[A-Za-z0-9_-]{1,120}$/.test(id)) throw new Error("유효하지 않은 XRF 측정 ID입니다.");
  return `PREC_${id}`;
}

function precisionUploadName(kind, fileName) {
  if (kind !== "source" && kind !== "report") throw new Error("지원하지 않는 파일 종류입니다.");
  const name = cleanText(fileName).trim();
  if (!name || name.length > 160 || /[\\/:*?"<>|\x00-\x1f]/.test(name) || name === "." || name === "..") {
    throw new Error("파일명을 확인하세요. 경로 문자와 160자를 넘는 이름은 사용할 수 없습니다.");
  }
  if (kind === "report" && !/\.pdf$/i.test(name)) throw new Error("정밀분석 성적서는 PDF만 업로드할 수 있습니다.");
  if (kind === "source" && !/\.(xlsx|xlsm|xls|csv|json|pdf)$/i.test(name)) {
    throw new Error("XRF 원본은 xlsx, xlsm, xls, csv, json, pdf 형식만 허용합니다.");
  }
  return `${kind}_${randomUUID()}_${name}`;
}

function precisionFileMetadata(item) {
  const match = /^(source|report)_[0-9a-f-]{36}_(.+)$/i.exec(cleanText(item?.name));
  if (!match || !item?.file) return null;
  return {
    id: item.id, kind: match[1].toLowerCase(), name: match[2],
    size: Number(item.size || 0), uploadedAt: item.lastModifiedDateTime || null,
    webUrl: item.webUrl || null,
  };
}

function itemPhotoFolderName(itemId) {
  const id = cleanText(itemId).trim();
  if (!/^[A-Za-z0-9_-]{1,120}$/.test(id)) throw new Error("유효하지 않은 품목 ID입니다.");
  return `ITEM_${id}`;
}

function itemPhotoUploadName(fileName) {
  const name = cleanText(fileName).trim();
  if (!name || name.length > 160 || /[\\/:*?"<>|\x00-\x1f]/.test(name) || name === "." || name === "..") {
    throw new Error("사진 파일명을 확인하세요. 경로 문자와 160자를 넘는 이름은 사용할 수 없습니다.");
  }
  if (!/\.(jpe?g|png|webp)$/i.test(name)) throw new Error("품목 사진은 JPG, PNG, WEBP 형식만 업로드할 수 있습니다.");
  return name;
}

function itemPhotoMetadata(item, itemId = "") {
  const storedName = cleanText(item?.name).trim();
  const legacyMatch = /^photo_[0-9a-f-]{36}_(.+)$/i.exec(storedName);
  const displayName = legacyMatch?.[1] || storedName;
  if (!item?.file || !/\.(jpe?g|png|webp)$/i.test(displayName)) return null;
  return {
    itemId,
    id: item.id,
    name: displayName,
    storedName,
    size: Number(item.size || 0),
    mimeType: item.file?.mimeType || "application/octet-stream",
    uploadedAt: item.lastModifiedDateTime || null,
    webUrl: item.webUrl || null,
  };
}

function validateUploadSessionUrl(raw) {
  const url = new URL(cleanText(raw));
  const host = url.hostname.toLowerCase();
  if (url.protocol !== "https:" || url.username || url.password || url.port
    || !(host.endsWith(".sharepoint.com") || host.endsWith(".1drv.com"))) {
    throw new Error("유효하지 않은 Graph 업로드 세션 URL입니다.");
  }
  return url.href;
}

async function precisionDriveId(siteId, accessToken) {
  const response = await fetchWithRetry(`${listUrl(siteId, "XRF_ReportLibrary")}/drive`, {
    method: "GET", headers: graphHeaders(accessToken),
  });
  const drive = await response.json();
  if (!drive?.id || drive.driveType !== "documentLibrary") throw new Error("XRF_ReportLibrary 문서 라이브러리를 찾지 못했습니다.");
  return drive.id;
}

async function itemPhotoDriveId(siteId, accessToken) {
  const response = await fetchWithRetry(`${listUrl(siteId, "XRF_ItemPhotoLibrary")}/drive`, {
    method: "GET", headers: graphHeaders(accessToken),
  });
  const drive = await response.json();
  if (!drive?.id || drive.driveType !== "documentLibrary") throw new Error("XRF_ItemPhotoLibrary 문서 라이브러리를 찾지 못했습니다.");
  return drive.id;
}

async function precisionFolder(driveId, measurementId, accessToken, create = false) {
  const folderName = precisionFileFolderName(measurementId);
  const root = `${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/root`;
  const path = `${root}:/${encodeURIComponent(folderName)}`;
  const find = async () => {
    const response = await fetch(path, { method: "GET", headers: graphHeaders(accessToken) });
    if (response.status === 404) return null;
    if (!response.ok) throw new Error(`문서 폴더 조회 실패: HTTP ${response.status}`);
    const item = await response.json();
    if (!item?.folder) throw new Error("정밀분석 문서 경로가 폴더가 아닙니다.");
    return item;
  };
  let folder = await find();
  if (folder || !create) return folder;
  const response = await fetch(`${root}/children`, {
    method: "POST", headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
    body: JSON.stringify({ name: folderName, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
  });
  if (response.status === 409) folder = await find();
  else if (response.ok) folder = await response.json();
  else throw new Error(`정밀분석 문서 폴더 생성 실패: HTTP ${response.status}`);
  if (!folder?.id) throw new Error("정밀분석 문서 폴더 ID를 받지 못했습니다.");
  return folder;
}

async function precisionFolderFiles(driveId, folder, accessToken) {
  if (!folder?.id) return [];
  return (await graphGetAll(`${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folder.id)}/children?$top=200`, accessToken))
    .map(item => ({ item, metadata: precisionFileMetadata(item) }))
    .filter(row => row.metadata);
}

async function namedDriveFolder(driveId, folderName, accessToken, create = false) {
  const root = `${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/root`;
  const path = `${root}:/${encodeURIComponent(folderName)}`;
  const find = async () => {
    const response = await fetch(path, { method: "GET", headers: graphHeaders(accessToken) });
    if (response.status === 404) return null;
    if (!response.ok) throw new Error(`문서 폴더 조회 실패: HTTP ${response.status}`);
    const item = await response.json();
    if (!item?.folder) throw new Error("문서 저장 경로가 폴더가 아닙니다.");
    return item;
  };
  let folder = await find();
  if (folder || !create) return folder;
  const response = await fetch(`${root}/children`, {
    method: "POST", headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
    body: JSON.stringify({ name: folderName, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
  });
  if (response.status === 409) folder = await find();
  else if (response.ok) folder = await response.json();
  else throw new Error(`문서 폴더 생성 실패: HTTP ${response.status}`);
  if (!folder?.id) throw new Error("문서 폴더 ID를 받지 못했습니다.");
  return folder;
}

async function itemPhotoFolder(driveId, itemId, accessToken, create = false) {
  return namedDriveFolder(driveId, itemPhotoFolderName(itemId), accessToken, create);
}

async function itemPhotoFolderFiles(driveId, folder, itemId, accessToken) {
  if (!folder?.id) return [];
  return (await graphGetAll(`${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folder.id)}/children?$top=200`, accessToken))
    .map(item => itemPhotoMetadata(item, itemId))
    .filter(Boolean);
}

function latestItemPhoto(files) {
  return (files || []).reduce((latest, file) => !latest || String(file.uploadedAt) > String(latest.uploadedAt) ? file : latest, null);
}

async function handleListItemPhotos(siteId, accessToken) {
  const driveId = await itemPhotoDriveId(siteId, accessToken);
  const rootRows = await graphGetAll(`${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/root/children?$top=200`, accessToken);
  const folders = rootRows.filter(row => row?.folder && /^ITEM_[A-Za-z0-9_-]{1,120}$/.test(cleanText(row?.name)));
  const entries = await Promise.all(folders.map(async folder => {
    const itemId = cleanText(folder.name).slice(5);
    const latest = latestItemPhoto(await itemPhotoFolderFiles(driveId, folder, itemId, accessToken));
    return latest ? [itemId, latest] : null;
  }));
  return Object.fromEntries(entries.filter(Boolean));
}

async function handleBeginItemPhotoUpload(siteId, accessToken, payload) {
  const itemId = cleanText(payload?.itemId).trim();
  const storedName = itemPhotoUploadName(payload?.fileName);
  const size = Number(payload?.size);
  if (!Number.isSafeInteger(size) || size < 1 || size > ITEM_PHOTO_MAX_BYTES) {
    throw new Error("품목 사진은 1바이트 이상, 10MB 이하만 업로드할 수 있습니다.");
  }
  const item = await findListItem(siteId, "XRF_Items", payload?.itemSpItemId, accessToken);
  if (!item || cleanText(item?.fields?.Title).trim() !== itemId) {
    throw new Error("실제 DB에 등록된 품목에만 사진을 업로드할 수 있습니다.");
  }
  const driveId = await itemPhotoDriveId(siteId, accessToken);
  const folder = await itemPhotoFolder(driveId, itemId, accessToken, true);
  const url = `${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(storedName)}:/createUploadSession`;
  const response = await fetchWithRetry(url, {
    method: "POST", headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "replace", name: storedName } }),
  }, { retry5xx: false, retryNetwork: false });
  const session = await response.json();
  if (!session?.uploadUrl) throw new Error("Graph 사진 업로드 세션을 시작하지 못했습니다.");
  return { uploadUrl: validateUploadSessionUrl(session.uploadUrl), storedName, chunkBytes: PRECISION_FILE_CHUNK_BYTES };
}

async function handleFinishItemPhotoUpload(siteId, accessToken, payload) {
  const itemId = cleanText(payload?.itemId).trim();
  const storedName = cleanText(payload?.storedName).trim();
  const driveItemId = cleanText(payload?.driveItemId).trim();
  itemPhotoFolderName(itemId);
  if (!driveItemId || !storedName) throw new Error("완료된 사진의 Graph ID가 없습니다.");
  const driveId = await itemPhotoDriveId(siteId, accessToken);
  const folder = await itemPhotoFolder(driveId, itemId, accessToken);
  const files = await itemPhotoFolderFiles(driveId, folder, itemId, accessToken);
  const photo = files.find(file => file.id === driveItemId);
  if (!photo || photo.size < 1) throw new Error("문서 라이브러리에 업로드된 사진을 확인하지 못했습니다.");
  const item = await findListItem(siteId, "XRF_Items", payload?.itemSpItemId, accessToken);
  if (!item) throw new Error("사진을 연결할 품목 행을 찾지 못했습니다.");
  if (cleanText(item?.fields?.Title).trim() !== itemId) throw new Error("사진 품목 ID와 SharePoint 행이 일치하지 않습니다.");
  await patchListItem(siteId, "XRF_Items", item.id, {
    Photo_File_Name: photo.name,
    Photo_File_ID: photo.id,
    Photo_File_URL: photo.webUrl,
  }, accessToken);
  return photo;
}

async function handleNormalizeItemPhotoName(siteId, accessToken, payload) {
  const itemId = cleanText(payload?.itemId).trim();
  const driveItemId = cleanText(payload?.driveItemId).trim();
  const desiredName = itemPhotoUploadName(payload?.fileName);
  itemPhotoFolderName(itemId);
  if (!driveItemId) throw new Error("정리할 사진의 Graph ID가 없습니다.");

  const item = await findListItem(siteId, "XRF_Items", payload?.itemSpItemId, accessToken);
  if (!item || cleanText(item?.fields?.Title).trim() !== itemId) {
    throw new Error("사진 파일명을 정리할 품목 행을 확인하지 못했습니다.");
  }
  const driveId = await itemPhotoDriveId(siteId, accessToken);
  const folder = await itemPhotoFolder(driveId, itemId, accessToken);
  const files = await itemPhotoFolderFiles(driveId, folder, itemId, accessToken);
  const current = files.find(file => file.id === driveItemId);
  if (!current) throw new Error("품목 폴더에서 정리할 사진을 찾지 못했습니다.");

  let normalized = current;
  if (current.storedName !== desiredName) {
    const response = await fetchWithRetry(`${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(driveItemId)}`, {
      method: "PATCH",
      headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
      body: JSON.stringify({ name: desiredName }),
    }, { retry5xx: false, retryNetwork: false });
    normalized = itemPhotoMetadata(await response.json(), itemId);
    if (!normalized) throw new Error("정리된 사진 파일 정보를 확인하지 못했습니다.");
  }

  await patchListItem(siteId, "XRF_Items", item.id, {
    Photo_File_Name: normalized.name,
    Photo_File_ID: normalized.id,
    Photo_File_URL: normalized.webUrl,
  }, accessToken);
  return { ...normalized, renamed: current.storedName !== desiredName };
}

async function sendItemPhoto(res, siteId, accessToken, itemId, requestedFileId = "") {
  itemPhotoFolderName(itemId);
  const driveId = await itemPhotoDriveId(siteId, accessToken);
  const fileId = cleanText(requestedFileId).trim();
  if (fileId && !/^[A-Za-z0-9!._-]{1,240}$/.test(fileId)) throw new Error("유효하지 않은 사진 파일 ID입니다.");
  let photo = null;
  if (!fileId) {
    const folder = await itemPhotoFolder(driveId, itemId, accessToken);
    photo = latestItemPhoto(await itemPhotoFolderFiles(driveId, folder, itemId, accessToken));
  }
  const resolvedFileId = fileId || photo?.id || "";
  if (!resolvedFileId) return sendJson(res, 404, { ok: false, error: "등록된 품목 사진이 없습니다." });
  const response = await fetchWithRetry(`${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(resolvedFileId)}/content`, {
    method: "GET", headers: graphHeaders(accessToken),
  });
  const buffer = Buffer.from(await response.arrayBuffer());
  res.statusCode = 200;
  res.setHeader("Content-Type", response.headers.get("content-type") || photo?.mimeType || "application/octet-stream");
  res.setHeader("Content-Length", String(buffer.length));
  res.setHeader("Cache-Control", "private, max-age=300");
  return res.end(buffer);
}

function latestPrecisionFiles(rows) {
  const output = { source: null, report: null };
  for (const row of rows) {
    const file = row.metadata;
    if (!output[file.kind] || String(file.uploadedAt) > String(output[file.kind].uploadedAt)) output[file.kind] = file;
  }
  return output;
}

async function handleListPrecisionFiles(siteId, accessToken, measurementId) {
  precisionFileFolderName(measurementId);
  const driveId = await precisionDriveId(siteId, accessToken);
  const folder = await precisionFolder(driveId, measurementId, accessToken);
  return latestPrecisionFiles(await precisionFolderFiles(driveId, folder, accessToken));
}

async function handleBeginPrecisionFileUpload(siteId, accessToken, payload) {
  const measurementId = cleanText(payload?.measurementId).trim();
  const kind = cleanText(payload?.kind).trim();
  const originalName = cleanText(payload?.fileName).trim();
  const storedName = precisionUploadName(kind, originalName);
  const size = Number(payload?.size);
  if (!Number.isSafeInteger(size) || size < 1 || size > PRECISION_FILE_MAX_BYTES) {
    throw new Error("파일은 1바이트 이상, 250MB 이하만 업로드할 수 있습니다.");
  }
  const measurements = await readList(siteId, "XRF_Measurements", accessToken);
  if (!measurements.some(row => cleanText(row?.fields?.Title).trim() === measurementId)) {
    throw new Error("실제 DB에 등록된 XRF 측정 건만 파일을 첨부할 수 있습니다.");
  }
  const precisionRows = await readList(siteId, "XRF_PrecisionAnalyses", accessToken);
  if (kind === "report" && !precisionRows.some(row => cleanText(row?.fields?.field_1).trim() === measurementId)) {
    throw new Error("정밀분석 인계가 DB에 저장된 뒤 파일을 업로드하세요.");
  }
  const driveId = await precisionDriveId(siteId, accessToken);
  const folder = await precisionFolder(driveId, measurementId, accessToken, true);
  if (kind === "report") {
    const files = latestPrecisionFiles(await precisionFolderFiles(driveId, folder, accessToken));
    if (!files.source) throw new Error("XRF 원본 파일을 먼저 업로드하세요.");
  }
  const url = `${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(storedName)}:/createUploadSession`;
  const response = await fetchWithRetry(url, {
    method: "POST", headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "fail", name: storedName } }),
  }, { retry5xx: false, retryNetwork: false });
  const session = await response.json();
  if (!session?.uploadUrl) throw new Error("Graph 업로드 세션을 시작하지 못했습니다.");
  return { uploadUrl: validateUploadSessionUrl(session.uploadUrl), storedName, chunkBytes: PRECISION_FILE_CHUNK_BYTES };
}

async function handlePrecisionFileChunk(payload) {
  const uploadUrl = validateUploadSessionUrl(payload?.uploadUrl);
  const offset = Number(payload?.offset);
  const total = Number(payload?.totalSize);
  const base64 = cleanText(payload?.base64);
  if (!Number.isSafeInteger(offset) || offset < 0 || !Number.isSafeInteger(total)
    || total < 1 || total > PRECISION_FILE_MAX_BYTES || !base64 || !/^[A-Za-z0-9+/]+={0,2}$/.test(base64)) {
    throw new Error("업로드 조각의 크기 또는 형식이 올바르지 않습니다.");
  }
  const bytes = Buffer.from(base64, "base64");
  if (!bytes.length || bytes.length > PRECISION_FILE_CHUNK_BYTES || offset + bytes.length > total
    || Buffer.from(bytes).toString("base64") !== base64) {
    throw new Error("업로드 조각 데이터가 올바르지 않습니다.");
  }
  if (offset + bytes.length < total && bytes.length % (320 * 1024) !== 0) {
    throw new Error("마지막 조각을 제외한 업로드 크기는 320KiB의 배수여야 합니다.");
  }
  const response = await fetch(uploadUrl, {
    method: "PUT", headers: {
      "Content-Length": String(bytes.length),
      "Content-Range": `bytes ${offset}-${offset + bytes.length - 1}/${total}`,
    }, body: bytes,
  });
  if (![200, 201, 202].includes(response.status)) {
    const detail = await response.text();
    throw new Error(`Graph 파일 업로드 HTTP ${response.status}: ${detail.slice(0, 500)}`);
  }
  const result = await response.json();
  return response.status === 202
    ? { complete: false, nextExpectedRanges: result?.nextExpectedRanges || [] }
    : { complete: true, driveItemId: result?.id || null };
}

async function handleFinishPrecisionFileUpload(siteId, accessToken, payload) {
  const measurementId = cleanText(payload?.measurementId).trim();
  const storedName = cleanText(payload?.storedName).trim();
  const driveItemId = cleanText(payload?.driveItemId).trim();
  precisionFileFolderName(measurementId);
  if (!driveItemId || !storedName) throw new Error("완료된 파일의 Graph ID가 없습니다.");
  const driveId = await precisionDriveId(siteId, accessToken);
  const folder = await precisionFolder(driveId, measurementId, accessToken);
  const rows = await precisionFolderFiles(driveId, folder, accessToken);
  const file = rows.find(row => row.item.id === driveItemId && row.item.name === storedName)?.metadata;
  if (!file || file.size < 1) throw new Error("문서 라이브러리에 업로드된 파일을 확인하지 못했습니다.");
  const measurementRows = await readList(siteId, "XRF_Measurements", accessToken);
  const measurement = measurementRows.find(row => cleanText(row?.fields?.Title).trim() === measurementId);
  if (!measurement) throw new Error("XRF 측정 행을 찾지 못해 원본 파일을 연결할 수 없습니다.");
  const precisionRows = await readList(siteId, "XRF_PrecisionAnalyses", accessToken);
  const header = precisionRows.find(row => cleanText(row?.fields?.field_1).trim() === measurementId);
  if (file.kind === "source") {
    await patchListItemExistingFields(siteId, "XRF_Measurements", measurement.id, {
      field_6: file.name,
      Source_File_ID: file.id,
      Source_File_URL: file.webUrl,
    }, accessToken);
  } else {
    if (!header) throw new Error("정밀분석 인계 기록을 찾지 못해 성적서 파일을 연결할 수 없습니다.");
    await patchListItemExistingFields(siteId, "XRF_PrecisionAnalyses", header.id, {
      field_5: file.name,
      field_6: false,
      Report_File_ID: file.id,
      Report_File_URL: file.webUrl,
    }, accessToken);
  }
  return file;
}

async function graphGetAll(firstUrl, accessToken) {
  const rows = [];
  let nextUrl = firstUrl;
  while (nextUrl) {
    const response = await fetchWithRetry(nextUrl, { method: "GET", headers: graphHeaders(accessToken) });
    const payload = await response.json();
    if (Array.isArray(payload?.value)) rows.push(...payload.value);
    nextUrl = payload?.["@odata.nextLink"] || null;
  }
  return rows;
}

async function patchListItemExistingFields(siteId, listName, itemId, fields, accessToken, options = {}) {
  const columns = await graphGetAll(`${listUrl(siteId, listName)}/columns?$top=200`, accessToken);
  const names = new Set(columns.map(column => cleanText(column?.name).trim()));
  const supported = Object.fromEntries(Object.entries(fields).filter(([name]) => names.has(name)));
  if (!Object.keys(supported).length) {
    if (options.allowEmpty) return null;
    throw new Error(`${listName}에 저장 가능한 파일 참조 열이 없습니다.`);
  }
  return patchListItem(siteId, listName, itemId, supported, accessToken);
}

async function handleEnsureFileReferenceColumns(siteId, accessToken) {
  const created = [];
  const existing = [];
  for (const [listName, definitions] of Object.entries(FILE_REFERENCE_COLUMNS)) {
    const columns = await graphGetAll(`${listUrl(siteId, listName)}/columns?$top=200`, accessToken);
    const names = new Set(columns.map(column => cleanText(column?.name).trim()));
    for (const definition of definitions) {
      if (names.has(definition.name)) {
        existing.push(`${listName}.${definition.name}`);
        continue;
      }
      const response = await fetchWithRetry(`${listUrl(siteId, listName)}/columns`, {
        method: "POST",
        headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
        body: JSON.stringify({
          description: "MXKPCS document-library file reference",
          enforceUniqueValues: false,
          hidden: false,
          indexed: false,
          ...definition,
        }),
      }, { retry5xx: false, retryNetwork: false });
      const column = await response.json();
      created.push(`${listName}.${column?.name || definition.name}`);
    }
  }
  return { created, existing };
}

async function handleListColumns(siteId, accessToken, listName) {
  if (!LIST_IDS[listName]) throw new Error("조회할 수 없는 SharePoint List입니다.");
  return (await graphGetAll(`${listUrl(siteId, listName)}/columns?$top=200`, accessToken)).map(column => ({
    id: column?.id || null,
    name: column?.name || null,
    displayName: column?.displayName || null,
    hidden: column?.hidden === true,
    readOnly: column?.readOnly === true,
    type: column?.text ? "text" : column?.number ? "number" : column?.dateTime ? "dateTime"
      : column?.boolean ? "boolean" : column?.hyperlinkOrPicture ? "hyperlinkOrPicture" : "other",
  }));
}

function listUrl(siteId, listName) {
  const listId = LIST_IDS[listName];
  if (!listId) throw new Error(`알 수 없는 SharePoint List: ${listName}`);
  return `${GRAPH_ROOT}/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}`;
}

function compactListItem(row) {
  return {
    id: row?.id ?? null,
    eTag: row?.eTag ?? row?.["@odata.etag"] ?? null,
    createdDateTime: row?.createdDateTime ?? null,
    lastModifiedDateTime: row?.lastModifiedDateTime ?? null,
    createdBy: row?.createdBy ?? null,
    lastModifiedBy: row?.lastModifiedBy ?? null,
    fields: row?.fields ?? {},
  };
}

async function readList(siteId, listName, accessToken) {
  const url = `${listUrl(siteId, listName)}/items?$expand=fields&$top=200`;
  return (await graphGetAll(url, accessToken)).map(compactListItem);
}

async function findListItem(siteId, listName, itemSpItemId, accessToken) {
  if (itemSpItemId) {
    const response = await fetchWithRetry(`${listUrl(siteId, listName)}/items/${encodeURIComponent(itemSpItemId)}?$expand=fields`, {
      method: "GET", headers: graphHeaders(accessToken),
    });
    return compactListItem(await response.json());
  }
  return null;
}

function graphOperationError(context, error) {
  const wrapped = new Error(`[${context}] ${error?.message || String(error)}`);
  wrapped.status = error?.status;
  wrapped.cause = error;
  return wrapped;
}

async function createListItem(siteId, listName, fields, accessToken) {
  if (!WRITABLE_LISTS.has(listName)) throw new Error(`쓰기 허용되지 않은 List: ${listName}`);
  try {
    const response = await fetchWithRetry(`${listUrl(siteId, listName)}/items`, {
      method: "POST",
      headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
      body: JSON.stringify({ fields: createFields(fields) }),
    });
    return compactListItem(await response.json());
  } catch (error) {
    throw graphOperationError(`${listName} CREATE`, error);
  }
}

async function patchListItem(siteId, listName, itemId, fields, accessToken) {
  if (!WRITABLE_LISTS.has(listName)) throw new Error(`쓰기 허용되지 않은 List: ${listName}`);
  if (!itemId) throw new Error(`${listName} patch itemId가 없습니다.`);
  try {
    const response = await fetchWithRetry(`${listUrl(siteId, listName)}/items/${encodeURIComponent(itemId)}/fields`, {
      method: "PATCH",
      headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
      body: JSON.stringify(fields),
    });
    return await response.json();
  } catch (error) {
    throw graphOperationError(`${listName} PATCH id=${itemId}`, error);
  }
}

async function deleteListItem(siteId, listName, itemId, accessToken) {
  if (!WRITABLE_LISTS.has(listName)) throw new Error(`삭제 허용되지 않은 List: ${listName}`);
  if (!itemId) throw new Error(`${listName} delete itemId가 없습니다.`);
  const response = await fetch(`${listUrl(siteId, listName)}/items/${encodeURIComponent(itemId)}`, {
    method: "DELETE",
    headers: graphHeaders(accessToken),
  });
  if (!response.ok && response.status !== 404) {
    const text = await response.text();
    throw new Error(`Graph DELETE ${response.status}: ${text.slice(0, 1200)}`);
  }
}

function makeItemId(partNumber) {
  const body = cleanText(partNumber).trim().replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  if (!body) throw new Error("Part Number가 없어 Item_ID를 만들 수 없습니다.");
  return `ITM_${body}`;
}

function nextRequestId(existingRows, requestType, targetPartNumber, requestedDate) {
  const date = yyyymmdd(requestedDate || new Date().toISOString().slice(0, 10));
  const titles = new Set(existingRows.map(r => cleanText(r?.fields?.Title).trim()).filter(Boolean));
  let base;
  if (requestType === "Discontinue") {
    base = `${cleanText(targetPartNumber || "DISC").trim() || "DISC"}-REQ-${date}`;
    if (!titles.has(base)) return base;
    for (let n = 2; n < 1000; n += 1) {
      const candidate = `${base}-${String(n).padStart(2, "0")}`;
      if (!titles.has(candidate)) return candidate;
    }
  } else {
    const prefix = requestType === "Replacement" ? "REQ_REP" : "REQ_NEW";
    const rx = new RegExp(`^${prefix}_${date}_(\\d{3})$`);
    let max = 0;
    for (const title of titles) {
      const m = title.match(rx);
      if (m) max = Math.max(max, Number(m[1]));
    }
    return `${prefix}_${date}_${String(max + 1).padStart(3, "0")}`;
  }
  throw new Error("Request_ID를 발번하지 못했습니다.");
}

function nextMeasurementId(rows, itemId, partNumber, measuredDate) {
  const date = dateOnly(measuredDate);
  const dateCompact = yyyymmdd(date);
  let maxSequence = 0;
  for (const row of rows) {
    const f = row?.fields || {};
    if (cleanText(f.field_1).trim() !== cleanText(itemId).trim()) continue;
    if (dateOnly(f.Measured_Date) !== date) continue;
    const m = cleanText(f.Title).match(/_(\d{2})$/);
    if (m) maxSequence = Math.max(maxSequence, Number(m[1]));
    else maxSequence += 1;
  }
  return `${cleanText(partNumber).trim()}_${dateCompact}_${String(maxSequence + 1).padStart(2, "0")}`;
}

function rawElementFields(measurementId, element, row) {
  return {
    Title: measurementId,
    field_1: element,
    // SharePoint의 Raw_* 컬럼은 모두 Text이므로 XLS/XLSX 숫자 셀도 문자열로 저장합니다.
    field_2: cleanText(row?.rawPpm ?? row?.ppm),
    field_3: cleanText(row?.rawSigma ?? row?.sigma),
    field_4: cleanText(row?.rawJudgement ?? row?.sourceJudgement ?? row?.judge),
  };
}

async function createMeasurementRecord(siteId, accessToken, payload) {
  const itemId = cleanText(payload?.itemId).trim();
  const partNumber = cleanText(payload?.partNumber).trim();
  const measuredDate = dateOnly(payload?.measuredDate);
  const role = cleanText(payload?.role || "Initial").trim();
  if (!itemId || !partNumber || !measuredDate) throw new Error("측정 저장에 Item_ID, Part_Number, Measured_Date가 필요합니다.");
  if (!["Initial", "Periodic", "Retest"].includes(role)) throw new Error(`허용되지 않은 Measurement_Role: ${role}`);

  const measurements = await readList(siteId, "XRF_Measurements", accessToken);
  const measurementId = nextMeasurementId(measurements, itemId, partNumber, measuredDate);
  const created = [];
  let writeStage = "XRF_Measurements Header CREATE";
  try {
    const header = await createListItem(siteId, "XRF_Measurements", {
      Title: measurementId,
      field_1: itemId,
      Measured_Date: measuredDate,
      field_3: role,
      field_4: Number(payload?.attemptNo || 1),
      field_5: nonEmpty(payload?.parentMeasurementId),
      field_6: nonEmpty(payload?.sourceFileName),
      field_7: nonEmpty(payload?.sourceSheet),
      Requester_Name: nonEmpty(payload?.requesterName),
    }, accessToken);
    created.push(["XRF_Measurements", header.id]);

    for (const element of XRF_SOURCE_ELEMENTS) {
      writeStage = `XRF_ElementResults ${element} CREATE`;
      const row = payload?.elements?.[element] || {};
      const elementItem = await createListItem(siteId, "XRF_ElementResults", rawElementFields(measurementId, element, row), accessToken);
      created.push(["XRF_ElementResults", elementItem.id]);
    }

    return { measurementId, measurementSpItemId: header.id };
  } catch (error) {
    for (const [listName, itemSpId] of created.reverse()) {
      try { await deleteListItem(siteId, listName, itemSpId, accessToken); } catch {}
    }
    const stagedError = new Error(`${writeStage} 실패: ${error?.message || String(error)}`);
    stagedError.status = error?.status;
    throw stagedError;
  }
}

function requestFieldsFromPayload(requestId, payload, itemId) {
  const requestType = payload.requestType;
  const form = payload.form || {};
  const isReplacement = requestType === "Replacement";
  const isDiscontinue = requestType === "Discontinue";
  const reasonLegacy = isReplacement
    ? (form.replacementReason === "기타" ? form.replacementReasonOther : form.replacementReason)
    : isDiscontinue
      ? (form.discontinueReason === "기타" ? form.discontinueReasonOther : form.discontinueReason)
      : form.requestReason;
  const dispositionLegacy = isReplacement ? form.oldItemDisposition : isDiscontinue ? form.stockDisposition : "";

  return {
    Title: requestId,
    field_1: requestType,
    field_2: payload.status || "Pending",
    field_3: nonEmpty(payload.targetItemId),
    field_4: nonEmpty(itemId),
    field_5: nonEmpty(payload.partNumber),
    field_6: nonEmpty(form.itemName),
    field_7: nonEmpty(form.type),
    field_8: nonEmpty(form.handlingDept),
    field_9: nonEmpty(form.materialCategory),
    field_10: nonEmpty(form.materialState),
    field_11: nonEmpty(form.manufacturer),
    field_13: nonEmpty(form.crType),
    field_14: isDiscontinue ? null : (payload.xrfInputMode === "upload" ? "Upload" : "Request"),
    field_15: nonEmpty(reasonLegacy),
    field_16: nonEmpty(dispositionLegacy),
    field_18: nonEmpty(form.note),
    Use_Date: dateOnly(form.useDate),
    Effective_Date: dateOnly(isReplacement ? form.replacementDate : isDiscontinue ? form.finalUseDate : null),
    Request_Reason: nonEmpty(form.requestReason),
    Replacement_Reason: isReplacement ? nonEmpty(form.replacementReason) : null,
    Replacement_Reason_Other: isReplacement && form.replacementReason === "기타" ? nonEmpty(form.replacementReasonOther) : null,
    Old_Item_Disposition: isReplacement ? nonEmpty(form.oldItemDisposition) : null,
    Replacement_Date: isReplacement ? dateOnly(form.replacementDate) : null,
    Discontinue_Reason: isDiscontinue ? nonEmpty(form.discontinueReason) : null,
    Discontinue_Reason_Other: isDiscontinue && form.discontinueReason === "기타" ? nonEmpty(form.discontinueReasonOther) : null,
    Final_Use_Date: isDiscontinue ? dateOnly(form.finalUseDate) : null,
    Stock_Disposition: isDiscontinue ? nonEmpty(form.stockDisposition) : null,
    Processed_At: payload.status && payload.status !== "Pending" ? (payload.processedAt || new Date().toISOString()) : null,
    Processed_By: payload.status && payload.status !== "Pending" ? nonEmpty(payload.actor || "admin") : null,
    Before_Target_Snapshot_JSON: payload.beforeTargetSnapshot ? JSON.stringify(payload.beforeTargetSnapshot) : null,
    Is_Migration: false,
    Requester_Name: nonEmpty(form.requesterName),
  };
}

function itemFieldsFromPayload(itemId, payload) {
  const item = payload.item || {};
  return {
    Title: itemId,
    field_1: cleanText(payload.partNumber).trim(),
    field_2: cleanText(item.name || payload.form?.itemName || "신규 등록 요청 품목").trim(),
    field_3: nonEmpty(item.nameEn),
    field_4: item.type || payload.form?.type || "Auxiliary Materials",
    field_5: item.dept || payload.form?.handlingDept || "Assembly",
    field_6: item.materialCategory || payload.form?.materialCategory || "Others",
    field_7: nonEmpty(item.materialState || payload.form?.materialState),
    field_8: item.crType || payload.form?.crType || "비접촉",
    field_9: item.lifecycle || (payload.requestType === "Replacement" ? "NewReplacement" : "NewItem"),
    field_10: item.isCurrent !== false,
    field_11: item.registrationPeriodNo == null ? null : Number(item.registrationPeriodNo),
    field_13: nonEmpty(item.actionNote),
    Replacement_Of: payload.requestType === "Replacement" ? nonEmpty(payload.targetPartNumber) : nonEmpty(item.replacementOf),
    Replacement_Date: payload.requestType === "Replacement"
      ? dateOnly(item.replacementDate || payload.form?.replacementDate)
      : null,
  };
}

async function handleCreateRequest(siteId, accessToken, payload) {
  const requestType = payload?.requestType;
  if (!["New", "Replacement", "Discontinue"].includes(requestType)) throw new Error("requestType은 New / Replacement / Discontinue 중 하나여야 합니다.");

  const requestRows = await readList(siteId, "XRF_Requests", accessToken);
  const requestId = nextRequestId(requestRows, requestType, payload?.targetPartNumber, payload?.requestedDate);
  const isDiscontinue = requestType === "Discontinue";
  const itemId = isDiscontinue ? null : (payload?.itemId || makeItemId(payload?.partNumber));
  const created = [];
  try {
    const request = await createListItem(siteId, "XRF_Requests", requestFieldsFromPayload(requestId, payload, itemId), accessToken);
    created.push(["XRF_Requests", request.id]);

    let item = null;
    if (!isDiscontinue) {
      item = await createListItem(siteId, "XRF_Items", itemFieldsFromPayload(itemId, payload), accessToken);
      created.push(["XRF_Items", item.id]);
    }

    let measurement = null;
    if (!isDiscontinue && payload?.initialMeasurement?.elements) {
      measurement = await createMeasurementRecord(siteId, accessToken, {
        itemId,
        partNumber: payload.partNumber,
        measuredDate: payload.initialMeasurement.measuredDate,
        role: "Initial",
        attemptNo: 1,
        parentMeasurementId: null,
        sourceFileName: payload.initialMeasurement.sourceFileName,
        sourceSheet: payload.initialMeasurement.sourceSheet,
        requesterName: payload.form?.requesterName,
        elements: payload.initialMeasurement.elements,
      });
      if (payload?.initialMeasurement?.requestStatus && payload.initialMeasurement.requestStatus !== "Pending") {
        await patchListItem(siteId, "XRF_Requests", request.id, {
          field_2: payload.initialMeasurement.requestStatus,
          Processed_At: new Date().toISOString(),
          Processed_By: nonEmpty(payload.actor || "admin"),
        }, accessToken);
      }
    }

    // Existing target is patched last. This keeps create/measurement failures from leaving the old master in a terminal state.
    if (requestType === "Replacement" && payload?.targetSpItemId) {
      await patchListItem(siteId, "XRF_Items", payload.targetSpItemId, {
        field_9: "ReplacedOld",
        field_10: false,
        Replacement_Date: dateOnly(payload?.form?.replacementDate),
        Lifecycle_End_Date: dateOnly(payload?.form?.replacementDate),
      }, accessToken);
    }

    if (isDiscontinue && payload?.status === "Applied" && payload?.targetSpItemId) {
      await patchListItem(siteId, "XRF_Items", payload.targetSpItemId, {
        field_9: "Discontinued",
        field_10: false,
        Replacement_Date: dateOnly(payload?.form?.finalUseDate || payload?.requestedDate),
      }, accessToken);
    }

    return {
      requestId,
      requestSpItemId: request.id,
      itemId,
      itemSpItemId: item?.id || null,
      measurement,
    };
  } catch (error) {
    // Roll back only rows created by this action. Existing target-item patches are performed last where possible.
    for (const [listName, itemSpId] of created.reverse()) {
      try { await deleteListItem(siteId, listName, itemSpId, accessToken); } catch {}
    }
    throw error;
  }
}

async function handleSaveMeasurement(siteId, accessToken, payload) {
  const measurement = await createMeasurementRecord(siteId, accessToken, payload);
  if (payload?.itemSpItemId && payload?.itemPatch) {
    const patch = {};
    if (payload.itemPatch.actionNote !== undefined) patch.field_13 = nonEmpty(payload.itemPatch.actionNote);
    if (payload.itemPatch.registrationPeriodNo !== undefined) patch.field_11 = payload.itemPatch.registrationPeriodNo == null ? null : Number(payload.itemPatch.registrationPeriodNo);
    if (Object.keys(patch).length) await patchListItem(siteId, "XRF_Items", payload.itemSpItemId, patch, accessToken);
  }
  if (payload?.requestSpItemId && payload?.requestStatus) {
    const fields = { field_2: payload.requestStatus };
    if (payload.requestStatus !== "Pending") {
      fields.Processed_At = new Date().toISOString();
      fields.Processed_By = nonEmpty(payload.actor || "admin");
    }
    await patchListItem(siteId, "XRF_Requests", payload.requestSpItemId, fields, accessToken);
  }
  return measurement;
}

async function handleUpdateItem(siteId, accessToken, payload) {
  const itemSpItemId = payload?.itemSpItemId;
  const oldCode = cleanText(payload?.oldCode).trim();
  const newCode = cleanText(payload?.newCode).trim();
  if (!itemSpItemId || !newCode) throw new Error("품목 수정에 itemSpItemId와 newCode가 필요합니다.");
  await patchListItem(siteId, "XRF_Items", itemSpItemId, {
    field_1: newCode,
    field_2: cleanText(payload?.name).trim(),
    field_3: nonEmpty(payload?.nameEn),
    field_5: cleanText(payload?.dept).trim(),
  }, accessToken);

  if (oldCode && oldCode !== newCode) {
    const itemRows = await readList(siteId, "XRF_Items", accessToken);
    for (const row of itemRows) {
      if (row.id === itemSpItemId) continue;
      if (cleanText(row?.fields?.Replacement_Of).trim() === oldCode) {
        await patchListItem(siteId, "XRF_Items", row.id, { Replacement_Of: newCode }, accessToken);
      }
    }
    const requestRows = await readList(siteId, "XRF_Requests", accessToken);
    for (const row of requestRows) {
      if (cleanText(row?.fields?.field_5).trim() === oldCode) {
        await patchListItem(siteId, "XRF_Requests", row.id, { field_5: newCode }, accessToken);
      }
    }
  }
  return { oldCode, newCode };
}

async function handleProcessDiscontinue(siteId, accessToken, payload) {
  if (!payload?.requestSpItemId) throw new Error("단종 처리 Request SharePoint Item ID가 없습니다.");
  const approved = payload?.status === "Applied";
  const rejected = payload?.status === "Rejected";
  if (!approved && !rejected) throw new Error("단종 처리 status는 Applied 또는 Rejected여야 합니다.");

  await patchListItem(siteId, "XRF_Requests", payload.requestSpItemId, {
    field_2: payload.status,
    Processed_At: payload.processedAt || new Date().toISOString(),
    Processed_By: nonEmpty(payload.actor || "admin"),
    Before_Target_Snapshot_JSON: approved && payload.beforeTargetSnapshot ? JSON.stringify(payload.beforeTargetSnapshot) : null,
  }, accessToken);

  if (approved && payload?.targetSpItemId) {
    await patchListItem(siteId, "XRF_Items", payload.targetSpItemId, {
      field_9: "Discontinued",
      field_10: false,
      Replacement_Date: dateOnly(payload.finalUseDate || payload.processedAt),
    }, accessToken);
  }
  return { status: payload.status };
}

async function handleRevertDiscontinue(siteId, accessToken, payload) {
  if (!payload?.requestSpItemId || !payload?.targetSpItemId) throw new Error("원복에 Request/Target SharePoint Item ID가 필요합니다.");
  const restored = payload?.restored || {};
  await patchListItem(siteId, "XRF_Items", payload.targetSpItemId, {
    field_9: restored.lifecycle || "ExistingActive",
    field_10: restored.isCurrent !== false,
    Replacement_Date: dateOnly(restored.replacementDate),
  }, accessToken);
  await patchListItem(siteId, "XRF_Requests", payload.requestSpItemId, {
    field_2: "Reverted",
    Reverted_At: dateOnly(payload.revertedAt || new Date().toISOString()),
    Reverted_By: nonEmpty(payload.actor || "admin"),
    Revert_Reason: nonEmpty(payload.reason),
    Restored_Target_Snapshot_JSON: payload.restored ? JSON.stringify(payload.restored) : null,
  }, accessToken);
  return { status: "Reverted" };
}

function precisionIdForTrigger(triggerMeasurementId) {
  const raw = `PREC_${cleanText(triggerMeasurementId).trim()}`;
  return raw.length <= 250 ? raw : `PREC_${Buffer.from(raw).toString("base64url").slice(0, 220)}`;
}

async function handleUpsertPrecision(siteId, accessToken, payload) {
  const triggerMeasurementId = cleanText(payload?.triggerMeasurementId).trim();
  if (!triggerMeasurementId) throw new Error("정밀분석 Trigger_Measurement_ID가 없습니다.");
  const precisionRows = await readList(siteId, "XRF_PrecisionAnalyses", accessToken);
  let header = precisionRows.find(r => cleanText(r?.fields?.field_1).trim() === triggerMeasurementId) || null;
  const precisionId = header?.fields?.Title || payload?.precisionId || precisionIdForTrigger(triggerMeasurementId);
  const override = payload?.override || {};
  const headerFields = {
    field_1: triggerMeasurementId,
    Requested_At: dateOnly(override.requestedAt || payload.requestedAt || new Date().toISOString()),
    Requested_By: nonEmpty(override.requestedBy || payload.actor || "admin"),
    field_5: nonEmpty(override.resultFileId || override.resultFileName),
    field_6: override.finalConfirm === "확인" || bool(override.finalConfirmed),
  };
  if (override.measuredDate) headerFields.Measured_Date = dateOnly(override.measuredDate);
  if (override.labName) headerFields.field_2 = nonEmpty(override.labName);
  if (override.reportNo) headerFields.field_3 = nonEmpty(override.reportNo);

  if (!header) {
    header = await createListItem(siteId, "XRF_PrecisionAnalyses", { Title: precisionId, ...headerFields }, accessToken);
  } else {
    await patchListItem(siteId, "XRF_PrecisionAnalyses", header.id, headerFields, accessToken);
  }

  const elementResults = override.elementResults || {};
  if (Object.keys(elementResults).length) {
    const elementRows = await readList(siteId, "XRF_PrecisionElementResults", accessToken);
    for (const [element, state] of Object.entries(elementResults)) {
      if (!element || !state || (!state.presence && state.ppm === undefined)) continue;
      const existing = elementRows.find(r => cleanText(r?.fields?.Title).trim() === precisionId && cleanText(r?.fields?.field_1).trim() === element);
      const fields = {
        Title: precisionId,
        field_1: element,
        field_2: state.presence || "함유",
        field_3: state.presence === "미함유" || state.ppm === "" || state.ppm == null ? null : Number(state.ppm),
      };
      if (existing) await patchListItem(siteId, "XRF_PrecisionElementResults", existing.id, fields, accessToken);
      else await createListItem(siteId, "XRF_PrecisionElementResults", fields, accessToken);
    }
  }

  if ((override.finalConfirm === "확인" || bool(override.finalConfirmed)) && payload?.requestSpItemId && payload?.finalResult) {
    const requestStatus = payload.finalResult === "OK" ? "Applied" : payload.finalResult === "NG" ? "Rejected" : null;
    if (requestStatus) {
      await patchListItem(siteId, "XRF_Requests", payload.requestSpItemId, {
        field_2: requestStatus,
        Processed_At: new Date().toISOString(),
        Processed_By: nonEmpty(payload.actor || "admin"),
      }, accessToken);
    }
  }

  return { precisionId, precisionSpItemId: header.id };
}

async function handleDeleteRecord(siteId, accessToken, payload) {
  const listName = payload?.listName;
  if (!WRITABLE_LISTS.has(listName)) throw new Error("삭제 대상 List가 허용되지 않았습니다.");
  await deleteListItem(siteId, listName, payload?.itemSpItemId, accessToken);
  return { listName, itemSpItemId: payload?.itemSpItemId };
}

async function handleWriteProbe(siteId, accessToken) {
  const id = `WRITE_PROBE_${Date.now()}`;
  const row = await createListItem(siteId, "XRF_Requests", {
    Title: id,
    field_1: "New",
    field_2: "Cancelled",
    field_18: "자동 WRITE 권한 확인용 임시 행 - 즉시 삭제",
    Is_Migration: true,
  }, accessToken);
  await deleteListItem(siteId, "XRF_Requests", row.id, accessToken);
  return { ok: true, probeId: id };
}

async function readBootstrap(siteId, accessToken) {
  const names = [
    "XRF_Items",
    "XRF_Requests",
    "XRF_Measurements",
    "XRF_ElementResults",
    "XRF_PrecisionAnalyses",
    "XRF_PrecisionElementResults",
  ];
  const entries = await Promise.all(names.map(async name => [name, await readList(siteId, name, accessToken)]));
  const lists = Object.fromEntries(entries);
  return {
    ok: true,
    generatedAt: new Date().toISOString(),
    siteId,
    lists,
    counts: Object.fromEntries(Object.entries(lists).map(([name, rows]) => [name, rows.length])),
  };
}

async function parseBody(req) {
  if (req.body && typeof req.body === "object") return req.body;
  if (typeof req.body === "string") return JSON.parse(req.body || "{}");
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  const raw = Buffer.concat(chunks).toString("utf8");
  return raw ? JSON.parse(raw) : {};
}

function isSharedReadOnlyDeployment() {
  const deploymentEnv = process.env.VERCEL_ENV || process.env.VERCEL_TARGET_ENV;
  return deploymentEnv === "production" || deploymentEnv === "preview"
    || (process.env.VERCEL === "1" && process.env.NODE_ENV === "production");
}

const SHARED_APPROVED_ACTIONS = new Set([
  "createRequest", "saveMeasurement", "updateItem", "processDiscontinue",
  "revertDiscontinue", "upsertPrecision", "beginPrecisionFileUpload",
  "uploadPrecisionFileChunk", "finishPrecisionFileUpload",
  "beginItemPhotoUpload", "finishItemPhotoUpload",
]);

export default async function handler(req, res) {
  const shared = isSharedReadOnlyDeployment();
  const session = shared ? verifyWriteSession(req) : null;
  const readOnly = shared && !session;
  if (req.method === "OPTIONS") {
    res.statusCode = 204;
    res.setHeader("Allow", shared ? "GET, POST, OPTIONS" : "GET, POST, DELETE, OPTIONS");
    return res.end();
  }

  if (shared && req.method === "DELETE") {
    return sendJson(res, 403, { ok: false, error: "공유 배포에서는 삭제를 허용하지 않습니다." });
  }
  if (readOnly && req.method !== "GET") {
    return sendJson(res, 403, { ok: false, error: "등록·수정·업로드하려면 승인된 계정의 접근 코드로 인증하세요." });
  }

  try {
    let checkedBody = null;
    if (shared && req.method === "POST") {
      const origin = req.headers?.origin;
      if (origin) {
        const host = req.headers?.host;
        const scheme = req.headers?.["x-forwarded-proto"] === "https" || process.env.VERCEL === "1" ? "https" : "http";
        if (!host || origin !== `${scheme}://${host}`) {
          return sendJson(res, 403, { ok: false, error: "요청 출처를 확인할 수 없습니다." });
        }
      }
      checkedBody = await parseBody(req);
      if (!SHARED_APPROVED_ACTIONS.has(cleanText(checkedBody?.action).trim())) {
        return sendJson(res, 403, { ok: false, error: "공유 배포에서 허용하지 않는 작업입니다." });
      }
    }
    const siteId = process.env.MXKPCS_SITE_ID || DEFAULT_SITE_ID;
    const accessToken = await acquireGraphToken();

    if (req.method === "GET") {
      const params = new URL(req.url || "/", "http://localhost").searchParams;
      if (params.has("columnsFor")) {
        if (shared) return sendJson(res, 403, { ok: false, error: "공유 배포에서는 스키마 조회를 허용하지 않습니다." });
        return sendJson(res, 200, { ok: true, columns: await handleListColumns(siteId, accessToken, params.get("columnsFor")) });
      }
      if (params.has("precisionFilesFor")) {
        return sendJson(res, 200, { ok: true, files: await handleListPrecisionFiles(siteId, accessToken, params.get("precisionFilesFor")) });
      }
      if (params.has("itemPhotoFor")) {
        return sendItemPhoto(res, siteId, accessToken, params.get("itemPhotoFor"), params.get("itemPhotoId"));
      }
      return sendJson(res, 200, { ...await readBootstrap(siteId, accessToken), readOnly, authRequired: shared, writeAccessEmail: session?.email || null });
    }

    if (req.method === "DELETE") {
      const body = await parseBody(req);
      const result = await handleDeleteRecord(siteId, accessToken, body);
      return sendJson(res, 200, { ok: true, action: "deleteRecord", result });
    }

    if (req.method !== "POST") {
      res.setHeader("Allow", "GET, POST, DELETE, OPTIONS");
      return sendJson(res, 405, { ok: false, error: "Method Not Allowed" });
    }

    const body = checkedBody || await parseBody(req);
    const action = cleanText(body?.action).trim();
    const payload = body?.payload || {};
    let result;
    if (action === "createRequest") result = await handleCreateRequest(siteId, accessToken, payload);
    else if (action === "saveMeasurement") result = await handleSaveMeasurement(siteId, accessToken, payload);
    else if (action === "updateItem") result = await handleUpdateItem(siteId, accessToken, payload);
    else if (action === "processDiscontinue") result = await handleProcessDiscontinue(siteId, accessToken, payload);
    else if (action === "revertDiscontinue") result = await handleRevertDiscontinue(siteId, accessToken, payload);
    else if (action === "upsertPrecision") result = await handleUpsertPrecision(siteId, accessToken, payload);
    else if (action === "beginPrecisionFileUpload") result = await handleBeginPrecisionFileUpload(siteId, accessToken, payload);
    else if (action === "uploadPrecisionFileChunk") result = await handlePrecisionFileChunk(payload);
    else if (action === "finishPrecisionFileUpload") result = await handleFinishPrecisionFileUpload(siteId, accessToken, payload);
    else if (action === "beginItemPhotoUpload") result = await handleBeginItemPhotoUpload(siteId, accessToken, payload);
    else if (action === "finishItemPhotoUpload") result = await handleFinishItemPhotoUpload(siteId, accessToken, payload);
    else if (action === "normalizeItemPhotoName") result = await handleNormalizeItemPhotoName(siteId, accessToken, payload);
    else if (action === "ensureFileReferenceColumns") result = await handleEnsureFileReferenceColumns(siteId, accessToken);
    else if (action === "deleteRecord") result = await handleDeleteRecord(siteId, accessToken, payload);
    else if (action === "writeProbe") result = await handleWriteProbe(siteId, accessToken);
    else throw new Error(`지원하지 않는 action: ${action}`);

    return sendJson(res, 200, { ok: true, action, result });
  } catch (error) {
    console.error("[xrf-sharepoint]", error);
    const status = Number(error?.status) || 500;
    return sendJson(res, status >= 400 && status < 600 ? status : 500, {
      ok: false,
      error: error?.message || String(error),
    });
  }
}
