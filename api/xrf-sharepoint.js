// Vercel Serverless Function
// SharePoint XRF DB READ + WRITE gateway
// Client secret must exist only in server-side environment variables.

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

async function fetchWithRetry(url, options = {}) {
  let lastError = null;
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt += 1) {
    try {
      const response = await fetch(url, options);
      if (response.ok) return response;
      const retryable = response.status === 429 || response.status >= 500;
      const text = await response.text();
      if (!retryable || attempt === MAX_RETRIES) {
        const error = new Error(`Graph HTTP ${response.status}: ${text.slice(0, 1800)}`);
        error.status = response.status;
        throw error;
      }
      const retryAfter = Number(response.headers.get("retry-after"));
      await sleep(Number.isFinite(retryAfter) && retryAfter > 0
        ? retryAfter * 1000
        : RETRY_BASE_MS * Math.pow(2, attempt));
    } catch (error) {
      lastError = error;
      if (attempt === MAX_RETRIES) throw error;
      await sleep(RETRY_BASE_MS * Math.pow(2, attempt));
    }
  }
  throw lastError || new Error("Graph request failed");
}

async function acquireGraphToken() {
  const tenant = process.env.MXKPCS_TENANT || DEFAULT_TENANT;
  const clientId = process.env.MXKPCS_CLIENT_ID || DEFAULT_CLIENT_ID;
  const clientSecret = process.env.MXKPCS_CLIENT_SECRET;
  if (!clientSecret) throw new Error("Vercel 환경변수 MXKPCS_CLIENT_SECRET가 설정되어 있지 않습니다.");

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const response = await fetchWithRetry(
    `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`,
    { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body }
  );
  const payload = await response.json();
  if (!payload?.access_token) throw new Error("Microsoft Graph access token을 받지 못했습니다.");
  return payload.access_token;
}

function graphHeaders(accessToken, extra = {}) {
  return { Authorization: `Bearer ${accessToken}`, Accept: "application/json", ...extra };
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

async function createListItem(siteId, listName, fields, accessToken) {
  if (!WRITABLE_LISTS.has(listName)) throw new Error(`쓰기 허용되지 않은 List: ${listName}`);
  const response = await fetchWithRetry(`${listUrl(siteId, listName)}/items`, {
    method: "POST",
    headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
    body: JSON.stringify({ fields: createFields(fields) }),
  });
  return compactListItem(await response.json());
}

async function patchListItem(siteId, listName, itemId, fields, accessToken) {
  if (!WRITABLE_LISTS.has(listName)) throw new Error(`쓰기 허용되지 않은 List: ${listName}`);
  if (!itemId) throw new Error(`${listName} patch itemId가 없습니다.`);
  const response = await fetchWithRetry(`${listUrl(siteId, listName)}/items/${encodeURIComponent(itemId)}/fields`, {
    method: "PATCH",
    headers: graphHeaders(accessToken, { "Content-Type": "application/json" }),
    body: JSON.stringify(fields),
  });
  return await response.json();
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
    field_2: row?.rawPpm ?? row?.ppm ?? "",
    field_3: row?.rawSigma ?? row?.sigma ?? "",
    field_4: row?.rawJudgement ?? row?.sourceJudgement ?? row?.judge ?? "",
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
  try {
    const header = await createListItem(siteId, "XRF_Measurements", {
      Title: measurementId,
      field_1: itemId,
      Measured_Date: measuredDate,
      field_3: role,
      field_4: Number(payload?.attemptNo || 1),
      field_5: nonEmpty(payload?.parentMeasurementId),
    }, accessToken);
    created.push(["XRF_Measurements", header.id]);

    for (const element of XRF_SOURCE_ELEMENTS) {
      const row = payload?.elements?.[element] || {};
      const elementItem = await createListItem(siteId, "XRF_ElementResults", rawElementFields(measurementId, element, row), accessToken);
      created.push(["XRF_ElementResults", elementItem.id]);
    }

    return { measurementId, measurementSpItemId: header.id };
  } catch (error) {
    for (const [listName, itemSpId] of created.reverse()) {
      try { await deleteListItem(siteId, listName, itemSpId, accessToken); } catch {}
    }
    throw error;
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

export default async function handler(req, res) {
  if (req.method === "OPTIONS") {
    res.statusCode = 204;
    res.setHeader("Allow", "GET, POST, DELETE, OPTIONS");
    return res.end();
  }

  try {
    const siteId = process.env.MXKPCS_SITE_ID || DEFAULT_SITE_ID;
    const accessToken = await acquireGraphToken();

    if (req.method === "GET") return sendJson(res, 200, await readBootstrap(siteId, accessToken));

    if (req.method === "DELETE") {
      const body = await parseBody(req);
      const result = await handleDeleteRecord(siteId, accessToken, body);
      return sendJson(res, 200, { ok: true, action: "deleteRecord", result });
    }

    if (req.method !== "POST") {
      res.setHeader("Allow", "GET, POST, DELETE, OPTIONS");
      return sendJson(res, 405, { ok: false, error: "Method Not Allowed" });
    }

    const body = await parseBody(req);
    const action = cleanText(body?.action).trim();
    const payload = body?.payload || {};
    let result;
    if (action === "createRequest") result = await handleCreateRequest(siteId, accessToken, payload);
    else if (action === "saveMeasurement") result = await handleSaveMeasurement(siteId, accessToken, payload);
    else if (action === "updateItem") result = await handleUpdateItem(siteId, accessToken, payload);
    else if (action === "processDiscontinue") result = await handleProcessDiscontinue(siteId, accessToken, payload);
    else if (action === "revertDiscontinue") result = await handleRevertDiscontinue(siteId, accessToken, payload);
    else if (action === "upsertPrecision") result = await handleUpsertPrecision(siteId, accessToken, payload);
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
