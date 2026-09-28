const API_URL = process.env.XRF_API_URL || "http://localhost:3000/api/xrf-sharepoint";
const SEED_PREFIX = "DB-UAT-PG-";
const SEED_COUNT = 40;
const SOURCE_ELEMENTS = ["Pb", "Hg", "Cr", "Cd", "Cl", "Br"];
const DEPARTMENTS = ["Assembly", "Common", "Molding", "Plating", "Stamping"];

function partNumber(number) {
  return `${SEED_PREFIX}${String(number).padStart(3, "0")}`;
}

function itemId(number) {
  return `ITM_DB_UAT_PG_${String(number).padStart(3, "0")}`;
}

function isSeedRow(listName, row) {
  const fields = row?.fields || {};
  if (listName === "XRF_Items") return String(fields.field_1 || "").startsWith(SEED_PREFIX);
  if (listName === "XRF_Requests") return String(fields.field_5 || "").startsWith(SEED_PREFIX);
  if (listName === "XRF_Measurements") return String(fields.Title || "").startsWith(SEED_PREFIX);
  if (listName === "XRF_ElementResults") return String(fields.Title || "").startsWith(SEED_PREFIX);
  return false;
}

function nonSeedState(lists) {
  return Object.fromEntries(Object.entries(lists || {}).map(([listName, rows]) => [
    listName,
    new Map((rows || [])
      .filter(row => !isSeedRow(listName, row))
      .map(row => [String(row.id), JSON.stringify(row.fields || {})]))
  ]));
}

function compareNonSeedState(before, after) {
  const changes = [];
  for (const [listName, beforeRows] of Object.entries(before)) {
    const afterRows = after[listName] || new Map();
    for (const [id, fields] of beforeRows) {
      if (!afterRows.has(id)) changes.push(`${listName} row ${id}: deleted`);
      else if (afterRows.get(id) !== fields) changes.push(`${listName} row ${id}: fields changed`);
    }
    for (const id of afterRows.keys()) {
      if (!beforeRows.has(id)) changes.push(`${listName} row ${id}: unexpected row added`);
    }
  }
  return changes;
}

async function api(method = "GET", body) {
  const response = await fetch(API_URL, {
    method,
    headers: body ? { "Content-Type": "application/json" } : undefined,
    body: body ? JSON.stringify(body) : undefined,
    signal: AbortSignal.timeout(method === "GET" ? 240_000 : 180_000),
  });
  const text = await response.text();
  let parsed;
  try { parsed = JSON.parse(text); }
  catch { throw new Error(`${method} HTTP ${response.status}: invalid JSON (${text.slice(0, 300)})`); }
  if (!response.ok || parsed?.ok !== true) {
    throw new Error(`${method} HTTP ${response.status}: ${parsed?.error || text.slice(0, 500)}`);
  }
  return parsed;
}

function initialElements(remeasureRequired) {
  return Object.fromEntries(SOURCE_ELEMENTS.map(element => [element,
    remeasureRequired && element === "Pb"
      ? { rawPpm: "45", rawSigma: "3", rawJudgement: "??" }
      : { rawPpm: "N.D.", rawSigma: "", rawJudgement: "N.D." }
  ]));
}

function createPayload(number) {
  const code = partNumber(number);
  const remeasureRequired = number % 10 === 0;
  const name = `[DB-UAT-PAGE] 페이지 이동 및 긴 제품명 줄바꿈 검증용 부자재 ${String(number).padStart(3, "0")}${remeasureRequired ? " · Pb 재측정" : ""}`;
  const today = new Date().toISOString().slice(0, 10);
  const department = DEPARTMENTS[(number - 1) % DEPARTMENTS.length];
  return {
    action: "createRequest",
    payload: {
      requestType: "New",
      status: "Pending",
      requestedDate: today,
      actor: "uat-pagination-seed",
      partNumber: code,
      itemId: itemId(number),
      xrfInputMode: remeasureRequired ? "upload" : "request",
      form: {
        itemName: name,
        type: "Auxiliary Materials",
        handlingDept: department,
        materialCategory: "Packaging Material",
        materialState: "",
        manufacturer: "DB UAT pagination fixture",
        crType: "비접촉",
        requestReason: "페이지 이동·표 줄바꿈·XRF 재측정 화면 검증",
        note: "DB-UAT-PG 전용 테스트 데이터",
        useDate: today,
      },
      item: {
        name,
        nameEn: `Database UAT pagination and long product-name wrapping fixture ${String(number).padStart(3, "0")}`,
        type: "Auxiliary Materials",
        dept: department,
        materialCategory: "Packaging Material",
        materialState: "",
        crType: "비접촉",
        lifecycle: "NewItem",
        isCurrent: true,
        actionNote: null,
      },
      ...(remeasureRequired ? {
        initialMeasurement: {
          measuredDate: today,
          sourceFileName: `${code}_REMEASURE_UAT.xlsx`,
          sourceSheet: "DB UAT",
          requestStatus: "Pending",
          elements: initialElements(true),
        }
      } : {})
    }
  };
}

const before = await api();
const beforeCounts = before.counts;
const beforeNonSeed = nonSeedState(before.lists);
const existingNumbers = new Set((before.lists.XRF_Items || [])
  .map(row => String(row?.fields?.field_1 || ""))
  .filter(code => code.startsWith(SEED_PREFIX))
  .map(code => Number(code.slice(SEED_PREFIX.length)))
  .filter(Number.isFinite));

const created = [];
const skipped = [];
for (let number = 1; number <= SEED_COUNT; number += 1) {
  if (existingNumbers.has(number)) {
    skipped.push(partNumber(number));
    continue;
  }
  const response = await api("POST", createPayload(number));
  created.push({
    partNumber: partNumber(number),
    itemId: response.result?.itemId,
    itemSpItemId: response.result?.itemSpItemId,
    requestId: response.result?.requestId,
    requestSpItemId: response.result?.requestSpItemId,
    measurementId: response.result?.measurement?.measurementId || null,
  });
  console.log(`[${number}/${SEED_COUNT}] created ${partNumber(number)}${number % 10 === 0 ? " (Pb remeasure)" : ""}`);
}

const after = await api();
const afterNonSeed = nonSeedState(after.lists);
const nonSeedChanges = compareNonSeedState(beforeNonSeed, afterNonSeed);
const seededItems = (after.lists.XRF_Items || []).filter(row => isSeedRow("XRF_Items", row));
const seededRequests = (after.lists.XRF_Requests || []).filter(row => isSeedRow("XRF_Requests", row));
const seededMeasurements = (after.lists.XRF_Measurements || []).filter(row => isSeedRow("XRF_Measurements", row));
const seededElements = (after.lists.XRF_ElementResults || []).filter(row => isSeedRow("XRF_ElementResults", row));
const remeasureCodes = seededMeasurements
  .filter(measurement => {
    const id = String(measurement?.fields?.Title || "");
    const pb = seededElements.find(row => row?.fields?.Title === id && row?.fields?.field_1 === "Pb");
    return String(pb?.fields?.field_4 || "").trim() === "??";
  })
  .map(row => row.fields.Title);

const result = {
  api: API_URL,
  beforeCounts,
  createdCount: created.length,
  skippedCount: skipped.length,
  afterCounts: after.counts,
  controlledRows: {
    items: seededItems.length,
    requests: seededRequests.length,
    measurements: seededMeasurements.length,
    elementResults: seededElements.length,
  },
  remeasureMeasurements: remeasureCodes,
  nonSeedChanges,
  pass: seededItems.length === SEED_COUNT
    && seededRequests.length === SEED_COUNT
    && seededMeasurements.length === 4
    && seededElements.length === 24
    && remeasureCodes.length === 4
    && nonSeedChanges.length === 0,
};

console.log(JSON.stringify(result, null, 2));
if (!result.pass) process.exitCode = 1;
