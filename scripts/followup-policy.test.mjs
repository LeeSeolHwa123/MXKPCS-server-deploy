import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createServer } from "vite";

const exportedPolicyFunctions = [
  "normalizeMeasurementRecord",
  "measurementXrfResult",
  "precisionRequiredElements",
  "pendingXrfRemeasureElements",
  "measurementFollowupInfo",
  "deriveInitialApprovalStatus",
  "deriveRuntimeWorkflow",
  "measurementAttemptsForPeriod",
];

const cacheDir = await mkdtemp(join(tmpdir(), "xrf-followup-policy-"));
const server = await createServer({
  appType: "custom",
  cacheDir,
  esbuild: { jsx: "transform", jsxFactory: "__testJsx", jsxFragment: "__testFragment" },
  logLevel: "silent",
  server: { middlewareMode: true },
  plugins: [{
    name: "followup-policy-test-exports",
    enforce: "pre",
    transform(code, id) {
      if (!id.replaceAll("\\", "/").endsWith("/app.jsx")) return null;
      const withoutRuntimeImports = code
        .replace(/^import .* from "react";\r?\n/m, "")
        .replace(/^import \* as XLSX from "xlsx";\r?\n/m, "const XLSX = {};\n");
      return `const __testJsx=()=>null; const __testFragment=Symbol("fragment");\n${withoutRuntimeImports}\nexport { ${exportedPolicyFunctions.join(", ")} };`;
    },
  }],
});

try {
  const policy = await server.ssrLoadModule("/app.jsx");
  const sourceElements = {
    Pb: { rawPpm: "??", rawSigma: "3.2", rawJudgement: "??" },
    Hg: { rawPpm: "80.7", rawSigma: "1.2", rawJudgement: "OK" },
    Cr: { rawPpm: "ND", rawSigma: "4.2", rawJudgement: "OK" },
    Cd: { rawPpm: "----", rawSigma: "17.9", rawJudgement: "----" },
    Cl: { rawPpm: "ND", rawSigma: "65.2", rawJudgement: "OK" },
    Br: { rawPpm: "890", rawSigma: "0.6", rawJudgement: "NG" },
  };
  const initial = policy.normalizeMeasurementRecord({
    id: "A-05_20260908_01",
    Item_ID: "ITM_A_05",
    date: "2026-09-08",
    role: "Initial",
    attemptNo: 1,
    elements: sourceElements,
  });
  const makeItem = latest => ({
    code: "A-05",
    itemId: "ITM_A_05",
    crLevel: "M",
    lifecycle: "NewItem",
    isCurrent: true,
    firstDate: initial.date,
    firstMeasurementId: initial.id,
    latestMeasurementId: latest.id,
    latest,
    __measurementMap: { [initial.id]: initial, [latest.id]: latest },
    history: [initial, ...(latest.id === initial.id ? [] : [latest])].map(m => ({
      ...m,
      measurementRole: m.role,
      periodNo: 0,
    })),
    periods: [{
      num: 0,
      label: "R",
      isReference: true,
      measurementId: initial.id,
      measurementIds: [initial.id],
      measured: initial.date,
      measuredDates: [initial.date],
      status: "reference",
    }],
  });

  assert.equal(policy.measurementXrfResult(initial), "NG");
  assert.deepEqual(policy.precisionRequiredElements(initial), ["Br"]);
  assert.deepEqual(policy.pendingXrfRemeasureElements(initial), ["Pb"]);
  assert.deepEqual(policy.measurementFollowupInfo(initial), {
    precisionElements: ["Br"],
    remeasureElements: ["Pb"],
    flaggedElements: ["Br", "Pb"],
    label: "정밀분석 및 XRF 재측정 필요",
    detail: "정밀분석: Br · XRF 재측정: Pb",
    hasFollowup: true,
  });
  assert.equal(policy.deriveInitialApprovalStatus(initial), "보류");

  const initialItem = makeItem(initial);
  assert.equal(policy.measurementAttemptsForPeriod(initialItem, initialItem.periods[0]).length, 1);
  assert.equal(policy.deriveRuntimeWorkflow(initialItem, {}, {}).stage, "XRF_REMEASURE_REQUIRED");

  const precisionOkBeforeRetest = {
    [initial.id]: {
      requestStatus: "REQUESTED",
      uploadStatus: "UPLOADED",
      elementResults: { Br: { presence: "미함유" } },
      finalConfirm: "확인",
    },
  };
  const blocked = policy.deriveRuntimeWorkflow(initialItem, precisionOkBeforeRetest, {});
  assert.equal(blocked.stage, "XRF_REMEASURE_REQUIRED");
  assert.equal(blocked.approval.status, "보류");

  const resolvedRetest = policy.normalizeMeasurementRecord({
    id: "A-05_20260909_02",
    Item_ID: "ITM_A_05",
    date: "2026-09-09",
    role: "Retest",
    attemptNo: 2,
    parentMeasurementId: initial.id,
    elements: { ...sourceElements, Pb: { rawPpm: "10", rawSigma: "1", rawJudgement: "OK" } },
  });
  assert.deepEqual(policy.pendingXrfRemeasureElements(resolvedRetest), []);
  assert.deepEqual(policy.precisionRequiredElements(resolvedRetest), ["Br"]);

  const resolvedItem = makeItem(resolvedRetest);
  const precisionOkAfterRetest = {
    [resolvedRetest.id]: {
      requestStatus: "REQUESTED",
      uploadStatus: "UPLOADED",
      elementResults: { Br: { presence: "미함유" } },
      finalConfirm: "확인",
    },
  };
  assert.equal(policy.deriveRuntimeWorkflow(resolvedItem, precisionOkAfterRetest, {}).stage, "APPROVED");

  const unresolvedRetest = policy.normalizeMeasurementRecord({
    id: "A-05_20260909_03",
    Item_ID: "ITM_A_05",
    date: "2026-09-09",
    role: "Retest",
    attemptNo: 2,
    parentMeasurementId: initial.id,
    elements: sourceElements,
  });
  assert.deepEqual(policy.pendingXrfRemeasureElements(unresolvedRetest), []);
  assert.deepEqual(policy.precisionRequiredElements(unresolvedRetest), ["Br", "Pb"]);
  assert.equal(policy.deriveInitialApprovalStatus(unresolvedRetest), "의뢰 필요");

  console.log("followup-policy: PASS");
} finally {
  await server.close();
  await rm(cacheDir, { recursive: true, force: true });
}
