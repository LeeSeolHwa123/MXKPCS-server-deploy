import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const app = await readFile(new URL("../app.jsx", import.meta.url), "utf8");
const api = await readFile(new URL("../api/xrf-sharepoint.js", import.meta.url), "utf8");

assert(api.includes('field_1: "PeriodSubstitution"'), "period substitution history must use a dedicated XRF_Requests event type");
assert(api.includes('Before_Target_Snapshot_JSON: JSON.stringify(event)'), "the complete substitution link must be persisted without changing the existing SharePoint schema");
assert(api.includes('eventType: "SUBSTITUTED"'), "substitution event type must be persisted explicitly");
assert(api.includes('linkedMeasurementId'), "substitution event must link to the actual substitute measurement");
assert(api.includes('if (!["Initial", "Periodic", "Retest"].includes(role))'), "substitution must not be stored as a fake XRF measurement role");
assert(api.includes('handleMigrateM02Cb100Substitution'), "the guarded M-02 / CB-100 migration must remain available");
assert(api.includes('handleCorrectM02Cb100PartNumber'), "the guarded M-02-01 correction must remain available");
assert(api.includes('const substitutePartNumber = "M-02-01"'), "CB-100 must use the M-02 replacement-family part number");
assert(api.includes('const newMeasurementId = "M-02-01_20260130_01"'), "the correction must update the linked measurement ID");
assert(api.includes('precisionRows.some(row => cleanText(row?.fields?.field_1).trim() === sourceMeasurementId)'), "migration must stop when precision references exist");

assert(app.includes('function activePeriodSubstitution(item,periodNo)'), "the UI must resolve a substitution by item and Pn");
assert(app.includes('status="substituted"'), "a substituted period must have its own terminal status");
assert(app.includes('substitutionEvent?[]:[...(explicitMeasurementsByPeriod.get(n)||[])]'), "the original item must not consume substitute measurements");
assert(app.includes('const substitutionMeasurement=periodSubstitutionMeasurement(p);'), "period details must resolve the linked substitute measurement");
assert(app.includes('function periodSubstitutionNotice(period,event,currentItem,lang="ko")'), "substitute periods must use one consistent explanatory sentence");
assert(app.includes('onOpenSubstituteResult={handleOpenSubstituteResult}'), "the compliance report cell must open the substitute item's own XRF screen");
assert(app.includes('if(measurementId && onOpenSubstituteResult) onOpenSubstituteResult(p,p.substitutionEvent,measurementId);'), "the substitute result button must pass the linked substitute event and measurement ID");
assert(app.includes('pendingXrfNavigationRef.current={code:substitute.code'), "cross-item navigation must preserve the linked measurement through the item-change reset");
assert(app.includes('setSelKey(substitute.code);'), "the substitute result button must switch the selected item to the substitute part number");
assert(app.indexOf(') : p.substitutionEvent ? (') < app.indexOf(') : sourceFileName ? ('), "substitute result navigation must remain available even when a report file name exists");
assert(app.includes('setXrfDetailView("analysis")'), "opening a substitute result must return to the measurement analysis view");
assert(app.includes('id="xrf-selected-measurement-detail"'), "the selected measurement details must expose a stable scroll target");
assert(app.includes('scrollIntoView({behavior:"smooth",block:"start"})'), "the substitute result button must visibly move to the selected measurement details");
assert(app.includes('currentSubstitutionEvent?"Substitute R"'), "the analysis summary must identify the linked result as substitute R");
assert(app.includes('currentMeasurement&&!currentSubstitutionEvent'), "the original item trend must exclude substitute XRF values");
assert(app.includes('&& !selected?.substitutionEvent'), "a closed substitution period must not allow another normal upload");
assert(app.includes('function orderMaterialListItemsByParent(items)'), "replacement and period-substitute rows must be grouped below their parent in the material list");
assert(app.includes('parent?.periodEvents||[]'), "temporary period substitutes must use PeriodSubstitution history for list placement");
assert(app.includes('function buildSharePointReplacementHistory(requestRows,items)'), "risk replacement history must read persisted request events");
assert(app.includes('key.includes("reactivat") || key.includes("revert")'), "replacement history must be ready to display future return-to-use events without implementing the reversion mutation");
assert(app.includes('kind:"periodSubstitution"'), "period-level substitutions must be included in the replacement history timeline");
assert(app.includes('const replacementHistoryGroups=useMemo'), "replacement history must group derived items under the same previous item");
assert(app.includes('const key=groupable?`derived:${entry.fromPartNumber}`'), "permanent replacement groups must use the previous part number as their parent key");
const riskSection=app.slice(app.indexOf('{tab==="risk"&&(()=>{'),app.indexOf('{tab==="reg"&&('));
assert(!riskSection.includes('<ChartCard title="측정 도래 예정"'), "risk overview must not duplicate the material-list upcoming-measurement chart");
assert(!riskSection.includes('<ChartCard title="이행 상태 구성"'), "risk overview must not duplicate the material-list compliance chart");
assert(!riskSection.includes('<ChartCard title="부서별 위험도 구성"'), "risk overview must not duplicate the material-list department-risk chart");
assert(riskSection.includes('<CycleManagementChart'), "risk overview must emphasize cycle management with a chart");
assert(riskSection.includes('<AnalysisFlowChart'), "risk overview must emphasize the analysis flow with a chart");
assert(riskSection.includes('replacementHistoryGroups.map((group,index)=>'), "risk overview must render one relation row per previous-item group");
assert(riskSection.includes('whiteSpace:"normal",wordBreak:"keep-all",overflowWrap:"break-word"'), "replacement reasons must wrap instead of being truncated");
assert(riskSection.includes('value:precisionFlowSummary.handoff'), "risk flow must use the same measurement-case summary as the precision tab");
assert(riskSection.includes('gotoPrecisionWithFlowFilter(step.filter)'), "risk flow steps must open the matching filtered precision cases");
assert(riskSection.includes('key:"upcoming"') && riskSection.includes('key:"notAvailable"'), "cycle status chart must include every material-list compliance state");
const flowChartSection=app.slice(app.indexOf('function AnalysisFlowChart'),app.indexOf('const SPLASH_BG_TOP'));
assert(!flowChartSection.includes('bottom:9,height:5'), "analysis flow must not render decorative bottom progress underlines");
assert(api.includes('handleCorrectS01ReplacementLifecycle'), "the guarded S-01 replacement lifecycle correction must remain available");

function sourceFunction(source,name){
  const start=source.indexOf(`function ${name}(`);
  assert(start>=0,`${name} source must exist`);
  const bodyStart=source.indexOf("{",start);
  let depth=0;
  for(let index=bodyStart;index<source.length;index+=1){
    if(source[index]==="{") depth+=1;
    if(source[index]==="}") depth-=1;
    if(depth===0) return source.slice(start,index+1);
  }
  throw new Error(`${name} source is incomplete`);
}
const orderMaterialListItemsByParent=Function(`return (${sourceFunction(app,"orderMaterialListItemsByParent")})`)();
const periodSubstitutionNotice=Function("periodLabel",`return (${sourceFunction(app,"periodSubstitutionNotice")})`)(period=>period?.displayLabel||`P${period?.num}`);
const precisionCaseFlowKey=Function(`return (${sourceFunction(app,"precisionCaseFlowKey")})`)();
const summarizePrecisionCases=Function("precisionCaseFlowKey",`return (${sourceFunction(app,"summarizePrecisionCases")})`)(precisionCaseFlowKey);
const precisionSummary=summarizePrecisionCases([
  {precision:{requestStatus:"NOT_REQUESTED"}},
  {precision:{requestStatus:"REQUESTED"}},
  {precision:{requestStatus:"REQUESTED",resultFileId:"report-1"}},
  {precision:{requestStatus:"REQUESTED",resultFileId:"report-2",result:"OK"},finalConfirm:"확인"},
  {precision:{requestStatus:"REQUESTED",resultFileId:"report-3",result:"NG"},finalConfirm:"확인"},
]);
assert.deepEqual(precisionSummary,{handoff:1,waiting:1,confirmation:1,ok:1,ng:1,exception:0,total:5,inProgress:4});
assert.equal(
  periodSubstitutionNotice(
    {num:1,displayLabel:"P1"},
    {substitutePartNumber:"M-02-01",substituteItemName:"CB-100 초음파세척액"},
    {code:"M-02"},
    "ko",
  ),
  "P1은 대체품 M-02-01(CB-100 초음파세척액) 사용 / 이 결과는 M-02 원소 추이 및 위험도에서 제외됩니다.",
);
const ordered=orderMaterialListItemsByParent([
  {code:"M-02",itemId:"ITM_M_02",periodEvents:[{eventType:"SUBSTITUTED",substituteItemId:"ITM_M_02_01"}]},
  {code:"X-01",itemId:"ITM_X_01"},
  {code:"M-02-01",itemId:"ITM_M_02_01",lifecycle:"Discontinued"},
  {code:"S-01-02",itemId:"ITM_S_01_02",replacementOf:"S-01"},
  {code:"S-01",itemId:"ITM_S_01"},
  {code:"S-01-01",itemId:"ITM_S_01_01",replacementOf:"S-01"},
]);
assert.deepEqual(ordered.map(item=>item.code),["M-02","M-02-01","X-01","S-01","S-01-02","S-01-01"]);

const event={eventType:"SUBSTITUTED",status:"ACTIVE",periodNo:1,effectiveDate:"2026-01-30",linkedMeasurementId:"M-02-01_20260130_01"};
const originalMeasurements=[
  {id:"M-02_20250519_01",role:"Initial",periodNo:0},
  {id:"M-02_20260710_01",role:"Periodic",periodNo:2},
  {id:"M-02_20261014_01",role:"Periodic",periodNo:3},
];
assert.equal(event.periodNo,1);
assert(!originalMeasurements.some(row=>row.periodNo===1), "M-02 P1 must be represented by the substitution event, not a Met Power measurement");
assert.equal(event.linkedMeasurementId,"M-02-01_20260130_01");

console.log("period substitution policy tests passed");
