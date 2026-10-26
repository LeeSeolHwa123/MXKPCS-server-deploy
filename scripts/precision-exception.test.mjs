import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createRequire } from "node:module";
import { transformWithOxc } from "vite";

const require=createRequire(import.meta.url);
const projectRoot=path.resolve(path.dirname(fileURLToPath(import.meta.url)),"..");
const sourcePath=path.join(projectRoot,"app.jsx");
const apiPath=path.join(projectRoot,"api","xrf-sharepoint.js");
let source=await fs.readFile(sourcePath,"utf8");
const apiSource=await fs.readFile(apiPath,"utf8");
const reactUrl=pathToFileURL(require.resolve("react")).href;
const reactJsxRuntimeUrl=pathToFileURL(require.resolve("react/jsx-runtime")).href;
const xlsxUrl=pathToFileURL(require.resolve("xlsx")).href;
source=source
  .replace('from "react"',`from "${reactUrl}"`)
  .replace('from "xlsx"',`from "${xlsxUrl}"`);
source+=`
export {
  buildRuntimeFromSharePoint,runtimePrecisionCasesForItem,deriveRuntimeWorkflow,
  precisionCaseFlowKey,summarizePrecisionCases,approvalFilterKey,workflowStageLabel,
  itemXrfWorst,isPrecisionHandoffExceptionCase,precisionRowsForListMode
};
`;

const transformed=await transformWithOxc(source,sourcePath,{lang:"jsx",jsx:{runtime:"automatic"}});
const executableCode=transformed.code.replace('from "react/jsx-runtime"',`from "${reactJsxRuntimeUrl}"`);
const moduleUrl=`data:text/javascript;base64,${Buffer.from(executableCode).toString("base64")}`;
const app=await import(`${moduleUrl}#precision-exception`);

const measurementId="EXCEPTION_CASE_20261026_01";
const runtime=app.buildRuntimeFromSharePoint({lists:{
  XRF_Items:[{id:"1",fields:{Title:"ITM_EXCEPTION",field_1:"EX-01",field_2:"Exception test item",field_4:"Auxiliary Materials",field_5:"Common",field_8:"비접촉",field_9:"ExistingActive",field_10:true}}],
  XRF_Measurements:[{id:"10",fields:{Title:measurementId,field_1:"ITM_EXCEPTION",Measured_Date:"2026-10-26",field_3:"Initial",field_4:1,field_7:"RawData [[MXKPCS:Period_No=0]]"}}],
  XRF_ElementResults:[
    {id:"20",fields:{Title:measurementId,field_1:"Pb",field_2:"1200",field_3:"1",field_4:"Exceeded"}},
    {id:"21",fields:{Title:measurementId,field_1:"Hg",field_2:"ND",field_3:"1",field_4:"N.D."}},
    {id:"22",fields:{Title:measurementId,field_1:"Cr",field_2:"ND",field_3:"1",field_4:"N.D."}},
    {id:"23",fields:{Title:measurementId,field_1:"Cd",field_2:"ND",field_3:"1",field_4:"N.D."}},
    {id:"24",fields:{Title:measurementId,field_1:"Cl",field_2:"ND",field_3:"1",field_4:"N.D."}},
    {id:"25",fields:{Title:measurementId,field_1:"Br",field_2:"ND",field_3:"1",field_4:"N.D."}},
  ],
  XRF_Requests:[{id:"30",fields:{
    Title:`PREC_EXC_${measurementId}`,
    field_1:"PrecisionHandoffException",
    field_2:"Applied",
    field_3:"ITM_EXCEPTION",
    field_4:measurementId,
    field_5:"HAZARDOUS_SUBSTANCE_SCREENING",
    field_15:"유해물질 확인 목적의 조사성 측정",
    Processed_At:"2026-10-26T09:00:00Z",
    Processed_By:"admin",
    Before_Target_Snapshot_JSON:JSON.stringify({
      eventType:"PRECISION_HANDOFF_EXCEPTION",status:"ACTIVE",triggerMeasurementId:measurementId,
      itemId:"ITM_EXCEPTION",reasonCode:"HAZARDOUS_SUBSTANCE_SCREENING",
      reasonDetail:"유해물질 확인 목적의 조사성 측정",evidence:"APPROVAL-2026-001",
      createdAt:"2026-10-26T09:00:00Z",createdBy:"admin"
    })
  }}],
  XRF_PrecisionAnalyses:[],
  XRF_PrecisionElementResults:[],
}});

const failures=[];
const assert=(condition,message)=>{if(!condition)failures.push(message);};
const item=runtime.items[0];
const precisionCase=app.runtimePrecisionCasesForItem(item,runtime.precisionOverrides)[0];
const workflow=app.deriveRuntimeWorkflow(item,runtime.precisionOverrides,{});
const summary=app.summarizePrecisionCases([precisionCase]);

assert(runtime.precisionOverrides[measurementId]?.requestStatus==="EXEMPTED","DB exception must map to EXEMPTED instead of REQUESTED");
assert(runtime.precisionOverrides[measurementId]?.exceptionReasonCode==="HAZARDOUS_SUBSTANCE_SCREENING","exception reason code must survive DB mapping");
assert(precisionCase?.precision?.handoffException===true,"measurement case must retain the exception flag");
assert(app.precisionCaseFlowKey(precisionCase)==="exception","precision tab and risk flow must share the exception bucket");
assert(summary.total===1 && summary.exception===1 && summary.inProgress===0,"exception history must remain in totals but leave the in-progress count");
assert(workflow.stage==="PRECISION_HANDOFF_EXEMPTED","latest-item workflow must expose the exception stage");
assert(workflow.approval?.status==="미인계 예외","other tabs must receive the same exception status");
assert(app.approvalFilterKey(workflow.approval?.status)==="exception","material-list approval filters must map the exception consistently");
assert(app.workflowStageLabel(workflow.stage)==="정밀분석 미인계 예외","current-action labels must expose the exception state");
assert(app.itemXrfWorst(item)==="NG","exception must not overwrite the source XRF NG result");
assert(app.isPrecisionHandoffExceptionCase(precisionCase)===true,"exception case classifier must recognize the DB-mapped row");
assert(app.precisionRowsForListMode([precisionCase],"targets").length===0,"active precision target list must exclude exception cases");
assert(app.precisionRowsForListMode([precisionCase],"exceptions").length===1,"exception list must retain exception cases for later review");
assert(source.includes('ko:"예외처리 목록"'),"precision tab must expose a separate exception list selector");
assert(source.includes('setPrecisionQuickException({row:r,mode:"apply"})'),"precision target rows must expose a direct exception action");
assert(source.includes('setPrecisionQuickException({row:r,mode:"remove"})'),"exception rows must expose exception removal");
assert(source.includes('row?.precision?.exceptionReasonDetail'),"exception-list search must include the saved detailed reason");
assert(source.includes('role="switch" aria-checked={checked}'),"exception actions must use an accessible ON/OFF switch");
assert(source.includes('background:checked?C.charcoalDk:C.card'),"active exception switches must use the requested black styling");
assert(!source.includes('openExceptionDialog'),"precision detail must not retain an exception action");
assert(source.includes('if(isPrecisionHandoffExceptionCase(selectedCase))'),"exception cases must be redirected away from precision detail");
assert(!source.includes('기존과 동일하게 NG 발생 건이 자동으로 누적됩니다.'),"removed precision-list guidance must not remain");
assert(!source.includes('정밀분석 인계에서 제외된 건입니다. XRF NG와 위험도 이력은 유지됩니다.'),"removed exception-list guidance must not remain");

assert(apiSource.includes('action === "setPrecisionHandoffException"'),"API must expose the exception mutation action");
assert(apiSource.includes('field_1: "PrecisionHandoffException"'),"API must store exceptions as isolated XRF_Requests audit events");
assert(apiSource.includes('Before_Target_Snapshot_JSON: JSON.stringify(event)'),"API must persist the complete exception audit payload");
assert(apiSource.includes('if (activeException) throw new Error'),"precision transfer must reject an active exception until it is removed");

if(failures.length){
  console.error("Precision exception regression failures:");
  failures.forEach(message=>console.error(`- ${message}`));
  process.exit(1);
}
console.log(JSON.stringify({
  PASS:true,
  measurementId,
  xrfResult:app.itemXrfWorst(item),
  precisionFlow:app.precisionCaseFlowKey(precisionCase),
  workflowStage:workflow.stage,
  inProgress:summary.inProgress,
  exceptionHistory:summary.exception
},null,2));
