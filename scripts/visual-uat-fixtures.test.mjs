import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createRequire } from "node:module";
import { transformWithOxc } from "vite";

const require=createRequire(import.meta.url);
const projectRoot=path.resolve(path.dirname(fileURLToPath(import.meta.url)),"..");
const sourcePath=path.join(projectRoot,"app.jsx");
let source=await fs.readFile(sourcePath,"utf8");
const reactUrl=pathToFileURL(require.resolve("react")).href;
const reactJsxRuntimeUrl=pathToFileURL(require.resolve("react/jsx-runtime")).href;
const xlsxUrl=pathToFileURL(require.resolve("xlsx")).href;
source=source
  .replace('from "react"',`from "${reactUrl}"`)
  .replace('from "xlsx"',`from "${xlsxUrl}"`);
source+=`
export {
  VISUAL_UAT_MODE,VISUAL_UAT_ITEMS,VISUAL_UAT_COMPLIANCE_ITEMS,VISUAL_UAT_PAGINATION_ITEMS,VISUAL_UAT_ALL_ITEMS,VISUAL_UAT_PRECISION_OVERRIDES,VISUAL_UAT_EXPECTATIONS,
  applyPeriodBasedCategories,deriveReplacementRelations,applyCurrentRiskPolicy,
  deriveRuntimeWorkflow,visualUatSelfCheckRow,workflowForItem,analysisWorkflowSummary,
  MATERIAL_LIST_PAGE_SIZE,isCancelledListItem,approvalStatusOfItemFallback,translateUiTextKoToEn
};
`;

const transformed=await transformWithOxc(source,sourcePath,{lang:"jsx",jsx:{runtime:"automatic"}});
const executableCode=transformed.code.replace('from "react/jsx-runtime"',`from "${reactJsxRuntimeUrl}"`);
const moduleUrl=`data:text/javascript;base64,${Buffer.from(executableCode).toString("base64")}`;
const app=await import(`${moduleUrl}#normal`);
globalThis.window={location:{search:"?uat=1"}};
const uatApp=await import(`${moduleUrl}#uat`);
delete globalThis.window;
const visualBase=app.applyPeriodBasedCategories(
  app.deriveReplacementRelations(app.VISUAL_UAT_ITEMS).map(app.applyCurrentRiskPolicy)
);
const visualItems=visualBase.map(item=>({
  ...item,
  __workflow:app.deriveRuntimeWorkflow(item,app.VISUAL_UAT_PRECISION_OVERRIDES,{})
}));
const byCode=new Map(visualItems.map(item=>[item.code,item]));
const approvalStatusOf=item=>app.workflowForItem(item)?.approval?.status||item?.approvalStatus||"측정 진행중";
const rows=app.VISUAL_UAT_EXPECTATIONS.map(manifest=>{
  const item=byCode.get(manifest.code);
  return item
    ? app.visualUatSelfCheckRow(item,manifest,approvalStatusOf)
    : {code:manifest.code,group:manifest.group,PASS:false,failedFields:["fixture missing"],expected:manifest.expected,actual:null};
});

const expectedGroups={XRF:17,Risk:9,Cycle:6,Period:7,Workflow:10,Lifecycle:6};
const actualGroups=Object.fromEntries(Object.keys(expectedGroups).map(group=>[
  group,app.VISUAL_UAT_EXPECTATIONS.filter(row=>row.group===group).length
]));
const structuralFailures=[];
const workflowItems=visualItems.filter(item=>item.code.startsWith("VU-W"));
const workflowSummary=app.analysisWorkflowSummary(workflowItems);
const expectedWorkflowSummary={
  xrfRemeasure:1,
  precisionTransferRequired:2,
  precisionResultWaiting:1,
  precisionFollowupRequired:0,
  precisionNgReview:1,
  total:5
};
for(const [key,expected] of Object.entries(expectedWorkflowSummary)){
  if(workflowSummary[key]!==expected) structuralFailures.push(`workflow summary ${key} ${workflowSummary[key]} !== ${expected}`);
}
const transferCompatibilitySummary=app.analysisWorkflowSummary([
  {__workflow:{stage:"PRECISION_TRANSFER_REQUEST_REQUIRED"}},
  {__workflow:{stage:"PRECISION_REQUEST_REQUIRED"}}
]);
if(transferCompatibilitySummary.precisionTransferRequired!==2 || transferCompatibilitySummary.total!==2){
  structuralFailures.push("precision transfer/request workflow stages must share the same dashboard bucket");
}
if(app.translateUiTextKoToEn("사진 선택 (JPG, PNG, WEBP / 최대 10MB)")!=="Choose Photo (JPG, PNG, WEBP / max. 10 MB)"){
  structuralFailures.push("photo picker translation mismatch");
}
if(app.VISUAL_UAT_MODE!==false) structuralFailures.push("SSR/normal mode must disable Visual UAT");
if(uatApp.VISUAL_UAT_MODE!==true) structuralFailures.push("?uat=1 must enable Visual UAT");
if(visualItems.length!==55) structuralFailures.push(`fixture count ${visualItems.length} !== 55`);
if(app.VISUAL_UAT_EXPECTATIONS.length!==55) structuralFailures.push(`manifest count ${app.VISUAL_UAT_EXPECTATIONS.length} !== 55`);
if(app.VISUAL_UAT_COMPLIANCE_ITEMS.length!==20) structuralFailures.push(`compliance fixture count ${app.VISUAL_UAT_COMPLIANCE_ITEMS.length} !== 20`);
if(app.VISUAL_UAT_PAGINATION_ITEMS.length!==40) structuralFailures.push(`pagination fixture count ${app.VISUAL_UAT_PAGINATION_ITEMS.length} !== 40`);
if(app.VISUAL_UAT_ALL_ITEMS.length!==115) structuralFailures.push(`total Visual UAT fixture count ${app.VISUAL_UAT_ALL_ITEMS.length} !== 115`);
const paginationRemeasureItems=app.VISUAL_UAT_PAGINATION_ITEMS.filter(item=>app.workflowForItem(item).stage==="XRF_REMEASURE_REQUIRED");
if(paginationRemeasureItems.length!==4) structuralFailures.push(`pagination remeasure fixture count ${paginationRemeasureItems.length} !== 4`);
const visibleVisualItems=app.VISUAL_UAT_ALL_ITEMS.filter(item=>
  !app.isCancelledListItem(item,app.approvalStatusOfItemFallback(item))
);
const visualPages=Array.from(
  {length:Math.ceil(visibleVisualItems.length/app.MATERIAL_LIST_PAGE_SIZE)},
  (_,page)=>visibleVisualItems.slice(page*app.MATERIAL_LIST_PAGE_SIZE,(page+1)*app.MATERIAL_LIST_PAGE_SIZE)
);
const pageRemeasureCounts=visualPages.map(page=>page.filter(item=>app.workflowForItem(item).stage==="XRF_REMEASURE_REQUIRED").length);
if(app.MATERIAL_LIST_PAGE_SIZE!==100) structuralFailures.push(`material list page size ${app.MATERIAL_LIST_PAGE_SIZE} !== 100`);
if(visibleVisualItems.length!==114) structuralFailures.push(`visible Visual UAT fixture count ${visibleVisualItems.length} !== 114`);
if(visualPages.length!==2 || visualPages[0]?.length!==100 || visualPages[1]?.length!==14){
  structuralFailures.push(`Visual UAT pages ${visualPages.map(page=>page.length).join("/")} !== 100/14`);
}
if(pageRemeasureCounts.some(count=>count<1)) structuralFailures.push(`remeasure fixtures missing by page: ${pageRemeasureCounts.join("/")}`);
if(byCode.size!==visualItems.length) structuralFailures.push("fixture codes are not unique");
for(const item of visualItems){
  if(!item.code.startsWith("VU-")) structuralFailures.push(`${item.code}: invalid code prefix`);
  if(!String(item.name||"").startsWith("[UAT-VIS]")) structuralFailures.push(`${item.code}: invalid name prefix`);
  if(item.__visualUat!==true || item.__readOnlyFixture!==true) structuralFailures.push(`${item.code}: missing read-only flags`);
}
for(const [group,expected] of Object.entries(expectedGroups)){
  if(actualGroups[group]!==expected) structuralFailures.push(`${group} count ${actualGroups[group]} !== ${expected}`);
}

const failed=rows.filter(row=>!row.PASS);
console.log(JSON.stringify({
  fixtureCount:visualItems.length,
  totalVisualFixtureCount:app.VISUAL_UAT_ALL_ITEMS.length,
  paginationFixtureCount:app.VISUAL_UAT_PAGINATION_ITEMS.length,
  paginationRemeasureCount:paginationRemeasureItems.length,
  visibleVisualFixtureCount:visibleVisualItems.length,
  pageLengths:visualPages.map(page=>page.length),
  pageRemeasureCounts,
  manifestCount:app.VISUAL_UAT_EXPECTATIONS.length,
  modeGate:{normal:app.VISUAL_UAT_MODE,uat:uatApp.VISUAL_UAT_MODE},
  groups:actualGroups,
  passed:rows.length-failed.length,
  failed:failed.length,
  structuralFailures,
  failedCases:failed.map(row=>({code:row.code,failedFields:row.failedFields,expected:row.expected,actual:row.actual}))
},null,2));

if(failed.length || structuralFailures.length) process.exitCode=1;
