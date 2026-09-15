import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createRequire } from "node:module";
import { transformWithOxc } from "vite";

const require=createRequire(import.meta.url);
const projectRoot=path.resolve(path.dirname(fileURLToPath(import.meta.url)),"..");
const sourcePath=path.join(projectRoot,"app.jsx");
let source=await fs.readFile(sourcePath,"utf8");
const originalSource=source;
const reactUrl=pathToFileURL(require.resolve("react")).href;
const reactJsxRuntimeUrl=pathToFileURL(require.resolve("react/jsx-runtime")).href;
const xlsxUrl=pathToFileURL(require.resolve("xlsx")).href;
source=source
  .replace('from "react"',`from "${reactUrl}"`)
  .replace('from "xlsx"',`from "${xlsxUrl}"`);
source+=`
export {
  buildRuntimeFromSharePoint,applyPeriodBasedCategories,deriveReplacementRelations,applyCurrentRiskPolicy,
  itemXrfDisplayPeriods,itemMeasurementRecords,measurementById,measurementAttemptsForPeriod,
  periodMeasurementsOf,xrfTrendMeasurement,xrfTrendMeasurementEntries,xrfTrendSeries,
  xrfTrendPeriodsThroughSelection
};
`;

const transformed=await transformWithOxc(source,sourcePath,{lang:"jsx",jsx:{runtime:"automatic"}});
const executableCode=transformed.code.replace('from "react/jsx-runtime"',`from "${reactJsxRuntimeUrl}"`);
const moduleUrl=`data:text/javascript;base64,${Buffer.from(executableCode).toString("base64")}`;
const app=await import(`${moduleUrl}#trend`);

const elements=ppm=>({
  Pb:{rawPpm:String(ppm),ppm,rawSigma:"1",sigma:1,rawJudgement:"OK",judge:"OK",legalLimit:1000,internalLimit:700},
});
const measurements=Array.from({length:8},(_,index)=>({
  id:`M${index}`,
  date:`2026-${String(index+1).padStart(2,"0")}-10`,
  // SharePoint 과거 행처럼 Initial 이외에는 저장 periodNo가 없어도
  // period.measurementIds 연결만으로 올바른 Pn에 귀속돼야 한다.
  periodNo:index===0?0:null,
  role:index===0?"Initial":"Periodic",
  attemptNo:1,
  elements:elements(index*10),
}));
const retest={id:"M3-R",date:"2026-04-12",periodNo:null,role:"Retest",attemptNo:2,parentMeasurementId:"M3",elements:elements(35)};
const allMeasurements=[...measurements,retest];
const syntheticItem={
  code:"TREND-TEST",
  history:allMeasurements.map(row=>({...row,measurementRole:row.role})),
  __measurementMap:Object.fromEntries(allMeasurements.map(row=>[row.id,row])),
  firstMeasurementId:"M0",
  latestMeasurementId:"M7",
  latest:measurements[7],
};
const syntheticPeriods=measurements.map((measurement,index)=>({
  num:index,
  displaySeq:index,
  label:index===0?"R":`P${index}`,
  displayLabel:index===0?"R":`P${index}`,
  isReference:index===0,
  measurementId:measurement.id,
  measurementIds:[measurement.id],
  measured:measurement.date,
  measuredDates:[measurement.date],
}));

const failures=[];
const assert=(condition,message)=>{ if(!condition) failures.push(message); };
assert(originalSource.includes('lang==="en"?"XRF Trend by Element":"원소별 XRF 트렌드"'),"trend detail title must react to the selected language");
assert(originalSource.includes("XRF_TREND_EL_NAMES[el]?.[lang]"),"trend element names must react to the selected language");
assert(originalSource.includes("displayItemName(sel,lang)"),"trend detail must use the localized item name");
assert(originalSource.includes('No XRF measurement is linked to the selected period.'),"empty selected period must have an English message");
assert(originalSource.includes("{currentMeasurement\n                      ? <XrfTrendChart"),"trend chart must be hidden when the selected period has no measurement");
const syntheticEntries=app.xrfTrendMeasurementEntries(syntheticItem,syntheticPeriods);
assert(syntheticEntries.length===8,`management periods: expected 8, actual ${syntheticEntries.length}`);
assert(syntheticEntries.map(row=>row.label).join(",")==="R,P1,P2,P3,P4,P5,P6,P7","axis labels must remain R/Pn");
assert(!syntheticEntries.some(row=>row.label.includes("#")),"trend labels must not use attempt suffixes");
assert(app.measurementAttemptsForPeriod(syntheticItem,syntheticPeriods[0]).length===1,"null periodNo must not be treated as R");
assert(app.periodMeasurementsOf(syntheticItem,syntheticPeriods[3]).length===2,"P3 must expose original + retest");
assert(app.xrfTrendMeasurement(syntheticItem,syntheticPeriods[3])?.id==="M3-R","P3 representative must be the latest retest");
const syntheticPb=app.xrfTrendSeries(syntheticItem,syntheticPeriods,"Pb",syntheticEntries);
assert(syntheticPb.length===8,"trend must not cap the x-axis at five periods");
assert(syntheticPb.find(point=>point.measurementId==="M3-R")?.date==="2026-04-12","retest date must come from its Measurement");
assert(syntheticPb.find(point=>point.measurementId==="M3-R")?.ppm===35,"retest value must come from the same Measurement as its date");
assert(syntheticPb.every(point=>point.label),"every measurement must have an x-axis label");
const throughP3=app.xrfTrendPeriodsThroughSelection(syntheticPeriods,syntheticPeriods[3]);
assert(throughP3.map(period=>period.label).join(",")==="R,P1,P2,P3","selected P3 trend must exclude later periods");

const exportDir=path.join(projectRoot,"db_export");
const exportFiles=(await fs.readdir(exportDir)).filter(name=>/^XRF_SHAREPOINT_PRODUCTION_EXPORT_.*\.json$/.test(name)).sort();
const exportFile=exportFiles.at(-1);
assert(!!exportFile,"production SharePoint export is required");
let auditedItems=0;
let auditedMeasurements=0;
if(exportFile){
  const payload=JSON.parse(await fs.readFile(path.join(exportDir,exportFile),"utf8"));
  const runtime=app.buildRuntimeFromSharePoint(payload);
  const items=app.applyPeriodBasedCategories(
    app.deriveReplacementRelations(runtime.items).map(app.applyCurrentRiskPolicy)
  );
  for(const item of items){
    const records=app.itemMeasurementRecords(item);
    if(!records.length) continue;
    auditedItems++;
    auditedMeasurements+=records.length;
    const periods=app.itemXrfDisplayPeriods(item);
    const entries=app.xrfTrendMeasurementEntries(item,periods);
    const expectedEntryCount=periods.filter(period=>app.xrfTrendMeasurement(item,period)).length;
    assert(entries.length===expectedEntryCount,`${item.code}: trend period count ${entries.length} !== ${expectedEntryCount}`);
    assert(entries.every(entry=>entry.label===entry.baseLabel && !entry.label.includes("#")),`${item.code}: altered R/Pn label`);
    const reference=periods.find(period=>period?.isReference||Number(period?.num)===0);
    const referenceIds=new Set(reference?.measurementIds||[reference?.measurementId].filter(Boolean));
    const referenceAttempts=reference?app.measurementAttemptsForPeriod(item,reference):[];
    assert(referenceAttempts.every(row=>{
      const role=String(row?.measurementRole||row?.role||"");
      return role==="Initial" || (role==="Retest" && referenceIds.has(row?.parentMeasurementId));
    }),`${item.code}: non-reference measurement leaked into R`);
    const pb=app.xrfTrendSeries(item,periods,"Pb",entries);
    for(const point of pb){
      const measurement=app.measurementById(item,point.measurementId);
      assert(!!measurement,`${item.code}: unresolved ${point.measurementId}`);
      assert(String(point.date||"")===String(measurement?.date||""),`${item.code}/${point.measurementId}: date mismatch`);
    }
  }
}

console.log(JSON.stringify({
  syntheticMeasurements:syntheticEntries.length,
  productionExport:exportFile||null,
  auditedItems,
  auditedMeasurements,
  failures,
  pass:failures.length===0,
},null,2));
if(failures.length) process.exitCode=1;
