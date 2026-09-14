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
  buildRuntimeFromSharePoint,applyPeriodBasedCategories,deriveReplacementRelations,applyCurrentRiskPolicy,
  itemXrfDisplayPeriods,itemMeasurementRecords,measurementById,measurementAttemptsForPeriod,
  periodMeasurementsOf,xrfTrendMeasurement,xrfTrendMeasurementEntries,xrfTrendSeries
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
  periodNo:index,
  role:index===0?"Initial":"Periodic",
  attemptNo:1,
  elements:elements(index*10),
}));
const retest={id:"M3-R",date:"2026-04-12",periodNo:3,role:"Retest",attemptNo:2,parentMeasurementId:"M3",elements:elements(35)};
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
const syntheticEntries=app.xrfTrendMeasurementEntries(syntheticItem,syntheticPeriods);
assert(syntheticEntries.length===9,`all measurements: expected 9, actual ${syntheticEntries.length}`);
assert(new Set(syntheticEntries.map(row=>row.measurement.id)).size===9,"measurement IDs must be unique");
assert(app.periodMeasurementsOf(syntheticItem,syntheticPeriods[3]).length===2,"P3 must expose original + retest");
assert(app.xrfTrendMeasurement(syntheticItem,syntheticPeriods[3])?.id==="M3-R","P3 representative must be the latest retest");
const syntheticPb=app.xrfTrendSeries(syntheticItem,syntheticPeriods,"Pb",syntheticEntries);
assert(syntheticPb.length===9,"trend must not cap the x-axis at five points");
assert(syntheticPb.find(point=>point.measurementId==="M3-R")?.date==="2026-04-12","retest date must come from its Measurement");
assert(syntheticPb.find(point=>point.measurementId==="M3-R")?.ppm===35,"retest value must come from the same Measurement as its date");
assert(syntheticPb.every(point=>point.label),"every measurement must have an x-axis label");

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
    const recordIds=new Set(records.map(row=>String(row.id||row.measurementId)));
    const entryIds=new Set(entries.map(row=>String(row.measurement.id||row.measurement.measurementId)));
    assert(entryIds.size===recordIds.size,`${item.code}: trend ${entryIds.size} !== records ${recordIds.size}`);
    for(const id of recordIds) assert(entryIds.has(id),`${item.code}: missing trend measurement ${id}`);
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
