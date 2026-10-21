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
  buildRollingPeriods,itemAllPeriods,itemListPeriods,currentCompliancePeriod,itemCycleFromRisk,
  isComplianceTarget,precisionCycleStopDate,itemHasConfirmedPrecisionResult
};
`;

const transformed=await transformWithOxc(source,sourcePath,{lang:"jsx",jsx:{runtime:"automatic"}});
const executableCode=transformed.code.replace('from "react/jsx-runtime"',`from "${reactJsxRuntimeUrl}"`);
const moduleUrl=`data:text/javascript;base64,${Buffer.from(executableCode).toString("base64")}`;
const app=await import(`${moduleUrl}#period-window-regression`);

const failures=[];
const checks=[];
const check=(condition,name,details={})=>{
  checks.push({name,pass:!!condition,...details});
  if(!condition) failures.push(name);
};
const iso=date=>`${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,"0")}-${String(date.getDate()).padStart(2,"0")}`;
const today=new Date();
today.setHours(12,0,0,0);
const addMonths=(value,months)=>{
  const date=new Date(`${value}T12:00:00`);
  const day=date.getDate();
  date.setDate(1);
  date.setMonth(date.getMonth()+months);
  const last=new Date(date.getFullYear(),date.getMonth()+1,0).getDate();
  date.setDate(Math.min(day,last));
  return iso(date);
};
const todayIso=iso(today);
const cycleMonths={Monthly:1,Semiannual:6,Annual:12};
const completeElements=ppm=>Object.fromEntries(["Pb","Hg","Cr","Cd","Cl","Br"].map(element=>[
  element,
  {rawPpm:String(ppm),ppm,rawSigma:"1",sigma:1,rawJudgement:ppm>=700?"NG":"OK",judge:ppm>=700?"NG":"OK",legalLimit:1000,internalLimit:700}
]));

const makeManagedItem=(cycle,completedCount=0)=>{
  const months=cycleMonths[cycle];
  const baseDate=addMonths(todayIso,-months*6);
  const history=[];
  for(let periodNo=1;periodNo<=completedCount;periodNo++){
    const measurement={
      id:`${cycle}-P${periodNo}`,
      date:addMonths(baseDate,months*periodNo),
      periodNo,
      role:"Periodic",
      measurementRole:"Periodic",
      attemptNo:1,
      elements:completeElements(100),
    };
    history.push(measurement);
  }
  return {
    code:`WINDOW-${cycle}-${completedCount}`,
    type:"Auxiliary Materials",
    lifecycle:"ExistingActive",
    category:"existing",
    isCurrent:true,
    cycle,
    firstDate:baseDate,
    periods:[{num:0,label:"R",displayLabel:"R",isReference:true,referenceDate:baseDate,date:baseDate,measured:baseDate,status:"reference"}],
    history,
    __measurementMap:Object.fromEntries(history.map(row=>[row.id,row])),
  };
};

const generationSummary={};
for(const cycle of Object.keys(cycleMonths)){
  const initial=makeManagedItem(cycle,0);
  const throughP3=makeManagedItem(cycle,3);
  const throughP5=makeManagedItem(cycle,5);
  const throughP6=makeManagedItem(cycle,6);
  const initialLabels=app.itemListPeriods(initial).map(period=>period.label);
  const p3Labels=app.itemListPeriods(throughP3).map(period=>period.label);
  const p5Labels=app.itemListPeriods(throughP5).map(period=>period.label);
  const p6Labels=app.itemListPeriods(throughP6).map(period=>period.label);

  check(JSON.stringify(initialLabels)===JSON.stringify(["R","P1","P2","P3"]),`${cycle}: initial four-dot window`,{actual:initialLabels});
  check(JSON.stringify(p3Labels)===JSON.stringify(["P3","P4","P5","P6"]),`${cycle}: fourth completed dot starts the next window`,{actual:p3Labels});
  check(JSON.stringify(p5Labels)===JSON.stringify(["P3","P4","P5","P6"]),`${cycle}: window remains stable before its fourth dot completes`,{actual:p5Labels});
  check(JSON.stringify(p6Labels)===JSON.stringify(["P6","P7","P8","P9"]),`${cycle}: second window boundary rolls by three`,{actual:p6Labels});

  const generated=app.buildRollingPeriods(throughP6);
  const targets=generated.filter(period=>!period.isReference);
  const futureOpen=targets.filter(period=>!period.measured && String(period.due)>=todayIso);
  const expectedFutureOpen=cycle==="Monthly"?24:4;
  check(futureOpen.length>=expectedFutureOpen,`${cycle}: required unmeasured future periods are retained`,{expectedAtLeast:expectedFutureOpen,actual:futureOpen.length});
  check(targets.every((period,index)=>period.label===`P${index+1}`),`${cycle}: generated Pn labels are continuous`);
  check(app.itemListPeriods(throughP6).length===4,`${cycle}: material-list window never loses a slot`);
  check(app.itemListPeriods(throughP6).length<=4,`${cycle}: material-list window never exceeds four slots`);
  generationSummary[cycle]={
    initial:initialLabels,
    afterP3:p3Labels,
    afterP5:p5Labels,
    afterP6:p6Labels,
    generatedThrough:targets.at(-1)?.label||null,
    futureOpen:futureOpen.length,
  };
}

const pendingItem=makeManagedItem("Monthly",3);
const pendingPeriods=app.itemAllPeriods(pendingItem);
check(app.precisionCycleStopDate(pendingItem)===null,"precision pending: no stop date");
check(app.isComplianceTarget(pendingItem),"precision pending: cycle remains active");
check(!!app.currentCompliancePeriod(pendingItem),"precision pending: next report target remains available");
check(pendingPeriods.some(period=>period.label==="P4"),"precision pending: future Pn continues");

const precisionItem=(result)=>{
  const item=makeManagedItem("Monthly",3);
  const precision={result,finalConfirm:"확인",confirmedAt:todayIso};
  item.__workflow={
    precision,
    precisionCases:[{precision,finalConfirm:"확인",measurement:{date:todayIso}}],
    triggerMeasurement:{date:todayIso},
  };
  return item;
};
const okItem=precisionItem("OK");
const ngItem=precisionItem("NG");
check(app.itemHasConfirmedPrecisionResult(okItem),"precision OK: result is detected");
check(app.precisionCycleStopDate(okItem)===null,"precision OK: no stop date");
check(app.isComplianceTarget(okItem),"precision OK: cycle remains active");
check(app.itemAllPeriods(okItem).some(period=>period.label==="P4"),"precision OK: future Pn continues");
check(app.precisionCycleStopDate(ngItem)===todayIso,"precision NG: stop date is confirmation date");
check(!app.isComplianceTarget(ngItem),"precision NG: cycle target stops");
check(app.currentCompliancePeriod(ngItem)===null,"precision NG: no next report target remains");
check(app.itemAllPeriods(ngItem).every(period=>period.isReference||period.measured||String(period.due)<=todayIso),"precision NG: no unmeasured Pn after confirmation survives");

const notAllowedPeriodic=[
  {id:"NA-P1",date:"2026-01-30",role:"Periodic",attemptNo:1,elements:completeElements(850)},
  {id:"NA-P2",date:"2026-07-14",role:"Periodic",attemptNo:1,elements:completeElements(850)},
];
const notAllowedItem={
  code:"C-07-SYNTHETIC",
  type:"Auxiliary Materials",
  lifecycle:"ExistingActive",
  category:"existing",
  isCurrent:true,
  finalRisk:"Not Allowed",
  firstDate:"2025-05-19",
  latestMeasurementId:notAllowedPeriodic[1].id,
  latest:notAllowedPeriodic[1],
  periods:[{num:0,label:"R",displayLabel:"R",isReference:true,referenceDate:"2025-05-19",date:"2025-05-19",measured:"2025-05-19",status:"reference"}],
  history:notAllowedPeriodic.map(row=>({...row,measurementRole:row.role})),
  __measurementMap:Object.fromEntries(notAllowedPeriodic.map(row=>[row.id,row])),
};
check(app.itemCycleFromRisk(notAllowedItem)==="Semiannual","Not Allowed: observed six-month cadence is preserved");
check(app.isComplianceTarget(notAllowedItem),"Not Allowed: pending precision result remains managed");
check(app.itemListPeriods(notAllowedItem).length===4,"Not Allowed: four dots are displayed");

const exportDir=path.join(projectRoot,"db_export");
const exportFiles=await fs.readdir(exportDir).catch(error=>error?.code==="ENOENT"?[]:Promise.reject(error));
const latestExport=exportFiles.filter(name=>/^XRF_SHAREPOINT_PRODUCTION_EXPORT_.*\.json$/.test(name)).sort().at(-1)||null;
let productionAudit={file:latestExport,activeWithReference:0,fourDotRows:0,maxDotCount:0,c07:null};
if(latestExport){
  const payload=JSON.parse(await fs.readFile(path.join(exportDir,latestExport),"utf8"));
  const runtime=app.buildRuntimeFromSharePoint(payload);
  const items=app.applyPeriodBasedCategories(app.deriveReplacementRelations(runtime.items).map(app.applyCurrentRiskPolicy));
  for(const item of items){
    const periods=app.buildRollingPeriods(item);
    const hasReference=periods.some(period=>period.isReference||Number(period.num)===0);
    if(!app.isComplianceTarget(item)||!hasReference) continue;
    const dots=app.itemListPeriods(item);
    productionAudit.activeWithReference++;
    productionAudit.maxDotCount=Math.max(productionAudit.maxDotCount,dots.length);
    if(dots.length===4) productionAudit.fourDotRows++;
    check(dots.length===4,`production ${item.code}: active row has four dots`,{actual:dots.map(period=>period.label)});
    check(dots.length<=4,`production ${item.code}: dot count does not exceed four`,{actual:dots.length});
  }
  const c07=items.find(item=>item.code==="C-07")||null;
  if(c07){
    productionAudit.c07={
      cycle:app.itemCycleFromRisk(c07),
      dots:app.itemListPeriods(c07).map(period=>period.label),
      generatedThrough:app.itemAllPeriods(c07).at(-1)?.label||null,
      nextDue:app.currentCompliancePeriod(c07)?.due||null,
    };
    check(productionAudit.c07.cycle==="Semiannual","production C-07: cycle is Semiannual",productionAudit.c07);
    check(productionAudit.c07.dots.length===4,"production C-07: four dots are displayed",productionAudit.c07);
  }
}

const report={
  pass:failures.length===0,
  today:todayIso,
  checks:checks.length,
  failures,
  generation:generationSummary,
  precision:{pending:"continues",OK:"continues",NG:"stops"},
  productionAudit,
};
console.log(JSON.stringify(report,null,2));
if(failures.length) process.exitCode=1;
