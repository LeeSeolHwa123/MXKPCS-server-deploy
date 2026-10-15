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
  buildRollingPeriods,itemAllPeriods,currentCompliancePeriod,itemCycleFromRisk,isComplianceTarget,
  precisionCycleStopDate,itemHasConfirmedPrecisionResult,
  runtimePrecisionCasesForItem,deriveRuntimeWorkflow,
  itemXrfDisplayPeriods,itemMeasurementRecords,measurementById,measurementAttemptsForPeriod,
  periodMeasurementsOf,xrfTrendMeasurement,xrfTrendMeasurementEntries,xrfTrendSeries,
  xrfTrendPeriodsThroughSelection,normalizeItemRecord,displayItemName
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
assert(/\{currentMeasurement\r?\n\s+\? <XrfTrendChart/.test(originalSource),"trend chart must be hidden when the selected period has no measurement");
assert(originalSource.includes('["itemNameEn","Item_Name_EN"'),"new-item forms must accept an English item name");
assert(originalSource.includes('nameEn:String(reqForm.itemNameEn||"").trim()'),"new-item creation must preserve the English item name");
assert(originalSource.includes('window.alert("Item_Name_EN을 입력하세요.")'),"new-item submission must reject a missing English item name");
assert(originalSource.includes('background:"#343A40"'),"element trend header must use #343A40");
assert(originalSource.includes("const canUploadPrecisionPdf=isRequested && hasSourceFile"),"precision PDF upload must follow transfer and source-file readiness");
assert(originalSource.includes('data-upload-arrow="true"'),"precision uploads must show an upload arrow");
assert(originalSource.includes('strokeWidth="1.35"'),"precision upload arrows must use a thin stroke");
assert(!originalSource.includes('background:enabled?"rgba(25,49,167,.32)"'),"precision upload arrows must not use a dark-blue circular background");
assert(originalSource.includes('gridTemplateColumns:"24px minmax(0,1fr) 24px"'),"precision upload button labels must remain optically centered");
assert(!originalSource.includes('canEditPrecisionResult?"결과 입력 가능"'),"precision detail header must not show the result-entry-ready label");
assert(originalSource.includes('border:"1px solid #343A40",borderRadius:6,background:"#343A40"'),"element trend button must match the compact edit-button shape and #343A40 color");
assert(!originalSource.includes("품목 등록·변경"),"registration/change labels must consistently use a slash");
assert(originalSource.includes('l:"품목 / 변경 관리"'),"registration/change tab must use the requested label");
assert(!originalSource.includes("Lifecycle = "),"request type cards must not display internal lifecycle codes");
assert(!originalSource.includes("의뢰자 등록은 먼저 요청 유형을 선택한 뒤"),"request type explanation must be removed");
assert(originalSource.includes('const registrationButtonStyle='),"registration navigation and action buttons must share one style helper");
assert(originalSource.includes('measurementSavePending||registrationSavePending?"SharePoint 저장 중...":"등록"'),"admin submit action must be labelled 등록");
const s11=app.normalizeItemRecord({code:"S-11",name:"S-05 접촉 영향 확인용 간지",nameEn:""});
assert(s11.nameEn==="Interleaf for S-05 Contact Impact Verification","S-11 must recover its missing legacy English name");
assert(app.displayItemName(s11,"en")===s11.nameEn,"English UI must prefer the S-11 English name");
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

const testLegalLimits={Pb:1000,Hg:1000,Cr:1000,Cd:100,Cl:900,Br:900,"Cl+Br":1500};
const completeElements=pbPpm=>Object.fromEntries(Object.entries(testLegalLimits).map(([element,legalLimit])=>[
  element,{rawPpm:String(element==="Pb"?pbPpm:10),rawSigma:"1",rawJudgement:"OK",legalLimit,internalLimit:legalLimit*.7}
]));
const historicalPrecisionMeasurement={id:"PREC-HISTORY-01",date:"2026-01-10",role:"Periodic",attemptNo:1,elements:completeElements(800)};
const latestOkMeasurement={id:"PREC-LATEST-OK",date:"2026-02-10",role:"Periodic",attemptNo:1,elements:completeElements(100)};
const accumulatedPrecisionItem={
  code:"PRECISION-ACCUMULATED",
  history:[historicalPrecisionMeasurement,latestOkMeasurement].map(row=>({...row,measurementRole:row.role})),
  __measurementMap:{
    [historicalPrecisionMeasurement.id]:historicalPrecisionMeasurement,
    [latestOkMeasurement.id]:latestOkMeasurement,
  },
  firstMeasurementId:historicalPrecisionMeasurement.id,
  latestMeasurementId:latestOkMeasurement.id,
  latest:latestOkMeasurement,
};
const accumulatedPrecisionCases=app.runtimePrecisionCasesForItem(accumulatedPrecisionItem,{});
const accumulatedPrecisionWorkflow=app.deriveRuntimeWorkflow(accumulatedPrecisionItem,{},{});
assert(accumulatedPrecisionCases.length===1,"an unresolved historical precision case must remain visible after a newer OK XRF");
assert(accumulatedPrecisionCases[0]?.key===historicalPrecisionMeasurement.id,"a precision case must keep its trigger Measurement_ID as the DB key");
assert(accumulatedPrecisionWorkflow.precisionCases.length===1,"workflow output must retain accumulated historical precision cases");
assert(accumulatedPrecisionWorkflow.activePrecisionCaseKey===null,"a historical case must not replace the latest XRF as the current approval basis");
assert(accumulatedPrecisionWorkflow.stage==="APPROVED","a newer OK XRF must keep the existing current-approval behavior");

const nextPeriodNgMeasurement={id:"PREC-NEXT-PERIOD-NG",date:"2026-03-10",role:"Periodic",attemptNo:1,elements:completeElements(850)};
const repeatedPrecisionItem={
  ...accumulatedPrecisionItem,
  history:[historicalPrecisionMeasurement,latestOkMeasurement,nextPeriodNgMeasurement].map(row=>({...row,measurementRole:row.role})),
  __measurementMap:{
    [historicalPrecisionMeasurement.id]:historicalPrecisionMeasurement,
    [latestOkMeasurement.id]:latestOkMeasurement,
    [nextPeriodNgMeasurement.id]:nextPeriodNgMeasurement,
  },
  latestMeasurementId:nextPeriodNgMeasurement.id,
  latest:nextPeriodNgMeasurement,
};
const firstPrecisionOkOverride={
  [historicalPrecisionMeasurement.id]:{
    requestStatus:"REQUESTED",uploadStatus:"UPLOADED",resultFileId:"precision-ok.pdf",finalConfirm:"확인",
    elementResults:{Pb:{presence:"미함유",ppm:""}},
  },
};
const repeatedPrecisionCases=app.runtimePrecisionCasesForItem(repeatedPrecisionItem,firstPrecisionOkOverride);
const repeatedPrecisionWorkflow=app.deriveRuntimeWorkflow(repeatedPrecisionItem,firstPrecisionOkOverride,{});
assert(repeatedPrecisionCases.length===2,"a later NG measurement must append a new precision case after an earlier precision OK");
assert(repeatedPrecisionCases.some(row=>row.key===historicalPrecisionMeasurement.id && row.precision.result==="OK"),"the earlier confirmed precision OK must remain linked to its Measurement_ID");
assert(repeatedPrecisionCases.some(row=>row.key===nextPeriodNgMeasurement.id && row.precision.requestStatus==="NOT_REQUESTED"),"the later untransferred precision case must appear automatically from XRF measurement history");
assert(repeatedPrecisionWorkflow.activePrecisionCaseKey===nextPeriodNgMeasurement.id,"the latest NG measurement must become the active precision case without replacing older cases");

const registeredCycleItem=(cycle,referenceDate="2025-01-15")=>({
  code:`CYCLE-${cycle}`,
  type:"Auxiliary Materials",
  lifecycle:"ExistingActive",
  category:"existing",
  isCurrent:false,
  approvalStatus:"반려",
  precisionRequired:true,
  finalRisk:"Not Allowed",
  cycle,
  firstDate:referenceDate,
  periods:[{
    num:0,label:"R",displayLabel:"R",isReference:true,
    referenceDate,date:referenceDate,measured:referenceDate,status:"reference"
  }],
  history:[]
});
const cycleExpectations=[
  ["Monthly","2025-02-15","2025-03-15"],
  ["Semiannual","2025-07-15","2026-01-15"],
  ["Annual","2026-01-15","2027-01-15"],
];
for(const [cycle,p1Due,p2Due] of cycleExpectations){
  const item=registeredCycleItem(cycle);
  const periods=app.buildRollingPeriods(item);
  const targets=periods.filter(period=>!period.isReference);
  assert(app.itemCycleFromRisk(item)===cycle,`${cycle}: registered cycle must survive rejected/precision state`);
  assert(app.isComplianceTarget(item),`${cycle}: rejected/precision item with a cycle must remain an XRF measurement target`);
  assert(app.itemAllPeriods(item).length===periods.length,`${cycle}: Is_Current=false must not truncate a registered active lifecycle`);
  assert(targets[0]?.label==="P1" && targets[0]?.due===p1Due,`${cycle}: P1 due date mismatch`);
  assert(targets[1]?.label==="P2" && targets[1]?.due===p2Due,`${cycle}: P2 due date mismatch`);
  assert(targets.every((period,index)=>period.label===`P${index+1}`),`${cycle}: report periods must grow as continuous P1..Pn rows`);
  assert(!!app.currentCompliancePeriod(item),`${cycle}: a report-upload target period must remain available`);
}

const annualSameYearMeasurements=[
  {id:"ANNUAL_20260120_01",date:"2026-01-20",role:"Periodic",attemptNo:1,elements:elements(20)},
  {id:"ANNUAL_20260714_01",date:"2026-07-14",role:"Periodic",attemptNo:1,elements:elements(30)},
  {id:"ANNUAL_20260714_02",date:"2026-07-14",role:"Periodic",attemptNo:1,elements:elements(40)},
];
const annualSameYearItem={
  ...registeredCycleItem("Annual"),
  history:annualSameYearMeasurements.map(row=>({...row,measurementRole:row.role})),
  __measurementMap:Object.fromEntries(annualSameYearMeasurements.map(row=>[row.id,row])),
  latestMeasurementId:annualSameYearMeasurements.at(-1).id,
  latest:annualSameYearMeasurements.at(-1),
};
const annualSameYearPeriods=app.buildRollingPeriods(annualSameYearItem);
const annualP1=annualSameYearPeriods.find(period=>period.label==="P1");
const annualP2=annualSameYearPeriods.find(period=>period.label==="P2");
assert(annualP1?.measurementIds?.length===3,"Annual: multiple standalone measurements in the same calendar year must stay in one Pn");
assert(annualP1?.measurementId==="ANNUAL_20260714_02","Annual: the highest same-day Measurement_ID sequence must be the representative measurement");
assert(!annualP2?.measured,"Annual: measurements made before the next due year must not consume the next Pn");

for(const precisionResult of ["OK","NG"]){
  const precision={result:precisionResult,finalConfirm:"확인",confirmedAt:"2025-04-20"};
  const confirmedItem={
    ...registeredCycleItem("Monthly"),
    isCurrent:true,
    __workflow:{
      precision,
      precisionCases:[{precision,finalConfirm:"확인",measurement:{date:"2025-03-10"}}],
      triggerMeasurement:{date:"2025-03-10"}
    }
  };
  const retained=app.itemAllPeriods(confirmedItem);
  const retainedTargets=retained.filter(period=>!period.isReference);
  assert(app.itemHasConfirmedPrecisionResult(confirmedItem),`${precisionResult}: entered precision result must be detected`);
  if(precisionResult==="OK"){
    assert(app.precisionCycleStopDate(confirmedItem)===null,"OK: confirmed precision result must not set a cycle stop date");
    assert(app.isComplianceTarget(confirmedItem),"OK: confirmed precision result must keep active cycle management");
    assert(!!app.currentCompliancePeriod(confirmedItem),"OK: the next XRF report input period must remain active");
    assert(retainedTargets.some(period=>period.label==="P4"),"OK: future periods must continue to be generated");
  }else{
    assert(app.precisionCycleStopDate(confirmedItem)==="2025-04-20","NG: precision stop date mismatch");
    assert(!app.isComplianceTarget(confirmedItem),"NG: confirmed precision result must stop active cycle management");
    assert(app.currentCompliancePeriod(confirmedItem)===null,"NG: no further XRF report input period may remain active");
    assert(retainedTargets.some(period=>period.label==="P3"),"NG: elapsed period history must be retained");
    assert(!retainedTargets.some(period=>period.label==="P4"),"NG: future periods after precision confirmation must not be generated");
  }
}

const isoAtMonthOffset=offset=>{
  const now=new Date();
  const value=new Date(now.getFullYear(),now.getMonth()+offset,10);
  return `${value.getFullYear()}-${String(value.getMonth()+1).padStart(2,"0")}-10`;
};
const recentMonthly=app.buildRollingPeriods(registeredCycleItem("Monthly",isoAtMonthOffset(-1)));
const olderMonthly=app.buildRollingPeriods(registeredCycleItem("Monthly",isoAtMonthOffset(-36)));
assert(olderMonthly.length>recentMonthly.length,"Monthly report rows must increase as more cycle periods elapse");
assert(originalSource.includes('["주기","기준·도래일","측정일","상태","XRF 결과","XRF 보고서 첨부"].map'),"every generated period table must retain the XRF report input column");

const exportDir=path.join(projectRoot,"db_export");
const exportFiles=await fs.readdir(exportDir).catch(error=>{
  if(error?.code==="ENOENT") return [];
  throw error;
});
const productionExportFiles=exportFiles.filter(name=>/^XRF_SHAREPOINT_PRODUCTION_EXPORT_.*\.json$/.test(name)).sort();
const exportFile=productionExportFiles.at(-1);
let auditedItems=0;
let auditedMeasurements=0;
let auditedPrecisionCases=0;
if(exportFile){
  const payload=JSON.parse(await fs.readFile(path.join(exportDir,exportFile),"utf8"));
  const runtime=app.buildRuntimeFromSharePoint(payload);
  const productionS11=runtime.items.find(item=>item.code==="S-11");
  assert(productionS11?.nameEn==="Interleaf for S-05 Contact Impact Verification","production S-11 must receive the English-name fallback");
  const items=app.applyPeriodBasedCategories(
    app.deriveReplacementRelations(runtime.items).map(app.applyCurrentRiskPolicy)
  );
  for(const item of items){
    const records=app.itemMeasurementRecords(item);
    if(!records.length) continue;
    auditedItems++;
    auditedMeasurements+=records.length;
    auditedPrecisionCases+=app.runtimePrecisionCasesForItem(item,runtime.precisionOverrides).length;
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
  auditedPrecisionCases,
  failures,
  pass:failures.length===0,
},null,2));
if(failures.length) process.exitCode=1;
