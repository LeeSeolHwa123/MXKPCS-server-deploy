import { useState, useMemo, useEffect, useLayoutEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

const TODAY = new Date();
TODAY.setHours(0,0,0,0);
const TODAY_STR = [
  TODAY.getFullYear(),
  String(TODAY.getMonth()+1).padStart(2,"0"),
  String(TODAY.getDate()).padStart(2,"0")
].join("-");

// Visual UAT fixtures are local, URL-gated display records. They are never sent to SharePoint.
const VISUAL_UAT_MODE =
  typeof window !== "undefined"
  && new URLSearchParams(window.location.search).get("uat") === "1";
const VISUAL_UAT_READ_ONLY_MESSAGE = "Visual UAT 테스트 데이터는 읽기 전용입니다.";
const VISUAL_UAT_MUTATION_ACTIONS = new Set([
  "createRequest","saveMeasurement","updateItem","processDiscontinue",
  "revertDiscontinue","upsertPrecision","deleteRecord",
  "beginPrecisionFileUpload","uploadPrecisionFileChunk","finishPrecisionFileUpload",
  "beginItemPhotoUpload","finishItemPhotoUpload"
]);

function containsVisualUatReference(value,depth=0){
  if(depth>5 || value==null) return false;
  if(typeof value==="string") return value.startsWith("VU-") || value.includes("VISUAL_UAT::VU-");
  if(Array.isArray(value)) return value.some(row=>containsVisualUatReference(row,depth+1));
  if(typeof value==="object"){
    if(value.__visualUat===true || value.__readOnlyFixture===true) return true;
    return Object.values(value).some(row=>containsVisualUatReference(row,depth+1));
  }
  return false;
}

// Prototype-only admin gate for item master editing.
// 실제 운영에서는 프론트엔드 상수 대신 로그인/권한 서버 검증으로 교체해야 합니다.
const ADMIN_ITEM_EDIT_PASSWORD = "admin1234";

const C = {
  // Molex red + PowerFX legend colors + neutral gray UI surface.
  red:"#E60A32", redDk:"#8C0A22", redBg:"#FCE8EC",
  charcoal:"#323C3C", charcoalDk:"#1A2020", charcoalMd:"#5C6A6A",
  charcoalLt:"#8B95A1", charcoalBg:"#EDF0F0",
  xOk:"#16A34A", xOkBg:"#DCFCE7",
  xWarn:"#D97706", xWarnBg:"#FEF3C7",
  blue:"#1D5F99", blueBd:"#BBD4EA", blueBg:"#DFF0FF",
  orange:"#D97706", orangeBd:"#F3C998", orangeBg:"#FFEDDB",
  text1:"#1A2020", text2:"#333333", text3:"#6B6B6B", text4:"#8B95A1",
  // Open Color Gray 0 — 기본 중립 배경(surface) 색상
  bg:"#F8F9FA", card:"#FFFFFF", alt:"#F8F9FA", bd:"#EFEFF1", bd2:"#DFDFE3",
};

// 첨부 레퍼런스(pill 버튼 / 상태칩)에서 가져온 파스텔 액센트 팔레트입니다.
// 각 항목은 라인(ln) · 배경(bg) · 글자(tx) 3색 한 세트로 구성해
// 버튼, 상태 배지, 차트 색상을 모두 같은 규칙으로 씁니다.
const A = {
  // 첨부 타임라인 UI의 색상만 사용합니다.
  // 보라(주요/진행) · 초록(완료) · 주황(주의) · 하늘(보조) · 회색(중립)
  purple: {ln:"#CFCBF7", bg:"#F2F0FE", tx:"#5A4FC4", solid:"#A9A2F0"},
  green:  {ln:"#A9E6C1", bg:"#EAF9F0", tx:"#1E8A50", solid:"#3FC177"},
  amber:  {ln:"#FFD79B", bg:"#FFF5E4", tx:"#B0700F", solid:"#FFAA33"},
  blue:   {ln:"#A9DCF3", bg:"#EAF6FD", tx:"#1F7CA6", solid:"#6BC6EC"},
  rose:   {ln:"#F3BDBD", bg:"#FDEFEF", tx:"#BF4A42", solid:"#EC8079"},
  slate:  {ln:"#DDE1E5", bg:"#F4F6F7", tx:"#5F686D", solid:"#AEB6BC"},
};

// 폰트: 영문·숫자·한글 모두 Pretendard 하나로 통일합니다.
const FONT_SANS = "'Pretendard Variable','Pretendard','Apple SD Gothic Neo','Malgun Gothic',system-ui,sans-serif";
const FONT_NUM  = FONT_SANS;

// 타이포 계층. 화면 전체가 같은 규칙을 쓰도록 한 곳에 모읍니다.
//   pageTitle  화면 제목        18 / 700
//   cardTitle  카드·패널 제목   15 / 600
//   section    구역 이름        12 / 600  (XRF 분석 탭의 모든 구역 헤딩)
//   label      필드 라벨        11 / 500
//   body       본문             12 / 400
//   caption    보조 설명        11 / 400
//   micro      단위·꼬리표      10 / 400
const T = {
  pageTitle: {fontSize:18, fontWeight:700, color:"#1A2020", letterSpacing:"-.2px"},
  cardTitle: {fontSize:15, fontWeight:600, color:"#1A2020", letterSpacing:"-.1px"},
  section:   {fontSize:12, fontWeight:600, color:"#333333", letterSpacing:0},
  label:     {fontSize:11, fontWeight:500, color:"#6B6B6B"},
  body:      {fontSize:12, fontWeight:400, color:"#333333"},
  caption:   {fontSize:11, fontWeight:400, color:"#6B6B6B"},
  micro:     {fontSize:10, fontWeight:400, color:"#8B95A1"},
};

// 레이아웃/디자인 토큰 (프레젠테이션 전용, 로직 비관여)
const UI = {
  r:  14,                                  // 카드 라운드
  rs: 10,                                  // 소형 컨트롤 라운드
  pill: 999,                               // pill 버튼/칩 라운드
  sh:  "0 1px 2px rgba(26,32,32,.04), 0 4px 16px rgba(26,32,32,.05)",
  shHover: "0 2px 6px rgba(26,32,32,.07), 0 8px 26px rgba(26,32,32,.08)",
  rowH: 56,
  hoverRow: "#F8F9FA",
  selRow:   "#FFF4F6",
};

const ELEMENTS = ["Pb","Hg","Cr","Cd","Cl","Br","Cl+Br"];
// XRF의 OK/NG는 원본 보고서의 Legal Limit이 아니라 사내 관리기준으로 판정합니다.
// 사내 관리기준(Internal Limit)은 각 원소 Legal Limit의 70%이며 Content(ppm)와 비교합니다.
// 원본 Judgment는 추적용 원본값과 재측정(?? / Cannot Judge) 식별에 사용합니다.
const XRF_REPORT_ELEMENTS = ["Pb","Hg","Cr","Cd","Cl","Br","Cl+Br"];
const XRF_ELEMENT_LIMITS = {Pb:1000, Hg:1000, Cr:1000, Cd:100, Cl:900, Br:900, "Cl+Br":1500};
const XRF_INTERNAL_LIMIT_RATIO = 0.70;
const XRF_INTERNAL_LIMITS = Object.fromEntries(
  Object.entries(XRF_ELEMENT_LIMITS).map(([el,legal])=>[el,Math.round(legal*XRF_INTERNAL_LIMIT_RATIO*10)/10])
);
const XRF_AXIS_MAX = 1500;

const XRF_NUMERIC_MISSING_MARKERS = new Set(["", "----", "---", "--", "-", "N/A", "NA", "NULL", "NONE"]);
const XRF_ND_MARKERS = new Set(["ND", "N.D.", "N.D", "NOT DETECTED", "NON DETECTED", "불검출"]);
// XRF OK/NG 판정의 source-of-truth는 Content(ppm)와 Internal Limit(법적 기준의 70%) 비교입니다.
// 원본 Judgment는 원본 추적과 ??/Cannot Judge 재측정 신호에 사용하며, Content의 dash는 2개 이상일 때만 N.D.로 인정합니다.
const XRF_JUDGEMENT_DASH_ND_RE = /^[\-–—−]+$/;

function normalizeXrfToken(v){
  return v == null ? "" : String(v).trim().toUpperCase();
}
function isXrfMissing(v){
  return v == null || XRF_NUMERIC_MISSING_MARKERS.has(normalizeXrfToken(v));
}
function isXrfJudgementDashND(v){
  const raw=String(v??"").trim();
  return XRF_JUDGEMENT_DASH_ND_RE.test(raw);
}
function isXrfContentDashND(v){
  // Content의 단일 '-'는 결측으로 유지하고, 보고서의 연속 dash 표기(---- 등)만 N.D.로 봅니다.
  const raw=String(v??"").trim();
  return /^[\-–—−]{2,}$/.test(raw);
}
function displayXrfEmpty(v="---"){
  return <span style={{fontSize:12,color:C.text4,fontWeight:600}}>{v}</span>;
}
function displayXrfText(v){
  return v==null || v==="" || v==="—" ? "---" : v;
}

// 가변 데이터 값 전용 자동 축소 텍스트.
// 고정 헤더/라벨/설명문에는 사용하지 않고, 품번·품목명·상태·날짜·측정값처럼
// 데이터에 따라 길이가 달라지는 값만 현재 기본 글씨 크기에서 필요한 만큼 축소합니다.
function AutoFitText({value, children, baseFontSize=11, minFontSize=7, style={}, title, align="inherit"}){
  const ref=useRef(null);
  const content=value!==undefined ? value : children;
  const dependency=typeof content==="string" || typeof content==="number" ? String(content) : "";

  useLayoutEffect(()=>{
    const el=ref.current;
    if(!el) return;
    let raf=0;
    let disposed=false;
    const maxSize=Number(baseFontSize)||11;
    const minSize=Math.min(maxSize,Number(minFontSize)||7);

    const fitNow=()=>{
      if(disposed || !el) return;
      cancelAnimationFrame(raf);
      raf=requestAnimationFrame(()=>{
        if(disposed || !el) return;
        el.style.fontSize=`${maxSize}px`;
        const available=el.clientWidth;
        if(!available) return;
        const natural=el.scrollWidth;
        if(natural<=available+0.5) return;

        // 우선 실제 초과 비율만큼 한 번에 줄인 뒤 소수점 단위로 미세 조정합니다.
        let next=Math.max(minSize,Math.floor((maxSize*(available/Math.max(natural,1)))*4)/4);
        el.style.fontSize=`${next}px`;
        while(next>minSize && el.scrollWidth>el.clientWidth+0.5){
          next=Math.max(minSize,Math.round((next-0.25)*100)/100);
          el.style.fontSize=`${next}px`;
        }
      });
    };

    fitNow();

    // 글씨 크기 조정 자체로 요소 높이가 바뀌어도 ResizeObserver가 반복 호출되지 않도록
    // 실제 너비가 변한 경우에만 다시 계산합니다.
    let observedWidth=el.clientWidth;
    const ro=typeof ResizeObserver!=="undefined" ? new ResizeObserver(entries=>{
      const width=entries?.[0]?.contentRect?.width ?? el.clientWidth;
      if(Math.abs(width-observedWidth)<=0.5) return;
      observedWidth=width;
      fitNow();
    }) : null;
    if(ro) ro.observe(el);
    const onResize=()=>fitNow();
    if(!ro && typeof window!=="undefined") window.addEventListener("resize",onResize);
    if(typeof document!=="undefined" && document.fonts?.ready) document.fonts.ready.then(fitNow).catch(()=>{});

    return ()=>{
      disposed=true;
      cancelAnimationFrame(raf);
      ro?.disconnect();
      if(!ro && typeof window!=="undefined") window.removeEventListener("resize",onResize);
    };
  },[dependency,baseFontSize,minFontSize]);

  return <span ref={ref} title={title||dependency||undefined} style={{
    display:"block",width:"100%",maxWidth:"100%",minWidth:0,boxSizing:"border-box",
    // minFontSize까지 줄여도 극단적으로 긴 값이 남는 경우 인접 셀을 침범하지 않도록
    // 최종 안전장치로 overflow를 차단합니다. 정상 데이터는 글씨 축소만으로 전부 표시됩니다.
    whiteSpace:"nowrap",overflow:"hidden",textOverflow:"clip",textAlign:align,lineHeight:1.2,
    fontSize:baseFontSize,...style
  }}>{content}</span>;
}
function isXrfND(v){
  return XRF_ND_MARKERS.has(normalizeXrfToken(v));
}
function isXrfContentND(v){
  return isXrfND(v) || isXrfContentDashND(v);
}
function xrfNumber(v){
  if(isXrfContentND(v)) return 0;
  if(isXrfMissing(v)) return null;
  if(typeof v === "number") return Number.isFinite(v) ? v : null;
  const matched=String(v).replace(/,/g, "").match(/[-+]?\d+(?:\.\d+)?/);
  return matched ? Number(matched[0]) : null;
}
function toNum(v){
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}
function sourcePpmValue(d){
  return d && Object.prototype.hasOwnProperty.call(d,"rawPpm") ? d.rawPpm : d?.ppm;
}
function sourceSigmaValue(d){
  return d && Object.prototype.hasOwnProperty.call(d,"rawSigma") ? d.rawSigma : d?.sigma;
}
function sourceJudgementValue(d){
  return d?.rawJudgement ?? d?.sourceJudgement ?? d?.judge ?? "";
}
function elementDataMissing(d){
  const ppm=sourcePpmValue(d);
  return !!d?.dataMissing || (isXrfMissing(ppm) && !isXrfContentND(ppm));
}
function xrfContentMissing(d){
  const ppm=sourcePpmValue(d);
  return isXrfMissing(ppm) && !isXrfContentND(ppm);
}
function xrfLegalLimitFromData(d){
  const explicit=xrfNumber(d?.legalLimit);
  if(explicit!=null) return explicit;
  // 기존 데이터의 limit 필드는 Legal Limit으로 저장되어 있으므로 호환성을 위해 그대로 읽습니다.
  return xrfNumber(d?.limit);
}
function xrfInternalLimitFromData(d){
  const explicit=xrfNumber(d?.internalLimit);
  if(explicit!=null) return explicit;
  const legal=xrfLegalLimitFromData(d);
  return legal==null ? null : Math.round(legal*XRF_INTERNAL_LIMIT_RATIO*10)/10;
}
function xrfLevelFromJudgement(d){
  if(!d || isXrfRemeasureRequired(d) || xrfContentMissing(d)) return null;
  const rawPpm=sourcePpmValue(d);
  const rawJudgement=sourceJudgementValue(d);
  // Content 자체가 N.D. 또는 2개 이상 연속 dash이면 불검출(L)로 봅니다.
  // Content 단일 '-'는 N.D.가 아니라 결측으로 처리합니다.
  if(isXrfContentND(rawPpm)) return "L";
  const ppm=xrfNumber(rawPpm);
  const internalLimit=xrfInternalLimitFromData(d);
  if(ppm==null || internalLimit==null) return null;
  // Content가 숫자 0이고 원본 Judgment가 N.D.인 구형 데이터만 L로 보존합니다.
  if(ppm===0 && isXrfND(rawJudgement)) return "L";
  // 사내기준 초과는 H, 검출값이 있으나 사내기준 이내이면 M입니다.
  return ppm>internalLimit ? "H" : "M";
}
function xrfJudgementDecision(d){
  if(!d || isXrfRemeasureRequired(d) || xrfContentMissing(d)) return null;
  const rawPpm=sourcePpmValue(d);
  if(isXrfContentND(rawPpm)) return "OK";
  const ppm=xrfNumber(rawPpm);
  const internalLimit=xrfInternalLimitFromData(d);
  if(ppm==null || internalLimit==null) return null;
  // OK/NG의 실제 판정 기준: Content <= Legal Limit의 70% → OK, 초과 → NG.
  return ppm<=internalLimit ? "OK" : "NG";
}
function isXrfRemeasureRequired(d){
  const raw=normalizeXrfToken(sourceJudgementValue(d));
  return raw==="??" || raw==="※" || raw==="CANNOT JUDGE" || raw==="CANNOT BE JUDGED";
}
function buildClBrResult(clRaw, brRaw){
  const clND=isXrfContentND(clRaw);
  const brND=isXrfContentND(brRaw);
  const cl=xrfNumber(clRaw);
  const br=xrfNumber(brRaw);
  const legalLimit=XRF_ELEMENT_LIMITS["Cl+Br"];
  const internalLimit=XRF_INTERNAL_LIMITS["Cl+Br"];
  if(cl==null || br==null){
    return {ppm:null, judgement:"", level:null, legalLimit, internalLimit};
  }
  const sum=Math.round((cl+br)*10)/10;
  if(clND && brND){
    const display=(isXrfContentDashND(clRaw) && isXrfContentDashND(brRaw)) ? "----" : "ND";
    return {ppm:display, judgement:"N.D.", level:"L", isND:true, legalLimit, internalLimit};
  }
  // Cl+Br도 사내기준을 적용합니다: Legal 1,500 ppm × 70% = Internal Limit 1,050 ppm.
  if(sum<=internalLimit){
    return {ppm:sum, judgement:"Less than", level:"M", legalLimit, internalLimit};
  }
  return {ppm:sum, judgement:"Exceeded", level:"H", legalLimit, internalLimit};
}
function axisPct(v){
  const n=xrfNumber(v);
  return n==null ? 0 : Math.max(0, Math.min((n / XRF_AXIS_MAX) * 100, 100));
}
const DEPTS = ["Assembly","Common","Molding","Plating","Stamping"];

// 마스터 DB의 Type 값을 그대로 사용합니다.
// Facility는 최초 XRF 이력만 관리하고 정기 P주기 대상에서는 제외합니다.
const ITEM_TYPE_OPTIONS = ["Auxiliary Materials","Facility"];
const ITEM_TYPE_KO = {"Auxiliary Materials":"부자재","Facility":"설비"};
const REQUESTER_MATERIAL_CATEGORY_OPTIONS = [
  "Metal","Plated Metal","Plastic","Rubber","Glass","Ceramics","Paper","Paint/Ink",
  "Adhesive","Silicon","Solder wire/paste/flux","Textile","Chemical","Vinyl","Tape","Label","Others"
];
const CHEMICAL_MATERIAL_STATE_OPTIONS = [
  "Liquid","Volatile Liquid","Viscous Liquid","Aerosol","Gas"
];
function isChemicalMaterialCategory(value){
  return String(value||"").trim().toLowerCase()==="chemical";
}
function isValidChemicalMaterialState(value){
  return CHEMICAL_MATERIAL_STATE_OPTIONS.includes(String(value||"").trim());
}
const CYCLE_KO = {Monthly:"월 1회", Semiannual:"반기 1회", Annual:"연 1회", "Not Available":"—"};
const CR_TYPE_KO = {"직접접촉+잔류":"직접접촉·잔류","직접접촉+비잔류":"직접접촉·비잔류","비접촉":"비접촉"};
const CR_LEVEL_KO = {H:"H", M:"M", L:"L"};

const CAT_META = {
  existing: {dot:C.charcoal,   label:"기존 품목"},
  new:      {dot:C.charcoalMd, label:"신규 등록"},
  changed:  {dot:C.charcoalLt, label:"변경·대체"},
};

const LIFECYCLE_KO = {
  ExistingActive: "기존 운영",
  NewReplacement: "신규 대체",
  ReplacedOld: "대체됨(이력만)",
  Discontinued: "단종",
  NewItem: "신규 도입",
};

const ST = {
  compliant:{fill:C.charcoal, border:C.charcoal, label:"이행"},
  measured: {fill:C.charcoal, border:C.charcoal, label:"측정 완료"},
  early:    {fill:C.charcoalLt, border:C.charcoalLt, label:"조기측정"},
  late:     {fill:C.redDk, border:C.redDk, label:"지연측정"},
  overdue:  {fill:C.red, border:C.red, label:"미이행"},
  dueSoon:  {fill:C.card, border:C.red, label:"측정권장"},
  due_soon: {fill:C.card, border:C.red, label:"측정권장"},
  upcoming: {fill:C.card, border:C.bd2, label:"예정"},
  reference:{fill:C.charcoal, border:C.charcoal, label:"기준"},
};

const FILTER_LABELS = {
  category: {existing:"기존 품목", new:"신규 등록", changed:"변경·대체"},
  code: {},
  name: {},
  dept: {},
  cycle: {Monthly:"월 1회", Semiannual:"반기 1회", Annual:"연 1회", "Not Available":"—"},
  firstDate: {},
  xrf: {NG:"NG · 정밀분석", OK:"OK", "재측정 필요":"재측정", "확인 필요":"결과 확인", H:"H · Exceeded", M:"M · Less than", L:"L · N.D.", "—":"—"},
  compliance: {overdue:"미이행", dueSoon:"측정권장", ok:"이행", upcoming:"예정", notAvailable:"대상 아님"},
  precision: {OK:"OK", NG:"NG", "—":"—"},
  crLevel: {H:"H · 직접접촉+잔류", M:"M · 직접접촉+비잔류", L:"L · 비접촉"},
  risk: {"Not Allowed":"사용불허", H:"H", M:"M", L:"L", "—":"—"},
  retest: {review:"재측정", trusted:"재측정 불필요"},
  approval: {approved:"사용/단종 승인", measuring:"측정 진행중", request:"의뢰 필요", rejected:"사용/단종 반려", hold:"보류/검토중", cancelled:"처리 취소"},
  lifecycle: {ExistingActive:"기존 운영", NewReplacement:"신규 대체", ReplacedOld:"대체됨", Discontinued:"단종", NewItem:"신규 도입", Cancelled:"처리 취소"},
  nextDue: {},
  dDay: {},
  dueMonth: {},
};
const COL_NAMES = {category:"분류", code:"품번", name:"품목명", dept:"부서", cycle:"주기", firstDate:"최초 등록", xrf:"XRF 결과", compliance:"현재 상태", precision:"정밀분석", crLevel:"C&R 등급", risk:"Risk", retest:"XRF 후속조치", approval:"승인 상태", lifecycle:"라이프사이클", nextDue:"다음 마감", dDay:"D-DAY", dueMonth:"도래월"};
const UNIQUE_VALS = {
  category: ["existing","new","changed"],
  dept: ["Assembly","Common","Molding","Plating","Stamping"],
  cycle: ["Monthly","Semiannual","Annual","Not Available"],
  xrf: ["NG","OK","재측정 필요","확인 필요","—"],
  compliance: ["overdue","dueSoon","ok","upcoming","notAvailable"],
  crLevel: ["H","M","L"],
  risk: ["Not Allowed","H","M","L","—"],
  retest: ["review","trusted"],
  approval: ["approved","measuring","request","rejected","hold","cancelled"],
  lifecycle: ["ExistingActive","NewReplacement","ReplacedOld","Discontinued","NewItem","Cancelled"],
};

// 초기 마이그레이션 원본부터 Item_Name_EN이 비어 있던 품목의 표시 보정값입니다.
// SharePoint 원본값이 추가되면 원본 영문명이 항상 이 값보다 우선합니다.
const ITEM_NAME_EN_FALLBACKS = Object.freeze({
  "S-11":"Interleaf for S-05 Contact Impact Verification",
});

// SharePoint DB는 브라우저 번들에 하드코딩하지 않습니다.
// 브라우저는 동일 호스트의 /api/xrf-sharepoint만 호출하고,
// Microsoft Graph용 Client Secret은 서버 환경변수에만 보관합니다.
const XRF_SHAREPOINT_API = "/api/xrf-sharepoint";

async function fetchSharePointBootstrap(){
  const response=await fetch(XRF_SHAREPOINT_API,{
    method:"GET",
    headers:{"Accept":"application/json"},
    cache:"no-store"
  });
  let payload=null;
  try{ payload=await response.json(); }catch{}
  if(!response.ok || !payload?.ok){
    const detail=payload?.error || payload?.message || `HTTP ${response.status}`;
    throw new Error(`SharePoint DB 조회 실패: ${detail}`);
  }
  return payload;
}

async function mutateSharePoint(action,payload={}){
  const response=await fetch(XRF_SHAREPOINT_API,{
    method:"POST",
    headers:{"Accept":"application/json","Content-Type":"application/json"},
    cache:"no-store",
    body:JSON.stringify({action,payload})
  });
  let result=null;
  try{ result=await response.json(); }catch{}
  if(!response.ok || !result?.ok){
    const detail=result?.error || result?.message || `HTTP ${response.status}`;
    throw new Error(`SharePoint DB 저장 실패: ${detail}`);
  }
  return result?.result ?? result;
}

async function fetchPrecisionStoredFiles(measurementId){
  const response=await fetch(`${XRF_SHAREPOINT_API}?precisionFilesFor=${encodeURIComponent(measurementId)}`,{cache:"no-store"});
  const payload=await response.json();
  if(!response.ok || !payload?.ok) throw new Error(payload?.error||`파일 조회 HTTP ${response.status}`);
  return payload.files||{source:null,report:null};
}

async function uploadPrecisionBinary(measurementId,kind,file,onProgress){
  if(!file?.size || file.size>250*1024*1024) throw new Error("파일은 1바이트 이상, 250MB 이하만 업로드할 수 있습니다.");
  let session;
  for(let attempt=0;attempt<3;attempt++){
    try{
      session=await mutateSharePoint("beginPrecisionFileUpload",{measurementId,kind,fileName:file.name,size:file.size});
      break;
    }catch(error){
      if(attempt===2 || !String(error?.message||"").includes("정밀분석 인계가 DB에 저장된 뒤")) throw error;
      await new Promise(resolve=>setTimeout(resolve,600));
    }
  }
  let offset=0;
  let driveItemId="";
  while(offset<file.size){
    const end=Math.min(offset+session.chunkBytes,file.size);
    const bytes=new Uint8Array(await file.slice(offset,end).arrayBuffer());
    let binary="";
    for(let i=0;i<bytes.length;i+=32768) binary+=String.fromCharCode(...bytes.subarray(i,i+32768));
    const chunk=await mutateSharePoint("uploadPrecisionFileChunk",{
      uploadUrl:session.uploadUrl,offset,totalSize:file.size,base64:btoa(binary)
    });
    offset=end;
    onProgress?.(Math.round(offset/file.size*100));
    if(chunk.complete){
      if(offset!==file.size || !chunk.driveItemId) throw new Error("업로드 완료 파일 ID를 확인하지 못했습니다.");
      driveItemId=chunk.driveItemId;
    }else if(offset===file.size || !chunk.nextExpectedRanges?.some(range=>Number(String(range).split("-")[0])===offset)){
      throw new Error("Graph가 예상한 업로드 위치가 일치하지 않습니다. 파일 저장 상태를 확인하세요.");
    }
  }
  return mutateSharePoint("finishPrecisionFileUpload",{
    measurementId,storedName:session.storedName,driveItemId
  });
}

async function uploadItemPhotoBinary(itemId,file,onProgress,itemSpItemId=null){
  if(!file?.size || file.size>10*1024*1024) throw new Error("품목 사진은 1바이트 이상, 10MB 이하만 업로드할 수 있습니다.");
  if(!/^image\/(jpeg|png|webp)$/i.test(file.type||"") && !/\.(jpe?g|png|webp)$/i.test(file.name||"")){
    throw new Error("품목 사진은 JPG, PNG, WEBP 형식만 업로드할 수 있습니다.");
  }
  const session=await mutateSharePoint("beginItemPhotoUpload",{itemId,itemSpItemId,fileName:file.name,size:file.size,mimeType:file.type||""});
  let offset=0;
  let driveItemId="";
  while(offset<file.size){
    const end=Math.min(offset+session.chunkBytes,file.size);
    const bytes=new Uint8Array(await file.slice(offset,end).arrayBuffer());
    let binary="";
    for(let i=0;i<bytes.length;i+=32768) binary+=String.fromCharCode(...bytes.subarray(i,i+32768));
    const chunk=await mutateSharePoint("uploadPrecisionFileChunk",{
      uploadUrl:session.uploadUrl,offset,totalSize:file.size,base64:btoa(binary)
    });
    offset=end;
    onProgress?.(Math.round(offset/file.size*100));
    if(chunk.complete){
      if(offset!==file.size || !chunk.driveItemId) throw new Error("사진 업로드 완료 파일 ID를 확인하지 못했습니다.");
      driveItemId=chunk.driveItemId;
    }else if(offset===file.size || !chunk.nextExpectedRanges?.some(range=>Number(String(range).split("-")[0])===offset)){
      throw new Error("Graph가 예상한 사진 업로드 위치와 일치하지 않습니다.");
    }
  }
  return mutateSharePoint("finishItemPhotoUpload",{itemId,itemSpItemId,storedName:session.storedName,driveItemId});
}

function requestStatusFromApproval(approvalStatus){
  const key=String(approvalStatus||"").trim();
  if(key==="승인" || key==="사용 승인" || key==="단종 승인") return "Applied";
  if(key==="반려" || key==="사용 반려" || key==="단종 반려") return "Rejected";
  if(key==="처리 취소") return "Reverted";
  return "Pending";
}

function sharePointFields(row){
  return row?.fields || row || {};
}
function sharePointPlainText(value){
  if(value==null) return "";
  const raw=String(value);
  if(!raw.includes("<")) return raw;
  if(typeof document!=="undefined"){
    const box=document.createElement("div");
    box.innerHTML=raw;
    return box.textContent || box.innerText || "";
  }
  return raw
    .replace(/<br\s*\/?>/gi,"\n")
    .replace(/<\/p>/gi,"\n")
    .replace(/<[^>]+>/g,"")
    .replace(/&quot;/g,'"')
    .replace(/&#39;/g,"'")
    .replace(/&amp;/g,"&")
    .replace(/&lt;/g,"<")
    .replace(/&gt;/g,">");
}
function parseSharePointJson(value){
  const raw=sharePointPlainText(value).trim();
  if(!raw) return null;
  try{ return JSON.parse(raw); }catch{ return null; }
}
function sharePointRequestApproval(status){
  const key=String(status||"").trim();
  if(key==="Applied") return "승인";
  if(key==="Rejected") return "반려";
  if(key==="Cancelled" || key==="Reverted") return "처리 취소";
  return "보류";
}
function discontinueRequestLifecycle(status,targetLifecycle="ExistingActive"){
  const key=String(status||"Pending").trim().toLowerCase();
  if(key==="applied") return "Discontinued";
  if(key==="rejected" || key==="cancelled" || key==="reverted") return "Cancelled";
  return targetLifecycle && targetLifecycle!=="Discontinued" && targetLifecycle!=="Cancelled"
    ? targetLifecycle
    : "ExistingActive";
}
function sourceClassificationFromLifecycle(lifecycle){
  if(lifecycle==="ReplacedOld" || lifecycle==="Discontinued") return "변경·대체";
  if(lifecycle==="NewItem" || lifecycle==="NewReplacement") return "신규 등록";
  return "기존 품목";
}

// SharePoint의 XRF_ElementResults(측정 1건 × 6원소)를 현재 JSX의
// measurement.elements 객체로 다시 조립합니다. Raw 값은 문자열 그대로 보존합니다.
function buildSharePointElementsByMeasurement(rows){
  const output={};
  for(const row of rows||[]){
    const f=sharePointFields(row);
    const measurementId=String(f.Title||"").trim();
    const element=String(f.field_1||"").trim();
    if(!measurementId || !element) continue;
    if(!output[measurementId]) output[measurementId]={};
    output[measurementId][element]={
      rawPpm:f.field_2 ?? "",
      rawSigma:f.field_3 ?? "",
      rawJudgement:f.field_4 ?? ""
    };
  }
  return output;
}

function measurementSequenceFromId(id){
  const matched=String(id||"").match(/_(\d{2})$/);
  return matched ? Number(matched[1]) : 1;
}
function buildSharePointMeasurementMap(measurementRows,elementRows){
  const elementsByMeasurement=buildSharePointElementsByMeasurement(elementRows);
  const output={};
  for(const row of measurementRows||[]){
    const f=sharePointFields(row);
    const id=String(f.Title||"").trim();
    if(!id) continue;
    const raw={
      id,
      Measurement_ID:id,
      itemId:String(f.field_1||"").trim(),
      Item_ID:String(f.field_1||"").trim(),
      date:dateOnly(f.Measured_Date),
      Measured_Date:f.Measured_Date||null,
      sequence:measurementSequenceFromId(id),
      role:f.field_3||"",
      Measurement_Role:f.field_3||"",
      attemptNo:Number(f.field_4||1),
      parentMeasurementId:f.field_5||null,
      sourceFileName:f.field_6||"",
      sourceFileId:f.Source_File_ID||"",
      sourceFileUrl:f.Source_File_URL||"",
      sourceSheet:f.field_7||"",
      requesterName:f.Requester_Name||"",
      elements:elementsByMeasurement[id]||{},
      _spItemId:row?.id||null
    };
    output[id]=normalizeMeasurementRecord(raw,id);
  }
  return output;
}

function measurementSortValue(m){
  return `${dateOnly(m?.date)||""}|${String(m?.id||"")}|${String(Number(m?.attemptNo||1)).padStart(3,"0")}`;
}
function decorateRuntimeItemWithMeasurements(baseItem,itemId,measurementMap){
  const measurements=Object.values(measurementMap||{})
    .filter(m=>String(m?.itemId||m?.Item_ID||"").trim()===String(itemId||"").trim())
    .sort((a,b)=>measurementSortValue(a).localeCompare(measurementSortValue(b)));

  const initial=measurements.find(m=>normalizeMeasurementRole(m?.role||m?.measurementRole)==="Initial")
    || measurements[0]
    || null;
  const latest=measurements[measurements.length-1]||null;
  const localMap=Object.fromEntries(measurements.map(m=>[m.id,m]));
  const history=measurements.map(m=>({
    id:m.id,
    measurementId:m.id,
    date:m.date,
    periodNo:normalizeMeasurementRole(m?.role)==="Initial" ? 0 : null,
    measurementRole:normalizeMeasurementRole(m?.role||m?.measurementRole||""),
    attemptNo:Number(m?.attemptNo||1),
    parentMeasurementId:m?.parentMeasurementId||null,
    xrf_worst:measurementXrfResult(m),
    xrfWorst:measurementXrfResult(m),
    xrfResult:measurementXrfResult(m),
    xrfLevel:measurementXrfLevel(m),
    elements:m?.elements||{},
    sourceFileName:m?.sourceFileName||"",
    sourceFileId:m?.sourceFileId||"",
    sourceFileUrl:m?.sourceFileUrl||"",
    sourceSheet:m?.sourceSheet||""
  }));

  const referencePeriods=initial ? [{
    num:0,
    label:"R",
    displayLabel:"R",
    displaySeq:0,
    isReference:true,
    referenceDate:initial.date,
    date:initial.date,
    start:null,
    due:null,
    targetDate:null,
    windowStart:null,
    measured:initial.date,
    measuredDates:[initial.date],
    status:"reference",
    statusLabel:"등록완료",
    daysOver:0,
    measurementId:initial.id,
    measurementIds:[initial.id],
    closeType:"registered",
    finalRisk:null,
    isDisplayTarget:false,
    sourceBasis:"SharePoint Initial XRF 기준 R",
    cycleMonths:null,
    windowDays:null,
    prevDue:null,
    nextDue:null,
    cycleVersion:1,
    isCurrentCycle:true,
    measurementRole:"Initial"
  }] : [];

  return {
    ...baseItem,
    itemId:itemId||baseItem?.itemId||null,
    __measurementMap:localMap,
    firstDate:initial?.date||baseItem?.firstDate||null,
    lastMeasured:latest?.date||null,
    firstMeasurementId:initial?.id||null,
    latestMeasurementId:latest?.id||null,
    measurementCount:measurements.length,
    latest,
    history,
    periods:referencePeriods,
    xrfWorst:latest ? measurementXrfResult(latest) : "—",
    latestXrfResult:latest ? measurementXrfResult(latest) : "—",
    latestXrfLevel:latest ? measurementXrfLevel(latest) : null
  };
}

function buildSharePointMasterItems(itemRows,measurementMap,itemPhotos={}){
  return (itemRows||[]).map(row=>{
    const f=sharePointFields(row);
    const itemId=String(f.Title||"").trim();
    const lifecycle=String(f.field_9||"ExistingActive").trim()||"ExistingActive";
    const crType=String(f.field_8||"").trim();
    const photo=itemPhotos?.[itemId]||null;
    const raw={
      itemId,
      _spItemId:row?.id||null,
      _sourceCode:String(f.field_1||"").trim(),
      code:String(f.field_1||"").trim(),
      name:f.field_2||"",
      nameEn:f.field_3||"",
      type:f.field_4||"",
      dept:f.field_5||"",
      respDept:"PQE",
      materialCategory:f.field_6||"",
      materialState:f.field_7||"",
      crType,
      crLevel:crLevelFromType(crType),
      lifecycle,
      originLifecycle:lifecycle,
      sourceClassification:sourceClassificationFromLifecycle(lifecycle),
      originCategory:sourceClassificationFromLifecycle(lifecycle)==="신규 등록" ? "new"
        : sourceClassificationFromLifecycle(lifecycle)==="변경·대체" ? "changed" : "existing",
      isCurrent:f.field_10!==false,
      registrationPeriodNo:Number(f.field_11)||null,
      actionNote:sharePointPlainText(f.field_13||""),
      replacementOf:f.Replacement_Of||null,
      replacementDate:dateOnly(f.Replacement_Date),
      lifecycleEndDate:dateOnly(f.Lifecycle_End_Date),
      photoFileName:f.Photo_File_Name||photo?.name||"",
      photoFileId:f.Photo_File_ID||photo?.id||"",
      photoFileUrl:(photo?.id||f.Photo_File_ID)
        ? `${XRF_SHAREPOINT_API}?itemPhotoFor=${encodeURIComponent(itemId)}&itemPhotoId=${encodeURIComponent(photo?.id||f.Photo_File_ID)}`
        : ""
    };
    return decorateRuntimeItemWithMeasurements(normalizeItemRecord(raw),itemId,measurementMap);
  });
}

function buildSharePointRequestItems(requestRows,masterItems,measurementMap){
  const masterByItemId=new Map((masterItems||[]).map(item=>[String(item?.itemId||""),item]));
  const masterByCode=new Map((masterItems||[]).map(item=>[String(item?.code||""),item]));
  const rows=[];
  for(const row of requestRows||[]){
    const f=sharePointFields(row);
    const requestId=String(f.Title||"").trim();
    const isMigration=f.Is_Migration===true || /^REQ_MIG_/i.test(requestId);
    if(isMigration) continue;

    const requestType=String(f.field_1||"New").trim();
    const requestTypeKey=requestType.toLowerCase();
    const status=String(f.field_2||"Pending").trim();
    const target=masterByItemId.get(String(f.field_3||"").trim())||null;
    const targetCode=target?.code||"";
    const proposedItemId=String(f.field_4||"").trim();
    const proposedCode=String(f.field_5||"").trim();

    // 신규/대체 요청의 Proposed Item이 이미 XRF_Items에 존재하면 Master 행을 화면 source로 사용합니다.
    // Request 상세정보는 mergeSharePointRequestMetadata()에서 Master에 다시 결합합니다.
    if((requestTypeKey==="new" || requestTypeKey==="replacement")
      && proposedCode
      && masterByCode.has(proposedCode)) continue;

    // 승인된 단종 요청은 별도 가상 품목으로 다시 만들지 않습니다.
    // 원본 Master 품목이 Discontinued 상태와 기존 측정 이력을 함께 표시합니다.
    if(requestTypeKey==="discontinue"
      && target
      && status.toLowerCase()==="applied") continue;

    const runtimeCode=requestTypeKey==="discontinue"
      ? requestId
      : (proposedCode||requestId);
    const lifecycle=requestTypeKey==="discontinue"
      ? discontinueRequestLifecycle(status,target?.lifecycle)
      : requestTypeKey==="replacement" ? "NewReplacement" : "NewItem";
    const crType=String(f.field_13||target?.crType||"비접촉").trim();
    const requestReason=sharePointPlainText(f.Request_Reason ?? (requestTypeKey==="new"?f.field_15:""));
    const replacementReason=sharePointPlainText(
      f.Replacement_Reason
      ?? (requestTypeKey==="replacement" ? f.field_15 : "")
    );
    const replacementReasonOther=sharePointPlainText(f.Replacement_Reason_Other||"");
    const oldItemDisposition=String(
      f.Old_Item_Disposition
      ?? (requestTypeKey==="replacement" ? f.field_16 : "")
      ?? ""
    ).trim();
    const discontinueReason=String(
      f.Discontinue_Reason
      ?? (requestTypeKey==="discontinue" ? f.field_15 : "")
      ?? ""
    ).trim();
    const stockDisposition=String(
      f.Stock_Disposition
      ?? (requestTypeKey==="discontinue" ? f.field_16 : "")
      ?? ""
    ).trim();

    const base=normalizeItemRecord({
      _isRequestRecord:true,
      itemId:proposedItemId||null,
      _spRequestItemId:row?.id||null,
      _sourceCode:runtimeCode,
      code:runtimeCode,
      name:requestTypeKey==="discontinue"
        ? `[단종 요청] ${target?.name||f.field_6||targetCode||"기존 품목"}`
        : (f.field_6||"신규 등록 요청 품목"),
      nameEn:target?.nameEn||"",
      type:f.field_7||target?.type||"Auxiliary Materials",
      dept:f.field_8||target?.dept||"",
      respDept:"PQE",
      materialCategory:f.field_9||target?.materialCategory||"",
      materialState:f.field_10||target?.materialState||"",
      manufacturer:sharePointPlainText(f.field_11||""),
      useDate:dateOnly(f.Use_Date),
      crType,
      crLevel:crLevelFromType(crType),
      approvalStatus:sharePointRequestApproval(status),
      actionNote:sharePointPlainText(f.field_18||f.Note||""),
      sourceClassification:requestTypeKey==="discontinue"?"변경·대체":"신규 등록",
      originCategory:requestTypeKey==="discontinue"?"changed":"new",
      category:requestTypeKey==="discontinue"?"changed":"new",
      categoryLabel:requestTypeKey==="discontinue"?"변경·대체":"신규 등록",
      lifecycle,
      originLifecycle:lifecycle,
      isCurrent:status==="Pending",
      requestId,
      requestType:requestTypeKey,
      processStatus:status,
      replacementTargetCode:requestTypeKey==="replacement"?targetCode:"",
      discontinueTargetCode:requestTypeKey==="discontinue"?targetCode:"",
      requestReason,
      replacementReason,
      replacementReasonOther,
      oldItemDisposition,
      replacementOf:requestTypeKey==="replacement"?targetCode:null,
      replacementDate:dateOnly(f.Replacement_Date||f.Effective_Date),
      discontinueReason,
      discontinueReasonOther:sharePointPlainText(f.Discontinue_Reason_Other||""),
      finalUseDate:dateOnly(f.Final_Use_Date||f.Effective_Date),
      stockDisposition,
      processedAt:f.Processed_At||null,
      processedBy:sharePointPlainText(f.Processed_By||""),
      revertedAt:f.Reverted_At||null,
      revertedBy:sharePointPlainText(f.Reverted_By||""),
      revertReason:sharePointPlainText(f.Revert_Reason||""),
      beforeTargetSnapshot:parseSharePointJson(f.Before_Target_Snapshot_JSON),
      restoredTargetSnapshot:parseSharePointJson(f.Restored_Target_Snapshot_JSON),
      note:sharePointPlainText(f.field_18||""),
      requesterName:sharePointPlainText(f.Requester_Name||"")
    });
    rows.push(decorateRuntimeItemWithMeasurements(base,proposedItemId,measurementMap));
  }
  return rows;
}

function mergeSharePointRequestMetadata(masterItems,requestRows){
  const byCode=new Map((masterItems||[]).map(item=>[String(item?.code||""),item]));
  const byItemId=new Map((masterItems||[]).map(item=>[String(item?.itemId||""),item]));
  const patches=new Map();
  const mergePatch=(item,patch)=>{
    if(!item) return;
    const key=String(item?.itemId||item?.code||"");
    patches.set(key,{...(patches.get(key)||{}),...patch});
  };

  for(const row of requestRows||[]){
    const f=sharePointFields(row);
    const requestId=String(f.Title||"").trim();
    if(!requestId || f.Is_Migration===true || /^REQ_MIG_/i.test(requestId)) continue;
    const requestType=String(f.field_1||"").trim().toLowerCase();
    const status=String(f.field_2||"Pending").trim();
    const requestCommon={
      _spRequestItemId:row?.id||null,
      requestId,
      requestType,
      processStatus:status,
      requestReason:sharePointPlainText(f.Request_Reason||""),
      requesterName:sharePointPlainText(f.Requester_Name||""),
      processedAt:f.Processed_At||null,
      processedBy:sharePointPlainText(f.Processed_By||""),
      revertedAt:f.Reverted_At||null,
      revertedBy:sharePointPlainText(f.Reverted_By||""),
      revertReason:sharePointPlainText(f.Revert_Reason||""),
      beforeTargetSnapshot:parseSharePointJson(f.Before_Target_Snapshot_JSON),
      restoredTargetSnapshot:parseSharePointJson(f.Restored_Target_Snapshot_JSON),
    };

    if(requestType==="new" || requestType==="replacement"){
      const proposedCode=String(f.field_5||"").trim();
      const proposedItemId=String(f.field_4||"").trim();
      const item=byItemId.get(proposedItemId)||byCode.get(proposedCode)||null;
      mergePatch(item,{
        ...requestCommon,
        manufacturer:sharePointPlainText(f.field_11||""),
        useDate:dateOnly(f.Use_Date),
        replacementTargetCode:requestType==="replacement"
          ? (byItemId.get(String(f.field_3||"").trim())?.code||item?.replacementOf||"")
          : "",
        replacementReason:sharePointPlainText(f.Replacement_Reason||""),
        replacementReasonOther:sharePointPlainText(f.Replacement_Reason_Other||""),
        oldItemDisposition:String(f.Old_Item_Disposition||"").trim(),
        replacementDate:dateOnly(f.Replacement_Date||f.Effective_Date)||item?.replacementDate||null,
        note:sharePointPlainText(f.field_18||""),
      });
    }

    if(requestType==="discontinue"){
      const target=byItemId.get(String(f.field_3||"").trim())||null;
      const statusKey=status.toLowerCase();
      const applied=statusKey==="applied";
      const pending=statusKey==="pending";
      const reverted=statusKey==="reverted" || statusKey==="cancelled";
      // Pending/Rejected 요청의 상태를 Master 품목에 덮어쓰지 않습니다.
      // 대상 품목은 승인된 경우에만 단종되고, 반려되면 기존 Lifecycle과 XRF 이력을 그대로 유지합니다.
      if(applied) mergePatch(target,{
          ...requestCommon,
          discontinueTargetCode:target?.code||"",
          pendingDiscontinueRequestCode:null,
          discontinuedByRequestCode:requestId,
          discontinueReason:String(f.Discontinue_Reason||f.field_15||"").trim(),
          discontinueReasonOther:sharePointPlainText(f.Discontinue_Reason_Other||""),
          finalUseDate:dateOnly(f.Final_Use_Date||f.Effective_Date),
          stockDisposition:String(f.Stock_Disposition||f.field_16||"").trim(),
          category:"changed",
          categoryLabel:"변경·대체",
          lifecycle:"Discontinued",
          isCurrent:false,
          nextDue:null,
          dDay:null,
          complianceStatus:"notAvailable",
      });
      else if(reverted) mergePatch(target,{
        pendingDiscontinueRequestCode:null,
        restoredFromDiscontinueRequestCode:requestId,
        restoredAt:dateOnly(f.Reverted_At),
        restoredBy:sharePointPlainText(f.Reverted_By||""),
        restoreReason:sharePointPlainText(f.Revert_Reason||""),
      });
      else if(pending) mergePatch(target,{pendingDiscontinueRequestCode:requestId});
      else mergePatch(target,{pendingDiscontinueRequestCode:null});
    }
  }

  return (masterItems||[]).map(item=>{
    const key=String(item?.itemId||item?.code||"");
    return patches.has(key)?{...item,...patches.get(key)}:item;
  });
}

function buildSharePointPrecisionOverrides(precisionRows,precisionElementRows){
  const elementsByPrecision={};
  for(const row of precisionElementRows||[]){
    const f=sharePointFields(row);
    const precisionId=String(f.Title||"").trim();
    const element=String(f.field_1||"").trim();
    if(!precisionId || !element) continue;
    if(!elementsByPrecision[precisionId]) elementsByPrecision[precisionId]={};
    elementsByPrecision[precisionId][element]={
      presence:f.field_2||"",
      ppm:f.field_3 ?? ""
    };
  }
  const output={};
  for(const row of precisionRows||[]){
    const f=sharePointFields(row);
    const precisionId=String(f.Title||"").trim();
    const triggerMeasurementId=String(f.field_1||"").trim();
    if(!precisionId || !triggerMeasurementId) continue;
    const resultFile=String(f.field_5||"").trim();
    output[triggerMeasurementId]={
      precisionId,
      _spItemId:row?.id||null,
      requestStatus:"REQUESTED",
      requestedAt:dateOnly(f.Requested_At||row?.createdDateTime),
      requestedBy:sharePointPlainText(f.Requested_By||row?.createdBy?.user?.displayName||""),
      resultFileId:resultFile,
      reportFileId:f.Report_File_ID||"",
      reportFileUrl:f.Report_File_URL||"",
      uploadStatus:resultFile?"UPLOADED":"WAITING",
      finalConfirm:f.field_6===true?"확인":"",
      elementResults:elementsByPrecision[precisionId]||{}
    };
  }
  return output;
}

function buildRuntimeFromSharePoint(payload){
  const lists=payload?.lists||{};
  const itemRows=lists.XRF_Items||lists.items||[];
  const requestRows=lists.XRF_Requests||lists.requests||[];
  const measurementRows=lists.XRF_Measurements||lists.measurements||[];
  const elementRows=lists.XRF_ElementResults||lists.elementResults||[];
  const precisionRows=lists.XRF_PrecisionAnalyses||lists.precisionAnalyses||[];
  const precisionElementRows=lists.XRF_PrecisionElementResults||lists.precisionElementResults||[];

  const measurementMap=buildSharePointMeasurementMap(measurementRows,elementRows);
  const baseMasterItems=deriveReplacementRelations(buildSharePointMasterItems(itemRows,measurementMap,payload?.itemPhotos||{}));
  const masterItems=mergeSharePointRequestMetadata(baseMasterItems,requestRows);
  const requestItems=buildSharePointRequestItems(requestRows,masterItems,measurementMap);
  const precisionOverrides=buildSharePointPrecisionOverrides(precisionRows,precisionElementRows);

  return {
    items:masterItems,
    requestItems,
    precisionOverrides,
    measurementMap,
    counts:{
      items:itemRows.length,
      requests:requestRows.length,
      measurements:measurementRows.length,
      elementResults:elementRows.length,
      precisionAnalyses:precisionRows.length,
      precisionElementResults:precisionElementRows.length
    }
  };
}

// SharePoint 리스트와 현재 분할 DB를 동일한 방식으로 읽기 위한 어댑터입니다.
// 운영 시 XRF_Measurements는 마스터 DB의 열 이름을 그대로 사용하고,
// XRF_Items는 Lifecycle/Replacement_Of/Replacement_Date를 명시적으로 저장합니다.
function pickField(row, ...names){
  for(const name of names){
    if(row && Object.prototype.hasOwnProperty.call(row,name) && row[name] !== undefined) return row[name];
  }
  return undefined;
}
function dateOnly(value){
  if(!value) return null;
  const text=String(value);
  return text.length>=10 ? text.slice(0,10) : text;
}
function normalizedClassification(value){
  const key=String(value||"").trim().toLowerCase().replace(/[\s·_-]/g,"");
  if(["변경대체","changedreplacement","replacement","changed"].includes(key)) return "changed";
  if(["신규등록","new","newitem"].includes(key)) return "new";
  return "existing";
}
function classificationMeta(value){
  const category=normalizedClassification(value);
  if(category==="changed") return {category:"changed",categoryLabel:"변경·대체",lifecycle:"ReplacedOld",isCurrent:false};
  if(category==="new") return {category:"new",categoryLabel:"신규 등록",lifecycle:"NewItem",isCurrent:true};
  return {category:"existing",categoryLabel:"기존 품목",lifecycle:"ExistingActive",isCurrent:true};
}
function normalizeElementRecord(record, element){
  const nested=record?.elements?.[element] || record?.[element] || {};
  const sourcePpm=pickField(nested,"rawPpm","sourceValue") ?? pickField(record,`${element}_ppm`,`${element}_PPM`) ?? nested.ppm;
  const rawSigma=pickField(nested,"rawSigma","sourceSigma") ?? pickField(record,`${element}_3sigma`,`${element}_3Sigma`) ?? nested.sigma;
  const rawJudgement=pickField(nested,"rawJudgement","sourceJudgement")
    ?? pickField(record,`${element}_Judgement`,`${element}_Judgment`)
    ?? nested.judge ?? "";
  // 내부 OK/NG는 실제 Content를 사용해야 하므로 Judgment 값으로 Content를 덮어쓰지 않습니다.
  const rawPpm=sourcePpm;
  const legacyLimit=pickField(nested,"limit","Limit")
    ?? pickField(record,`${element}_limit`,`${element}_Limit`,`${element}_LIMIT`);
  const legalLimit=pickField(nested,"legalLimit","Legal_Limit","LegalLimit")
    ?? pickField(record,`${element}_legal_limit`,`${element}_Legal_Limit`)
    ?? XRF_ELEMENT_LIMITS[element]
    ?? legacyLimit;
  const internalLimit=pickField(nested,"internalLimit","Internal_Limit","InternalLimit")
    ?? pickField(record,`${element}_internal_limit`,`${element}_Internal_Limit`)
    ?? (xrfNumber(legalLimit)==null ? null : Math.round(xrfNumber(legalLimit)*XRF_INTERNAL_LIMIT_RATIO*10)/10);
  const ppm=xrfNumber(rawPpm);
  const sigma=xrfNumber(rawSigma);
  const dataMissing=!!nested.dataMissing
    || (isXrfMissing(rawPpm) && !isXrfContentND(rawPpm))
    || (element!=="Cl+Br" && isXrfMissing(rawSigma) && !isXrfContentND(rawSigma));
  const normalized={
    ...nested,
    rawPpm,
    rawSigma,
    rawJudgement,
    // limit는 기존 데이터 호환을 위해 Legal Limit을 유지하고, 실제 XRF 판정은 internalLimit을 사용합니다.
    limit:legalLimit,
    legalLimit,
    internalLimit,
    ppm,
    sigma,
    isND:isXrfContentND(rawPpm),
    dataMissing,
  };
  // 저장된 과거 level 값보다 현재 Internal Limit 70% 재판정값을 항상 우선합니다.
  normalized.xrfLevel=xrfLevelFromJudgement(normalized);
  return normalized;
}
function normalizeMeasurementRole(value){
  const role=String(value??"").trim();
  return role==="Cycle_Reset" ? "Periodic" : role;
}
function normalizeMeasurementRecord(raw, fallbackId=""){
  const id=String(pickField(raw,"id","Measurement_ID","measurementId") || fallbackId);
  const sourceRole=pickField(raw,"role","Measurement_Role","measurementRole");
  const elements={};
  ELEMENTS.filter(element=>element!=="Cl+Br").forEach(element=>{ elements[element]=normalizeElementRecord(raw,element); });
  const clBr=buildClBrResult(sourcePpmValue(elements.Cl),sourcePpmValue(elements.Br));
  elements["Cl+Br"]={
    ...normalizeElementRecord(raw,"Cl+Br"),
    rawPpm:clBr.ppm,
    rawSigma:null,
    rawJudgement:clBr.judgement,
    ppm:xrfNumber(clBr.ppm),
    sigma:null,
    judge:clBr.judgement,
    level:clBr.level,
    xrfLevel:clBr.level,
    limit:clBr.legalLimit,
    legalLimit:clBr.legalLimit,
    internalLimit:clBr.internalLimit,
    primaryClass:clBr.level==="H" ? "NG" : clBr.level ? "OK" : null,
    calculated:true,
    decisionSource:"cl_br_content_sum_internal_limit_70pct",
    dataMissing:clBr.ppm==null,
    isND:!!clBr.isND,
  };
  return {
    ...raw,
    id,
    date:dateOnly(pickField(raw,"date","Measured_Date","measuredDate")),
    sequence:pickField(raw,"sequence","Measurement_Sequence","measurementSequence"),
    role:normalizeMeasurementRole(sourceRole),
    sourceRole:sourceRole||null,
    xrfWorst:pickField(raw,"xrfWorst","XRF_Worst") || raw?.xrf_worst,
    approvalStatus:pickField(raw,"approvalStatus","Approval_Status"),
    retestRequiredElements:pickField(raw,"retestRequiredElements","Retest_Elements") || "",
    elements,
  };
}
function normalizeMeasurementCollection(source){
  const output={};
  if(Array.isArray(source)){
    source.forEach(row=>{ const m=normalizeMeasurementRecord(row); if(m.id) output[m.id]=m; });
  }else{
    Object.entries(source||{}).forEach(([id,row])=>{ const m=normalizeMeasurementRecord(row,id); if(m.id) output[m.id]=m; });
  }
  return output;
}
function normalizeItemRecord(raw){
  const code=String(pickField(raw,"code","Part_Number","Part Number","PartNumber") || "").trim();
  const sourceClassification=pickField(raw,"sourceClassification","분류","Classification","classification")
    ?? pickField(raw,"categoryLabel") ?? "기존 품목";
  const sourceMeta=classificationMeta(sourceClassification);
  const explicitCategory=pickField(raw,"category","Category");
  const explicitLifecycle=pickField(raw,"lifecycle","Lifecycle","Lifecycle_Status");
  const originCategory=normalizedClassification(
    pickField(raw,"originCategory","Origin_Category","Original_Category")
      ?? explicitCategory
      ?? sourceClassification
  );
  return {
    ...raw,
    _sourceCode:pickField(raw,"_sourceCode","Source_Code","sourceCode") || code,
    code,
    name:pickField(raw,"name","Item_Name","Item Name") || "",
    nameEn:pickField(raw,"nameEn","Item_Name_EN","Item Name EN","English_Name","English Name") || raw?.nameEn || ITEM_NAME_EN_FALLBACKS[code] || "",
    type:pickField(raw,"type","Type","Item_Type") || "",
    dept:pickField(raw,"dept","Handling_Dept") || "",
    materialCategory:pickField(raw,"materialCategory","Material_Category") || "",
    materialState:pickField(raw,"materialState","Material_State") || "",
    requesterName:pickField(raw,"requesterName","Requester_Name") || "",
    photoFileName:pickField(raw,"photoFileName","Photo_File_Name") || "",
    photoFileId:pickField(raw,"photoFileId","Photo_File_ID") || "",
    photoFileUrl:pickField(raw,"photoFileUrl","Photo_File_URL") || "",
    crType:pickField(raw,"crType","CR_Type") || "",
    crLevel:pickField(raw,"crLevel","CR_Type_Level") || crLevelFromType(pickField(raw,"crType","CR_Type") || ""),
    approvalStatus:pickField(raw,"approvalStatus","Approval_Status") || "측정 진행중",
    approvalDetail:pickField(raw,"approvalDetail","Approval_Detail","Retest_Elements") || "",
    sourceClassification,
    originCategory,
    originLifecycle:pickField(raw,"originLifecycle","Origin_Lifecycle")
      || explicitLifecycle
      || sourceMeta.lifecycle,
    // 신규/기존 판정은 더 이상 최초 2025년 조사 리스트나 고정 날짜 만료일을 기준으로 하지 않습니다.
    // 등록 당시의 관리 주기(Pn)를 저장하고, 현재 관리 주기와 비교하여 신규 여부를 판정합니다.
    registrationPeriodNo:Number(pickField(raw,"registrationPeriodNo","Registration_Period_No","Registered_Period_No")) || null,
    newCategoryUntil:dateOnly(pickField(
      raw,
      "newCategoryUntil","New_Category_Until","New_Until","Registration_Period_End"
    )),
    category:explicitCategory || sourceMeta.category,
    categoryLabel:pickField(raw,"categoryLabel","Category_Label") || sourceMeta.categoryLabel,
    lifecycle:explicitLifecycle || sourceMeta.lifecycle,
    isCurrent:pickField(raw,"isCurrent","Is_Current") ?? sourceMeta.isCurrent,
    replacementOf:pickField(raw,"replacementOf","Replacement_Of") || null,
    replacedBy:pickField(raw,"replacedBy","Replaced_By") || null,
    replacementDate:dateOnly(pickField(raw,"replacementDate","Replacement_Date")),
    firstDate:dateOnly(pickField(raw,"firstDate","First_Measured_Date","Registered_Date")),
  };
}
function deriveReplacementRelations(sourceItems){
  const items=(sourceItems||[]).map(item=>({...normalizeItemRecord(item)}));
  const byCode=new Map(items.map(item=>[item.code,item]));
  const relations=new Map();
  const addRelation=(parentCode,childCode,source)=>{
    if(!parentCode || !childCode || parentCode===childCode || !byCode.has(parentCode) || !byCode.has(childCode)) return;
    if(!relations.has(parentCode)) relations.set(parentCode,new Map());
    const current=relations.get(parentCode).get(childCode);
    if(!current || source==="explicit") relations.get(parentCode).set(childCode,source);
  };

  // 1순위: SharePoint XRF_Items / XRF_Requests에 저장된 명시적 관계
  items.forEach(item=>{
    if(item.replacementOf) addRelation(item.replacementOf,item.code,"explicit");
    String(item.replacedBy||"").split(",").map(v=>v.trim()).filter(Boolean)
      .forEach(childCode=>addRelation(item.code,childCode,"explicit"));
  });

  // 2순위: 관계 열이 없던 기존 마스터를 최초 이관할 때만 사용하는 규칙 기반 보완
  // 자식 품번이 부모-NN 형식이고 실제 부모 품번이 존재할 때,
  // ① 부모가 변경대체이거나 ② 자식이 신규등록이면 대체 관계로 복원합니다.
  // 부서가 서로 다르면 잘못 연결하지 않습니다. 운영 SharePoint에서는 Replacement_Of를 우선 사용합니다.
  items.forEach(child=>{
    if(child.replacementOf) return;
    const matched=child.code.match(/^(.*)-(\d{2})$/);
    if(!matched) return;
    const parent=byCode.get(matched[1]);
    if(!parent) return;
    const parentIsReplacement=normalizedClassification(parent.sourceClassification)==="changed"
      || parent.lifecycle==="ReplacedOld" || parent.category==="changed";
    const childIsNew=normalizedClassification(child.sourceClassification)==="new"
      || child.lifecycle==="NewItem" || child.lifecycle==="NewReplacement" || child.category==="new";
    const sameDepartment=!parent.dept || !child.dept || parent.dept===child.dept;
    if(sameDepartment && (parentIsReplacement || childIsNew)){
      addRelation(parent.code,child.code,"inferred");
    }
  });

  relations.forEach((children,parentCode)=>{
    const parent=byCode.get(parentCode);
    const childItems=Array.from(children.keys()).map(code=>byCode.get(code)).filter(Boolean);
    if(!parent || !childItems.length) return;
    childItems.forEach(child=>{
      const relationSource=children.get(child.code);
      child.replacementOf=parentCode;
      child.relationshipSource=relationSource;
      child.replacementDate=child.replacementDate || child.firstDate || null;
      child.originCategory=child.originCategory || "new";
      child.originLifecycle=child.originLifecycle || "NewReplacement";
      // 대체 관계는 유지하되 이미 단종/취소된 자식 품목의 현재 Lifecycle을 다시 NewReplacement로 되돌리지 않습니다.
      const terminalChild=["Discontinued","Cancelled"].includes(child.lifecycle);
      if(!terminalChild){
        child.category="new";
        child.categoryLabel="신규 등록";
        child.lifecycle="NewReplacement";
        child.isCurrent=true;
      }else{
        child.isCurrent=false;
      }
    });
    const dates=childItems.map(child=>child.replacementDate||child.firstDate).filter(Boolean).sort();
    parent.replacedBy=childItems.map(child=>child.code).sort().join(",");
    parent.replacementDate=parent.replacementDate || dates[0] || null;
    parent.relationshipSource=Array.from(children.values()).includes("explicit") ? "explicit" : "inferred";
    parent.category="changed";
    parent.categoryLabel="변경·대체";
    // 단종이 명시된 품목은 terminal 상태를 유지하고, 대체 관계는 replacedBy로 별도 보존합니다.
    if(parent.lifecycle!=="Discontinued" && parent.lifecycle!=="Cancelled") parent.lifecycle="ReplacedOld";
    parent.isCurrent=false;
  });
  return items;
}

function replacementCodeList(value){
  return String(value||"").split(",").map(v=>v.trim()).filter(Boolean);
}
function buildReplacementChains(items){
  const rows=(items||[]).filter(i=>i?.code && !isDiscontinueRequestItem(i));
  const byCode=new Map(rows.map(i=>[i.code,i]));
  const children=new Map();
  const parents=new Map();
  const addEdge=(from,to)=>{
    if(!from || !to || from===to || !byCode.has(from) || !byCode.has(to)) return;
    if(!children.has(from)) children.set(from,new Set());
    children.get(from).add(to);
    if(!parents.has(to)) parents.set(to,new Set());
    parents.get(to).add(from);
  };
  rows.forEach(item=>{
    if(item.replacementOf) addEdge(item.replacementOf,item.code);
    replacementCodeList(item.replacedBy).forEach(code=>addEdge(item.code,code));
  });

  const involved=new Set();
  children.forEach((set,from)=>{ involved.add(from); set.forEach(to=>involved.add(to)); });
  rows.filter(i=>i.lifecycle==="ReplacedOld" || i.lifecycle==="Discontinued").forEach(i=>involved.add(i.code));

  const paths=[];
  const signature=new Set();
  const pushPath=(codes)=>{
    if(!codes.length) return;
    const sig=codes.join(">");
    if(signature.has(sig)) return;
    signature.add(sig);
    paths.push(codes);
  };
  const walk=(code,path,seen)=>{
    if(seen.has(code)){ pushPath([...path,code]); return; }
    const nextSeen=new Set(seen); nextSeen.add(code);
    const nextPath=[...path,code];
    const next=[...(children.get(code)||[])].filter(c=>!nextSeen.has(c));
    if(!next.length){ pushPath(nextPath); return; }
    next.forEach(child=>walk(child,nextPath,nextSeen));
  };
  const roots=[...involved].filter(code=>!(parents.get(code)?.size));
  roots.forEach(code=>walk(code,[],new Set()));
  // 잘못된 순환 데이터나 부모가 누락된 관계도 위험도 화면에서 사라지지 않게 보완합니다.
  [...involved].forEach(code=>{
    if(!paths.some(path=>path.includes(code))) walk(code,[],new Set());
  });

  return paths.map((codes,index)=>{
    const chainItems=codes.map(code=>byCode.get(code)).filter(Boolean);
    const last=chainItems[chainItems.length-1]||null;
    const eventDates=chainItems.map(i=>i.replacementDate||i.finalUseDate).filter(Boolean).sort();
    return {
      id:`replacement-chain-${index}-${codes.join("-")}`,
      codes,
      items:chainItems,
      last,
      eventDate:eventDates[eventDates.length-1]||null,
      hasDiscontinued:chainItems.some(i=>i.lifecycle==="Discontinued"),
      hasReplaced:chainItems.some(i=>i.lifecycle==="ReplacedOld"),
    };
  });
}

function buildReplacementRelationGroups(items){
  // 첨부 코드의 "기존 품목 → 대체품" 표현을 그대로 사용할 수 있도록
  // Replacement_Of와 Replaced_By 양쪽 데이터를 한 관계 그룹으로 묶습니다.
  const rows=(items||[]).filter(item=>item?.code && !isDiscontinueRequestItem(item));
  const byCode=new Map(rows.map(item=>[item.code,item]));
  const replacementMap=new Map();
  const add=(oldCode,newCode)=>{
    if(!oldCode || !newCode || oldCode===newCode) return;
    if(!byCode.has(oldCode) || !byCode.has(newCode)) return;
    if(!replacementMap.has(oldCode)) replacementMap.set(oldCode,new Set());
    replacementMap.get(oldCode).add(newCode);
  };
  rows.forEach(item=>{
    if(item.replacementOf) add(item.replacementOf,item.code);
  });
  rows.forEach(old=>{
    replacementCodeList(old.replacedBy).forEach(code=>add(old.code,code));
  });
  return Array.from(replacementMap.entries()).map(([oldCode,codes])=>{
    const old=byCode.get(oldCode);
    const successors=Array.from(codes)
      .map(code=>byCode.get(code))
      .filter(Boolean)
      .sort((a,b)=>String(a.replacementDate||a.firstDate||"").localeCompare(String(b.replacementDate||b.firstDate||"")));
    return old&&successors.length ? {old,successors} : null;
  }).filter(Boolean).sort((a,b)=>String(a.old.replacementDate||a.old.finalUseDate||"").localeCompare(String(b.old.replacementDate||b.old.finalUseDate||"")));
}

// 부자재 리스트의 신규/기존 분류는 "최초 2025년 조사 리스트"와의 차이가 아니라
// 현재 관리 주기(Pn)를 기준으로 다시 산정합니다.
//
// 판정 순서
// 1) 대체됨/단종/취소 품목은 Lifecycle이 최우선이며 신규/기존으로 되돌리지 않습니다.
// 2) 아직 XRF 측정값이 없는 의뢰 품목은 항상 신규 등록입니다.
// 3) XRF가 있더라도 현재 Pn에서 R로 처음 등록된 품목은 그 Pn 동안 신규 등록입니다.
// 4) 다음 Pn으로 넘어가면 해당 품목은 기존 품목으로 자동 전환됩니다.
// 5) 대체품은 위 신규 규칙을 따르고, 대체된 기존 품목은 ReplacedOld로 주기 전개를 종료합니다.
function originalCategoryOf(item){
  const explicit=pickField(item,"originCategory","Origin_Category","Original_Category");
  if(explicit) return normalizedClassification(explicit);
  if(item?.lifecycle==="NewItem" || item?.lifecycle==="NewReplacement") return "new";
  return normalizedClassification(
    pickField(item,"sourceClassification","분류","Classification")
      ?? item?.categoryLabel
      ?? item?.category
  );
}
function itemStateKey(item){
  return String(item?._sourceCode || item?.sourceCode || item?.code || "").trim();
}
function hasAnyXrfMeasurement(item){
  if(!item) return false;
  if(Number(item.measurementCount||0)>0) return true;
  if(item.latestMeasurementId || item.firstMeasurementId || item.lastMeasured || item.latest?.id) return true;
  return (item.history||[]).some(h=>!!(h?.id || h?.measurementId || h?.date));
}
function portfolioPeriodCandidate(item){
  if(!item || isDiscontinueRequestItem(item) || isHistoryRecordItem(item) || isEquipmentItem(item)) return null;
  // 새 의뢰/신규 등록품이 스스로 전체 기준 Pn을 바꾸지 않도록 최초 기준이 기존 운영품인 항목만 사용합니다.
  if(originalCategoryOf(item)==="new") return null;

  // "현재 Pn"은 다음 미래 예정 주기가 아니라 실제로 도래한 가장 최근 주기입니다.
  // 측정 완료된 Pn, 도래일이 지난 Pn, 또는 측정권장 window가 시작된 Pn 중 가장 큰 번호를 사용합니다.
  // 예: P1 측정 완료 후 P2가 아직 4개월 뒤라면 여전히 P1이고, P2 권장 window가 시작되면 P2로 전환됩니다.
  const arrived=(itemAllPeriods(item)||[])
    .filter(p=>!(p?.isReference || Number(p?.num)===0))
    .filter(p=>{
      if(p?.measured) return true;
      const due=dateOnly(p?.due || p?.targetDate);
      const windowStart=dateOnly(p?.windowStart);
      return (!!windowStart && dateGte(TODAY_STR,windowStart)) || (!!due && dateGte(TODAY_STR,due));
    })
    .sort((a,b)=>Number(a?.num||0)-Number(b?.num||0));
  const arrivedNo=Number(arrived[arrived.length-1]?.num);
  if(Number.isFinite(arrivedNo) && arrivedNo>0) return arrivedNo;

  const stored=Number(item?.currentPeriodNo);
  return Number.isFinite(stored) && stored>0 ? stored : 1;
}
function derivePortfolioCurrentPeriodNo(items){
  const counts=new Map();
  (items||[]).forEach(item=>{
    const n=portfolioPeriodCandidate(item);
    if(!n) return;
    counts.set(n,(counts.get(n)||0)+1);
  });
  if(!counts.size) return 1;
  // 서로 다른 Risk 주기가 섞여 있으므로 단순 max가 아니라 가장 많은 운영품이 위치한 Pn을 사용합니다.
  // 동률이면 더 최근 단계인 큰 Pn을 선택합니다.
  return Array.from(counts.entries())
    .sort((a,b)=>b[1]-a[1] || b[0]-a[0])[0][0];
}
function applyPeriodBasedCategories(items){
  const currentPeriodNo=derivePortfolioCurrentPeriodNo(items);
  return (items||[]).map(item=>{
    const next={...item};
    next.originCategory=originalCategoryOf(next);
    next.classificationPeriodNo=currentPeriodNo;

    // 대체된 기존품/단종/취소는 관계 상태가 신규·기존 분류보다 우선합니다.
    if(
      next.lifecycle==="ReplacedOld"
      || next.lifecycle==="Discontinued"
      || next.lifecycle==="Cancelled"
      || normalizedClassification(next.category)==="changed"
    ) return next;

    const noXrf=!hasAnyXrfMeasurement(next);
    // 현재 정적 마스터의 기존 신규품(S-01-01, S-01-02, S-11 등)은
    // sourceClassification/lifecycle에는 신규 정보가 있지만 registrationPeriodNo가 비어 있습니다.
    // 이 경우 마스터에 이미 저장된 currentPeriodNo를 최초 등록 관리주기의 legacy fallback으로 사용합니다.
    // 이후 전체 관리주기가 다음 Pn으로 넘어가면 정상적으로 기존 품목으로 전환됩니다.
    const legacyRegistrationPeriod=(next.originCategory==="new" && !Number(next.registrationPeriodNo||0))
      ? Number(next.currentPeriodNo||0)
      : 0;
    const registeredPeriod=Number(next.registrationPeriodNo||legacyRegistrationPeriod||0);
    if(!next.registrationPeriodNo && registeredPeriod>0 && next.originCategory==="new") next.registrationPeriodNo=registeredPeriod;
    const registeredInCurrentPeriod=registeredPeriod>0 && registeredPeriod===Number(currentPeriodNo);
    const isNewNow=noXrf || registeredInCurrentPeriod;

    next.newCategoryActive=isNewNow;
    next.category=isNewNow ? "new" : "existing";
    next.categoryLabel=isNewNow ? "신규 등록" : "기존 품목";

    // 대체 관계 자체는 Replacement_Of / Replaced_By에 남기고,
    // 다음 Pn부터는 현재 운영품으로 분류합니다.
    next.originLifecycle=next.originLifecycle || next.lifecycle;
    if(!isNewNow && ["NewItem","NewReplacement"].includes(next.lifecycle)){
      next.lifecycle="ExistingActive";
      next.isCurrent=true;
    }
    return next;
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// UI language layer (Korean / English)
// 업무 데이터의 key/value와 판정 로직은 건드리지 않고, 실제 화면에 렌더링되는
// 고정 UI 문구·상태 문구·placeholder/title만 표시 단계에서 번역합니다.
// React의 비교/필터/Workflow는 기존 한국어 원본값을 계속 사용하므로 언어 전환이
// 데이터 로직에 영향을 주지 않습니다.
// ─────────────────────────────────────────────────────────────────────────────
const UI_EN_MAP = {
  // ── 탭 / 전역 ──
  "부자재 리스트":"Auxiliary Material List","XRF 분석":"XRF Analysis","정밀분석":"Precision Analysis",
  "위험도 현황":"Risk Overview","품목 / 변경 관리":"Item / Change Management",
  "부자재 XRF 관리 시스템":"Auxiliary Material XRF Management System",
  "기준일":"As of","한국어":"Korean","영어":"English",

  // ── 공통 라벨 ──
  "분류":"Category","품번":"Part No.","품목명":"Item Name","부서":"Department","주기":"Cycle",
  "최초 등록":"First Registered","이행 현황":"Compliance","현재 상태":"Current Status",
  "승인 상태":"Approval Status","다음 마감":"Next Due","상세 보기":"View Details",
  "검사주기":"Inspection Cycle","공정명":"Process","담당자":"Owner",
  "요청일":"Request Date","측정일":"Measurement Date","최종 사용일":"Final Use Date",
  "최종 사용 예정일":"Planned Final Use Date","사용 예정 일자":"Planned Use Date","사용 일자":"Use Date",
  "전환 예정일":"Planned Transition Date","비고":"Note","사유":"Reason","상태":"Status",
  "결과":"Results","건":" items","총":"Total","이력":"History","운영":"Active",
  "취소":"Cancel","선택":"Select","적용":"Apply","초기화":"Reset","전체 초기화":"Reset All",
  "전체 선택":"Select All","필터":"Filter","필터 초기화":"Reset Filters","필터:":"Filters:",
  "이전":"Previous","다음":"Next","목록으로":"Back to List","저장":"Save","확인":"Confirm",
  "읽는 중":"Reading","대기":"Waiting","없음":"None","해당 없음":"N/A","기타":"Other",
  "요청 번호":"Request No.","원복일":"Restored On","원복 사유":"Restore Reason",
  "설비":"Facility","부자재":"Auxiliary Material","측정":"Measurement","컨펌":"Confirm",
  "인계":"Transfer","성적서":"Report","파일 선택":"Choose File","PDF 선택":"Choose PDF",
  "제조사":"Manufacturer","제품 적재용 릴":"Product loading reel","원자재 이송용 레일":"Material transfer rail",

  // ── 상태 값 ──
  "기준":"Cycle Reference","이행":"Compliant","예정":"Upcoming","지연":"Late","지연측정":"Late Measurement",
  "조기측정":"Early Measurement","미이행":"Overdue","측정권장":"Measurement Recommended",
  "측정 완료":"Measured","측정완료":"Measured","미측정":"Not Measured","대상 아님":"Not Applicable",
  "등록완료":"Registered","처리 완료":"Complete","처리 취소":"Cancelled","진행 중":"In Progress",
  "진행중":"In Progress","완료":"Complete","승인":"Approved","반려":"Rejected","보류":"On Hold",
  "측정 진행중":"Measuring","측정중":"Measuring","의뢰 필요":"Request Required",
  "확인 필요":"Review Required","재측정 필요":"Remeasurement Required","재측정":"Remeasure",
  "재측정 불필요":"No Remeasurement","필요":"Required","불필요":"Not Required","필요 없음":"Not Required",
  "사용 가능":"Allowed","사용불허":"Not Allowed","사용 불허":"Not Allowed","사용불가":"Not Allowed",
  "사용 승인":"Use Approved","사용 반려":"Use Rejected","단종 승인":"Discontinuation Approved",
  "단종 반려":"Discontinuation Rejected","단종 검토중":"Discontinuation Under Review",
  "단종":"Discontinued","단종 처리":"Discontinuation","대체됨":"Replaced","대체됨(이력만)":"Replaced (History Only)",
  "기존 운영":"Existing Active","기존 품목":"Existing Item","신규 등록":"New Registration",
  "신규 대체":"New Replacement","신규 도입":"New Item","변경·대체":"Change / Replacement",
  "변경대체":"Change / Replacement","신규등록":"New Registration","불검출":"N.D.","함유":"Detected",
  "미함유":"Not Detected","적합":"Pass","부적합":"Fail","미등록":"Not Registered","미선택":"Not Selected",
  "PDF 등록":"PDF Uploaded","저장 완료":"Saved","오류 확인":"Check Error","파싱 완료":"Parsing Complete",
  "파싱 대기":"Awaiting Parsing","시트":"Sheet","데이터 부족":"Insufficient Data",
  "결과 없음":"No Result","결과 대기":"Awaiting Result","측정 없음":"No Measurement",
  "측정 대기":"Awaiting Measurement","결과 확인":"Review Result","오류":"Error",

  // ── 주기 / C&R ──
  "월 1회":"Monthly","반기 1회":"Semiannual","연 1회":"Annual","반기":"Semiannual",
  "직접접촉+잔류":"Direct Contact + Residual","직접접촉·잔류":"Direct Contact + Residual",
  "직접접촉+비잔류":"Direct Contact + Non-residual","직접접촉·비잔류":"Direct Contact + Non-residual",
  "비접촉":"Non-contact","직접+잔류":"Direct + Residual","직접+비잔류":"Direct + Non-residual",
  "염소+브롬":"Cl+Br","납":"Lead","수은":"Mercury","크롬":"Chromium","카드뮴":"Cadmium",
  "염소":"Chlorine","브롬":"Bromine",

  // ── 후속조치 / Workflow ──
  "후속조치":"Follow-up","후속조치 없음":"No Follow-up","후속조치 필요":"Follow-up Required",
  "추가 조치 없음":"No Additional Action","추가 후속조치 없음":"No Additional Follow-up",
  "XRF 후속조치":"XRF Follow-up","XRF 재측정":"XRF Remeasurement","XRF 결과":"XRF Result",
  "XRF 방식":"XRF Method","XRF 방식 선택":"Select XRF Method","XRF 측정 의뢰":"XRF Measurement Request",
  "XRF 측정 대기":"Awaiting XRF Measurement","XRF 결과 업로드":"Upload XRF Result",
  "XRF 원본 업로드":"Upload Original XRF","XRF 원본 파일":"Original XRF File",
  "XRF 측정 ID":"XRF Measurement ID","XRF 원본 보고서":"Original XRF Report",
  "XRF NG 판정":"XRF NG","재측정 후 ?? 판정":"?? After Remeasurement",
  "XRF 재측정 필요":"XRF Remeasurement Required",
  "정밀분석 필요":"Precision Analysis Required","정밀분석 및 XRF 재측정 필요":"Precision Analysis and XRF Remeasurement Required","정밀분석 인계":"Transfer to Precision Analysis",
  "정밀분석 인계 필요":"Transfer Required","정밀분석 인계 완료":"Transfer Complete",
  "정밀분석 의뢰 필요":"Precision Request Required","정밀분석 결과":"Precision Result",
  "정밀 결과":"Precision Result","정밀분석 결과 대기":"Awaiting Precision Result",
  "정밀분석 결과 없음":"No Precision Result","정밀분석 결과 확인":"Review Precision Result",
  "정밀분석 결과 확인 완료":"Precision Result Confirmed","정밀분석 대상":"Precision Analysis Target",
  "정밀분석 상세":"Precision Analysis Details","정밀분석 상세 열기":"Open Precision Details",
  "정밀분석 NG · 반려":"Precision NG · Rejected","정밀분석 후속조치 필요":"Precision Follow-up Required",
  "정밀분석 성적서 PDF":"Precision Report PDF","정밀분석 관리":"Precision Analysis Management",
  "정밀분석 인계 관리":"Transfer Management","인계 상태":"Transfer Status","인계 사유":"Transfer Reason",
  "인계일":"Transfer Date","인계 완료":"Transfer Complete","인계 필요":"Transfer Required",
  "현재 단계":"Current Stage","현재 조치":"Current Action","현재 주기":"Current Period",
  "선택 주기":"Selected Period","대상 원소":"Target Elements","대상 품목":"Target Item",
  "원소 함유 결과":"Element Detection","함유 선택":"Select Detection","결과 입력":"Enter Results",
  "결과 입력 가능":"Result Entry Enabled","결과 저장":"Save Result","상세 확인":"Detailed Review",
  "내부 판정":"Internal Judgment","XRF 내부 판정 기준":"XRF Internal Judgment Criteria",
  "내부 판정 조건":"Internal Judgment Condition","처리 기준":"Processing Rule",
  "함유량 (ppm)":"Content (ppm)","법적 기준치 (ppm)":"Legal Limit (ppm)",
  "측정 방법":"Measurement Method","측정 도래 예정":"Upcoming Measurements",
  "이행 상태 구성":"Compliance Distribution","부서별 위험도 구성":"Risk by Department",
  "주기 이행 상태":"Cycle Compliance","주기별 이행 현황":"Cycle Compliance Status",
  "이행 상세":"Compliance Details","주기별 산정 요약":"Assessment Summary by Period",
  "산정 요약":"Assessment Summary","산정 기준":"Basis","최근 측정으로":"Latest Measurement",
  "원소별 XRF 분석":"XRF Analysis by Element","최신 XRF":"Latest XRF","현재 XRF":"Current XRF",
  "원소별 XRF 트렌드":"XRF Trend by Element","원소별 트렌드 확인":"View Element Trends",
  "XRF 분석으로 돌아가기":"Back to XRF Analysis","주기별 XRF 추세":"XRF Trend by Period",
  "기준 포함":"Full Scale","측정값 확대":"Zoom to Measurements","전체":"All",
  "측정값":"Measured Values","현재 선택 Measurement 기준":"Based on Selected Measurement",
  "전체 측정 이력":"Complete Measurement History","추세를 그릴 관리주기가 아직 없습니다.":"No measurement history is available for the trend.",
  "정기 XRF 결과":"Periodic XRF Result","최초 위험도 평가":"Initial Risk Assessment",
  "적용 Risk":"Applied Risk","라이프사이클":"Lifecycle","C&R 등급":"C&R Level","도래월":"Due Month",
  "기준·도래일":"Reference / Due Date","주기 측정 필요":"Cycle Measurement Required",
  "주기 측정 권장":"Cycle Measurement Recommended","주기 미이행":"Cycle Overdue",
  "측정권장 (월 D-15 / 반기·연 D-30)":"Recommended (Monthly D-15 / Semiannual·Annual D-30)","기한 내 미측정":"Not Measured by Due Date",
  "과거 주기의 지연 측정 이력":"Late Measurements in Past Cycles",
  "첫 관리주기 내 등록":"Registered in Current Cycle","운영 관리 대상":"Actively Managed",
  "전체 품목":"All Items","활성 품목":"Active Items","고위험":"High Risk",
  "후속분석 진행중":"Follow-up In Progress","분석 Workflow 현황":"Analysis Workflow Status",
  "주기 관리 현황":"Cycle Management Status","현재 이행":"Currently Compliant",
  "고위험 / 후속분석 대상 품목":"High-risk / Follow-up Items",
  "위험성 평가 매트릭스 · R 기준":"Risk Assessment Matrix · R Basis",
  "품목 라이프사이클 참고":"Item Lifecycle Reference","대체 관계 체인":"Replacement Chains",
  "대체품":"Replacement Item","기존 측정이력":"Existing Measurements","보존":"Preserved","전환":"Transition",
  "R 검사주기":"R Inspection Cycle","XRF 분석 열기":"Open XRF Analysis",

  // ── 등록 / 변경 ──
  "요청 유형":"Request Type","요청 유형 선택":"Select Request Type","요청 완료":"Request Complete",
  "요청 요약":"Request Summary","기본정보 입력":"Enter Basic Information","요청 사유":"Request Reason",
  "요청 사유 / 비고":"Request Reason / Note","요청 부서":"Requesting Department",
  "기존 품목 대체 등록":"Replace Existing Item",
  "기존 품목 단종 / 사용 중지":"Discontinue / Stop Using Existing Item",
  "단종 / 사용 중지":"Discontinue / Stop Use","단종 / 사용 중지 사유":"Discontinuation Reason",
  "단종 / 사용 중지 정보 입력":"Enter Discontinuation Information",
  "단종 / 사용 중지 요청 제출":"Submit Discontinuation Request",
  "단종 정보 입력":"Enter Discontinuation Information","단종 대상":"Discontinuation Target",
  "단종 대상 품목":"Item to Discontinue","단종 대상 품목 선택":"Select Item to Discontinue",
  "단종 대상 기존 품목":"Existing Item to Discontinue","단종 사유":"Discontinuation Reason",
  "대체 대상":"Replacement Target","대체 사유":"Replacement Reason",
  "대체 대상 품목 선택":"Select Item to Replace","대체 대상 기존 품목":"Existing Item to Replace",
  "기존 품목 처리":"Existing Item Disposition","기존 품목 처리 방식":"Existing Item Disposition",
  "잔여 재고 처리":"Remaining Stock","잔여 재고 처리 방식":"Remaining Stock Disposition",
  "사용 중지":"Stop Use","공정 변경":"Process Change","고객 요구":"Customer Request",
  "공급 중단":"Supply Discontinued","대체품 적용 완료":"Replacement Applied","폐기":"Scrap",
  "반품":"Return","별도 보관":"Store Separately","기존 품목 단종":"Existing Item Discontinued",
  "공급업체 변경":"Supplier Change","제조사 변경":"Manufacturer Change","사양 변경":"Specification Change",
  "품질 개선":"Quality Improvement","원가 절감":"Cost Reduction",
  "의뢰자 등록":"Requester Registration","관리자 직접 등록":"Admin Direct Registration",
  "관리자 품목정보 수정":"Admin Item Info Edit","품목정보 수정":"Edit Item Info",
  "변경사항 저장":"Save Changes","관리자 인증 완료":"Administrator Verified",
  "한글 품목명":"Korean Item Name","영문 품목명":"English Item Name",
  "자동 발번":"Auto-numbered","자동 생성":"Auto-generated","임시 저장":"Save Draft",
  "승인 · 등록":"Approve · Register","신규 등록 요청 제출":"Submit New Registration Request",
  "유형 다시 선택":"Choose Another Type","단종 처리 원복":"Restore Discontinuation",
  "단종 원복 완료":"Discontinuation Restored","단종 승인":"Approve Discontinuation",
  "기존 주기 도래 / 측정 추가":"Existing Cycle Due / Add Measurement",
  "측정 추가 대상 품목":"Item for Additional Measurement",
  "변경 / 대체 등록":"Change / Replacement Registration",
  "사진":"Photo","품목 사진":"Item Photo","사진 없음":"No Photo",
  "사진 선택":"Choose Photo","사진 선택 (JPG, PNG, WEBP / 최대 10MB)":"Choose Photo (JPG, PNG, WEBP / max. 10 MB)",
  "선택한 품목 사진":"Selected Item Photo","품목 사진을 선택하세요.":"Select an item photo.",
  "품목 사진은 10MB 이하 파일만 선택하세요.":"Select an item photo no larger than 10 MB.",
  "품목 사진은 JPG, PNG, WEBP 형식만 선택하세요.":"Select an item photo in JPG, PNG, or WEBP format.",
  "XRF 결과 Excel 업로드":"Upload XRF Result Excel","XRF 원본 Excel 업로드":"Upload Original XRF Excel",
  "자동 산출 XRF":"Auto-derived XRF","업로드 파일":"Uploaded File","업로드된 파일 없음":"No Uploaded File",
  "업로드된 PDF 없음":"No Uploaded PDF","최근 업로드":"Latest Upload",
  "단종 처리 요청 리스트":"Discontinuation Request List",
  "품목 선택":"Select Item","-- 선택하세요 --":"-- Select --",
  "단종 / 사용 중지 처리 요청":"Discontinuation / Stop-use Request",
  "현재 처리 상태":"Current Processing Status",
  "이 품목은 단종 승인 처리됨":"This item has an approved discontinuation",

  // ── 문장 ──
  "조건에 맞는 품목이 없습니다.":"No items match the current filters.",
  "품목을 선택하면 XRF 분석 결과가 표시됩니다.":"Select an item to view its XRF analysis.",
  "정밀분석 대상이 없습니다.":"There are no precision analysis targets.",
  "선택된 정밀분석 항목이 없습니다.":"No precision analysis case is selected.",
  "정밀분석 대상 원소가 없습니다.":"There are no target elements for precision analysis.",
  "현재 조치 대상이 없습니다.":"There is no item requiring action.",
  "선택한 주기에 연결된 XRF 측정 이력이 없습니다.":"No XRF measurement is linked to the selected period.",
  "선택한 주기에 연결된 측정 이력이 없어 XRF 결과를 표시하지 않습니다. 최초 등록 시 산정한 Final Risk는 유지됩니다.":"No measurement is linked to the selected period, so no XRF result is shown. The Final Risk set at initial registration is retained.",
  "R 단계에서 XRF Level과 C&R Level로 Final Risk와 검사주기를 확정하며, P1은 R 등록일에서 해당 주기만큼 지난 시점에 생성됩니다.":"At the R stage, Final Risk and the inspection cycle are fixed from the XRF Level and C&R Level. P1 is created one cycle after the R registration date.",
  "Judgment와 후속조치만 갱신하며 Final Risk와 검사주기는 R 최초 평가 결과를 유지합니다.":"Only the judgment and follow-up action are updated. Final Risk and the inspection cycle keep the initial R assessment.",
  "강조된 원소만 후속 분석 대상으로 전달됩니다.":"Only highlighted elements are sent for follow-up analysis.",
  "원본 보고서 Judgment 기준 후속조치 대상이 없습니다.":"No follow-up target based on the original report judgment.",
  "같은 주기에 여러 측정이 있으면 같은 Pn으로 각각 표시되며, 측정 이력을 선택하면 상단 값과 우측 원소 분석도 해당 측정으로 전환됩니다.":"Multiple measurements in the same period are listed under the same Pn. Selecting one switches the values above and the element analysis on the right.",
  "Internal Limit = Legal Limit의 70%":"Internal Limit = 70% of Legal Limit",
  "Content ≤ Internal Limit 또는 N.D.":"Content ≤ Internal Limit or N.D.",
  "XRF 1회 재측정\n* 재측정 결과가 ?? / NG이면 정밀분석":"One XRF remeasurement\n* If the remeasurement result is ?? / NG, proceed to Precision Analysis",
  "미함유 또는 함유량 ≤ 법적 기준치 → OK · 법적 기준치 초과 → NG":"Not detected or content ≤ legal limit → OK · above the legal limit → NG",
  "법적 기준치를 초과한 대상 원소가 있습니다.":"One or more target elements exceed the legal limit.",
  "모든 대상 원소가 미함유이거나 법적 기준치 이내입니다.":"All target elements are not detected or within the legal limit.",
  "성적서 내용을 기준으로 대상 원소의 함유 여부와 함유량을 입력하면 정밀 결과가 자동 산출됩니다.":"Enter the detection status and content of each target element from the report; the precision result is calculated automatically.",
  "정밀분석 인계 버튼을 먼저 클릭해야 합니다. 인계 후 성적서 PDF를 업로드하면 결과 입력이 활성화됩니다.":"Click the transfer button first. Result entry is enabled after uploading the report PDF.",
  "1. 정밀분석 인계 버튼을 클릭하세요.  2. 성적서 PDF를 업로드하세요.  3. 원소별 결과를 입력하세요.":"1. Click the transfer button.  2. Upload the report PDF.  3. Enter results by element.",
  "인계 완료 후 정밀분석 성적서 PDF를 등록해야 원소별 결과 입력이 활성화됩니다.":"After the transfer, upload the precision report PDF to enable element-level result entry.",
  "인계가 완료되었습니다. 정밀분석 성적서 PDF를 업로드하면 원소별 결과 입력이 가능합니다.":"Transfer complete. Upload the precision report PDF to enter element-level results.",
  "정밀분석 성적서 PDF를 업로드해야 원소별 결과를 입력할 수 있습니다.":"Upload the precision report PDF to enter element-level results.",
  "정밀분석 인계, 원본/결과 파일, 원소 함유 확인을 항목별로 관리합니다.":"Manage the transfer, source and result files, and element detection review for each case.",
  "현재 상태: 보류 · XRF 결과 대기":"Current status: On hold · Awaiting XRF result",
  "의뢰자가 XRF 측정을 요청했지만 아직 판정 지표가 없어 보류 상태입니다. XRF 결과 업로드 후 R 위험도 평가와 주기별 이행현황이 생성됩니다.":"The requester asked for an XRF measurement, but no evaluation data exists yet, so the item is on hold. Uploading the XRF result creates the R risk assessment and the cycle plan.",
  "XRF 측정 결과가 아직 업로드되지 않았습니다.":"The XRF measurement result has not been uploaded yet.",
  "관리자가 결과 파일을 업로드하면 원소별 분석, Final Risk, 재측정 여부가 자동 표시됩니다.":"Once the administrator uploads the result file, element analysis, Final Risk, and remeasurement status appear automatically.",
  "관리자만 XRF 결과를 업로드할 수 있습니다.":"Only administrators can upload XRF results.",
  "관리자 XRF 결과 업로드 (.xlsx)":"Admin XRF Result Upload (.xlsx)",
  "업로드 후 XRF 결과, 재측정 여부, Approval_Status가 자동 반영됩니다.":"After upload, the XRF result, remeasurement status, and Approval_Status are applied automatically.",
  "단종 요청은 XRF 측정 대상이 아니므로 원소 분석, Final Risk, 주기별 이행현황을 계산하지 않습니다. 관리자는 요청 사유만 검토하여 승인 또는 반려를 적용합니다.":"A discontinuation request is not an XRF target, so element analysis, Final Risk, and cycle compliance are not calculated. The administrator reviews the reason and approves or rejects it.",
  "단종 요청 상세 정보가 없습니다.":"No discontinuation request details.",
  "저장 즉시 부자재 리스트와 XRF 분석 화면에 같은 값이 반영됩니다.":"Saved values apply immediately to the auxiliary material list and the XRF analysis screen.",
  "기존 관리 품목과 관계없이 새로운 부자재/공정부자재를 등록합니다.":"Register a new auxiliary or process material unrelated to existing managed items.",
  "현재 사용 중인 품목을 새로운 품목으로 변경하거나 대체합니다.":"Change or replace an item currently in use with a new item.",
  "신규 품목 추가 없이 기존 품목의 사용 중지 또는 단종을 요청합니다.":"Request stop-use or discontinuation of an existing item without adding a new item.",
  "신규 대체품 품번은 선택한 기존 품번 뒤에 순번을 붙여 자동 생성합니다. 예: 기존 품번 뒤에 -01, -02 순번을 추가합니다. 기존 품목은 변경·대체 이력으로 전환되고, 신규 대체품은 신규 등록 품목으로 주기 관리를 시작합니다.":"The replacement part number is generated by appending a sequence to the selected part number (e.g. -01, -02). The existing item moves to change / replacement history, and the replacement starts cycle management as a newly registered item.",
  "방식 A — XRF 사전 완료":"Option A — XRF Already Completed",
  "방식 B — XRF 측정 의뢰":"Option B — Request XRF Measurement",
  "XRF 원본 보고서가 이미 있을 때":"When the original XRF report already exists",
  "파일 업로드 → 값 추출 → 즉시 Risk 판정":"Upload file → extract values → immediate risk assessment",
  "아래 업로드 영역에서 파일 선택":"Choose a file in the upload area below",
  "보고서가 없을 때":"When there is no report",
  "의뢰서 작성 → 관리자 알림 → 측정 완료 후 판정":"Submit request → notify administrator → assess after measurement",
  "XRF 파일 선택 (.xlsx 또는 파싱 결과 .json)":"Choose XRF File (.xlsx or parsed .json)",
  "XRF 결과 파일 선택 (.xlsx)":"Choose XRF Result File (.xlsx)",
  "업로드한 XRF Report의 Content(ppm), Std.Deviation(ppm), Judgment를 자동 추출합니다. XRF OK/NG는 Content를 원소별 Internal Limit(법적 기준의 70%)과 비교해 산정하고, 원본 Judgment는 ??/Cannot Judge 재측정 판단과 추적용으로 보존합니다.":"Content(ppm), Std.Deviation(ppm), and Judgment are extracted automatically from the uploaded report. OK/NG is decided by comparing Content with each element's Internal Limit (70% of the legal limit); the original judgment is kept for ?? / Cannot Judge decisions and traceability.",
  "브라우저에서 원본 보고서의 O5 측정일과 Element / Content / Std.Deviation / Judgment 행을 직접 읽습니다. 등록 후에는 같은 결과를 SharePoint XRF_Measurements와 XRF_ElementResults에 저장하도록 연결합니다.":"The browser reads the O5 measurement date and the Element / Content / Std.Deviation / Judgment rows directly from the original report. After registration the same result is stored in SharePoint XRF_Measurements and XRF_ElementResults.",
  "XRF Report Content와 Internal Limit(법적 기준의 70%)을 비교해 XRF Result와 R 위험도 평가를 자동 산정합니다.":"The XRF result and the R risk assessment are derived by comparing the report Content with the Internal Limit (70% of the legal limit).",
  "측정 추가 대상 품목을 선택한 뒤 파일을 넣으면 해당 주기에 바로 저장되고 XRF 분석 화면으로 이동합니다.":"Select the target item, then upload the file. It is saved to that period and opens in the XRF analysis screen.",
  "기존 품목의 정기 XRF 결과를 추가 등록":"Add a periodic XRF result for an existing item",
  "새 품목을 즉시 등록하고 XRF 결과까지 반영":"Register a new item immediately and apply the XRF result",
  "기존 품목과 신규 대체품을 연결":"Link the existing item to its replacement",
  "기존 품목의 사용 중지 또는 이력 전환":"Stop use of the existing item or move it to history",
  "관리자가 수동으로 신규 도입, 기존 주기 측정 추가, 변경·대체, 단종 처리를 직접 선택합니다.":"The administrator directly selects new registration, an additional periodic measurement, replacement, or discontinuation.",
  "기본 정보, C&R 등급, XRF 결과, 위험도 산정, 승인 상태를 관리자가 직접 확인합니다.":"The administrator reviews basic information, C&R level, XRF result, risk assessment, and approval status directly.",
  "승인 시 선택 품목은 Discontinued로 전환되고, 리스트에서는 이력 상태로 관리됩니다.":"On approval the selected item becomes Discontinued and is managed as a history record in the list.",
  "최초 등록 시 R XRF Level과 C&R 등급으로 산정한 위험성 평가 결과입니다. 후속 정밀분석 승인 결과와는 별도로 유지합니다.":"This risk assessment is calculated at initial registration from the R XRF Level and the C&R level. It is kept separate from later precision analysis approvals.",
  "기존 품목이 신규 항목으로 대체된 이력":"History of existing items replaced by new items",
  "R 단계에서 확정된 Final Risk 기준":"Based on the Final Risk fixed at the R stage",
  "단종·취소·요청 기록":"Discontinuation / cancellation / request records",
  "기존 품목의 사용 중지 또는 단종을 요청합니다.":"Request stop-use or discontinuation of an existing item.",
  "단종 또는 사용 중지할 대상 품목을 선택하세요.":"Select the item to discontinue or stop using.",
  "단종 또는 사용 중지할 기존 품목을 선택하세요":"Select the existing item to discontinue or stop using",
  "대체할 기존 품목을 선택하세요":"Select the existing item to replace",
  "대체할 기존 품목을 선택하세요.":"Select the existing item to replace.",
  "기존 품목을 선택하세요":"Select an existing item",
  "단종/사용 중지 관련 특이사항을 입력하세요":"Enter any notes about the discontinuation or stop use",
  "신규 도입 또는 대체 등록 사유를 입력하세요":"Enter the reason for the new item or replacement",
  "기타 단종 사유를 구체적으로 입력하세요":"Describe the other discontinuation reason",
  "기타 대체 사유를 구체적으로 입력하세요":"Describe the other replacement reason",
  "전체 열 검색":"Search all columns",
  "관리자 비밀번호를 입력하세요.":"Enter the administrator password.",
  "관리자 비밀번호가 올바르지 않습니다.":"The administrator password is incorrect.",
  "품목 정보 편집은 관리자만 가능합니다.":"Only administrators can edit item information.",
  "품번, 한글 품목명, 공정명은 필수 입력입니다.":"Part No., Korean item name, and process are required.",
  "감사 추적을 위해 원복 사유를 입력해야 합니다.":"A restore reason is required for the audit trail.",
  "원복 사유를 입력하세요.":"Enter the restore reason.",
  "관리자 오선택":"Administrator selection error",
  "단종 처리를 승인하시겠습니까?":"Approve this discontinuation?",
  "승인된 단종 처리를 원복하시겠습니까?":"Restore the approved discontinuation?",
  "승인 후에도 '단종 처리 원복'으로 복구할 수 있으며 처리 이력이 남습니다.":"You can restore it later with Restore Discontinuation, and the history is retained.",
  "기존 XRF 측정 이력은 그대로 유지되고, 품목의 사용 상태와 이행주기 판정만 승인 전 상태로 복구됩니다.":"Existing XRF measurements are kept. Only the item's usage status and cycle judgment return to the pre-approval state.",
  "원복할 대상 품목을 찾을 수 없습니다. 단종 요청의 대상 품번을 확인하세요.":"The target item could not be found. Check the part number in the discontinuation request.",
  "측정 추가 대상 품목을 선택하세요.":"Select the item for the additional measurement.",
  "측정 추가 대상 품목을 먼저 선택하세요.":"Select the item for the additional measurement first.",
  "XRF 원본 보고서를 선택하고 파싱을 완료하세요.":"Select the original XRF report and complete parsing.",
  "XRF 원본 보고서를 파싱하지 못했습니다.":"Failed to parse the original XRF report.",
  "XRF 보고서를 파싱하지 못했습니다.":"Failed to parse the XRF report.",
  "XRF 원본 보고서 파싱을 완료한 뒤 등록 요청을 제출하세요.":"Complete parsing of the original XRF report before submitting the request.",
  "관리자 직접 등록은 XRF 원본 보고서를 먼저 선택하고 파싱을 완료해야 합니다.":"For admin direct registration, select and parse the original XRF report first.",
  "측정 결과를 연결할 주기를 찾지 못했습니다.":"No period was found to link the measurement result to.",
  "해당 주기에는 이미 XRF 측정값이 있습니다. 추가 등록은 재측정(??)이 필요한 경우에만 가능합니다.":"This period already has an XRF measurement. Additional entries are allowed only when a remeasurement (??) is required.",
  "XRF 재측정은 1회까지만 가능합니다. 이후에는 정밀분석 단계에서 처리하세요.":"Only one XRF remeasurement is allowed. Handle further cases in precision analysis.",
  "정밀분석 인계를 먼저 완료한 뒤 성적서 PDF를 업로드하세요.":"Complete the transfer before uploading the report PDF.",
  "XRF 원본 파일을 먼저 선택하세요.":"Select the original XRF file first.",
  "인계 완료 후 선택 가능":"Available after transfer","XRF 원본 파일 선택 후 업로드 가능":"Available after selecting the original XRF file",
  "정밀분석 성적서는 PDF 파일만 업로드할 수 있습니다.":"Only PDF files can be uploaded as the precision report.",
  "Material_Category가 Chemical인 경우 Material_State를 필수로 선택하세요.":"Material_State is required when Material_Category is Chemical.",
  "Material_State를 선택하세요":"Select Material_State","Item_Name을 입력하세요.":"Enter the item name.",
  "Part Number를 확인하세요.":"Check the part number.",
  "단종 사유에서 기타를 선택한 경우 상세 설명을 입력하세요.":"Describe the reason when Other is selected for discontinuation.",
  "대체 사유에서 기타를 선택한 경우 상세 설명을 입력하세요.":"Describe the reason when Other is selected for replacement.",
  "읽을 수 있는 워크시트가 없습니다.":"No readable worksheet was found.",
  "파일을 읽는 중이거나 파싱 결과가 없습니다.":"The file is being read or no parsed result is available.",
  "Element, Content(ppm), Std.Deviation(ppm), Judgment 행을 찾지 못했습니다.":"Could not find the Element, Content(ppm), Std.Deviation(ppm), and Judgment rows.",
  "판정 불가 - 일부 원소 데이터 부족":"Cannot judge - some element data is missing",
  "관리자 최초 XRF 업로드":"Admin initial XRF upload",
  "클릭하여 품목 상세 열기":"Click to open item details",
  "XRF 분석으로 이동":"Open XRF analysis","XRF 분석 상세로 이동":"Open XRF analysis details",
  "XRF 결과 상세로 이동":"Open XRF result details","정밀분석 상세로 이동":"Open precision analysis details",
  "XRF 상세로 이동":"Open XRF details",
  "R 기준 위험도와 XRF 상세로 이동":"Open R risk and XRF details",
  "해당 품목의 이행/XRF 상세로 이동":"Open compliance and XRF details",
  "현재 조치의 정밀분석 상세로 이동":"Open precision details for the current action",
  "현재 조치의 주기/XRF 상세로 이동":"Open cycle and XRF details for the current action",
  "정밀분석 대상이 없어 XRF 상세로 이동":"No precision target, opening XRF details",
  "현재 Workflow 상세로 이동":"Open current workflow details",
  "선택 주기의 XRF 결과 파일 업로드":"Upload the XRF result file for the selected period",
  "후속 분석 필요":"Follow-up analysis required","Internal Limit 이내":"Within Internal Limit",
  "XRF Report 원본 Content":"Original XRF report Content",
  "XRF Report 원본 Std.Deviation":"Original XRF report Std.Deviation",
  "Cl Content + Br Content 합계":"Sum of Cl and Br content",
  "Cl과 Br의 원본 Content 값을 합산":"Sum of the original Cl and Br content values",
  "XRF 재측정 필요 원소 없음":"No elements require XRF remeasurement",
  "데이터 확인 필요":"Data review required","판정 확인 필요":"Judgment review required",
  "Judgment 확인 필요":"Judgment review required",
  "인계 완료 후 업로드 가능":"Available after transfer",
  "성적서 필요":"Report required","성적서 등록 완료":"Report uploaded",
  "T-11 화면 페이지/롤링 검증 전용 테스트 데이터":"T-11 UI pagination / rolling validation test data",
  "[단종 요청]":"[Discontinuation Request]",
  "후속조치 대상 아님":"No follow-up target",
  "이미 사용 중인 품번입니다":"Part number already in use",
  "지원 형식은 .xlsx, .xlsm, .xls, .csv, .json 입니다.":"Supported formats: .xlsx, .xlsm, .xls, .csv, .json.",

  // ── 2차 보강 : 화면에 실제 렌더링되는 나머지 문구 ──
  "측정 의뢰":"Measurement Request","최종 컨펌":"Final Confirmation",
  "사용/단종 승인":"Use / Discontinuation Approved","사용/단종 반려":"Use / Discontinuation Rejected",
  "보류/검토중":"On Hold / Under Review",
  "신규 도입 기본정보":"New Item Basic Information",
  "대체 대상 및 신규 대체품 기본정보":"Replacement Target and New Item Information",
  "신규 대체품 Item_Name":"Replacement Item Name","신규 등록 요청 품목":"New Registration Request Item",
  "예: 제품 포장용 엠보스":"e.g. Packaging emboss","예: ABC Chemical":"e.g. ABC Chemical",
  "예: Solid / Powder":"e.g. Solid / Powder","기타 상세 미입력":"Other details not entered",
  "예상 Approval_Status":"Expected Approval_Status","측정 후 산출":"Calculated after measurement",
  "관리자 직접 등록 / 기존 주기 측정 추가":"Admin Direct Registration / Add Periodic Measurement",
  "관리자 직접 등록 / 변경/대체 등록":"Admin Direct Registration / Change / Replacement",
  "관리자 직접 등록 / 신규 도입":"Admin Direct Registration / New Item",
  "관리자 직접 등록 / 단종 처리":"Admin Direct Registration / Discontinuation",
  "관리자 직접 등록 / 기존 품목 대체 등록":"Admin Direct Registration / Replace Existing Item",
  "의뢰자 등록 · 신규 도입":"Requester Registration · New Item",
  "의뢰자 등록 · 기존 품목 대체 등록":"Requester Registration · Replace Existing Item",
  "Cl + Br Content 합계 기준 · Internal Limit 1,050 ppm":"Sum of Cl and Br content · Internal Limit 1,050 ppm",
  "신규 품목 승인 후 기존 재고 소진 후 단종":"Discontinue after the new item is approved and stock is used up",
  "승인 후 즉시 단종":"Discontinue immediately after approval",
  "신규 품목 승인 후 병행 사용":"Use in parallel after the new item is approved",
  "기존 품목 유지, 신규 품목 추가 등록만 진행":"Keep the existing item and only register the new item",
  "재고 소진 후 사용 중지":"Stop use after the stock is used up",
  "측정일 미확인":"Measurement date unknown",
  "원소별 후속조치 상태":"Follow-up status by element","ppm 입력":"Enter ppm",
  "도래일":"Due Date","추가 측정일":"Additional Measurement Date","기준일":"Reference Date",
  "사용 여부":"Usage","고객 사전 신고":"Customer pre-notice","측정 연결":"measurement linked",
  "관리자 처리":"Administrator action","관리자 단종 승인":"Administrator approved discontinuation",
  "재고 처리":"Stock disposition","성적서 등록 완료":"Report uploaded",
  "원소별 결과를 입력할 수 있습니다.":"element-level results can now be entered.",
  "관리자 입력 내용과 파싱 결과를 브라우저 임시 저장소에 저장했습니다.":"The entered values and parsed result were saved to the browser's local storage.",
  "임시 저장 실패":"Failed to save draft","XRF 파싱 실패":"XRF parsing failed",
  "XRF 재측정 필요 원소":"Elements requiring XRF remeasurement",
  "필수 원소 결과를 찾지 못했습니다":"Required element results were not found",
  "XRF 보고서 형식을 충분히 찾지 못했습니다. 제목 일치 점수":"The XRF report format could not be recognized. Header match score",
  "파일 품번":"File part no.","등록 품번":"Registered part no.","선택 품번":"Selected part no.",
  "업로드 파일 품번":"Uploaded file part no.",
  "이 일치하지 않습니다.":" do not match.","과 등록 품번":" and registered part no.",
  "관리자 최초 XRF 업로드":"Admin initial XRF upload","최초 XRF 측정 및 R 위험도 평가 기준":"Initial XRF measurement and R risk assessment basis",
  "정기 Pn과 별도 표시":"shown separately from the regular Pn",
  "R 이후 실제 XRF 측정 이력":"Actual XRF measurements after R",
  "사용 자제":"Limit use","관리 필요":"Management required","정상":"Normal",
  "사용 불허 Risk":"Not Allowed risk","기한초과":"overdue","고정 주기":"fixed cycle",
  "R 등록일":"R registration date","측정 이력은":"measurements are grouped into",
  "에 통합":"","년 측정 이력":" measurement history",
  "파일":"File","의 70%":" × 70%","완료":"complete",

  // ── 3차 보강 : 템플릿 문자열로 조합되는 런타임 문구의 연결어 ──
  "최신 XRF":"Latest XRF","최신 Level":"Latest Level","최신":"Latest","측정값":"Measured value",
  "과 선택 품번":" and selected part no.","과 등록 품번":" and registered part no.",
  "관리자 직접 주기":"Admin direct period","관리자 주기":"Admin period","관리자 직접":"Admin direct",
  "관리자":"Administrator","의뢰자":"Requester","XRF 업로드":"XRF upload","업로드":"upload",
  "요청":"request","등록":"registration","처리":"processing","기한 초과":"overdue",
};

// 정확 일치가 실패했을 때 사용할 부분 치환 규칙입니다. 긴 표현부터 순서대로 적용합니다.
const UI_EN_PAIRS = Object.entries(UI_EN_MAP).sort((a,b)=>b[0].length-a[0].length);

function translateUiTextKoToEn(value){
  if(value==null) return value;
  let text=String(value);
  if(!/[가-힣]/.test(text)) return text;

  // 한/영 병기로 저장된 관리방법은 영문 괄호문만 표시합니다.
  const bilingual=text.match(/^([\s\S]*?)\n\(\s*([\s\S]*?)\s*\)\s*$/);
  if(bilingual && /[A-Za-z]/.test(bilingual[2])) return bilingual[2].trim();

  const prefix=text.match(/^\s*/)?.[0]||"";
  const suffix=text.match(/\s*$/)?.[0]||"";
  const trimmed=text.trim();
  if(Object.prototype.hasOwnProperty.call(UI_EN_MAP,trimmed)) return prefix+UI_EN_MAP[trimmed]+suffix;

  // 날짜·건수·차수·지연일 등 동적 문구를 먼저 정규화합니다.
  text=text.replace(/(\d{4})년\s*측정\s*이력/g,"$1 measurement history");
  text=text.replace(/(\d+)월/g,(_,m)=>{
    const names=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
    const n=Number(m); return n>=1&&n<=12?names[n-1]:`${m} mo`;
  });
  text=text.replace(/(\d+)차/g,"Attempt $1");
  text=text.replace(/총\s*(\d+)건/g,"Total $1 items");
  text=text.replace(/기한초과\s*(\d+)건/g,"$1 overdue");
  text=text.replace(/(\d+)건/g,"$1 items");
  text=text.replace(/(\d+)번 페이지/g,"Page $1");
  text=text.replace(/\+(\d+)\s*일/g,"+$1 days");
  text=text.replace(/(\d+)\s*일\s*지연/g,"$1 days late");
  text=text.replace(/기준 향후 6개월/g,"· Next 6 months");
  text=text.replace(/기존 측정이력\s*(\d+)건\s*보존\s*·\s*전환\s*(\d{4}-\d{2}-\d{2}|—)/g,"Existing Measurements: $1 preserved · Transition $2");
  text=text.replace(/기존 품목\s*([A-Z0-9-]+)\s*대체/g,"Replaces Existing Item $1");
  text=text.replace(/([가-힣A-Za-z()·\s]+?):\s*/g,(m,label)=>{
    const key=label.trim();
    return UI_EN_MAP[key] ? `${UI_EN_MAP[key]}: ` : m;
  });

  for(const [ko,en] of UI_EN_PAIRS){
    if(text.includes(ko)) text=text.split(ko).join(en);
  }
  return text;
}

const UI_TEXT_NODE_SOURCE = new WeakMap();
const UI_ATTR_SOURCE = new WeakMap();
const UI_TRANSLATABLE_ATTRS = ["placeholder","title","aria-label"];

function shouldSkipI18nNode(node){
  const el=node?.nodeType===1 ? node : node?.parentElement;
  return !!el?.closest?.('[data-i18n-skip="true"]');
}

function localizeUiTextNode(node,lang){
  if(!node || node.nodeType!==3 || shouldSkipI18nNode(node)) return;
  const current=node.nodeValue??"";
  let source=UI_TEXT_NODE_SOURCE.get(node);
  if(lang==="en"){
    const expected=source==null?null:translateUiTextKoToEn(source);
    // React가 언어 전환 후 새 한국어 동적 값을 다시 렌더한 경우 source를 갱신합니다.
    if(source==null || current!==expected) source=current;
    UI_TEXT_NODE_SOURCE.set(node,source);
    const next=translateUiTextKoToEn(source);
    if(current!==next) node.nodeValue=next;
  }else if(source!=null){
    const expectedEn=translateUiTextKoToEn(source);
    if(current===expectedEn && current!==source){
      node.nodeValue=source;
    }else if(current!==source){
      // 한국어 모드에서 React가 새로운 동적 원본값을 렌더하면 과거 source로 되돌리지 않습니다.
      UI_TEXT_NODE_SOURCE.set(node,current);
    }
  }
}

function localizeUiElementAttributes(el,lang){
  if(!el || el.nodeType!==1 || shouldSkipI18nNode(el)) return;
  let sources=UI_ATTR_SOURCE.get(el);
  if(!sources){ sources={}; UI_ATTR_SOURCE.set(el,sources); }
  for(const attr of UI_TRANSLATABLE_ATTRS){
    if(!el.hasAttribute(attr)) continue;
    const current=el.getAttribute(attr)??"";
    let source=sources[attr];
    if(lang==="en"){
      const expected=source==null?null:translateUiTextKoToEn(source);
      if(source==null || current!==expected) source=current;
      sources[attr]=source;
      const next=translateUiTextKoToEn(source);
      if(current!==next) el.setAttribute(attr,next);
    }else if(source!=null){
      const expectedEn=translateUiTextKoToEn(source);
      if(current===expectedEn && current!==source){
        el.setAttribute(attr,source);
      }else if(current!==source){
        sources[attr]=current;
      }
    }
  }
}

function applyUiLanguage(root,lang){
  if(!root || typeof document==="undefined") return;
  if(!shouldSkipI18nNode(root)) localizeUiElementAttributes(root,lang);
  const walker=document.createTreeWalker(root,NodeFilter.SHOW_ELEMENT|NodeFilter.SHOW_TEXT);
  let node=walker.nextNode();
  while(node){
    if(node.nodeType===3) localizeUiTextNode(node,lang);
    else localizeUiElementAttributes(node,lang);
    node=walker.nextNode();
  }
}

// -----------------------------------------------------------------------------
// Workflow mock data
// 실제 회사 DB 연결 전 XRF → 정밀분석 → 정밀분석 → 승인 화면/상태머신을 검증하기 위한 임시 데이터입니다.
// 운영 전에는 false로 바꾸고 AnalysisCase/XRF/정밀분석/Precision/Approval API로 대체하세요.
const ENABLE_WORKFLOW_MOCK_DATA = false;

function mockElementSet(kind="OK"){
  const base={};
  ELEMENTS.forEach(el=>{
    const limit=XRF_ELEMENT_LIMITS[el]||1000;
    base[el]={rawPpm:0,ppm:0,rawSigma:1,sigma:1,rawJudgement:"N.D.",judge:"N.D.",limit,dataMissing:false,isND:true};
  });
  if(kind==="NG"){
    base.Cd={...base.Cd,rawPpm:82,ppm:82,rawSigma:3,sigma:3,rawJudgement:"NG",judge:"NG",dataMissing:false,isND:false};
  }else if(kind==="REMEASURE"){
    base.Cd={...base.Cd,rawPpm:45,ppm:45,rawSigma:20,sigma:20,rawJudgement:"??",judge:"??",dataMissing:false,isND:false};
  }else if(kind==="DASH_ND"){
    base.Cd={...base.Cd,rawPpm:0,ppm:0,rawSigma:0,sigma:0,rawJudgement:"----",judge:"----",dataMissing:false,isND:true};
  }
  const cl=base.Cl, br=base.Br;
  const clBr=buildClBrResult(sourcePpmValue(cl),sourcePpmValue(br));
  base["Cl+Br"]={...base["Cl+Br"],rawPpm:clBr.ppm,ppm:clBr.ppm,rawSigma:0,sigma:0,rawJudgement:clBr.judgement,judge:clBr.judgement,level:clBr.level,xrfLevel:clBr.level,limit:clBr.legalLimit,legalLimit:clBr.legalLimit,internalLimit:clBr.internalLimit,dataMissing:clBr.ppm==null,isND:!!clBr.isND};
  return base;
}
function makeMockWorkflowItem(code,name,{xrf="OK",approvalStatus="승인",finalRisk="M",due="2026-08-31",measured="2026-08-05",history=[]}={}){
  const latestId=history.length ? history[history.length-1].id : (measured?`${code}_20260805_01`:null);
  const latestHistory=history.length ? history[history.length-1] : null;
  const latest=latestHistory ? {
    id:latestHistory.id,date:latestHistory.date,role:latestHistory.measurementRole,attemptNo:latestHistory.attemptNo,
    xrfWorst:latestHistory.xrfResult, xrfResult:latestHistory.xrfResult, approvalStatus, elements:latestHistory.elements
  } : (measured ? {id:latestId,date:measured,role:"Periodic",attemptNo:1,xrfWorst:xrf,xrfResult:xrf,approvalStatus,elements:mockElementSet(xrf==="NG"?"NG":xrf==="재측정 필요"?"REMEASURE":"OK")} : null);
  const hist=history.length ? history : (latest ? [{date:latest.date,id:latest.id,periodNo:1,label:"P1",measurementRole:latest.role,attemptNo:1,xrfWorst:xrf,xrfResult:xrf,approvalStatus,elements:latest.elements,status:"measured"}] : []);
  return {
    code,name,nameEn:`Workflow mock ${code}`,dept:"Assembly",respDept:"PQE",type:"Auxiliary Materials",
    materialCategory:"Plastic",materialState:"Solid",crType:"직접접촉+비잔류",crLevel:"M",
    approvalStatus,category:"existing",categoryLabel:"기존 품목",lifecycle:"ExistingActive",isCurrent:true,
    firstDate:"2026-02-28",cycleResetDate:"2026-02-28",cycle:"Semiannual",cycleMonths:6,finalRisk,
    latestMeasurementId:latestId,firstMeasurementId:hist[0]?.id||null,lastMeasured:latest?.date||null,measurementCount:hist.length,
    xrfWorst:xrf,latestXrfResult:xrf,latest,history:hist,
    periods:[
      {num:0,label:"R",displayLabel:"R",isReference:true,referenceDate:"2026-02-28",date:"2026-02-28",measured:"2026-02-28",status:"reference",measurementId:null,closeType:"registered",displaySeq:0,isDisplayTarget:false,isCurrentCycle:true},
      {num:1,label:"P1",displayLabel:"P1",isReference:false,due,targetDate:due,windowStart:"2026-08-01",measured:latest?.date||null,status:latest?"compliant":(dateGt(TODAY_STR,due)?"overdue":"due_soon"),measurementId:latestId,closeType:latest?"measured":"system_open",displaySeq:1,isDisplayTarget:true,isCurrentCycle:true},
      {num:2,label:"P2",displayLabel:"P2",isReference:false,due:"2027-02-28",targetDate:"2027-02-28",windowStart:"2027-01-29",measured:null,status:"upcoming",measurementId:null,closeType:"system_open",displaySeq:2,isDisplayTarget:true,isCurrentCycle:true}
    ]
  };
}

const MOCK_WORKFLOW_DB = {
  analysisCases:{
    CASE_T01_P1:{id:"CASE_T01_P1",itemCode:"T-01",periodNo:1,stage:"APPROVED",xrfAttemptIds:["XRF_T01_01"],approvalId:"APR_T01"},
    CASE_T02_P1:{id:"CASE_T02_P1",itemCode:"T-02",periodNo:1,stage:"APPROVED",xrfAttemptIds:["XRF_T02_01"],approvalId:"APR_T02"},
    CASE_T03_P1:{id:"CASE_T03_P1",itemCode:"T-03",periodNo:1,stage:"APPROVED",xrfAttemptIds:["XRF_T03_01","XRF_T03_02"],approvalId:"APR_T03"},
    CASE_T04_P1:{id:"CASE_T04_P1",itemCode:"T-04",periodNo:1,stage:"PRECISION_TRANSFER_REQUEST_REQUIRED",xrfAttemptIds:["XRF_T04_01","XRF_T04_02"],precisionTransferId:"정밀분석_T04_01",approvalId:"APR_T04"},
    CASE_T05_P1:{id:"CASE_T05_P1",itemCode:"T-05",periodNo:1,stage:"APPROVED",xrfAttemptIds:["XRF_T05_01"],precisionTransferId:"정밀분석_T05_01",approvalId:"APR_T05"},
    CASE_T06_P1:{id:"CASE_T06_P1",itemCode:"T-06",periodNo:1,stage:"PRECISION_FOLLOWUP_REQUIRED",xrfAttemptIds:["XRF_T06_01"],precisionTransferId:"정밀분석_T06_01",precisionId:"PREC_T06_01",followUpId:"FU_T06_01",approvalId:"APR_T06"},
    CASE_T07_P1:{id:"CASE_T07_P1",itemCode:"T-07",periodNo:1,stage:"APPROVED",xrfAttemptIds:["XRF_T07_01"],precisionTransferId:"정밀분석_T07_01",precisionId:"PREC_T07_01",followUpId:"FU_T07_01",approvalId:"APR_T07"},
    CASE_T08_P1:{id:"CASE_T08_P1",itemCode:"T-08",periodNo:1,stage:"PRECISION_NG_REVIEW",xrfAttemptIds:["XRF_T08_01"],precisionTransferId:"정밀분석_T08_01",precisionId:"PREC_T08_01",followUpId:"FU_T08_01",approvalId:"APR_T08"},
    CASE_T09_P1:{id:"CASE_T09_P1",itemCode:"T-09",periodNo:1,stage:"XRF_REQUIRED",xrfAttemptIds:[],approvalId:"APR_T09"},
    CASE_T10_P1:{id:"CASE_T10_P1",itemCode:"T-10",periodNo:1,stage:"XRF_REQUIRED",xrfAttemptIds:[],approvalId:"APR_T10"}
  },
  xrfMeasurements:{
    XRF_T01_01:{id:"XRF_T01_01",caseId:"CASE_T01_P1",attemptNo:1,role:"Periodic",measuredDate:"2026-08-05",reportJudgement:"OK",normalizedResult:"OK",targetElements:[],fileId:"FX_T01"},
    XRF_T02_01:{id:"XRF_T02_01",caseId:"CASE_T02_P1",attemptNo:1,role:"Periodic",measuredDate:"2026-08-05",reportJudgement:"----",normalizedResult:"OK",targetElements:[],fileId:"FX_T02"},
    XRF_T03_01:{id:"XRF_T03_01",caseId:"CASE_T03_P1",attemptNo:1,role:"Periodic",measuredDate:"2026-08-03",reportJudgement:"??",normalizedResult:"REMEASURE",targetElements:["Cd"],fileId:"FX_T03_1"},
    XRF_T03_02:{id:"XRF_T03_02",caseId:"CASE_T03_P1",attemptNo:2,role:"Retest",measuredDate:"2026-08-05",reportJudgement:"OK",normalizedResult:"OK",targetElements:[],fileId:"FX_T03_2",parentMeasurementId:"XRF_T03_01"},
    XRF_T04_01:{id:"XRF_T04_01",caseId:"CASE_T04_P1",attemptNo:1,role:"Periodic",measuredDate:"2026-08-03",reportJudgement:"??",normalizedResult:"REMEASURE",targetElements:["Cd"],fileId:"FX_T04_1"},
    XRF_T04_02:{id:"XRF_T04_02",caseId:"CASE_T04_P1",attemptNo:2,role:"Retest",measuredDate:"2026-08-05",reportJudgement:"??",normalizedResult:"REMEASURE",targetElements:["Cd"],fileId:"FX_T04_2",parentMeasurementId:"XRF_T04_01"},
    XRF_T05_01:{id:"XRF_T05_01",caseId:"CASE_T05_P1",attemptNo:1,role:"Periodic",measuredDate:"2026-08-05",reportJudgement:"NG",normalizedResult:"NG",targetElements:["Cd"],fileId:"FX_T05"},
    XRF_T06_01:{id:"XRF_T06_01",caseId:"CASE_T06_P1",attemptNo:1,role:"Periodic",measuredDate:"2026-08-05",reportJudgement:"NG",normalizedResult:"NG",targetElements:["Cd"],fileId:"FX_T06"},
    XRF_T07_01:{id:"XRF_T07_01",caseId:"CASE_T07_P1",attemptNo:1,role:"Periodic",measuredDate:"2026-08-05",reportJudgement:"NG",normalizedResult:"NG",targetElements:["Cd"],fileId:"FX_T07"},
    XRF_T08_01:{id:"XRF_T08_01",caseId:"CASE_T08_P1",attemptNo:1,role:"Periodic",measuredDate:"2026-08-05",reportJudgement:"NG",normalizedResult:"NG",targetElements:["Cd"],fileId:"FX_T08"}
  },
  precisionTransferAnalyses:{
    정밀분석_T04_01:{id:"정밀분석_T04_01",caseId:"CASE_T04_P1",triggerXrfMeasurementId:"XRF_T04_02",targetElements:["Cd"],requestStatus:"NOT_REQUESTED",uploadStatus:"WAITING",overallResult:null,evidenceComplete:false},
    정밀분석_T05_01:{id:"정밀분석_T05_01",caseId:"CASE_T05_P1",triggerXrfMeasurementId:"XRF_T05_01",targetElements:["Cd"],requestStatus:"REQUESTED",requestedAt:"2026-08-06",requestedBy:"admin",uploadStatus:"UPLOADED",resultFileId:"FE_T05",measuredDate:"2026-08-07",uploadedAt:"2026-08-07",uploadedBy:"admin",elements:{Cd:{ppm:65,legalLimit:100,internalLimit:70,result:"OK"}},overallResult:"OK",evidenceComplete:true},
    정밀분석_T06_01:{id:"정밀분석_T06_01",caseId:"CASE_T06_P1",triggerXrfMeasurementId:"XRF_T06_01",targetElements:["Cd"],requestStatus:"REQUESTED",requestedAt:"2026-08-06",requestedBy:"admin",uploadStatus:"UPLOADED",resultFileId:"FE_T06",measuredDate:"2026-08-07",uploadedAt:"2026-08-07",uploadedBy:"admin",elements:{Cd:{ppm:82,legalLimit:100,internalLimit:70,result:"NG"}},overallResult:"NG",evidenceComplete:true},
    정밀분석_T07_01:{id:"정밀분석_T07_01",caseId:"CASE_T07_P1",triggerXrfMeasurementId:"XRF_T07_01",targetElements:["Cd"],requestStatus:"REQUESTED",requestedAt:"2026-08-06",requestedBy:"admin",uploadStatus:"UPLOADED",resultFileId:"FE_T07",measuredDate:"2026-08-07",uploadedAt:"2026-08-07",uploadedBy:"admin",elements:{Cd:{ppm:82,legalLimit:100,internalLimit:70,result:"NG"}},overallResult:"NG",evidenceComplete:true},
    정밀분석_T08_01:{id:"정밀분석_T08_01",caseId:"CASE_T08_P1",triggerXrfMeasurementId:"XRF_T08_01",targetElements:["Cd"],requestStatus:"REQUESTED",requestedAt:"2026-08-06",requestedBy:"admin",uploadStatus:"UPLOADED",resultFileId:"FE_T08",measuredDate:"2026-08-07",uploadedAt:"2026-08-07",uploadedBy:"admin",elements:{Cd:{ppm:82,legalLimit:100,internalLimit:70,result:"NG"}},overallResult:"NG",evidenceComplete:true}
  },
  precisionAnalyses:{
    PREC_T06_01:{id:"PREC_T06_01",caseId:"CASE_T06_P1",required:true,requestStatus:"NOT_REQUESTED",uploadStatus:"WAITING",elements:{},result:null,evidenceComplete:false},
    // 정밀분석 결과는 원소 함유 여부/함유량과 법적 기준치로 자동 산출합니다.
    PREC_T07_01:{id:"PREC_T07_01",caseId:"CASE_T07_P1",required:true,requestStatus:"REQUESTED",requestedAt:"2026-08-08",labName:"KTR",reportNo:"KTR-MOCK-007",uploadStatus:"UPLOADED",resultFileId:"FP_T07",measuredDate:"2026-08-09",elements:{Cd:{presence:"함유",ppm:65,legalLimit:100}},result:"OK",evidenceComplete:true},
    PREC_T08_01:{id:"PREC_T08_01",caseId:"CASE_T08_P1",required:true,requestStatus:"REQUESTED",requestedAt:"2026-08-08",labName:"SGS",reportNo:"SGS-MOCK-008",uploadStatus:"UPLOADED",resultFileId:"FP_T08",measuredDate:"2026-08-09",elements:{Cd:{presence:"함유",ppm:120,legalLimit:100}},result:"NG",evidenceComplete:true}
  },
  followUps:{
    FU_T06_01:{id:"FU_T06_01",caseId:"CASE_T06_P1",boundarySample:{required:true,status:"PENDING"},customerPreNotice:{required:true,status:"PENDING"}},
    FU_T07_01:{id:"FU_T07_01",caseId:"CASE_T07_P1",boundarySample:{required:true,status:"COMPLETED",completedAt:"2026-08-08"},customerPreNotice:{required:true,status:"COMPLETED",completedAt:"2026-08-08",reference:"CS-MOCK-007"}},
    FU_T08_01:{id:"FU_T08_01",caseId:"CASE_T08_P1",boundarySample:{required:true,status:"COMPLETED",completedAt:"2026-08-08"},customerPreNotice:{required:true,status:"COMPLETED",completedAt:"2026-08-08",reference:"CS-MOCK-008"}}
  },
  approvals:{
    APR_T01:{status:"승인",basis:"XRF_OK",basisRecordId:"XRF_T01_01"},
    APR_T02:{status:"승인",basis:"XRF_OK",basisRecordId:"XRF_T02_01"},
    APR_T03:{status:"승인",basis:"XRF_REMEASURE_OK",basisRecordId:"XRF_T03_02"},
    APR_T04:{status:"측정 진행중",basis:null,basisRecordId:null},
    APR_T05:{status:"승인",basis:"PRECISION_TRANSFER_OK",basisRecordId:"정밀분석_T05_01"},
    APR_T06:{status:"측정 진행중",basis:null,basisRecordId:null},
    APR_T07:{status:"승인",basis:"PRECISION_OK",basisRecordId:"PREC_T07_01"},
    APR_T08:{status:"반려",basis:"PRECISION_NG_REJECTED",basisRecordId:"PREC_T08_01"},
    APR_T09:{status:"측정 진행중",basis:null,basisRecordId:null},
    APR_T10:{status:"측정 진행중",basis:null,basisRecordId:null}
  }
};

function workflowCaseForItem(item){
  if(!item?.code) return null;
  const rows=Object.values(MOCK_WORKFLOW_DB.analysisCases).filter(c=>c.itemCode===item.code);
  return rows.sort((a,b)=>Number(b.periodNo||0)-Number(a.periodNo||0))[0]||null;
}
function workflowXrfAttempts(caseData){
  return (caseData?.xrfAttemptIds||[]).map(id=>MOCK_WORKFLOW_DB.xrfMeasurements[id]).filter(Boolean).sort((a,b)=>Number(a.attemptNo||0)-Number(b.attemptNo||0));
}
function workflowPrecision(caseData){ return caseData?.precisionId ? MOCK_WORKFLOW_DB.precisionAnalyses[caseData.precisionId]||null : null; }
function workflowPrecisionTransfer(caseData){ return caseData?.precisionTransferId ? MOCK_WORKFLOW_DB.precisionTransferAnalyses[caseData.precisionTransferId]||null : null; }
function workflowFollowUp(caseData){ return caseData?.followUpId ? MOCK_WORKFLOW_DB.followUps[caseData.followUpId]||null : null; }
function workflowApproval(caseData){ return caseData?.approvalId ? MOCK_WORKFLOW_DB.approvals[caseData.approvalId]||null : null; }
function workflowStageLabel(stage){
  const map={XRF_REQUIRED:"확인 필요",XRF_REMEASURE_REQUIRED:"XRF 재측정",PRECISION_TRANSFER_REQUEST_REQUIRED:"정밀분석 인계 필요",PRECISION_TRANSFER_RESULT_WAITING:"정밀분석 결과 대기",PRECISION_FOLLOWUP_REQUIRED:"후속조치 필요",PRECISION_REQUEST_REQUIRED:"정밀분석 의뢰 필요",PRECISION_RESULT_WAITING:"정밀분석 결과 대기",APPROVED:"처리 완료",PRECISION_NG_REVIEW:"정밀분석 NG · 반려",RISK_NOT_ALLOWED:"사용불허"};
  return map[stage]||stage||"—";
}
function baseWorkflowForItem(item){
  const caseData=workflowCaseForItem(item);
  if(caseData){
    const xrfAttempts=workflowXrfAttempts(caseData);
    return {
      caseData,xrfAttempts,latestXrf:xrfAttempts[xrfAttempts.length-1]||null,
      precisionTransfer:workflowPrecisionTransfer(caseData),precision:workflowPrecision(caseData),
      followUp:workflowFollowUp(caseData),approval:workflowApproval(caseData),stage:caseData.stage
    };
  }
  const latest=item?latestMeasurementOf(item):null;
  const xrfResult=latest?measurementXrfResult(latest):"—";
  const precisionTriggers=item?precisionTriggerMeasurementsOf(item):[];
  const remeasureTargets=pendingXrfRemeasureElements(latest);
  let stage="XRF_REQUIRED";
  if(remeasureTargets.length) stage="XRF_REMEASURE_REQUIRED";
  else if(precisionTriggers.length) stage="PRECISION_REQUEST_REQUIRED";
  else if(xrfResult==="OK") stage="APPROVED";
  else if(xrfResult==="재측정 필요") stage="XRF_REMEASURE_REQUIRED";
  else if(xrfResult==="NG") stage="PRECISION_REQUEST_REQUIRED";
  return {
    caseData:null,xrfAttempts:[],latestXrf:null,
    latestMeasurement:latest,
    precisionTriggerMeasurements:precisionTriggers,
    precisionTransfer:null,precision:null,followUp:null,approval:null,stage
  };
}
function workflowForItem(item){
  return item?.__workflow || baseWorkflowForItem(item);
}
const ANALYSIS_WORKFLOW_GROUP_STAGES=Object.freeze({
  xrfRemeasure:["XRF_REMEASURE_REQUIRED"],
  precisionTransferRequired:["PRECISION_TRANSFER_REQUEST_REQUIRED","PRECISION_REQUEST_REQUIRED"],
  precisionResultWaiting:["PRECISION_TRANSFER_RESULT_WAITING","PRECISION_RESULT_WAITING"],
  precisionFollowupRequired:["PRECISION_FOLLOWUP_REQUIRED"],
  precisionNgReview:["PRECISION_NG_REVIEW"]
});
function analysisWorkflowSummary(items=[]){
  const summary={
    xrfRemeasure:0,
    precisionTransferRequired:0,
    precisionResultWaiting:0,
    precisionFollowupRequired:0,
    precisionNgReview:0,
    total:0
  };
  (items||[]).forEach(item=>{
    const stage=workflowForItem(item)?.stage;
    for(const [key,stages] of Object.entries(ANALYSIS_WORKFLOW_GROUP_STAGES)){
      if(stages.includes(stage)){
        summary[key]++;
        summary.total++;
        break;
      }
    }
  });
  return summary;
}
function workflowOverrideKey(item,wf){
  return wf?.caseData?.id || latestMeasurementOf(item)?.id || item?.latestMeasurementId || item?.code || "";
}
function precisionTargetsForWorkflow(item,wf){
  const trigger=wf?.triggerMeasurement || wf?.latestXrf || latestMeasurementOf(item);
  return [...new Set([
    ...(wf?.precision?.targetElements||[]),
    ...(wf?.precisionTransfer?.targetElements||[]),
    ...(trigger?.targetElements||[]),
    ...(precisionRequiredElements(trigger)||[])
  ].filter(Boolean))];
}
function precisionLegalLimit(el,precision){
  const fromRecord=xrfNumber(precision?.elements?.[el]?.legalLimit);
  if(fromRecord!=null) return fromRecord;
  return xrfNumber(XRF_ELEMENT_LIMITS[el]);
}
function precisionElementState(precision,elementResults,el){
  const override=elementResults?.[el]||{};
  const base=precision?.elements?.[el]||{};
  const rawPresence=String(override.presence??base.presence??"").trim();
  const inferredPresence=rawPresence || (isXrfContentND(base.ppm)?"미함유":(base.ppm!==undefined&&base.ppm!==null&&String(base.ppm).trim()!==""?"함유":""));
  const ppm=override.ppm!==undefined ? override.ppm : (base.ppm!==undefined ? base.ppm : "");
  return {presence:inferredPresence,ppm,legalLimit:precisionLegalLimit(el,precision)};
}
function precisionElementDecision(state){
  if(state?.presence==="미함유") return "OK";
  if(state?.presence!=="함유") return "";
  const ppm=xrfNumber(state?.ppm);
  const limit=xrfNumber(state?.legalLimit);
  if(ppm==null || limit==null) return "";
  return ppm<=limit ? "OK" : "NG";
}
function derivePrecisionResult(targets,precision,elementResults){
  if(!targets?.length) return "";
  const decisions=targets.map(el=>precisionElementDecision(precisionElementState(precision,elementResults,el)));
  if(decisions.some(v=>!v)) return "";
  return decisions.some(v=>v==="NG") ? "NG" : "OK";
}
function deriveRuntimeWorkflow(item,precisionOverrides,approvalOverrides){
  const base=baseWorkflowForItem(item);

  // 실제 품목은 "최신 XRF"를 현재 승인/정밀분석 Workflow의 기준으로 사용합니다.
  // R Final Risk H는 관리상 고위험이지만 최신 XRF가 OK이면 승인할 수 있습니다.
  // 최신 XRF가 NG(또는 2차 재측정에서도 ??)이면 해당 최신 측정 건의 정밀분석 결과가 확정되기 전에는 승인할 수 없습니다.
  if(!base.caseData && !isDiscontinueRequestItem(item)){
    const precisionCases=runtimePrecisionCasesForItem(item,precisionOverrides);
    const latest=latestMeasurementOf(item);
    const latestXrfState=latest ? measurementXrfResult(latest) : "—";
    const latestRemeasureTargets=pendingXrfRemeasureElements(latest);
    const latestNeedsRemeasure=latestRemeasureTargets.length>0;
    const latestCase=precisionCases[0] || null;
    const latestNeedsPrecision=!!latestCase;
    const finalRisk=itemFinalRisk(item);
    let stage="XRF_REQUIRED";
    let approval=null;
    let precision=latestCase?.precision||null;

    if(latestNeedsPrecision){
      const p=latestCase.precision;
      if(p?.result==="NG"){
        stage="PRECISION_NG_REVIEW";
        approval={status:"반려",basis:"PRECISION_NG_REJECTED",basisRecordId:latestCase.key};
      }else if(latestNeedsRemeasure){
        // NG 정밀분석과 ?? 재측정은 서로 독립 의무입니다. 정밀분석 진행 여부와 관계없이
        // 1차 ??의 Retest가 완료되기 전에는 승인하지 않습니다.
        stage="XRF_REMEASURE_REQUIRED";
        approval={status:"보류",basis:"XRF_REMEASURE_AND_PRECISION_REQUIRED",basisRecordId:latestCase.key};
      }else if(p?.result==="OK"){
        // 최신 XRF NG라도 정밀분석 결과가 OK로 확정되면 승인 가능합니다.
        // 단, R Final Risk가 사용불허인 품목은 위험도 기준상 승인하지 않습니다.
        if(finalRisk==="Not Allowed"){
          stage="RISK_NOT_ALLOWED";
          approval={status:"반려",basis:"RISK_NOT_ALLOWED",basisRecordId:latestCase.key};
        }else{
          stage="APPROVED";
          approval={status:"승인",basis:"PRECISION_OK",basisRecordId:latestCase.key};
        }
      }else if(p?.requestStatus==="REQUESTED"){
        stage="PRECISION_RESULT_WAITING";
        approval={status:"측정 진행중",basis:"PRECISION_REQUESTED",basisRecordId:latestCase.key};
      }else{
        stage="PRECISION_REQUEST_REQUIRED";
        approval={status:"의뢰 필요",basis:"PRECISION_REQUEST_REQUIRED",basisRecordId:latestCase.key};
      }
    }else if(latestNeedsRemeasure){
      stage="XRF_REMEASURE_REQUIRED";
      approval={status:"보류",basis:"XRF_REMEASURE_REQUIRED",basisRecordId:latest?.id||item?.code};
    }else if(latestXrfState==="OK"){
      if(finalRisk==="Not Allowed"){
        stage="RISK_NOT_ALLOWED";
        approval={status:"반려",basis:"RISK_NOT_ALLOWED",basisRecordId:latest?.id||item?.code};
      }else{
        stage="APPROVED";
        approval={status:"승인",basis:"XRF_OK",basisRecordId:latest?.id||item?.code};
      }
    }else if(latestXrfState==="NG"){
      // 방어 코드: 최신 NG는 위 latestCase로 반드시 생성되어야 하며, 결과 입력 전 승인되지 않습니다.
      stage="PRECISION_REQUEST_REQUIRED";
      approval={status:"의뢰 필요",basis:"PRECISION_REQUEST_REQUIRED",basisRecordId:latest?.id||item?.code};
    }else if(latestXrfState==="재측정 필요"){
      stage="XRF_REMEASURE_REQUIRED";
      approval={status:"보류",basis:"XRF_REMEASURE_REQUIRED",basisRecordId:latest?.id||item?.code};
    }else{
      stage="XRF_REQUIRED";
      approval={status:"보류",basis:"XRF_REVIEW_REQUIRED",basisRecordId:latest?.id||item?.code};
    }

    const manualApproval=String(approvalOverrides?.[item?.code]||"").trim();
    // XRF 결과 자체가 없는 경우에만 수동 승인값을 보조적으로 허용합니다.
    // 최신 XRF NG/정밀분석 대기/사용불허 상태는 수동값으로 승인 우회할 수 없습니다.
    if(manualApproval && latestXrfState==="—") approval={...(approval||{}),status:manualApproval};

    return {
      ...base,
      precision,
      precisionCases,
      activePrecisionCaseKey:latestCase?.key||null,
      triggerMeasurement:latestCase?.measurement||latest||null,
      approval,
      stage
    };
  }

  // 기존 Workflow mock / 단종 요청은 기존 상태머신을 유지합니다.
  const key=workflowOverrideKey(item,base);
  const ov=precisionOverrides?.[key]||{};
  const latestXrfRaw=base.latestXrf?.normalizedResult || itemXrfWorst(item);
  const xrfState=latestXrfRaw==="REMEASURE" ? "재측정 필요"
    : latestXrfRaw==="REVIEW_REQUIRED" ? "확인 필요"
    : (latestXrfRaw||"—");

  const targets=precisionTargetsForWorkflow(item,base);
  const xrfNeedsPrecision=xrfState==="NG";
  const precisionSeed=base.precision || ((targets.length || xrfNeedsPrecision || String(base.stage||"").startsWith("PRECISION")) ? {required:true,requestStatus:"NOT_REQUESTED",uploadStatus:"WAITING"} : null);
  let precision=precisionSeed ? {...precisionSeed,...ov} : null;
  if(precision){
    const elementResults=ov.elementResults || precision.elementResults || {};
    const derived=derivePrecisionResult(targets,precision,elementResults);
    precision={...precision,targetElements:targets,elementResults,result:derived || precision.result || ""};
  }

  const precisionRequested=precision?.requestStatus==="REQUESTED";
  let stage=base.stage;
  let approval=base.approval ? {...base.approval} : null;
  const manualApproval=String(approvalOverrides?.[item?.code]||"").trim();

  if(isDiscontinueRequestItem(item)){
    if(manualApproval) approval={...(approval||{}),status:manualApproval};
    return {...base,precision,approval,stage};
  }

  if(xrfState==="OK"){
    stage="APPROVED";
    approval={...(approval||{}),status:"승인",basis:"XRF_OK",basisRecordId:base.latestXrf?.id||latestMeasurementOf(item)?.id||key};
  }else if(xrfState==="NG"){
    if(precision?.result==="OK"){
      stage="APPROVED";
      approval={...(approval||{}),status:"승인",basis:"PRECISION_OK",basisRecordId:precision.id||key};
    }else if(precision?.result==="NG"){
      stage="PRECISION_NG_REVIEW";
      approval={...(approval||{}),status:"반려",basis:"PRECISION_NG_REJECTED",basisRecordId:precision.id||key};
    }else if(precisionRequested){
      stage="PRECISION_RESULT_WAITING";
      approval={...(approval||{}),status:"측정 진행중",basis:"PRECISION_REQUESTED",basisRecordId:precision?.id||key};
    }else{
      stage="PRECISION_REQUEST_REQUIRED";
      approval={...(approval||{}),status:"의뢰 필요",basis:"PRECISION_REQUEST_REQUIRED",basisRecordId:precision?.id||key};
    }
  }else if(xrfState==="재측정 필요" || xrfState==="확인 필요" || xrfState==="—"){
    stage=xrfState==="재측정 필요" ? "XRF_REMEASURE_REQUIRED" : "XRF_REQUIRED";
    approval={...(approval||{}),status:"보류",basis:xrfState==="재측정 필요"?"XRF_REMEASURE_REQUIRED":"XRF_REVIEW_REQUIRED",basisRecordId:base.latestXrf?.id||latestMeasurementOf(item)?.id||key};
  }else if(manualApproval){
    approval={...(approval||{}),status:manualApproval};
  }

  return {...base,precision,approval,stage};
}
function workflowFollowUpProgress(item){
  const f=workflowForItem(item).followUp;
  if(!f) return null;
  const tasks=[f.boundarySample,f.customerPreNotice].filter(t=>t?.required);
  const done=tasks.filter(t=>String(t?.status||"").toUpperCase()==="COMPLETED").length;
  return {done,total:tasks.length};
}

const MOCK_WORKFLOW_ITEMS = [
  makeMockWorkflowItem("T-01","Mock · XRF OK",{xrf:"OK",approvalStatus:"승인"}),
  makeMockWorkflowItem("T-02","Mock · Judgment dash ND",{xrf:"OK",approvalStatus:"승인",history:[{date:"2026-08-05",id:"T-02_20260805_01",periodNo:1,label:"P1",measurementRole:"Periodic",attemptNo:1,xrfWorst:"OK",xrfResult:"OK",approvalStatus:"승인",elements:mockElementSet("DASH_ND"),status:"measured"}]}),
  makeMockWorkflowItem("T-03","Mock · ?? → 재측정 OK",{xrf:"OK",approvalStatus:"승인",history:[{date:"2026-08-03",id:"T-03_20260803_01",periodNo:1,label:"P1",measurementRole:"Periodic",attemptNo:1,xrfWorst:"재측정 필요",xrfResult:"재측정 필요",approvalStatus:"측정 진행중",elements:mockElementSet("REMEASURE"),status:"measured"},{date:"2026-08-05",id:"T-03_20260805_01",periodNo:1,label:"P1",measurementRole:"Retest",attemptNo:2,xrfWorst:"OK",xrfResult:"OK",approvalStatus:"승인",elements:mockElementSet("OK"),status:"measured"}]}),
  makeMockWorkflowItem("T-04","Mock · 재측정 ?? → 정밀분석",{xrf:"재측정 필요",approvalStatus:"측정 진행중",history:[{date:"2026-08-03",id:"T-04_20260803_01",periodNo:1,label:"P1",measurementRole:"Periodic",attemptNo:1,xrfWorst:"재측정 필요",xrfResult:"재측정 필요",approvalStatus:"측정 진행중",elements:mockElementSet("REMEASURE"),status:"measured"},{date:"2026-08-05",id:"T-04_20260805_01",periodNo:1,label:"P1",measurementRole:"Retest",attemptNo:2,xrfWorst:"재측정 필요",xrfResult:"재측정 필요",approvalStatus:"측정 진행중",elements:mockElementSet("REMEASURE"),status:"measured"}]}),
  makeMockWorkflowItem("T-05","Mock · XRF NG → 정밀분석 OK",{xrf:"NG",approvalStatus:"승인"}),
  makeMockWorkflowItem("T-06","Mock · 정밀분석 NG → 후속조치",{xrf:"NG",approvalStatus:"측정 진행중",finalRisk:"H"}),
  makeMockWorkflowItem("T-07","Mock · 정밀분석 OK",{xrf:"NG",approvalStatus:"승인",finalRisk:"H"}),
  makeMockWorkflowItem("T-08","Mock · 정밀분석 NG",{xrf:"NG",approvalStatus:"반려",finalRisk:"H"}),
  makeMockWorkflowItem("T-09","Mock · D-30 측정 필요",{xrf:"—",approvalStatus:"측정 진행중",due:"2026-09-02",measured:null}),
  makeMockWorkflowItem("T-10","Mock · 기한 초과",{xrf:"—",approvalStatus:"측정 진행중",due:"2026-08-01",measured:null})
];

function daysUntil(s){ if(!s) return null; return Math.round((new Date(s)-TODAY)/86400000); }
function normStatus(s){
  const raw=String(s||"").trim();
  const key=raw.toLowerCase().replace(/[\s_-]/g,"");
  const map={
    duesoon:"due_soon", due:"due_soon", due_soon:"due_soon",
    "측정권장":"due_soon",
    late:"late", delaypending:"late", measuredlate:"late", "지연측정":"late",
    overdue:"overdue", missed:"overdue", notdone:"overdue", noncompliant:"overdue", "미이행":"overdue",
    early:"early", "조기측정":"early",
    compliant:"compliant", measured:"measured", "이행":"compliant", "측정완료":"measured",
    upcoming:"upcoming", open:"upcoming", "예정":"upcoming",
    reference:"reference", registered:"reference", "기준":"reference"
  };
  return map[raw] || map[key] || raw || "upcoming";
}

function elementXrfResult(d){
  if(!d) return "확인 필요";
  if(isXrfRemeasureRequired(d)) return "재측정 필요";
  const judgement=xrfJudgementDecision(d);
  if(judgement==="NG") return "NG";
  if(judgement==="OK") return "OK";
  // Content/Std.Deviation 값이 존재해도 Judgment가 비어 있거나 알 수 없는 값이면 자동 OK 처리하지 않습니다.
  return "확인 필요";
}
function retestRequiredElements(m){
  if(!m?.elements) return [];
  return XRF_REPORT_ELEMENTS.filter(el=>isXrfRemeasureRequired(m.elements?.[el]));
}
function pendingXrfRemeasureElements(m){
  const targets=retestRequiredElements(m);
  // 1차 측정의 ??만 XRF 재측정 의무입니다. Retest에서도 남은 ??는
  // precisionRequiredElements()가 정밀분석 대상으로 전환하므로 재측정을 다시 요구하지 않습니다.
  const isRetest=Number(m?.attemptNo||0)>=2 || String(m?.role||m?.measurementRole||"").toLowerCase()==="retest";
  return isRetest ? [] : targets;
}
function measurementFollowupInfo(m){
  const precisionElements=precisionRequiredElements(m);
  const remeasureElements=pendingXrfRemeasureElements(m);
  const flaggedElements=[...new Set([...precisionElements,...remeasureElements])];
  const label=precisionElements.length && remeasureElements.length
    ? "정밀분석 및 XRF 재측정 필요"
    : precisionElements.length
      ? "정밀분석 필요"
      : remeasureElements.length
        ? "XRF 재측정"
        : "후속조치 없음";
  const detail=[
    precisionElements.length?`정밀분석: ${precisionElements.join(", ")}`:"",
    remeasureElements.length?`XRF 재측정: ${remeasureElements.join(", ")}`:""
  ].filter(Boolean).join(" · ") || "추가 조치 없음";
  return {precisionElements,remeasureElements,flaggedElements,label,detail,hasFollowup:flaggedElements.length>0};
}
function measurementFollowupVisibleDetail(info){
  // 정밀분석 대상 원소는 바로 위 원소별 상태에서 이미 강조되므로 같은 목록을 문장으로 반복하지 않습니다.
  // 재측정 의무가 함께 있는 경우에만 사용자가 처리 대상을 놓치지 않도록 별도 문구를 유지합니다.
  return info?.remeasureElements?.length ? `XRF 재측정: ${info.remeasureElements.join(", ")}` : "";
}
function retestShortText(list){
  return list?.length ? `${list.join(", ")} 필요` : "필요 없음";
}
function retestDetailText(list){
  return list?.length
    ? `XRF 재측정 필요 원소: ${list.join(", ")} · ??`
    : "XRF 재측정 필요 원소 없음";
}
function itemRetestRequired(item){
  const m=latestMeasurementOf(item);
  return retestRequiredElements(m).length>0;
}
function measurementXrfLevel(m){
  if(!m?.elements) return null;
  const rows=XRF_REPORT_ELEMENTS.map(el=>m.elements?.[el]);
  // 모든 평가 원소는 Content(ppm)를 Legal Limit의 70%인 Internal Limit과 비교합니다.
  // ??/Cannot Judge 또는 Content 결측이 있으면 위험도 산정을 보류합니다.
  if(rows.some(d=>!d || isXrfRemeasureRequired(d) || !xrfJudgementDecision(d))) return null;
  const levels=rows.map(d=>xrfLevelFromJudgement(d));
  if(levels.some(v=>!v)) return null;
  if(levels.includes("H")) return "H";
  if(levels.includes("M")) return "M";
  if(levels.includes("L")) return "L";
  return null;
}
function measurementXrfResult(m){
  if(!m?.elements) return "—";
  const rows=XRF_REPORT_ELEMENTS.map(el=>m.elements?.[el]);
  if(rows.some(d=>d && xrfJudgementDecision(d)==="NG")) return "NG";
  if(rows.some(d=>d && isXrfRemeasureRequired(d))) return "재측정 필요";
  // Content 결측 또는 사내기준을 계산할 수 없는 원소가 있으면 확인 필요입니다.
  // 원본 Judgment는 재측정 신호(??/Cannot Judge)와 추적용 원본값으로 유지합니다.
  if(rows.some(d=>!d || !xrfJudgementDecision(d))) return "확인 필요";
  return "OK";
}
function judgementDisplayValue(d){
  if(isXrfRemeasureRequired(d)) return "??";
  // Content의 연속 dash(2개 이상)는 N.D.로 보지만 단일 '-'는 결측으로 유지합니다.
  if(isXrfContentDashND(sourcePpmValue(d))) return "OK";
  const decision=xrfJudgementDecision(d);
  if(decision==="OK") return "OK";
  if(decision) return decision;
  const raw=sourceJudgementValue(d);
  return raw==null || String(raw).trim()==="" ? "확인 필요" : String(raw).trim();
}
function measurementJudgementSummary(m){
  if(!m?.elements) return "—";
  const rows=XRF_REPORT_ELEMENTS.map(el=>m.elements?.[el]);
  if(rows.some(d=>d && xrfJudgementDecision(d)==="NG")) return "NG";
  if(rows.some(d=>d && isXrfRemeasureRequired(d))) return "??";
  if(rows.some(d=>!d || !xrfJudgementDecision(d))) return "확인 필요";
  return "OK";
}
function xrfExceededElements(m){
  if(!m?.elements) return [];
  return XRF_REPORT_ELEMENTS.filter(el=>xrfJudgementDecision(m.elements?.[el])==="NG");
}
function precisionRequiredElements(m){
  const targets=[...xrfExceededElements(m)];
  // 1차 ??는 XRF 재측정으로 끝내고, 2차(Retest)에서도 ??이면 해당 원소를 정밀분석 대상으로 전환합니다.
  const isRetest=Number(m?.attemptNo||0)>=2 || String(m?.role||"").toLowerCase()==="retest";
  if(isRetest) targets.push(...retestRequiredElements(m));
  return [...new Set(targets)];
}

// XRF/정밀분석 Workflow는 더 이상 "최신 측정 1건"만 보지 않습니다.
// R/Pn/Retest 전체 측정 이력을 현재 Internal Limit 70% 정책으로 다시 판정한 뒤,
// 실제 정밀분석 대상이 된 측정 ID를 Case의 기준키로 사용합니다.
function itemMeasurementRecords(item){
  if(!item) return [];
  const rows=[];
  const seen=new Set();
  const addMeasurement=(measurement,idHint="")=>{
    if(!measurement) return;
    const id=String(measurement?.id||measurement?.measurementId||idHint||"").trim();
    const date=dateOnly(measurement?.date||measurement?.measuredDate||"")||"";
    const key=id || `${date}|${normalizeMeasurementRole(measurement?.role||measurement?.measurementRole||"")}|${Number(measurement?.attemptNo||1)}`;
    if(!key || seen.has(key)) return;
    seen.add(key);
    rows.push({...measurement,id:id||measurement?.id||null,date:date||measurement?.date||null});
  };
  const addId=id=>{ if(id){ const m=measurementById(item,id); if(m) addMeasurement(m,id); } };

  addId(item?.firstMeasurementId);
  (item?.periods||[]).forEach(period=>periodMeasurementIds(period).forEach(addId));
  (item?.history||[]).forEach(row=>{
    const id=row?.id||row?.measurementId;
    const m=id?measurementById(item,id):null;
    if(m) addMeasurement(m,id);
    else if(row?.elements) addMeasurement(row,id||"");
  });
  addId(item?.latestMeasurementId);
  if(item?.latest) addMeasurement(item.latest,item?.latestMeasurementId||"");

  return rows.sort((a,b)=>{
    const byDate=String(a?.date||"").localeCompare(String(b?.date||""));
    if(byDate) return byDate;
    return Number(a?.attemptNo||1)-Number(b?.attemptNo||1);
  });
}
function precisionTriggerMeasurementsOf(item){
  // 현재 승인/정밀분석 필요 여부는 가장 최근 XRF 측정 결과를 기준으로 판단합니다.
  // R Final Risk가 H/Not Allowed여도 최신 XRF가 OK이면 정밀분석 미입력을 이유로 승인을 막지 않습니다.
  // 반대로 최신 XRF가 NG이거나 2차 재측정에서도 ??이면 해당 최신 측정 건의 정밀분석 결과가 반드시 필요합니다.
  const latest=latestMeasurementOf(item);
  if(!latest) return [];
  return precisionRequiredElements(latest).length>0 ? [latest] : [];
}
function itemHasPrecisionTrigger(item){
  return precisionTriggerMeasurementsOf(item).length>0;
}
function precisionCaseKeyForMeasurement(item,measurement){
  return String(measurement?.id||measurement?.measurementId||`${item?.code||"ITEM"}::${dateOnly(measurement?.date)||"UNKNOWN"}`).trim();
}
function persistedPrecisionResultOf(measurement,item){
  // 정밀분석 결과는 해당 XRF Measurement에 귀속된 값만 우선 사용합니다.
  // 최신 XRF NG가 과거 품목 단위 정밀분석 결과를 재사용해 자동 승인되는 것을 방지합니다.
  const measurementRaw=normalizeAnalysisResult(firstResultValue(
    measurement?.precisionAnalysisResult,measurement?.precisionResult,measurement?.detailedAnalysisResult,measurement?.detailedResult,measurement?.Precision_Analysis_Result
  ));
  if(measurementRaw==="OK" || measurementRaw==="NG") return measurementRaw;

  // 구형 단일 측정 데이터만 품목 단위 저장값을 호환용으로 허용합니다.
  const records=itemMeasurementRecords(item);
  if(records.length<=1){
    const legacyRaw=normalizeAnalysisResult(firstResultValue(
      item?.precisionAnalysisResult,item?.precisionResult,item?.detailedAnalysisResult,item?.detailedResult,item?.Precision_Analysis_Result
    ));
    if(legacyRaw==="OK" || legacyRaw==="NG") return legacyRaw;
  }
  return "";
}
function runtimePrecisionCasesForItem(item,precisionOverrides){
  return precisionTriggerMeasurementsOf(item).map(measurement=>{
    const key=precisionCaseKeyForMeasurement(item,measurement);
    const targets=precisionRequiredElements(measurement);
    const ov=precisionOverrides?.[key]||{};
    const persistedResult=persistedPrecisionResultOf(measurement,item);
    const elementResults=ov.elementResults||{};
    const precision={
      required:true,
      targetElements:targets,
      requestStatus:ov.requestStatus || (persistedResult?"REQUESTED":"NOT_REQUESTED"),
      uploadStatus:ov.uploadStatus || (ov.resultFileId||persistedResult?"UPLOADED":"WAITING"),
      resultFileId:ov.resultFileId||"",
      requestedAt:ov.requestedAt||"",
      requestedBy:ov.requestedBy||"",
      elementResults,
      finalConfirm:ov.finalConfirm || (persistedResult?"확인":"")
    };
    const derived=derivePrecisionResult(targets,precision,elementResults);
    const confirmed=precision.finalConfirm==="확인" || !!persistedResult;
    precision.result=confirmed ? (derived||persistedResult||"") : "";
    return {
      key,item,measurement,targets,precision,
      triggerLabel:precisionTriggerReason(measurement),
      sourceFileName:ov.sourceFileName||measurement?.sourceFileName||"",
      finalConfirm:precision.finalConfirm||"",
      elementResults
    };
  });
}
function precisionTriggerReason(m){
  const targets=precisionRequiredElements(m);
  if(!targets.length) return "대상 아님";
  const ng=xrfExceededElements(m);
  const retest=retestRequiredElements(m);
  if(ng.length) return "XRF NG 판정";
  if(retest.length) return "재측정 후 ?? 판정";
  return "정밀분석 필요";
}
function riskFromMatrix(crLevel, xrfLevel){
  const MX={
    H:{H:"Not Allowed",M:"Not Allowed",L:"Not Allowed"},
    M:{H:"H",M:"M",L:"M"},
    L:{H:"M",M:"M",L:"L"}
  };
  return MX[xrfLevel]?.[crLevel] || null;
}
function measurementFinalRisk(m, item){
  // Final Risk의 단일 기준은 최초 R 원본 측정입니다.
  // 과거 DB의 finalRisk / riskBasisXrfLevel / 개별 Period finalRisk는 현재 Internal Limit 70% 정책과
  // 불일치할 수 있으므로 R 원본이 있는 품목에서는 어떤 저장 Risk 값도 판정 근거로 사용하지 않습니다.
  const basis=riskBasisMeasurementOf(item);
  if(!basis) return null;
  const xrfLevel=measurementXrfLevel(basis);
  const crLevel=riskBasisCrLevelOf(item);
  if(!xrfLevel || !crLevel) return null;
  return riskFromMatrix(crLevel, xrfLevel);
}
function latestMeasurementOf(item){
  const measurementMap=item?.__measurementMap||{};
  return item?.latestMeasurementId ? (measurementMap[item.latestMeasurementId] || item.latest) : item?.latest;
}
function measurementById(item,id){
  if(!id) return null;
  const measurementMap=item?.__measurementMap||{};
  if(measurementMap[id]) return measurementMap[id];
  if(item?.latest?.id===id) return item.latest;
  const historyRow=(item?.history||[]).find(h=>(h.id||h.measurementId)===id);
  if(historyRow){
    return {
      id,
      date:historyRow.date,
      periodNo:historyRow.periodNo,
      role:normalizeMeasurementRole(historyRow.measurementRole||historyRow.role||"Periodic"),
      attemptNo:Number(historyRow.attemptNo||1),
      xrfWorst:historyRow.xrf_worst||historyRow.xrfWorst||item?.xrfWorst,
      xrfResult:historyRow.xrfResult||historyRow.xrf_worst||historyRow.xrfWorst||item?.xrfWorst,
      finalRisk:historyRow.finalRisk||item?.finalRisk,
      approvalStatus:historyRow.approvalStatus||item?.approvalStatus,
      elements:historyRow.elements||{},
      retestRequired:!!historyRow.retestRequired,
      retestRequiredElements:historyRow.retestRequiredElements||""
    };
  }
  return null;
}
function itemXrfWorst(item){
  if(isDiscontinueRequestItem(item)) return "—";
  return itemXrfResult(item);
}
function itemXrfResult(item){
  if(isDiscontinueRequestItem(item)) return "—";
  return measurementXrfResult(latestMeasurementOf(item)) || item?.latestXrfResult || item?.xrfWorst || "—";
}
function itemLatestXrfLevel(item){
  const latest=latestMeasurementOf(item);
  return latest ? measurementXrfLevel(latest) : null;
}
function riskBasisMeasurementOf(item){
  if(!item) return null;
  // R 측정 ID를 가장 명시적인 소스부터 찾습니다. 저장 Risk 값은 사용하지 않습니다.
  const referencePeriod=(item?.periods||[]).find(p=>p?.isReference || Number(p?.num)===0 || String(p?.label||p?.displayLabel||"").trim()==="R");
  const referenceId=referencePeriod?.measurementId||null;
  const explicitRiskId=item?.riskBasisMeasurementId||null;
  for(const id of [referenceId,explicitRiskId].filter(Boolean)){
    const measurement=measurementById(item,id);
    if(measurement?.elements && Object.keys(measurement.elements).length) return measurement;
  }
  // firstMeasurementId는 실제 Initial/R 측정일 때만 사용합니다. P1을 R로 오인하지 않게 역할을 확인합니다.
  if(item?.firstMeasurementId){
    const first=measurementById(item,item.firstMeasurementId);
    if(normalizeMeasurementRole(first?.role||first?.measurementRole)==="Initial" && first?.elements && Object.keys(first.elements).length) return first;
  }
  // 구형 데이터에서 R ID 필드가 비어 있는 경우 Initial/periodNo=0 이력을 마지막 복구 경로로 사용합니다.
  const initialHistory=[...(item?.history||[])]
    .filter(h=>Number(h?.periodNo)===0 || normalizeMeasurementRole(h?.measurementRole||h?.role)==="Initial")
    .sort((a,b)=>String(a?.date||"").localeCompare(String(b?.date||"")));
  for(const row of initialHistory){
    const measurement=measurementById(item,row?.id||row?.measurementId);
    if(measurement?.elements && Object.keys(measurement.elements).length) return measurement;
    if(row?.elements && Object.keys(row.elements).length) return {...row,role:"Initial"};
  }
  const latest=latestMeasurementOf(item);
  if(normalizeMeasurementRole(latest?.role||latest?.measurementRole)==="Initial" && latest?.elements && Object.keys(latest.elements).length) return latest;
  return null;
}
function riskBasisXrfLevelOf(item){
  if(!item) return null;
  const basis=riskBasisMeasurementOf(item);
  // R 원본을 현재 Internal Limit 70% 정책으로 매번 재판정합니다.
  // 과거 저장 riskBasisXrfLevel은 더 이상 fallback으로 사용하지 않습니다.
  return basis ? measurementXrfLevel(basis) : null;
}
function riskBasisCrLevelOf(item){
  const saved=String(item?.riskBasisCrLevel||"").trim();
  if(["H","M","L"].includes(saved)) return saved;
  const current=String(item?.crLevel||"").trim();
  return ["H","M","L"].includes(current) ? current : null;
}
function itemFinalRisk(item){
  if(isDiscontinueRequestItem(item)) return "—";
  const basis=riskBasisMeasurementOf(item);
  if(basis){
    const xrfLevel=measurementXrfLevel(basis);
    const crLevel=riskBasisCrLevelOf(item);
    return riskFromMatrix(crLevel,xrfLevel) || "—";
  }
  // R 원본 자체가 없는 테스트/구형 임시 레코드만 기존 값으로 표시하고,
  // R이 존재하는 정상 품목에서는 이 경로를 절대 사용하지 않습니다.
  return ["Not Allowed","H","M","L"].includes(item?.finalRisk) ? item.finalRisk : "—";
}
function dateGt(a,b){ return a && b && new Date(a) > new Date(b); }
function dateGte(a,b){ return a && b && new Date(a) >= new Date(b); }
function dateLt(a,b){ return a && b && new Date(a) < new Date(b); }
function cleanStatus(s){ return String(s||"").toLowerCase().replace(/[\s_-]/g,""); }
function isDiscontinueRequestItem(item){
  return item?._isRequestRecord===true || String(item?.name||"").startsWith("[단종 요청]");
}
function hasDiscontinueRequestMetadata(item){
  return String(item?.requestType||"").toLowerCase()==="discontinue"
    && !!(item?._spRequestItemId || item?.requestId || item?.discontinuedByRequestCode);
}
function discontinueRequestCodeOf(item){
  return item?.requestId || item?.discontinuedByRequestCode || item?.code || "—";
}
function isDiscontinuedItem(item){
  return item?.lifecycle==="Discontinued";
}
function isHistoryRecordItem(item){
  if(!item) return false;
  const approvalKey=approvalFilterKey(item.approvalStatus);
  const processKey=String(item.processStatus||"").trim().toLowerCase();
  const lifecycleKey=String(item.lifecycle||"").trim();
  return isDiscontinueRequestItem(item)
    || lifecycleKey==="Discontinued"
    || lifecycleKey==="Cancelled"
    || lifecycleKey==="ReplacedOld"
    || approvalKey==="cancelled"
    || processKey==="reverted";
}
function isOperationalManagedItem(item){
  return !!item
    && !isHistoryRecordItem(item)
    && item.isCurrent!==false;
}
function cycleFromFinalRisk(finalRisk){
  if(finalRisk==="H") return "Monthly";
  if(finalRisk==="M") return "Semiannual";
  if(finalRisk==="L") return "Annual";
  return "Not Available";
}
// 검사주기는 품목 Master에 남아 있는 과거 cycle 값을 기준으로 하지 않고,
// R 단계에서 확정된 Final Risk를 단일 기준으로 다시 산정합니다.
// H → 월 1회 / M → 반기 1회 / L → 연 1회 / 사용불허·판정불가 → 주기 없음
function itemCycleFromRisk(item){
  if(!item || isDiscontinueRequestItem(item) || isEquipmentItem(item)) return "Not Available";
  return cycleFromFinalRisk(itemFinalRisk(item));
}
function applyCurrentRiskPolicy(item){
  if(!item) return item;
  const basis=riskBasisMeasurementOf(item);
  // R 원본이 있는 정상 품목만 현재 정책으로 강제 정규화합니다.
  // R이 없는 mock/임시 레코드는 테스트용 저장값을 보존합니다.
  if(!basis) return item;
  const xrfLevel=measurementXrfLevel(basis);
  const finalRisk=itemFinalRisk(item);
  const cycle=itemCycleFromRisk(item);
  // 모든 탭이 과거 Master 저장값이 아니라 동일한 R 재산정 결과를 읽도록 runtime snapshot을 정규화합니다.
  return {
    ...item,
    riskBasisXrfLevel:xrfLevel,
    finalRisk,
    cycle,
    cycleMonths:cycleMonthsFromCycle(cycle),
  };
}
function cycleMonthsFromCycle(cycle){
  if(cycle==="Monthly") return 1;
  if(cycle==="Semiannual") return 6;
  if(cycle==="Annual") return 12;
  return null;
}
function nextPeriod(item,p){
  return (item?.periods||[]).find(x=>Number(x.num)===Number(p?.num)+1);
}
function normalizeItemType(value){
  const key=String(value||"").trim().toLowerCase().replace(/[\s_-]/g,"");
  if(["facility","equipment","설비"].includes(key)) return "Facility";
  if(["auxiliarymaterials","auxiliary","부자재"].includes(key)) return "Auxiliary Materials";
  return String(value||"").trim();
}
function isEquipmentItem(item){
  return normalizeItemType(item?.type)==="Facility";
}
function isComplianceTarget(item){
  if(isDiscontinueRequestItem(item) || isDiscontinuedItem(item)) return false;
  const risk=itemFinalRisk(item);
  const riskCycle=itemCycleFromRisk(item);
  return item?.lifecycle!=="ReplacedOld"
    && item?.category!=="changed"
    && !isEquipmentItem(item)
    && ["H","M","L"].includes(risk)
    && riskCycle!=="Not Available";
}
function approvalFilterKey(status){
  const s=String(status||"").trim();
  if(s==="승인" || s==="사용 승인" || s==="단종 승인" || s==="Approved") return "approved";
  if(s==="측정 진행중" || s==="측정중" || s==="Measuring") return "measuring";
  if(s==="의뢰 필요" || s==="정밀분석 의뢰 필요" || s==="Request Required") return "request";
  if(s==="반려" || s==="사용 반려" || s==="단종 반려" || s==="Rejected") return "rejected";
  if(s==="보류" || s==="단종 검토중" || s==="Hold") return "hold";
  if(s==="처리 취소" || s==="취소" || s==="Cancelled" || s==="Reverted") return "cancelled";
  return "unknown";
}
function crLevelFromType(crType){
  if(crType==="직접접촉+잔류") return "H";
  if(crType==="직접접촉+비잔류") return "M";
  return "L";
}
function requestXrfWorstFromMode(mode, uploadResult){
  return mode==="upload" && uploadResult?.parsed ? uploadResult.xrfWorst : "—";
}
function deriveXrfApprovalStatus(xrfResult){
  if(xrfResult==="OK") return "승인";
  if(xrfResult==="NG") return "의뢰 필요";
  if(xrfResult==="확인 필요" || xrfResult==="재측정 필요") return "보류";
  // XRF 측정 의뢰만 있고 아직 판정 지표가 없으면 승인 판단이 불가능하므로 보류합니다.
  return "보류";
}
function deriveInitialApprovalStatus(measurement){
  if(!measurement) return "보류";
  // NG와 ??가 동시에 있어도 XRF 재측정 의무가 끝나기 전에는 승인/정밀분석 단독 상태로 넘기지 않습니다.
  if(pendingXrfRemeasureElements(measurement).length) return "보류";
  // Retest에서도 남은 ??는 정밀분석 대상으로 전환되므로 저장 상태도 Workflow와 동일하게 맞춥니다.
  if(precisionRequiredElements(measurement).length) return "의뢰 필요";
  const xrfResult=measurement.xrfResult || measurementXrfResult(measurement);
  return deriveXrfApprovalStatus(xrfResult);
}
function evaluateInitialRisk(measurement, crLevel, isFacility=false){
  const xrfLevel=measurementXrfLevel(measurement);
  if(!xrfLevel){
    return {xrfLevel:null,finalRisk:null,cycle:"Not Available",cycleMonths:null,evaluable:false};
  }
  const finalRisk=riskFromMatrix(crLevel,xrfLevel);
  const cycle=isFacility ? "Not Available" : cycleFromFinalRisk(finalRisk);
  return {
    xrfLevel,
    finalRisk,
    cycle,
    cycleMonths:cycleMonthsFromCycle(cycle),
    evaluable:true
  };
}
function emptyUploadedXrfResult(fileName="", extra={}){
  return {
    fileName,
    parsed:false,
    pending:!!fileName,
    error:"",
    sourceSheet:"",
    partNumberCandidate:"",
    measuredDate:null,
    xrfWorst:null,
    xrfResult:null,
    xrfLevel:null,
    retestRequired:false,
    retestRequiredElements:"",
    precisionRequired:false,
    precisionRequiredElements:[],
    dataMissingElements:[],
    reviewRequiredElements:[],
    approvalStatus:"측정 진행중",
    elements:{},
    ...extra,
  };
}
function normalizeUploadedXrfElements(rawRecord){
  const measurement=normalizeMeasurementRecord(rawRecord||{});
  const out={};
  ["Pb","Hg","Cr","Cd","Cl","Br"].forEach(el=>{
    const d={...(measurement.elements?.[el]||{})};
    d.xrfLevel=xrfLevelFromJudgement(d);
    d.remeasureRequired=isXrfRemeasureRequired(d);
    d.retestRequired=d.remeasureRequired;
    d.primaryClass=elementXrfResult(d);
    out[el]=d;
  });
  const clBr=buildClBrResult(sourcePpmValue(out.Cl),sourcePpmValue(out.Br));
  out["Cl+Br"]={
    rawPpm:clBr.ppm,
    ppm:clBr.ppm,
    rawSigma:null,
    sigma:null,
    rawJudgement:clBr.judgement,
    judge:clBr.judgement,
    xrfLevel:clBr.level,
    limit:clBr.legalLimit,
    legalLimit:clBr.legalLimit,
    internalLimit:clBr.internalLimit,
    calculated:true,
    decisionSource:"cl_br_content_sum_internal_limit_70pct",
    dataMissing:clBr.ppm==null,
    isND:!!clBr.isND,
    remeasureRequired:false,
    retestRequired:false,
    primaryClass:clBr.level==="H" ? "NG" : clBr.level ? "OK" : null
  };
  return out;
}
function uploadedPartNumberFromFileName(fileName){
  const stem=String(fileName||"").replace(/\.[^.]+$/,"").trim();
  const matched=stem.match(/^([A-Za-z](?:-?\d+)+)/);
  if(!matched) return "";
  const prefix=matched[1].charAt(0).toUpperCase();
  const groups=matched[1].slice(1).match(/\d+/g)||[];
  return groups.length?`${prefix}-${groups.join("-")}`:"";
}
function samePartNumber(a,b){
  return String(a||"").toUpperCase().replace(/[^A-Z0-9]/g,"")===String(b||"").toUpperCase().replace(/[^A-Z0-9]/g,"");
}
const GENERATED_ITEM_ACTION_NOTE_LINES=new Set([
  "의뢰자 등록 · 신규 도입",
  "관리자 직접 등록 / 신규 도입",
  "XRF 측정 의뢰 상태",
  "XRF 원본 업로드 완료",
]);
function withoutGeneratedItemActionNote(value){
  return String(value||"")
    .replace(/\\n/g,"\n")
    .split(/\r?\n/)
    .map(line=>line.trim())
    .filter(line=>{
      if(!line) return false;
      // 과거에 저장된 자동 문구도 현재 명칭으로 정규화해 계속 숨김 처리합니다.
      const normalizedLine=line.replace(/완전\s+(?=신규\s+도입$)/,"");
      return !GENERATED_ITEM_ACTION_NOTE_LINES.has(normalizedLine);
    })
    .join("\n");
}
function uploadedDateToISO(value){
  if(value==null || value==="") return null;
  if(value instanceof Date && !Number.isNaN(value.getTime())) return fmtISODate(value);
  if(typeof value==="number" && Number.isFinite(value)){
    const parts=XLSX.SSF.parse_date_code(value);
    if(parts?.y && parts?.m && parts?.d){
      return `${parts.y}-${String(parts.m).padStart(2,"0")}-${String(parts.d).padStart(2,"0")}`;
    }
  }
  const text=String(value).trim();
  const direct=text.match(/(20\d{2})[.\/-](\d{1,2})[.\/-](\d{1,2})/);
  if(direct){
    return `${direct[1]}-${String(direct[2]).padStart(2,"0")}-${String(direct[3]).padStart(2,"0")}`;
  }
  const parsed=new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : fmtISODate(parsed);
}
function uploadedNormalizeLabel(value){
  return String(value??"")
    .normalize("NFKC")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9가-힣]/g,"");
}
function uploadedCell(grid,row,col){
  return row>=0 && col>=0 && row<grid.length && col<(grid[row]?.length||0) ? grid[row][col] : "";
}
function uploadedFindLabel(grid,aliases){
  const targets=new Set((aliases||[]).map(uploadedNormalizeLabel).filter(Boolean));
  for(let r=0;r<grid.length;r++){
    for(let c=0;c<(grid[r]?.length||0);c++){
      if(targets.has(uploadedNormalizeLabel(grid[r][c]))) return {row:r,col:c};
    }
  }
  return null;
}
const UPLOADED_ELEMENT_ALIASES={
  Pb:["Pb","Lead","납"],
  Hg:["Hg","Mercury","수은"],
  Cr:["Cr","Chromium","Hexavalent Chromium","Chromium VI","크롬","육가크롬"],
  Cd:["Cd","Cadmium","카드뮴"],
  Cl:["Cl","Chlorine","Chloride","염소"],
  Br:["Br","Bromine","Bromide","브롬"],
};
function uploadedElementCode(value){
  const normalized=uploadedNormalizeLabel(value);
  return Object.entries(UPLOADED_ELEMENT_ALIASES)
    .find(([,aliases])=>aliases.some(alias=>uploadedNormalizeLabel(alias)===normalized))?.[0] || null;
}
function uploadedWorkbookGrid(sheet){
  return XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:"",blankrows:true});
}
function uploadedGridScore(grid){
  let score=uploadedDateToISO(uploadedCell(grid,4,14)) ? 1 : 0;
  [
    ["Element","Elements"],
    ["Content(ppm)","Content (ppm)","Content"],
    ["Std.Deviation(ppm)","Std. Deviation(ppm)","Std Deviation(ppm)","Standard Deviation(ppm)"],
    ["Judgment","Judgement"],
  ].forEach(aliases=>{ if(uploadedFindLabel(grid,aliases)) score+=1; });
  return score;
}
function uploadedRecordFromGrid(grid,sheetName=""){
  const elementHeader=uploadedFindLabel(grid,["Element","Elements"]);
  const contentHeader=uploadedFindLabel(grid,["Content(ppm)","Content (ppm)","Content"]);
  const sigmaHeader=uploadedFindLabel(grid,[
    "Std.Deviation(ppm)","Std. Deviation(ppm)","Std Deviation(ppm)",
    "Std.Deviation (ppm)","Standard Deviation(ppm)","Standard Deviation (ppm)"
  ]);
  const judgementHeader=uploadedFindLabel(grid,["Judgment","Judgement"]);
  if(!elementHeader || !contentHeader || !sigmaHeader || !judgementHeader){
    throw new Error("Element, Content(ppm), Std.Deviation(ppm), Judgment 행을 찾지 못했습니다.");
  }

  const record={Measured_Date:uploadedDateToISO(uploadedCell(grid,4,14)),Source_Sheet:sheetName};
  const found=[];
  const maxCols=Math.max(...grid.map(row=>row?.length||0),0);
  for(let c=elementHeader.col+1;c<maxCols;c++){
    const code=uploadedElementCode(uploadedCell(grid,elementHeader.row,c));
    if(!code) continue;
    record[`${code}_ppm`]=uploadedCell(grid,contentHeader.row,c);
    record[`${code}_3sigma`]=uploadedCell(grid,sigmaHeader.row,c);
    record[`${code}_Judgement`]=uploadedCell(grid,judgementHeader.row,c);
    found.push(code);
  }
  const missingRequired=["Pb","Hg","Cr","Cd","Cl","Br"].filter(code=>!found.includes(code));
  if(missingRequired.length){
    throw new Error(`필수 원소 결과를 찾지 못했습니다: ${missingRequired.join(", ")}`);
  }
  const clBr=buildClBrResult(record.Cl_ppm,record.Br_ppm);
  record["Cl+Br_ppm"]=clBr.ppm;
  record["Cl+Br_3sigma"]=null;
  record["Cl+Br_Judgement"]=clBr.judgement;
  return record;
}
function deriveUploadedXrfResultFromRecord(rawRecord,fileName="",extra={}){
  const source=rawRecord?.measurement || rawRecord?.item || rawRecord || {};
  const elements=normalizeUploadedXrfElements(source);
  const measurement={elements};
  const xrfResult=measurementXrfResult(measurement);
  const xrfWorst=xrfResult;
  const xrfLevel=measurementXrfLevel(measurement);
  const retestList=retestRequiredElements(measurement);
  const retestRequired=retestList.length>0;
  const precisionElements=precisionRequiredElements(measurement);
  const precisionRequired=precisionElements.length>0;
  const dataMissingElements=XRF_REPORT_ELEMENTS.filter(el=>elementDataMissing(elements[el]));
  const reviewRequiredElements=XRF_REPORT_ELEMENTS.filter(el=>!isXrfRemeasureRequired(elements[el]) && !xrfJudgementDecision(elements[el]));
  // 업로드 파일 안의 XRF_Worst/Retest_Required/Approval_Status는 판정 source로 사용하지 않습니다.
  // Pb/Hg/Cr/Cd/Cl/Br은 Content를 원소별 Internal Limit(법적 기준의 70%)과 비교하고,
  // Cl+Br은 Cl+Br Content 합계를 Internal Limit 1,050 ppm과 비교하여 XRF 결과를 산출합니다.
  const approvalStatus=deriveInitialApprovalStatus({...measurement,xrfResult});
  return {
    fileName,
    parsed:true,
    pending:false,
    error:"",
    sourceSheet:extra.sourceSheet || pickField(source,"Source_Sheet","sourceSheet") || "",
    partNumberCandidate:extra.partNumberCandidate || pickField(source,"Part Number","Part_Number","partNumber") || uploadedPartNumberFromFileName(fileName),
    measuredDate:uploadedDateToISO(pickField(source,"Measured_Date","Meas.Date","date","measuredDate")) || extra.measuredDate || null,
    xrfWorst:xrfWorst||null,
    xrfResult,
    xrfLevel,
    retestRequired,
    retestRequiredElements:retestList.join(", "),
    precisionRequired,
    precisionRequiredElements:precisionElements,
    dataMissingElements,
    reviewRequiredElements,
    approvalStatus,
    elements,
    rawRecord:source,
  };
}
async function parseUploadedXrfFile(file){
  if(!file) return emptyUploadedXrfResult();
  const lower=String(file.name||"").toLowerCase();
  try{
    if(lower.endsWith(".json")){
      const parsed=JSON.parse(await file.text());
      return deriveUploadedXrfResultFromRecord(parsed,file.name);
    }
    if(!/\.(xlsx|xlsm|xls|csv)$/.test(lower)){
      return emptyUploadedXrfResult(file.name,{pending:false,error:"지원 형식은 .xlsx, .xlsm, .xls, .csv, .json 입니다."});
    }
    const workbook=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true,raw:true});
    if(!workbook.SheetNames?.length) throw new Error("읽을 수 있는 워크시트가 없습니다.");
    const candidates=workbook.SheetNames.map(name=>{
      const grid=uploadedWorkbookGrid(workbook.Sheets[name]);
      return {name,grid,score:uploadedGridScore(grid)};
    }).sort((a,b)=>b.score-a.score);
    const selected=candidates[0];
    if(!selected || selected.score<4){
      throw new Error(`XRF 보고서 형식을 충분히 찾지 못했습니다. 제목 일치 점수: ${selected?.score||0}/5`);
    }
    const record=uploadedRecordFromGrid(selected.grid,selected.name);
    return deriveUploadedXrfResultFromRecord(record,file.name,{sourceSheet:selected.name,measuredDate:record.Measured_Date,partNumberCandidate:uploadedPartNumberFromFileName(file.name)});
  }catch(error){
    return emptyUploadedXrfResult(file.name,{pending:false,error:`XRF 파싱 실패: ${error?.message||error}`});
  }
}

function parseISODate(s){
  if(!s) return null;
  const [y,m,d]=String(s).slice(0,10).split("-").map(Number);
  if(!y||!m||!d) return null;
  return new Date(y,m-1,d);
}
function fmtISODate(d){
  if(!d) return null;
  const y=d.getFullYear();
  const m=String(d.getMonth()+1).padStart(2,"0");
  const day=String(d.getDate()).padStart(2,"0");
  return `${y}-${m}-${day}`;
}
function addDaysISO(s,n){
  const d=parseISODate(s);
  if(!d) return null;
  d.setDate(d.getDate()+n);
  return fmtISODate(d);
}
function addMonthsISO(s,n){
  const d=parseISODate(s);
  if(!d) return null;
  const day=d.getDate();
  d.setDate(1);
  d.setMonth(d.getMonth()+n);
  const last=new Date(d.getFullYear(),d.getMonth()+1,0).getDate();
  d.setDate(Math.min(day,last));
  return fmtISODate(d);
}
function diffDaysISO(a,b){
  const da=parseISODate(a), db=parseISODate(b);
  if(!da||!db) return 0;
  return Math.round((da-db)/86400000);
}
function cycleMonthsOf(item){
  return cycleMonthsFromCycle(itemCycleFromRisk(item)) || 0;
}
function windowDaysOf(item){
  return itemCycleFromRisk(item)==="Monthly" ? 15 : 30;
}
function periodLabel(p){
  if(!p) return "—";
  if(p.displayLabel) return p.displayLabel;
  if(p.isReference || Number(p.num)===0) return "R";
  return `P${Number(p.num||1)}`;
}
function periodPointDate(p){
  if(!p) return null;
  return p.isReference ? (p.referenceDate || p.date || null) : (p.due || p.targetDate || null);
}
function periodMeasurementIds(p){
  if(!p) return [];
  const ids=Array.isArray(p.measurementIds) ? p.measurementIds : (p.measurementId?[p.measurementId]:[]);
  return [...new Set(ids.filter(Boolean))];
}
function periodMeasuredDates(p){
  if(!p) return [];
  const dates=Array.isArray(p.measuredDates) ? p.measuredDates : (p.measured?[p.measured]:[]);
  return [...new Set(dates.map(dateOnly).filter(Boolean))].sort();
}
function periodMeasuredText(p){
  const dates=periodMeasuredDates(p);
  return dates.length ? dates.join(", ") : "—";
}
function periodMeasuredShortText(p){
  const dates=periodMeasuredDates(p);
  return dates.map(d=>d.slice(5)).join(", ");
}
function periodMeasurementsOf(item,p){
  return measurementAttemptsForPeriod(item,p)
    .map(row=>measurementById(item,row?.id||row?.measurementId)||row)
    .filter(row=>row && (row.id||row.measurementId))
    .sort((a,b)=>String(a?.date||"").localeCompare(String(b?.date||"")));
}
function appendPeriodMeasurementLink(period,measurementId,measuredDate){
  const ids=periodMeasurementIds(period);
  const dates=periodMeasuredDates(period);
  const pairs=[];
  ids.forEach((id,idx)=>pairs.push({id,date:dates[idx]||""}));
  if(measurementId && !pairs.some(row=>row.id===measurementId)) pairs.push({id:measurementId,date:dateOnly(measuredDate)||""});
  pairs.sort((a,b)=>String(a.date||"").localeCompare(String(b.date||"")));
  const cleanPairs=pairs.filter(row=>row.id||row.date);
  return {
    measurementIds:cleanPairs.map(row=>row.id).filter(Boolean),
    measuredDates:cleanPairs.map(row=>row.date).filter(Boolean),
    measurementId:cleanPairs[cleanPairs.length-1]?.id||measurementId||period?.measurementId||null,
    measured:cleanPairs[cleanPairs.length-1]?.date||dateOnly(measuredDate)||period?.measured||null,
  };
}
function dateLte(a,b){ return !dateGt(a,b); }
function collectPeriodMeasurements(item){
  const map=new Map();
  const add=(date,id,status,role,attemptNo)=>{
    const d=String(date||"").slice(0,10);
    if(!d) return;
    // 재측정(Retest)은 동일 Pn의 판정 보완이므로 다음 Pn 이행 실적으로 소비하지 않습니다.
    // 과거 데이터의 Cycle_Reset은 별도 기준점으로 쓰지 않고, R 이후 실제 XRF 측정 1건으로만 취급합니다.
    if(String(role||"").toLowerCase()==="retest") return;
    const key=`${d}|${id||""}`;
    if(!map.has(key)) map.set(key,{date:d,id:id||null,status:status||"",role:normalizeMeasurementRole(role),attemptNo:Number(attemptNo||1)});
  };
  (item?.periods||[]).forEach(p=>add(p.measured,p.measurementId,p.status,p.measurementRole,p.attemptNo));
  (item?.history||[]).forEach(h=>add(h.date,h.id||h.measurementId,h.status,h.measurementRole||h.role,h.attemptNo));
  if(item?.latestMeasurementId){
    const m=(item?.__measurementMap||{})[item.latestMeasurementId]||item.latest;
    add(m?.date||item.lastMeasured,item.latestMeasurementId,m?.status,m?.role||m?.measurementRole,m?.attemptNo);
  }
  return Array.from(map.values()).sort((a,b)=>String(a.date).localeCompare(String(b.date)));
}
function normalizeDbPeriods(item){
  const raw=item?.periods || [];
  if(!raw.length) return [];
  const referenceRaw=raw.find(p=>p?.isReference || Number(p?.num)===0 || String(p?.label||p?.displayLabel||"").trim()==="R");
  if(!referenceRaw) return [];
  // 기존 DB에 저장된 num=-1 "이행/지연" 또는 Cycle_Reset 기반 Pn은 화면 주기로 재사용하지 않습니다.
  // R만 기준점으로 보존하고 P1, P2...는 아래 buildRollingPeriods()에서 R 등록일 + R Final Risk 주기로 다시 생성합니다.
  return [{
    ...referenceRaw,
    num:0,
    displaySeq:0,
    label:"R",
    displayLabel:"R",
    isReference:true,
    referenceDate:referenceRaw.referenceDate || referenceRaw.date || referenceRaw.measured || referenceRaw.start || item?.firstDate || null,
    date:referenceRaw.date || referenceRaw.referenceDate || referenceRaw.measured || referenceRaw.start || item?.firstDate || null,
    due:null,
    targetDate:null,
    windowStart:null,
    status:"reference",
    closeType:referenceRaw.closeType || "registered",
    isDisplayTarget:false,
    isCurrentCycle:true,
    measurementRole:"Initial"
  }];
}

function buildRollingPeriods(item){
  const fromDb=normalizeDbPeriods(item);
  const fromDbReference=fromDb[0] || null;

  // Pn 생성의 유일한 기준은 최초 위험도 평가 R입니다.
  // R 등록일 + R Final Risk로 산정된 고정 주기로 P1, P2 ...를 생성합니다.
  // Annual(연 1회)은 같은 달력연도에 여러 XRF가 있어도 다음 Pn으로 넘기지 않고
  // 해당 연도의 동일 Pn에 모두 묶습니다. 예: P1 도래연도 2026 → 2026-01-30, 2026-07-14 모두 P1.
  if(!fromDbReference) return [];
  if(!isComplianceTarget(item)) return fromDb;

  const months=cycleMonthsOf(item);
  const baseDate=String(fromDbReference.referenceDate || fromDbReference.date || item?.firstDate || "").slice(0,10);
  if(!months || !baseDate) return fromDb;

  const riskCycle=itemCycleFromRisk(item);
  const isAnnual=months===12;
  const windowDays=windowDaysOf(item);
  const measurements=collectPeriodMeasurements(item);
  const out=[];
  let mi=0;
  let futureCount=0;

  // R은 최초 위험도 평가 기준점일 뿐 P1 이행값으로 재사용하지 않습니다.
  let referenceMeasurement=null;
  while(mi<measurements.length && dateLte(measurements[mi].date,baseDate)){
    referenceMeasurement=measurements[mi++];
  }
  const referenceDate=fromDbReference.referenceDate || fromDbReference.date || baseDate;
  const referenceMeasured=fromDbReference.measured || referenceMeasurement?.date || baseDate;
  const referenceMeasurementId=fromDbReference.measurementId || referenceMeasurement?.id || item?.firstMeasurementId || null;
  out.push({
    ...fromDbReference,
    num:0,displaySeq:0,displayLabel:"R",label:"R",isReference:true,
    referenceDate,date:referenceDate,start:null,due:null,targetDate:null,windowStart:null,
    measured:referenceMeasured,measuredDates:referenceMeasured?[referenceMeasured]:[],
    measurementId:referenceMeasurementId,measurementIds:referenceMeasurementId?[referenceMeasurementId]:[],
    status:"reference",statusLabel:"기준",daysOver:0,closeType:"registered",
    finalRisk:itemFinalRisk(item),
    isDisplayTarget:false,isCurrentCycle:true,measurementRole:"Initial",
    cycle:riskCycle,cycleMonths:months,windowDays,
    sourceBasis:"최초 XRF 측정 및 R 위험도 평가 기준"
  });

  // 연 1회 품목은 달력연도를 주기 bucket으로 사용합니다.
  // R 이후의 실제 XRF를 YYYY 기준으로 묶어 동일 연도 측정값이 하나의 Pn에 함께 표시되게 합니다.
  const annualMeasurementsByYear=new Map();
  if(isAnnual){
    measurements.slice(mi).forEach(m=>{
      const year=String(m?.date||"").slice(0,4);
      if(!year) return;
      if(!annualMeasurementsByYear.has(year)) annualMeasurementsByYear.set(year,[]);
      annualMeasurementsByYear.get(year).push(m);
    });
    annualMeasurementsByYear.forEach(rows=>rows.sort((a,b)=>String(a.date).localeCompare(String(b.date))));
  }

  for(let n=1;n<=60;n++){
    const prevDue=n===1 ? baseDate : addMonthsISO(baseDate,months*(n-1));
    const due=addMonthsISO(baseDate,months*n);
    if(!due) break;
    const nextDue=addMonthsISO(baseDate,months*(n+1));
    const windowStart=addDaysISO(due,-windowDays);

    let hits=[];
    let hitLate=false;
    if(isAnnual){
      // 연 1회: Pn의 도래연도와 실제 측정연도가 같으면 모두 동일 Pn에 귀속합니다.
      hits=[...(annualMeasurementsByYear.get(String(due).slice(0,4))||[])];
      if(hits.length) hitLate=!hits.some(hit=>dateLte(hit.date,due));
    }else{
      // 월/반기: 기존대로 하나의 실제 측정을 하나의 Pn에 순차 연결합니다.
      while(mi<measurements.length && prevDue && dateLte(measurements[mi].date,prevDue)) mi++;
      if(mi<measurements.length){
        const md=measurements[mi].date;
        if(dateLte(md,due)){
          hits=[measurements[mi++]];
        }else if(nextDue && dateLt(md,nextDue)){
          hits=[measurements[mi++]];
          hitLate=true;
        }
      }
    }

    let status="upcoming";
    let statusLabel="예정";
    let closeType="system_open";
    let measured=null;
    let measurementId=null;
    let measurementRole=null;
    let daysOver=0;
    const measuredDates=hits.map(hit=>hit.date).filter(Boolean).sort();
    const measurementIds=hits.map(hit=>hit.id).filter(Boolean);

    if(hits.length){
      const latestHit=hits[hits.length-1];
      measured=latestHit.date;
      measurementId=latestHit.id;
      measurementRole="Periodic";
      if(hitLate){
        const firstLate=hits.find(hit=>dateGt(hit.date,due)) || hits[0];
        status="late";
        statusLabel="지연측정";
        daysOver=diffDaysISO(firstLate.date,due);
        closeType=hits.length>1?"measured_multiple_late":"measured_late";
      }else{
        status="compliant";
        statusLabel="이행";
        closeType=hits.length>1?"measured_multiple":"measured";
      }
    }else if(dateGt(TODAY_STR,due)){
      status="overdue";
      statusLabel="미이행";
      closeType="missed";
      daysOver=diffDaysISO(TODAY_STR,due);
    }else if(windowStart && !dateLt(TODAY_STR,windowStart)){
      status="due_soon";
      statusLabel="측정권장";
    }

    out.push({
      num:n,displaySeq:n,displayLabel:`P${n}`,label:`P${n}`,isReference:false,
      referenceDate:null,date:null,start:prevDue,due,targetDate:due,prevDue,nextDue,windowStart,
      measured,measuredDates,measurementId,measurementIds,
      status,statusLabel,daysOver:Math.max(0,daysOver),closeType,
      finalRisk:itemFinalRisk(item),isDisplayTarget:true,isCurrentCycle:true,measurementRole,
      cycle:riskCycle,cycleMonths:months,windowDays,cycleVersion:1,
      groupingBasis:isAnnual?"CALENDAR_YEAR":"DUE_INTERVAL",
      sourceBasis:isAnnual
        ? `R 등록일 ${baseDate} 기준 ${CYCLE_KO[riskCycle]||riskCycle} · ${String(due).slice(0,4)}년 측정 이력은 P${n}에 통합`
        : `R 등록일 ${baseDate} 기준 ${CYCLE_KO[riskCycle]||riskCycle} 고정 주기`
    });

    if(!dateGt(TODAY_STR,due)) futureCount++;
    if(months===1){
      if(out.length>=25 && futureCount>=24) break;
    }else if(out.length>=5 && futureCount>=2) break;
  }
  return out;
}

function lifecyclePeriodEndDate(item){
  return dateOnly(item?.finalUseDate || item?.replacementDate || item?.processedAt || item?.approvalCompletedAt || null);
}
function truncateTerminalLifecyclePeriods(item,periods){
  const lifecycle=String(item?.lifecycle||"");
  const terminal=["ReplacedOld","Discontinued","Cancelled"].includes(lifecycle) || item?.isCurrent===false;
  if(!terminal) return periods||[];
  const cutoff=lifecyclePeriodEndDate(item);
  return (periods||[]).filter(p=>{
    if(p?.isReference || Number(p?.num)===0 || p?.measured) return true;
    // 종료일이 있으면 그 날짜까지 도래한 과거 관리주기만 이력으로 보존합니다.
    // 종료일을 알 수 없으면 측정 이력이 없는 미래 P주기는 즉시 제거합니다.
    if(!cutoff) return false;
    const due=dateOnly(p?.due || p?.targetDate);
    return !!due && dateLte(due,cutoff);
  });
}
function itemAllPeriods(item){
  return truncateTerminalLifecyclePeriods(item,buildRollingPeriods(item));
}

// XRF 분석 탭은 "정기 관리주기"와 "실제 XRF 측정 이력"을 같은 데이터 원본에서 보여 줍니다.
// Final Risk가 Not Allowed이면 정기 Pn은 생성하지 않지만, R 이후 실제로 수행된 XRF까지 숨기면
// 위험도 현황의 최신 XRF와 XRF 분석 탭이 서로 다른 것처럼 보이므로 미연결 측정을 F1, F2...로 표시합니다.
// F(Follow-up)는 실제 측정 이력 표시용이며 정기 Pn/이행률/다음 도래일 계산에는 절대 포함하지 않습니다.
function itemXrfDisplayPeriods(item){
  const base=itemAllPeriods(item);
  if(!item || isDiscontinueRequestItem(item)) return base;

  const linkedIds=new Set(base.flatMap(periodMeasurementIds));
  const records=itemMeasurementRecords(item);
  const riskBasis=riskBasisMeasurementOf(item);
  const riskBasisId=String(riskBasis?.id||riskBasis?.measurementId||item?.firstMeasurementId||"").trim();

  // 정기 Pn이 존재하는 품목의 실제 측정은 buildRollingPeriods()가 Pn에 연결합니다.
  // 여기서 추가 노드를 만드는 경우는 주기 자체가 없는 품목(대표적으로 R Final Risk = Not Allowed)뿐입니다.
  if(itemCycleFromRisk(item)!=="Not Available") return base;

  const extras=records.filter(m=>{
    const id=String(m?.id||m?.measurementId||"").trim();
    if(!id || id===riskBasisId || linkedIds.has(id)) return false;
    // Retest는 독립 관리주기가 아니라 직전 XRF 판정 보완이므로 별도 F 노드로 만들지 않습니다.
    if(normalizeMeasurementRole(m?.role||m?.measurementRole)==="Retest") return false;
    return true;
  }).sort((a,b)=>String(a?.date||"").localeCompare(String(b?.date||"")));

  if(!extras.length) return base;
  const followups=extras.map((m,idx)=>{
    const id=String(m?.id||m?.measurementId||"").trim();
    const measured=dateOnly(m?.date||m?.measuredDate)||null;
    return {
      num:900001+idx,
      displaySeq:900001+idx,
      label:`F${idx+1}`,
      displayLabel:`F${idx+1}`,
      isReference:false,
      isSupplementalMeasurement:true,
      isCompliancePeriod:false,
      referenceDate:null,
      date:null,
      start:null,
      due:null,
      targetDate:null,
      prevDue:null,
      nextDue:null,
      windowStart:null,
      measured,
      measuredDates:measured?[measured]:[],
      measurementId:id||null,
      measurementIds:id?[id]:[],
      status:"measured",
      statusLabel:"측정 완료",
      daysOver:0,
      closeType:"supplemental_measurement",
      finalRisk:itemFinalRisk(item),
      isDisplayTarget:true,
      isCurrentCycle:false,
      measurementRole:normalizeMeasurementRole(m?.role||m?.measurementRole||"Periodic"),
      cycle:"Not Available",
      cycleMonths:null,
      windowDays:null,
      sourceBasis:"R 이후 실제 XRF 측정 이력 · 정기 Pn과 별도 표시"
    };
  });
  return [...base,...followups];
}
// 선택한 Pn의 XRF 시도 이력을 반환합니다.
// 과거 데이터의 Cycle_Reset/num=-1 측정도 새 R 기준 Pn에 measurementId가 연결되면 해당 Pn의 1차 측정으로 취급합니다.
function measurementAttemptsForPeriod(item,period){
  if(!item || !period) return [];
  const history=Array.isArray(item.history)?item.history:[];
  const periodNo=Number(period.num);
  const isReference=!!period?.isReference || periodNo===0;
  const linkedIds=periodMeasurementIds(period);
  const rows=[];
  const seen=new Set();
  const add=(h)=>{
    if(!h) return;
    const id=h.id||h.measurementId;
    if(!id || seen.has(id)) return;
    if(String(h.measurementRole||h.role||"")==="Initial" && !isReference) return;
    seen.add(id);
    rows.push(h);
  };
  linkedIds.forEach(id=>add(history.find(h=>(h.id||h.measurementId)===id)));
  // 새로 저장된 데이터의 periodNo도 함께 읽고, Retest는 parent가 동일 Pn 측정에 연결된 경우 포함합니다.
  history.forEach(h=>{
    const id=h.id||h.measurementId;
    const parentId=h.parentMeasurementId||null;
    const rawStoredPeriod=h.periodNo;
    const hasStoredPeriod=rawStoredPeriod!==null
      && rawStoredPeriod!==undefined
      && String(rawStoredPeriod).trim()!=="";
    const role=normalizeMeasurementRole(h.measurementRole||h.role||"");
    // Number(null)과 Number("")는 0이므로 빈 periodNo를 R로 오인하지 않도록 먼저 존재 여부를 확인합니다.
    const sameStoredPeriod=hasStoredPeriod
      && Number(rawStoredPeriod)===periodNo
      && (isReference ? role==="Initial" : !["Initial","Cycle_Reset"].includes(role));
    const linkedRetest=role==="Retest" && parentId && linkedIds.includes(parentId);
    if(sameStoredPeriod || linkedRetest || linkedIds.includes(id)) add(h);
  });
  return rows.sort((a,b)=>String(a.date||"").localeCompare(String(b.date||"")) || Number(a.attemptNo||1)-Number(b.attemptNo||1));
}

// XRF 분석 탭의 모든 주기 기반 영역은 동일한 12개 window를 공유합니다.
// 한 페이지는 최대 12개 노드이며, 페이지 전환 시 직전 페이지의 마지막 노드를 다음 페이지 첫 노드로 1개 중복합니다.
// page 1: 1~12 / page 2: 12~23 / page 3: 23~34 ...
const XRF_PERIOD_WINDOW_SIZE=12;
const XRF_PERIOD_WINDOW_STEP=XRF_PERIOD_WINDOW_SIZE-1;

function isCompletedXrfPeriod(p){
  const st=normStatus(p?.status);
  return !!p?.measured || !!p?.isReference || ["reference","compliant","measured","late","early"].includes(st);
}

function xrfPeriodPageMeta(periods, selectedPeriodNum=null, manualPage=null){
  const safePeriods=periods||[];
  const totalCount=safePeriods.length;
  const totalPages=totalCount<=XRF_PERIOD_WINDOW_SIZE
    ? 1
    : 1+Math.ceil((totalCount-XRF_PERIOD_WINDOW_SIZE)/XRF_PERIOD_WINDOW_STEP);

  const pageForIndex=(idx)=>{
    if(idx<0 || idx<XRF_PERIOD_WINDOW_STEP) return 0;
    return Math.min(totalPages-1,Math.floor(idx/XRF_PERIOD_WINDOW_STEP));
  };

  let latestCompletedIdx=-1;
  for(let i=safePeriods.length-1;i>=0;i--){
    if(isCompletedXrfPeriod(safePeriods[i])){ latestCompletedIdx=i; break; }
  }
  if(latestCompletedIdx<0) latestCompletedIdx=0;
  const autoPage=pageForIndex(latestCompletedIdx);

  const selectedIdx=(selectedPeriodNum!==null && selectedPeriodNum!==undefined)
    ? safePeriods.findIndex(p=>Number(p.num)===Number(selectedPeriodNum))
    : -1;
  const selectedPage=selectedIdx>=0 ? pageForIndex(selectedIdx) : autoPage;
  const requestedPage=(manualPage===null || manualPage===undefined) ? selectedPage : Number(manualPage);
  const activePage=Math.max(0,Math.min(totalPages-1,Number.isFinite(requestedPage)?requestedPage:selectedPage));
  const windowStart=activePage*XRF_PERIOD_WINDOW_STEP;
  const visiblePeriods=safePeriods.slice(windowStart,windowStart+XRF_PERIOD_WINDOW_SIZE);

  return {
    totalCount,totalPages,activePage,windowStart,visiblePeriods,
    windowSize:XRF_PERIOD_WINDOW_SIZE,windowStep:XRF_PERIOD_WINDOW_STEP,
    autoPage,selectedPage
  };
}

function itemTimelinePeriods(item){
  if(isDiscontinueRequestItem(item)) return [];
  const all=itemXrfDisplayPeriods(item);
  if(!all.length) return all;

  const isMonthly=itemCycleFromRisk(item)==="Monthly" || cycleMonthsOf(item)===1;

  const ref=all.find(p=>p.isReference || Number(p.num)===0) || null;
  const ps=all
    .filter(p=>!(p.isReference || Number(p.num)===0))
    .sort((a,b)=>Number(a.num||0)-Number(b.num||0));
  const displayTargets=ps.filter(p=>p.isDisplayTarget);
  const targets=(displayTargets.length ? displayTargets : ps)
    .sort((a,b)=>Number(a.num||0)-Number(b.num||0));

  // 월 1회 타임라인은 DotFull에서 한 줄 최대 12개 노드로 windowing 합니다.
  // 여기서는 R + 사용 가능한 전체 P주기를 넘겨, 12번째 완료 노드가 다음 window의 첫 노드로 이어지게 합니다.
  if(isMonthly) return ref ? [ref, ...targets] : targets;

  // XRF 상세 타임라인은 반기/연간 주기의 전체 흐름을 보여준다.
  // DotFull에서 완료된 과거 주기를 압축하고 연도별로 묶어 세로 길이를 제어한다.
  return ref ? [ref, ...targets] : targets;
}

function itemListPeriods(item){
  const all=itemAllPeriods(item);
  if(!all.length) return all;

  const ref=all.find(p=>p.isReference || Number(p.num)===0) || null;
  const ps=all
    .filter(p=>!(p.isReference || Number(p.num)===0))
    .sort((a,b)=>{
      const da=Number.isFinite(Number(a?.displaySeq)) ? Number(a.displaySeq) : Number(a?.num||0);
      const db=Number.isFinite(Number(b?.displaySeq)) ? Number(b.displaySeq) : Number(b?.num||0);
      return da-db;
    });
  const displayTargets=ps.filter(p=>p.isDisplayTarget);
  const targets=(displayTargets.length ? displayTargets : ps);
  const sequence=ref ? [ref,...targets] : targets;
  if(!sequence.length) return [];

  // 부자재 리스트의 이행 현황은 최대 4개의 원만 표시합니다.
  // 1~4번째를 보이다가 4번째가 완료되면 4번째를 다음 구간의 첫 원으로 재사용합니다.
  // 따라서 1~4 → 4~7 → 7~10 → 10~13 ... 순서로 롤링하며 원의 개수는 4개를 넘지 않습니다.
  const WINDOW_SIZE=4;
  const WINDOW_STEP=WINDOW_SIZE-1;

  const isCompletedPeriod=(p)=>{
    const st=normStatus(p?.status);
    return !!p?.measured || !!p?.isReference || ["reference","compliant","measured","late","early"].includes(st);
  };

  let latestCompletedIdx=-1;
  for(let i=sequence.length-1;i>=0;i--){
    if(isCompletedPeriod(sequence[i])){ latestCompletedIdx=i; break; }
  }
  if(latestCompletedIdx<0) latestCompletedIdx=0;

  // index 3(화면의 4번째 원)이 완료되는 즉시 start=3이 됩니다.
  // 이때 직전 마지막 원이 새 화면의 첫 원이 되고, 뒤의 새 주기들이 같은 4개 슬롯을 채웁니다.
  const windowStart=latestCompletedIdx<WINDOW_STEP ? 0 : Math.floor(latestCompletedIdx/WINDOW_STEP)*WINDOW_STEP;
  return sequence.slice(windowStart,windowStart+WINDOW_SIZE);
}
function itemStatusPeriods(item){
  if(isDiscontinueRequestItem(item) || isDiscontinuedItem(item)) return [];
  // 과거 누락은 이력에 보존하되 현재 KPI/이행 상태에는 포함하지 않습니다.
  return itemAllPeriods(item).filter(p=>
    !(p?.isReference || Number(p?.num)===0)
    && p?.isCurrentCycle!==false
  );
}
function currentCompliancePeriod(item){
  if(!isComplianceTarget(item)) return null;
  const periods=itemStatusPeriods(item)
    .filter(p=>p?.due)
    .sort((a,b)=>String(a.due).localeCompare(String(b.due)));
  if(!periods.length) return null;

  // 마감일이 지난 가장 최근 주기가 미측정이면 현재 상태는 계속 미이행이다.
  const latestPast=periods.filter(p=>dateLt(p.due,TODAY_STR)).slice(-1)[0] || null;
  if(latestPast && !latestPast.measured) return latestPast;

  // 오늘이 마감일이거나 아직 마감 전인 첫 주기가 현재 판정 대상이다.
  const currentOrNext=periods.find(p=>!dateLt(p.due,TODAY_STR));
  return currentOrNext || latestPast;
}
function nextDueOfItem(item){
  if(!isComplianceTarget(item)) return null;
  const periods=itemStatusPeriods(item)
    .filter(p=>p?.due)
    .sort((a,b)=>String(a.due).localeCompare(String(b.due)));
  if(!periods.length) return null;

  // 이미 측정이 완료된 주기의 도래일이 아니라, 실제로 다음에 관리해야 할 미측정 주기의 도래일을 반환합니다.
  // 마감 초과 미측정 건이 있으면 해당 기한을 우선 보여 주어 D+ 상태를 유지합니다.
  const overdue=periods.filter(p=>!p.measured && dateLt(p.due,TODAY_STR)).slice(-1)[0] || null;
  if(overdue) return overdue.due;

  const nextOpen=periods.find(p=>!p.measured && !dateLt(p.due,TODAY_STR));
  return nextOpen?.due || null;
}
function isPeriodOverdue(item,p){
  if(p?.isReference || !p?.due) return false;
  const st=normStatus(p.status);
  if(st==="overdue") return true;
  const ps=cleanStatus(p.status);
  const pc=cleanStatus(p.closeType);
  return !p.measured && (dateGt(TODAY_STR,p.due) || ["overdue","missed","미이행","notdone","noncompliant"].includes(ps) || pc==="missed");
}
function isPeriodLate(item,p){
  if(p?.isReference || !p?.due) return false;
  const st=normStatus(p.status);
  if(st==="late") return true;
  const ps=cleanStatus(p.status);
  return (p.measured && dateGt(p.measured,p.due)) || ["late","delaypending","지연측정"].includes(ps);
}
function hasOverduePeriod(item){
  return isComplianceTarget(item) && itemStatusPeriods(item).some(p=>isPeriodOverdue(item,p));
}
function hasLatePeriod(item){
  return isComplianceTarget(item) && itemAllPeriods(item).some(p=>
    !(p?.isReference || Number(p?.num)===0)
    && isPeriodLate(item,p)
  );
}
function getCS(item){
  if(!isComplianceTarget(item)) return "notAvailable";
  const periods=itemStatusPeriods(item)
    .filter(p=>p?.due)
    .sort((a,b)=>String(a.due).localeCompare(String(b.due)));
  if(!periods.length) return "notAvailable";

  // 현재 상태는 "다음 Pn의 이름"이 아니라 현재 시점의 관리 상태를 표시합니다.
  // 1) 과거 도래 Pn 중 미측정이 하나라도 남아 있으면 미이행
  // 2) 다음 미측정 Pn의 측정권장 window가 시작됐으면 측정권장
  // 3) R 이후 Pn 측정 이력이 있고 다음 window 전이면 이행
  // 4) 아직 Pn 측정 이력이 전혀 없고 다음 window 전이면 예정
  const unresolvedOverdue=periods
    .filter(p=>!p.measured && dateLt(p.due,TODAY_STR))
    .slice(-1)[0] || null;
  if(unresolvedOverdue) return "overdue";

  const nextOpen=periods.find(p=>!p.measured && !dateLt(p.due,TODAY_STR)) || null;
  if(nextOpen?.windowStart && !dateLt(TODAY_STR,nextOpen.windowStart)) return "dueSoon";

  const hasCompletedPeriodic=periods.some(p=>!!p.measured);
  return hasCompletedPeriodic ? "ok" : "upcoming";
}

// -----------------------------------------------------------------------------
// Visual UAT fixture pack
// Raw element/measurement inputs are supplied here; the existing production
// normalization, XRF, Risk, Cycle, Pn and Workflow functions derive all results.
const VISUAL_UAT_DEPTS=["Assembly","Plating","Stamping","Common"];
const VISUAL_UAT_MATERIALS=["Chemical","Metal","Plastic","Others"];
const VISUAL_UAT_CHEMICAL_STATES=["Liquid","Volatile Liquid","Viscous Liquid","Aerosol","Gas"];
const VISUAL_UAT_CR_TYPES=["직접접촉+잔류","직접접촉+비잔류","비접촉"];

function visualUatSequence(code){
  return String(code||"").split("").reduce((sum,ch)=>sum+ch.charCodeAt(0),0);
}
function visualUatElementInputs(overrides={}){
  const output={};
  ["Pb","Hg","Cr","Cd","Cl","Br"].forEach((element,index)=>{
    const override=overrides[element]||{};
    output[element]={
      rawPpm:"ND",
      rawSigma:index+1,
      rawJudgement:"ND",
      ...override
    };
  });
  return output;
}
function makeVisualUatMeasurement(id,date,{role="Initial",attemptNo=1,parentMeasurementId=null,elements={},periodNo=null,precisionAnalysisResult=""}={}){
  return normalizeMeasurementRecord({
    id,date,role,attemptNo,parentMeasurementId,periodNo,
    precisionAnalysisResult,
    sourceFileName:`${id}.xlsx`,
    sourceSheet:"Visual UAT",
    elements:visualUatElementInputs(elements)
  },id);
}
function visualUatRawForLevel(level){
  if(level==="H") return {Pb:{rawPpm:701,rawSigma:3,rawJudgement:"NG"}};
  if(level==="M") return {Pb:{rawPpm:100,rawSigma:3,rawJudgement:"OK"}};
  return {};
}
function makeVisualUatItem(code,name,{measurements=[],crType="직접접촉+비잔류",type="Auxiliary Materials",lifecycle="ExistingActive",isCurrent=true,category="existing",replacementOf=null,replacedBy=null,extra={}}={}){
  const seq=visualUatSequence(code);
  const materialCategory=extra.materialCategory||VISUAL_UAT_MATERIALS[seq%VISUAL_UAT_MATERIALS.length];
  const materialState=materialCategory==="Chemical"
    ? (extra.materialState||VISUAL_UAT_CHEMICAL_STATES[seq%VISUAL_UAT_CHEMICAL_STATES.length])
    : (extra.materialState||"");
  const normalizedMeasurements=measurements.map(row=>normalizeMeasurementRecord(row,row?.id||""));
  const measurementMap=Object.fromEntries(normalizedMeasurements.map(row=>[row.id,row]));
  const initial=normalizedMeasurements.find(row=>normalizeMeasurementRole(row.role||row.measurementRole)==="Initial")||null;
  const latest=normalizedMeasurements[normalizedMeasurements.length-1]||null;
  const history=normalizedMeasurements.map((row,index)=>({
    id:row.id,date:row.date,periodNo:row.periodNo??(normalizeMeasurementRole(row.role)==="Initial"?0:index),
    measurementRole:normalizeMeasurementRole(row.role),attemptNo:Number(row.attemptNo||1),
    parentMeasurementId:row.parentMeasurementId||null,elements:row.elements,
    sourceFileName:row.sourceFileName||"",sourceSheet:row.sourceSheet||"Visual UAT"
  }));
  const base={
    __visualUat:true,__readOnlyFixture:true,
    itemId:`VISUAL_UAT::${code}`,_sourceCode:code,code,name,nameEn:`Visual UAT fixture ${code}`,
    type,dept:extra.dept||VISUAL_UAT_DEPTS[seq%VISUAL_UAT_DEPTS.length],respDept:"PQE",
    materialCategory,materialState,manufacturer:"Visual UAT only",
    crType,crLevel:crLevelFromType(crType),riskBasisCrLevel:crLevelFromType(crType),
    sourceClassification:category==="new"?"신규 등록":"기존 품목",originCategory:category,
    category,categoryLabel:category==="new"?"신규 등록":"기존 품목",lifecycle,isCurrent,
    replacementOf,replacedBy,replacementDate:extra.replacementDate||null,
    firstDate:initial?.date||extra.firstDate||null,lastMeasured:latest?.date||null,
    firstMeasurementId:initial?.id||null,latestMeasurementId:latest?.id||null,
    riskBasisMeasurementId:initial?.id||null,measurementCount:normalizedMeasurements.length,
    periods:initial?[{
      num:0,label:"R",displayLabel:"R",displaySeq:0,isReference:true,
      referenceDate:initial.date,date:initial.date,measured:initial.date,
      measurementId:initial.id,measurementIds:[initial.id],measuredDates:[initial.date],
      status:"reference",closeType:"registered",isDisplayTarget:false,isCurrentCycle:true,
      measurementRole:"Initial"
    }]:[],
    history,latest,__measurementMap:measurementMap,
    ...extra
  };
  return normalizeItemRecord(base);
}
function visualUatInitial(code,date,level="M",options={}){
  return makeVisualUatMeasurement(`${code}_R`,date,{...options,elements:{...visualUatRawForLevel(level),...(options.elements||{})}});
}
function visualUatPeriodic(code,date,level="M",options={}){
  return makeVisualUatMeasurement(`${code}_${options.suffix||"P1"}`,date,{...options,role:options.role||"Periodic",periodNo:options.periodNo||1,elements:{...visualUatRawForLevel(level),...(options.elements||{})}});
}

function buildVisualUatFixtures(){
  const fixtures=[];
  const push=(code,name,options)=>fixtures.push(makeVisualUatItem(code,name,options));
  const xrfDate=addMonthsISO(TODAY_STR,-2);

  const xrfCases=[
    ["VU-X01","All ND → L / OK",{}],
    ["VU-X02","N.D. marker → L / OK",{Pb:{rawPpm:"N.D.",rawSigma:11,rawJudgement:"N.D."}}],
    ["VU-X03","---- marker → L / OK",{Cr:{rawPpm:"----",rawSigma:"----",rawJudgement:"----"}}],
    ["VU-X04","Single dash → Missing",{Cd:{rawPpm:"-",rawSigma:4,rawJudgement:""}}],
    ["VU-X05","Blank Content → Missing",{Pb:{rawPpm:"",rawSigma:5,rawJudgement:""}}],
    ["VU-X06","Judgment ?? → Remeasure",{Pb:{rawPpm:45,rawSigma:6,rawJudgement:"??"}}],
    ["VU-X07","Cannot Judge → Remeasure",{Hg:{rawPpm:100,rawSigma:7,rawJudgement:"Cannot Judge"}}],
    ["VU-X08","Cd 69.9 → M / OK",{Cd:{rawPpm:69.9,rawSigma:8,rawJudgement:"OK"}}],
    ["VU-X09","Cd 70.0 boundary → M / OK",{Cd:{rawPpm:70,rawSigma:9,rawJudgement:"OK"}}],
    ["VU-X10","Cd 70.1 → H / NG",{Cd:{rawPpm:70.1,rawSigma:10,rawJudgement:"NG"}}],
    ["VU-X11","Pb 700 boundary → M / OK",{Pb:{rawPpm:700,rawSigma:11,rawJudgement:"OK"}}],
    ["VU-X12","Pb 700.1 → H / NG",{Pb:{rawPpm:700.1,rawSigma:12,rawJudgement:"NG"}}],
    ["VU-X13","Cl+Br = 1050 boundary",{Cl:{rawPpm:630,rawSigma:13,rawJudgement:"OK"},Br:{rawPpm:420,rawSigma:14,rawJudgement:"OK"}}],
    ["VU-X14","Cl+Br = 1050.1 → H",{Cl:{rawPpm:630,rawSigma:13,rawJudgement:"OK"},Br:{rawPpm:420.1,rawSigma:14,rawJudgement:"OK"}}],
    ["VU-X15","NG + ?? mixed precedence",{Pb:{rawPpm:45,rawSigma:15,rawJudgement:"??"},Br:{rawPpm:890,rawSigma:16,rawJudgement:"NG"}}],
    ["VU-X16","Zero + Judgment ND",{Pb:{rawPpm:0,rawSigma:16,rawJudgement:"ND"}}],
    ["VU-X17","Zero + Judgment OK",{Pb:{rawPpm:0,rawSigma:17,rawJudgement:"OK"}}]
  ];
  xrfCases.forEach(([code,label,elements],index)=>push(code,`[UAT-VIS][XRF] ${label}`,{
    measurements:[visualUatInitial(code,xrfDate,"L",{elements})],
    crType:VISUAL_UAT_CR_TYPES[index%VISUAL_UAT_CR_TYPES.length]
  }));

  const riskCases=[
    ["VU-R01","H","직접접촉+잔류","Not Allowed"],
    ["VU-R02","H","직접접촉+비잔류","Not Allowed"],
    ["VU-R03","H","비접촉","Not Allowed"],
    ["VU-R04","M","직접접촉+잔류","H"],
    ["VU-R05","M","직접접촉+비잔류","M"],
    ["VU-R06","M","비접촉","M"],
    ["VU-R07","L","직접접촉+잔류","M"],
    ["VU-R08","L","직접접촉+비잔류","M"],
    ["VU-R09","L","비접촉","L"]
  ];
  riskCases.forEach(([code,xrfLevel,crType,risk])=>push(code,`[UAT-VIS][RISK] XRF ${xrfLevel} + CR ${crLevelFromType(crType)} → ${risk}`,{
    measurements:[visualUatInitial(code,xrfDate,xrfLevel)],crType
  }));

  push("VU-C01","[UAT-VIS][CYCLE] Risk H → Monthly",{measurements:[visualUatInitial("VU-C01",xrfDate,"M")],crType:"직접접촉+잔류"});
  push("VU-C02","[UAT-VIS][CYCLE] Risk M → Semiannual",{measurements:[visualUatInitial("VU-C02",xrfDate,"M")],crType:"직접접촉+비잔류"});
  push("VU-C03","[UAT-VIS][CYCLE] Risk L → Annual",{measurements:[visualUatInitial("VU-C03",xrfDate,"L")],crType:"비접촉"});
  push("VU-C04","[UAT-VIS][CYCLE] Not Allowed → No Cycle",{measurements:[visualUatInitial("VU-C04",xrfDate,"H")],crType:"직접접촉+잔류"});
  push("VU-C05","[UAT-VIS][CYCLE] Unresolved → No Cycle",{measurements:[visualUatInitial("VU-C05",xrfDate,"L",{elements:{Pb:{rawPpm:"-",rawSigma:2,rawJudgement:""}}})],crType:"비접촉"});
  push("VU-C06","[UAT-VIS][CYCLE] Facility → No Cycle",{measurements:[visualUatInitial("VU-C06",xrfDate,"M")],crType:"직접접촉+비잔류",type:"Facility"});

  const p1DueSoon=addDaysISO(TODAY_STR,3);
  const p1Overdue=addDaysISO(TODAY_STR,-3);
  const p1CompletedDue=addDaysISO(TODAY_STR,3);
  const p1LateDue=addDaysISO(TODAY_STR,-10);
  const p1FutureDue=addDaysISO(TODAY_STR,20);
  const p1SequenceDue=addDaysISO(TODAY_STR,-10);
  const monthlyItem=(code,name,due,periodic=[])=>push(code,name,{
    measurements:[visualUatInitial(code,addMonthsISO(due,-1),"M"),...periodic],crType:"직접접촉+잔류"
  });
  monthlyItem("VU-P01","[UAT-VIS][PERIOD] 측정권장",p1DueSoon);
  monthlyItem("VU-P02","[UAT-VIS][PERIOD] 미이행",p1Overdue);
  monthlyItem("VU-P03","[UAT-VIS][PERIOD] 이행",p1CompletedDue,[visualUatPeriodic("VU-P03",TODAY_STR,"M")]);
  monthlyItem("VU-P04","[UAT-VIS][PERIOD] 지연측정",p1LateDue,[visualUatPeriodic("VU-P04",addDaysISO(p1LateDue,2),"M")]);
  monthlyItem("VU-P05","[UAT-VIS][PERIOD] 예정",p1FutureDue);
  monthlyItem("VU-P06","[UAT-VIS][PERIOD] R + P1 + P2",p1SequenceDue,[visualUatPeriodic("VU-P06",addDaysISO(p1SequenceDue,-2),"M")]);
  const currentYear=TODAY.getFullYear();
  const annualR=`${currentYear-1}-01-31`;
  push("VU-P07","[UAT-VIS][PERIOD] Annual same-year multiple",{
    measurements:[
      visualUatInitial("VU-P07",annualR,"L"),
      visualUatPeriodic("VU-P07",`${currentYear}-01-15`,"L",{suffix:"P1A"}),
      visualUatPeriodic("VU-P07",`${currentYear}-07-15`,"L",{suffix:"P1B"})
    ],crType:"비접촉"
  });

  const workflowR=addMonthsISO(TODAY_STR,-2);
  push("VU-W01","[UAT-VIS][WF] 신규 · XRF 미측정",{measurements:[],category:"new",lifecycle:"NewItem"});
  push("VU-W02","[UAT-VIS][WF] Initial OK → Approved",{measurements:[visualUatInitial("VU-W02",workflowR,"M")],crType:"비접촉"});
  push("VU-W03","[UAT-VIS][WF] Initial ?? → Retest required",{measurements:[visualUatInitial("VU-W03",workflowR,"L",{elements:{Pb:{rawPpm:45,rawSigma:3,rawJudgement:"??"}}})]});
  const w04Initial=visualUatInitial("VU-W04",workflowR,"L",{elements:{Pb:{rawPpm:45,rawSigma:3,rawJudgement:"??"}}});
  const w04Retest=makeVisualUatMeasurement("VU-W04_RET2",addDaysISO(workflowR,2),{role:"Retest",attemptNo:2,parentMeasurementId:w04Initial.id,elements:{Pb:{rawPpm:45,rawSigma:3,rawJudgement:"OK"}}});
  push("VU-W04","[UAT-VIS][WF] ?? → Retest OK",{measurements:[w04Initial,w04Retest]});
  const w05Initial=visualUatInitial("VU-W05",workflowR,"L",{elements:{Pb:{rawPpm:45,rawSigma:3,rawJudgement:"??"}}});
  const w05Retest=makeVisualUatMeasurement("VU-W05_RET2",addDaysISO(workflowR,2),{role:"Retest",attemptNo:2,parentMeasurementId:w05Initial.id,elements:{Pb:{rawPpm:45,rawSigma:3,rawJudgement:"??"}}});
  push("VU-W05","[UAT-VIS][WF] ?? → Retest ??",{measurements:[w05Initial,w05Retest]});
  push("VU-W06","[UAT-VIS][WF] Initial NG → Precision required",{measurements:[visualUatInitial("VU-W06",workflowR,"L",{elements:{Br:{rawPpm:701,rawSigma:6,rawJudgement:"NG"}}})]});
  push("VU-W07","[UAT-VIS][WF] Precision requested → Waiting",{measurements:[visualUatInitial("VU-W07",workflowR,"L",{elements:{Br:{rawPpm:701,rawSigma:7,rawJudgement:"NG"}}})]});
  const safeR=(code)=>visualUatInitial(code,addMonthsISO(workflowR,-6),"M");
  const ngPeriodic=(code)=>visualUatPeriodic(code,workflowR,"L",{elements:{Br:{rawPpm:701,rawSigma:8,rawJudgement:"NG"}}});
  push("VU-W08","[UAT-VIS][WF] Precision OK → Approved",{measurements:[safeR("VU-W08"),ngPeriodic("VU-W08")],crType:"직접접촉+잔류"});
  push("VU-W09","[UAT-VIS][WF] Precision NG → Rejected",{measurements:[safeR("VU-W09"),ngPeriodic("VU-W09")],crType:"직접접촉+잔류"});
  push("VU-W10","[UAT-VIS][POLICY] Not Allowed + Precision OK",{measurements:[visualUatInitial("VU-W10",workflowR,"H")],crType:"직접접촉+잔류"});

  push("VU-L01","[UAT-VIS][LIFE] NewItem",{measurements:[],category:"new",lifecycle:"NewItem",isCurrent:true});
  push("VU-L03","[UAT-VIS][LIFE] ReplacedOld",{measurements:[visualUatInitial("VU-L03",xrfDate,"M")],lifecycle:"ExistingActive",isCurrent:true,replacedBy:"VU-L02"});
  push("VU-L02","[UAT-VIS][LIFE] NewReplacement",{measurements:[],category:"new",lifecycle:"NewReplacement",isCurrent:true,replacementOf:"VU-L03",extra:{firstDate:addDaysISO(xrfDate,1),replacementDate:addDaysISO(xrfDate,1)}});
  push("VU-L04","[UAT-VIS][LIFE] Discontinue Pending",{measurements:[],lifecycle:"ExistingActive",isCurrent:false,extra:{
    _isRequestRecord:true,requestId:"VU-L04-REQUEST",requestType:"Discontinue",requestStatus:"Pending",
    approvalStatus:"단종 검토중",processStatus:"Pending",discontinueTargetCode:"VU-L05",
    discontinueReason:"Visual UAT only",finalUseDate:addDaysISO(TODAY_STR,30)
  }});
  push("VU-L05","[UAT-VIS][LIFE] Discontinued",{measurements:[visualUatInitial("VU-L05",xrfDate,"M")],lifecycle:"Discontinued",isCurrent:false,extra:{processedAt:addDaysISO(TODAY_STR,-1),processedBy:"admin",finalUseDate:addDaysISO(TODAY_STR,-1)}});
  push("VU-L06","[UAT-VIS][LIFE] Discontinue Reverted",{measurements:[visualUatInitial("VU-L06",xrfDate,"M")],lifecycle:"ExistingActive",isCurrent:true,extra:{
    processStatus:"Reverted",approvalStatus:"처리 취소",restoredFromDiscontinueRequestCode:"VU-L06-REQUEST",
    restoredAt:TODAY_STR,restoredBy:"admin",restoreReason:"Visual UAT 원복 상태 확인"
  }});
  return fixtures;
}

const VISUAL_UAT_ITEMS=Object.freeze(buildVisualUatFixtures());

// Compliance Visual UAT fixtures are kept separate from the original 55-case pack.
// Only raw dates and measurements are supplied; production period functions derive every status.
function buildVisualUatComplianceFixtures(){
  const fixtures=[];
  const push=(code,label,options)=>fixtures.push(makeVisualUatItem(code,`[UAT-VIS][CMP] ${label}`,options));
  const monthly=(code,label,p1Due,periodic=[])=>push(code,label,{
    measurements:[visualUatInitial(code,addMonthsISO(p1Due,-1),"M"),...periodic],
    crType:"직접접촉+잔류"
  });
  const semiannual=(code,label,p1Due)=>push(code,label,{
    measurements:[visualUatInitial(code,addMonthsISO(p1Due,-6),"M")],
    crType:"직접접촉+비잔류"
  });
  const annual=(code,label,p1Due,periodic=[])=>push(code,label,{
    measurements:[visualUatInitial(code,addMonthsISO(p1Due,-12),"L"),...periodic],
    crType:"비접촉"
  });

  monthly("VU-CMP-L01","R 기준",addMonthsISO(TODAY_STR,1));
  monthly("VU-CMP-L02","측정권장 · Monthly D-15 경계",addDaysISO(TODAY_STR,15));
  monthly("VU-CMP-L03","측정권장 · Monthly D-7 경계",addDaysISO(TODAY_STR,7));
  monthly("VU-CMP-L04","측정권장 · Monthly D-1",addDaysISO(TODAY_STR,1));
  monthly("VU-CMP-L05","측정권장 · D-Day",TODAY_STR);
  monthly("VU-CMP-L06","미이행 · D+1",addDaysISO(TODAY_STR,-1));
  monthly("VU-CMP-L07","미이행 · D+10",addDaysISO(TODAY_STR,-10));

  const l08Due=addDaysISO(TODAY_STR,3);
  monthly("VU-CMP-L08","정상 이행",l08Due,[visualUatPeriodic("VU-CMP-L08",TODAY_STR,"M")]);
  monthly("VU-CMP-L09","도래일 당일 이행",TODAY_STR,[visualUatPeriodic("VU-CMP-L09",TODAY_STR,"M")]);
  const l10Due=addDaysISO(TODAY_STR,20);
  monthly("VU-CMP-L10","매우 이른 측정",l10Due,[visualUatPeriodic("VU-CMP-L10",TODAY_STR,"M")]);
  const l11Due=addDaysISO(TODAY_STR,-10);
  monthly("VU-CMP-L11","지연측정 · D+5에 측정",l11Due,[visualUatPeriodic("VU-CMP-L11",addDaysISO(l11Due,5),"M")]);
  monthly("VU-CMP-L12","Monthly D-20 Compact 경계검증",addDaysISO(TODAY_STR,20));

  semiannual("VU-CMP-L13","Semiannual D-30 경계",addDaysISO(TODAY_STR,30));
  semiannual("VU-CMP-L14","Semiannual D-31",addDaysISO(TODAY_STR,31));
  annual("VU-CMP-L15","Annual D-30 경계",addDaysISO(TODAY_STR,30));
  annual("VU-CMP-L16","Annual D-31",addDaysISO(TODAY_STR,31));

  const l17P4Due=addDaysISO(TODAY_STR,15);
  const l17RDate=addMonthsISO(l17P4Due,-4);
  const l17P1Due=addMonthsISO(l17RDate,1);
  const l17P2Due=addMonthsISO(l17RDate,2);
  const l17P3Due=addMonthsISO(l17RDate,3);
  push("VU-CMP-L17","4-slot Rolling",{
    measurements:[
      visualUatInitial("VU-CMP-L17",l17RDate,"M"),
      visualUatPeriodic("VU-CMP-L17",addDaysISO(l17P1Due,-2),"M",{suffix:"P1",periodNo:1}),
      visualUatPeriodic("VU-CMP-L17",addDaysISO(l17P2Due,-2),"M",{suffix:"P2",periodNo:2}),
      visualUatPeriodic("VU-CMP-L17",addDaysISO(l17P3Due,-2),"M",{suffix:"P3",periodNo:3})
    ],
    crType:"직접접촉+잔류"
  });

  const currentYear=TODAY.getFullYear();
  const l18Due=`${currentYear}-12-15`;
  annual("VU-CMP-L18","Annual same-year multiple",l18Due,[
    visualUatPeriodic("VU-CMP-L18",`${currentYear}-03-15`,"L",{suffix:"P1A",periodNo:1}),
    visualUatPeriodic("VU-CMP-L18",`${currentYear}-09-01`,"L",{suffix:"P1B",periodNo:1})
  ]);

  push("VU-CMP-L19","Facility · 대상 아님",{
    measurements:[visualUatInitial("VU-CMP-L19",addMonthsISO(TODAY_STR,-1),"M")],
    crType:"직접접촉+비잔류",type:"Facility"
  });
  push("VU-CMP-L20","Supplemental Measured F1",{
    measurements:[
      visualUatInitial("VU-CMP-L20",addMonthsISO(TODAY_STR,-2),"H"),
      visualUatPeriodic("VU-CMP-L20",addMonthsISO(TODAY_STR,-1),"L",{suffix:"F1",periodNo:1})
    ],
    crType:"직접접촉+잔류"
  });

  return fixtures;
}

const VISUAL_UAT_COMPLIANCE_ITEMS=Object.freeze(buildVisualUatComplianceFixtures());

// 부자재 리스트의 100건 단위 페이지 전환을 실제 화면에서 검증하기 위한 읽기 전용 데이터입니다.
// 10번째 항목마다 Initial Pb Judgment를 ??로 두어 1·2페이지 모두에 XRF 재측정 항목이 나타납니다.
function buildVisualUatPaginationFixtures(){
  const referenceDate=addMonthsISO(TODAY_STR,-2);
  return Array.from({length:40},(_,index)=>{
    const number=index+1;
    const code=`VU-PG-${String(number).padStart(3,"0")}`;
    const remeasureRequired=number%10===0;
    const elements=remeasureRequired
      ? {Pb:{rawPpm:45,rawSigma:3,rawJudgement:"??"}}
      : {};
    return makeVisualUatItem(
      code,
      `[UAT-VIS][PAGE] ${String(number).padStart(3,"0")}${remeasureRequired?" · Pb 재측정":" · 정상"}`,
      {
        measurements:[visualUatInitial(code,referenceDate,"L",{elements})],
        crType:"비접촉",
        extra:{dept:VISUAL_UAT_DEPTS[index%VISUAL_UAT_DEPTS.length]}
      }
    );
  });
}

const VISUAL_UAT_PAGINATION_ITEMS=Object.freeze(buildVisualUatPaginationFixtures());
const VISUAL_UAT_ALL_ITEMS=Object.freeze([
  ...VISUAL_UAT_ITEMS,
  ...VISUAL_UAT_COMPLIANCE_ITEMS,
  ...VISUAL_UAT_PAGINATION_ITEMS
]);

const VISUAL_UAT_COMPLIANCE_EXPECTATIONS=Object.freeze([
  {code:"VU-CMP-L01",expected:{cycle:"Monthly",referenceStatus:"reference",referenceMeasured:true}},
  {code:"VU-CMP-L02",expected:{periodStatus:"due_soon",currentStatus:"dueSoon",dDay:15}},
  {code:"VU-CMP-L03",expected:{periodStatus:"due_soon",currentStatus:"dueSoon",dDay:7}},
  {code:"VU-CMP-L04",expected:{periodStatus:"due_soon",currentStatus:"dueSoon",dDay:1}},
  {code:"VU-CMP-L05",expected:{periodStatus:"due_soon",currentStatus:"dueSoon",dDay:0}},
  {code:"VU-CMP-L06",expected:{periodStatus:"overdue",currentStatus:"overdue",dDay:-1,daysOver:1}},
  {code:"VU-CMP-L07",expected:{periodStatus:"overdue",currentStatus:"overdue",dDay:-10,daysOver:10}},
  {code:"VU-CMP-L08",expected:{periodStatus:"compliant",currentStatus:"ok",p1MeasurementCount:1}},
  {code:"VU-CMP-L09",expected:{periodStatus:"compliant",currentStatus:"ok",p1MeasurementCount:1}},
  {code:"VU-CMP-L10",expected:{periodStatus:"compliant",currentStatus:"ok",p1MeasurementCount:1}},
  {code:"VU-CMP-L11",expected:{periodStatus:"late",currentStatus:"ok",daysOver:5,p1MeasurementCount:1}},
  {code:"VU-CMP-L12",expected:{periodStatus:"upcoming",currentStatus:"upcoming",dDay:20}},
  {code:"VU-CMP-L13",expected:{cycle:"Semiannual",periodStatus:"due_soon",currentStatus:"dueSoon",dDay:30}},
  {code:"VU-CMP-L14",expected:{cycle:"Semiannual",periodStatus:"upcoming",currentStatus:"upcoming",dDay:31}},
  {code:"VU-CMP-L15",expected:{cycle:"Annual",periodStatus:"due_soon",currentStatus:"dueSoon",dDay:30}},
  {code:"VU-CMP-L16",expected:{cycle:"Annual",periodStatus:"upcoming",currentStatus:"upcoming",dDay:31}},
  {code:"VU-CMP-L17",expected:{cycle:"Monthly",compactLabels:["P3","P4","P5","P6"],compactCount:4}},
  {code:"VU-CMP-L18",expected:{cycle:"Annual",p1MeasurementCount:2}},
  {code:"VU-CMP-L19",expected:{cycle:"Not Available",currentStatus:"notAvailable",regularPeriodCount:0,allLabels:["R"]}},
  {code:"VU-CMP-L20",expected:{cycle:"Not Available",currentStatus:"notAvailable",displayLabels:["R","F1"],supplementalStatus:"measured"}}
]);

function visualUatComplianceSelfCheckRow(item,manifest){
  const periods=itemAllPeriods(item);
  const displayPeriods=itemXrfDisplayPeriods(item);
  const compactPeriods=itemListPeriods(item);
  const reference=periods.find(period=>period?.isReference||Number(period?.num)===0)||null;
  const p1=periods.find(period=>Number(period?.num)===1)||null;
  const supplemental=displayPeriods.find(period=>period?.isSupplementalMeasurement)||null;
  const dueDays=p1?.due?diffDaysISO(p1.due,TODAY_STR):null;
  const actual={
    cycle:itemCycleFromRisk(item),
    p1Due:p1?.due||null,
    p1Measured:p1?.measured||null,
    periodStatus:p1?normStatus(p1.status):null,
    currentStatus:getCS(item),
    dDay:dueDays,
    daysOver:Number(p1?.daysOver||0),
    referenceStatus:reference?normStatus(reference.status):null,
    referenceMeasured:!!reference?.measured,
    p1MeasurementCount:periodMeasurementIds(p1).length,
    compactLabels:compactPeriods.map(period=>periodLabel(period)),
    compactCount:compactPeriods.length,
    allLabels:periods.map(period=>periodLabel(period)),
    regularPeriodCount:periods.filter(period=>!(period?.isReference||Number(period?.num)===0)).length,
    displayLabels:displayPeriods.map(period=>periodLabel(period)),
    supplementalStatus:supplemental?normStatus(supplemental.status):null,
    compactPotential30DayOverride:false,
  };
  const failedFields=[];
  Object.entries(manifest.expected||{}).forEach(([key,expected])=>{
    const value=actual[key];
    if(Array.isArray(expected)){
      if(JSON.stringify(value)!==JSON.stringify(expected)) failedFields.push(key);
    }else if(value!==expected) failedFields.push(key);
  });
  return {
    code:item.code,cycle:actual.cycle,p1Due:actual.p1Due,p1Measured:actual.p1Measured,
    expectedPeriodStatus:manifest.expected?.periodStatus??"",actualPeriodStatus:actual.periodStatus,
    expectedCurrentStatus:manifest.expected?.currentStatus??"",actualCurrentStatus:actual.currentStatus,
    compactLabels:actual.compactLabels.join(", "),compactCount:actual.compactCount,
    businessStatus:actual.periodStatus,
    compactVisualExpected:actual.periodStatus,
    compactVisualActual:actual.periodStatus,
    compactPotential30DayOverride:actual.compactPotential30DayOverride,
    PASS:failedFields.length===0,failedFields,expected:manifest.expected,actual
  };
}

const VISUAL_UAT_COMPLIANCE_RENDER_PERIODS=Object.freeze([
  {num:0,displayLabel:"R",isReference:true,referenceDate:TODAY_STR,measured:TODAY_STR,measuredDates:[TODAY_STR],status:"reference"},
  {num:1,displayLabel:"P1",due:addDaysISO(TODAY_STR,-30),measured:addDaysISO(TODAY_STR,-35),measuredDates:[addDaysISO(TODAY_STR,-35)],status:"compliant"},
  {num:2,displayLabel:"F1",measured:addDaysISO(TODAY_STR,-25),measuredDates:[addDaysISO(TODAY_STR,-25)],status:"measured",isSupplementalMeasurement:true},
  {num:3,displayLabel:"P3",due:addDaysISO(TODAY_STR,20),measured:TODAY_STR,measuredDates:[TODAY_STR],status:"early"},
  {num:4,displayLabel:"P4",due:addDaysISO(TODAY_STR,-10),measured:addDaysISO(TODAY_STR,-5),measuredDates:[addDaysISO(TODAY_STR,-5)],status:"late",daysOver:5},
  {num:5,displayLabel:"P5",due:addDaysISO(TODAY_STR,-5),measured:null,status:"overdue",daysOver:5},
  {num:6,displayLabel:"P6",due:addDaysISO(TODAY_STR,3),measured:null,status:"due_soon"},
  {num:7,displayLabel:"P7",due:addDaysISO(TODAY_STR,40),measured:null,status:"upcoming"}
]);

const VISUAL_UAT_PRECISION_OVERRIDES=Object.freeze({
  "VU-W07_R":{requestStatus:"REQUESTED",requestedAt:TODAY_STR,requestedBy:"visual-uat",uploadStatus:"WAITING"},
  "VU-W08_P1":{requestStatus:"REQUESTED",uploadStatus:"UPLOADED",resultFileId:"VU-W08-precision.pdf",finalConfirm:"확인",elementResults:{Br:{presence:"함유",ppm:100,legalLimit:900}}},
  "VU-W09_P1":{requestStatus:"REQUESTED",uploadStatus:"UPLOADED",resultFileId:"VU-W09-precision.pdf",finalConfirm:"확인",elementResults:{Br:{presence:"함유",ppm:1000,legalLimit:900}}},
  "VU-W10_R":{requestStatus:"REQUESTED",uploadStatus:"UPLOADED",resultFileId:"VU-W10-precision.pdf",finalConfirm:"확인",elementResults:{Pb:{presence:"함유",ppm:600,legalLimit:1000}}}
});

const VISUAL_UAT_EXPECTATIONS=Object.freeze([
  ...[
    ["VU-X01","OK","L"],["VU-X02","OK","L"],["VU-X03","OK","L"],
    ["VU-X04","확인 필요",null],["VU-X05","확인 필요",null],
    ["VU-X06","재측정 필요",null],["VU-X07","재측정 필요",null],
    ["VU-X08","OK","M"],["VU-X09","OK","M"],["VU-X10","NG","H"],
    ["VU-X11","OK","M"],["VU-X12","NG","H"],["VU-X13","OK","M"],
    ["VU-X14","NG","H"],["VU-X15","NG",null],["VU-X16","OK","L"],["VU-X17","OK","M"]
  ].map(([code,xrfResult,xrfLevel])=>({code,group:"XRF",expected:{xrfResult,xrfLevel}})),
  ...[
    ["VU-R01","Not Allowed"],["VU-R02","Not Allowed"],["VU-R03","Not Allowed"],
    ["VU-R04","H"],["VU-R05","M"],["VU-R06","M"],
    ["VU-R07","M"],["VU-R08","M"],["VU-R09","L"]
  ].map(([code,finalRisk])=>({code,group:"Risk",expected:{finalRisk}})),
  ...[
    ["VU-C01","H","Monthly",1,15],["VU-C02","M","Semiannual",6,30],
    ["VU-C03","L","Annual",12,30],["VU-C04","Not Allowed","Not Available",null,null],
    ["VU-C05","—","Not Available",null,null],["VU-C06","M","Not Available",null,null]
  ].map(([code,finalRisk,cycle,cycleMonths,recommendDays])=>({code,group:"Cycle",expected:{finalRisk,cycle,cycleMonths,recommendDays}})),
  {code:"VU-P01",group:"Period",expected:{periodStatus:"due_soon",dDay:3}},
  {code:"VU-P02",group:"Period",expected:{periodStatus:"overdue",dDay:-3}},
  {code:"VU-P03",group:"Period",expected:{periodStatus:"compliant"}},
  {code:"VU-P04",group:"Period",expected:{periodStatus:"late"}},
  {code:"VU-P05",group:"Period",expected:{periodStatus:"upcoming"}},
  {code:"VU-P06",group:"Period",expected:{periodLabelsPrefix:["R","P1","P2"]}},
  {code:"VU-P07",group:"Period",expected:{p1MeasurementCount:2}},
  {code:"VU-W01",group:"Workflow",expected:{xrfResult:"—",approval:"보류",stage:"XRF_REQUIRED"}},
  {code:"VU-W02",group:"Workflow",expected:{xrfResult:"OK",approval:"승인",stage:"APPROVED"}},
  {code:"VU-W03",group:"Workflow",expected:{xrfResult:"재측정 필요",approval:"보류",stage:"XRF_REMEASURE_REQUIRED"}},
  {code:"VU-W04",group:"Workflow",expected:{xrfResult:"OK",approval:"승인",stage:"APPROVED"}},
  {code:"VU-W05",group:"Workflow",expected:{xrfResult:"재측정 필요",approval:"의뢰 필요",stage:"PRECISION_REQUEST_REQUIRED"}},
  {code:"VU-W06",group:"Workflow",expected:{xrfResult:"NG",approval:"의뢰 필요",stage:"PRECISION_REQUEST_REQUIRED"}},
  {code:"VU-W07",group:"Workflow",expected:{xrfResult:"NG",approval:"측정 진행중",stage:"PRECISION_RESULT_WAITING"}},
  {code:"VU-W08",group:"Workflow",expected:{xrfResult:"NG",approval:"승인",stage:"APPROVED"}},
  {code:"VU-W09",group:"Workflow",expected:{xrfResult:"NG",approval:"반려",stage:"PRECISION_NG_REVIEW"}},
  {code:"VU-W10",group:"Workflow",expected:{finalRisk:"Not Allowed",approval:"반려",stage:"RISK_NOT_ALLOWED"}},
  {code:"VU-L01",group:"Lifecycle",expected:{lifecycle:"NewItem",isCurrent:true}},
  {code:"VU-L02",group:"Lifecycle",expected:{lifecycle:"NewReplacement",isCurrent:true,replacementOf:"VU-L03"}},
  {code:"VU-L03",group:"Lifecycle",expected:{lifecycle:"ReplacedOld",isCurrent:false,replacedByIncludes:"VU-L02"}},
  {code:"VU-L04",group:"Lifecycle",expected:{requestStatus:"Pending",approval:"단종 검토중"}},
  {code:"VU-L05",group:"Lifecycle",expected:{lifecycle:"Discontinued",isCurrent:false}},
  {code:"VU-L06",group:"Lifecycle",expected:{lifecycle:"ExistingActive",isCurrent:true,processStatus:"Reverted"}}
]);

function visualUatSelfCheckRow(item,manifest,approvalStatusOf){
  const periods=itemAllPeriods(item);
  const p1=periods.find(period=>Number(period?.num)===1)||null;
  const workflow=workflowForItem(item);
  const actual={
    xrfResult:itemXrfResult(item),
    xrfLevel:riskBasisXrfLevelOf(item)||itemLatestXrfLevel(item),
    finalRisk:itemFinalRisk(item),
    cycle:itemCycleFromRisk(item),
    cycleMonths:cycleMonthsOf(item)||null,
    recommendDays:itemCycleFromRisk(item)==="Not Available"?null:windowDaysOf(item),
    approval:workflow?.approval?.status||approvalStatusOf(item),stage:workflow?.stage||null,
    periodStatus:normStatus(p1?.status),
    dDay:p1?.due?diffDaysISO(p1.due,TODAY_STR):null,
    periodLabels:periods.map(period=>periodLabel(period)),
    p1MeasurementCount:periodMeasurementIds(p1).length,
    lifecycle:item.lifecycle,isCurrent:item.isCurrent,replacementOf:item.replacementOf||null,
    replacedBy:String(item.replacedBy||""),requestStatus:item.requestStatus||null,
    processStatus:item.processStatus||null
  };
  const failedFields=[];
  Object.entries(manifest.expected||{}).forEach(([key,expected])=>{
    if(key==="periodLabelsPrefix"){
      if(!expected.every((label,index)=>actual.periodLabels[index]===label)) failedFields.push(key);
    }else if(key==="replacedByIncludes"){
      if(!String(actual.replacedBy).split(",").map(v=>v.trim()).includes(expected)) failedFields.push(key);
    }else if(actual[key]!==expected){
      failedFields.push(key);
    }
  });
  return {
    code:item.code,group:manifest.group,
    expectedXrf:manifest.expected?.xrfResult??"",actualXrf:actual.xrfResult,
    expectedRisk:manifest.expected?.finalRisk??"",actualRisk:actual.finalRisk,
    expectedCycle:manifest.expected?.cycle??"",actualCycle:actual.cycle,
    expectedApproval:manifest.expected?.approval??"",actualApproval:actual.approval,
    expectedStage:manifest.expected?.stage??"",actualStage:actual.stage,
    PASS:failedFields.length===0,failedFields,expected:manifest.expected,actual
  };
}

function FilterPanel({column, pos, selected, values, onToggle, onClear, onClose}){
  const safeValues = values||UNIQUE_VALS[column]||[];
  const labels = FILTER_LABELS[column]||{};
  const allSel = selected.length===0;
  const handleAll = ()=>{
    if(allSel){ safeValues.forEach(v=>onToggle(column,v)); }
    else { onClear(column); }
  };
  return(
    <div data-filter-panel="1" style={{position:"fixed",top:pos.top,left:pos.left,zIndex:9999,
      background:C.card,border:`1px solid ${C.bd2}`,borderRadius:4,
      boxShadow:"0 6px 20px rgba(0,0,0,.18)",minWidth:210,maxWidth:280}}>
      <div style={{padding:"9px 12px",background:C.charcoalDk,color:"#fff",fontSize:12,fontWeight:600,borderRadius:"4px 4px 0 0"}}>
        {COL_NAMES[column]} 필터
      </div>
      <div onClick={handleAll} style={{padding:"8px 12px",borderBottom:`1px solid ${C.bd}`,cursor:"pointer",display:"flex",alignItems:"center",gap:8,background:C.alt}}>
        <div style={{width:14,height:14,borderRadius:2,border:`1.5px solid ${C.charcoal}`,background:allSel?C.charcoal:C.card,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
          {allSel&&<span style={{fontSize:8,color:"#fff",fontWeight:700}}>●</span>}
          {!allSel&&selected.length>0&&<span style={{fontSize:9,color:C.charcoalMd}}>—</span>}
        </div>
        <span style={{fontSize:12,fontWeight:500,color:C.text2}}>전체 선택</span>
        {!allSel&&<span style={{fontSize:10,color:C.text4,marginLeft:"auto"}}>{selected.length}/{safeValues.length}</span>}
      </div>
      <div style={{maxHeight:220,overflowY:"auto"}}>
        {safeValues.map(v=>{
          const isSel=selected.includes(v);
          return(
            <div key={v} onClick={()=>onToggle(column,v)}
              style={{padding:"7px 12px",cursor:"pointer",display:"flex",alignItems:"center",gap:8,
                background:isSel?C.charcoalBg:C.card,borderBottom:`1px solid ${C.bd}`}}>
              <div style={{width:14,height:14,borderRadius:2,border:`1.5px solid ${isSel?C.charcoal:C.bd2}`,
                background:isSel?C.charcoal:C.card,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>
                {isSel&&<span style={{fontSize:8,color:"#fff",fontWeight:700}}>●</span>}
              </div>
              <span style={{fontSize:12,color:isSel?C.text1:C.text3,overflowWrap:"anywhere",wordBreak:"keep-all"}}>{labels[v]||v}</span>
            </div>
          );
        })}
      </div>
      <div style={{padding:"8px 10px",background:C.bg,borderTop:`1px solid ${C.bd}`,display:"flex",gap:6,borderRadius:"0 0 4px 4px"}}>
        <button onClick={onClose} style={{flex:1,padding:"7px 0",background:C.charcoalDk,color:"#fff",border:"none",borderRadius:6,fontSize:11,fontWeight:600,cursor:"pointer"}}>적용</button>
        <button onClick={()=>onClear(column)} className="pill-btn" style={{padding:"7px 13px",background:C.card,border:`1px solid ${C.bd2}`,fontSize:11,color:C.text3}}>초기화</button>
      </div>
    </div>
  );
}
function FilterChips({filterState, onClear, onClearAll}){
  const active=Object.entries(filterState).filter(([,v])=>v.length>0);
  if(!active.length) return null;
  return(
    <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap",padding:"6px 0",marginBottom:4}}>
      <span style={{fontSize:11,color:C.text4,flexShrink:0}}>필터:</span>
      {active.map(([col,vals])=>{
        const lbs=FILTER_LABELS[col]||{};
        const isRed=col==="compliance"&&vals.includes("overdue");
        return(
          <div key={col} onClick={()=>onClear(col)}
            style={{display:"flex",alignItems:"center",gap:5,padding:"4px 11px",
              background:isRed?A.rose.bg:A.purple.bg,
              border:`1px solid ${isRed?A.rose.ln:A.purple.ln}`,
              borderRadius:UI.pill,fontSize:11,cursor:"pointer",color:isRed?A.rose.tx:A.purple.tx}}>
            <span style={{fontWeight:500}}>{COL_NAMES[col]}: {vals.map(v=>lbs[v]||v).join(", ")}</span>
            <span style={{fontWeight:700,opacity:.6}}>×</span>
          </div>
        );
      })}
      <button onClick={onClearAll} className="pill-btn" style={{padding:"4px 12px",background:C.card,border:`1px solid ${C.bd2}`,fontSize:11,color:C.text3}}>전체 초기화</button>
    </div>
  );
}

function FH({label, col, filterState, openFilter, onOpen}){
  const isAct=filterState[col]?.length>0;
  const cnt=filterState[col]?.length||0;
  return(
    <div data-filter-header={col} onClick={e=>onOpen(col,e)}
      style={{position:"relative",display:"inline-grid",gridTemplateColumns:"minmax(0,auto) auto",alignItems:"center",justifyContent:"center",columnGap:4,cursor:"pointer",userSelect:"none",width:"100%",minWidth:0}}>
      <span style={{minWidth:0,color:isAct?C.text1:C.charcoalLt,fontWeight:isAct?700:600,fontSize:10,letterSpacing:.5,textTransform:"uppercase",whiteSpace:"normal",lineHeight:1.15,textAlign:"center",overflowWrap:"normal"}}>{label}</span>
      <span style={{fontSize:9,color:isAct?C.red:C.charcoalLt,fontWeight:700,lineHeight:1,alignSelf:"center",flexShrink:0}}>{isAct?`▼${cnt}`:"▼"}</span>
      {isAct&&col!=="category"&&<span style={{width:4,height:4,borderRadius:"50%",background:C.red,flexShrink:0,position:"absolute",right:2}}/>}
    </div>
  );
}
function PlainH({label, align="center"}){
  return <span style={{display:"block",width:"100%",textAlign:align,color:C.charcoalLt,fontWeight:600,fontSize:10,letterSpacing:.5,textTransform:"uppercase",whiteSpace:"normal",lineHeight:1.15}}>{label}</span>;
}

function DotCompact({periods, selectedPeriodNum, onSelectPeriod}){
  if(!periods?.length) return <span style={{color:C.text4,fontSize:11}}>—</span>;
  const compact=periods.length>7;
  const dot=compact?6:13;
  const line=compact?2:8;
  const bd=compact?1:1.5;
  const labelFs=compact?7:8;
  // 리스트 이행현황은 최대 4개 슬롯을 한 묶음으로 사용합니다.
  // 실제 노드가 R 하나뿐이어도 4개 슬롯 기준의 첫 위치를 유지하여,
  // 다른 행의 R과 동일한 왼쪽 시작점에 정렬합니다.
  const maxVisibleSlots=4;
  const trackWidth=dot*maxVisibleSlots + line*(maxVisibleSlots-1);
  return(
    <div style={{display:"flex",alignItems:"flex-start",justifyContent:"center",gap:0,width:"100%",overflow:"hidden",paddingTop:1}}>
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"flex-start",width:trackWidth,flexShrink:0}}>
      {periods.map((p,i)=>{
        const st=normStatus(p.status);
        const s=ST[st]||ST.upcoming;
        const isLast=i===periods.length-1;
        const isSel=selectedPeriodNum===p.num;
        const label=periodLabel(p);
        const isDueWindow=st==="due_soon";
        return(
          <div key={i} style={{display:"flex",flexDirection:"column",alignItems:"flex-start",flexShrink:0}}>
            <div style={{display:"flex",alignItems:"center",height:10,marginBottom:2}}>
              <span style={{width:dot,textAlign:"center",fontSize:labelFs,lineHeight:"9px",fontWeight:600,color:C.text4,letterSpacing:-.2,whiteSpace:"nowrap",overflow:"visible"}}>{label}</span>
              {!isLast&&<span style={{width:line,flexShrink:0}}/>}
            </div>
            <div style={{display:"flex",alignItems:"center",flexShrink:0}}>
              <div onClick={onSelectPeriod?(e)=>{e.stopPropagation();onSelectPeriod(p.num);}:undefined}
                style={{width:dot,height:dot,borderRadius:"50%",background:isDueWindow?C.card:s.fill,
                  border:isDueWindow?`${bd}px dashed ${C.red}`:`${bd}px solid ${s.border}`,
                  flexShrink:0,cursor:onSelectPeriod?"pointer":"default",
                  boxShadow:isSel?`0 0 0 2px ${C.card}, 0 0 0 3.5px ${C.charcoalDk}`:"none"}}
                title={`${label} | ${p.isReference?"기준일":"도래일"}: ${periodPointDate(p)||"—"}\n${p.measured||"미측정"} | ${isDueWindow?"측정권장":s.label}`}/>
              {!isLast&&<div style={{width:line,height:1.2,background:C.bd2,flexShrink:0}}/>}
            </div>
          </div>
        );
      })}
      </div>
    </div>
  );
}

function DotFull({periods, windowDays, selectedPeriodNum, onSelectPeriod, pageMeta=null, onPageChange=null}){
  const safePeriods=periods||[];
  const meta=pageMeta||xrfPeriodPageMeta(safePeriods,selectedPeriodNum,null);
  const {totalPages,activePage,windowStart,visiblePeriods}=meta;

  // 타임라인, 이행 상세, 주기별 산정 요약이 모두 같은 visiblePeriods를 사용합니다.
  // 따라서 페이지를 바꾸면 XRF 분석 탭의 주기 관련 영역 전체가 동일한 12개 구간으로 함께 전환됩니다.
  if(!safePeriods.length) return null;

  const DOT=28;
  const bd=2;
  const labelFs=9;
  const dateFs=8;
  // 12개의 슬롯이 한 줄에서 최소한의 가독성을 유지하도록 고정 최소폭을 둡니다.
  // 화면이 더 좁으면 두 줄로 내리지 않고 가로 스크롤합니다.
  const TRACK_MIN_WIDTH=780;
  const slotPct=100/XRF_PERIOD_WINDOW_SIZE;

  return(
    <div style={{width:"100%",minWidth:0,textAlign:"center"}}>
      <div style={{width:"100%",overflowX:"auto",overflowY:"visible",paddingBottom:2,scrollbarWidth:"thin"}}>
        <div style={{position:"relative",width:"100%",minWidth:TRACK_MIN_WIDTH,height:98,boxSizing:"border-box",overflow:"visible"}}>
          {/* 연결선도 absolute로 배치해 DOM이 어떤 폭에서도 다음 줄로 내려갈 수 없도록 합니다. */}
          {visiblePeriods.slice(0,-1).map((p,i)=>{
            const next=visiblePeriods[i+1];
            const nextBd=next ? (ST[normStatus(next.status)]||ST.upcoming).border : C.bd2;
            return <div key={`period-line-${windowStart+i}`} aria-hidden="true" style={{
              position:"absolute",
              left:`${(i+.5)*slotPct}%`,
              top:55,
              width:`${slotPct}%`,
              height:2,
              background:nextBd||C.bd2,
              transform:"translateY(-50%)",
              zIndex:0,
              pointerEvents:"none"
            }}/>;
          })}

          {visiblePeriods.map((p,i)=>{
            const st=normStatus(p.status);
            const s=ST[st]||ST.upcoming;
            const isSel=selectedPeriodNum===p.num;
            const canClick=!!onSelectPeriod;
            return(
              <div key={`period-node-${windowStart+i}`} style={{
                position:"absolute",
                left:`${(i+.5)*slotPct}%`,
                top:0,
                width:`${slotPct}%`,
                height:98,
                transform:"translateX(-50%)",
                display:"flex",
                flexDirection:"column",
                alignItems:"center",
                justifyContent:"flex-start",
                overflow:"visible",
                zIndex:1
              }}>
                <div style={{height:14,display:"flex",alignItems:"center",justifyContent:"center",width:"100%"}}>
                  <span style={{fontSize:labelFs,fontWeight:600,color:C.text4,letterSpacing:-.15,whiteSpace:"nowrap",lineHeight:"10px"}}>{periodLabel(p)}</span>
                </div>

                <div style={{height:26,display:"flex",alignItems:"flex-end",justifyContent:"center",width:"100%",paddingBottom:8,boxSizing:"border-box"}}>
                  {periodMeasuredDates(p).length>0&&<div style={{width:"calc(100% - 4px)",minWidth:0}}><AutoFitText value={periodMeasuredShortText(p)} title={periodMeasuredText(p)} baseFontSize={periodMeasuredDates(p).length>1?7.2:dateFs} minFontSize={6} align="center" style={{color:s.fill===C.card?s.border:s.fill,fontWeight:600,lineHeight:1}}/></div>}
                </div>

                <div style={{position:"relative",width:DOT,height:DOT,flex:"0 0 auto",zIndex:2,overflow:"visible"}}>
                  {st==="due_soon"&&<span aria-label="측정권장" title="측정권장" style={{
                    position:"absolute",left:"50%",bottom:"calc(100% + 3px)",transform:"translateX(-50%)",
                    fontSize:9,color:C.red,fontWeight:700,whiteSpace:"nowrap",lineHeight:1,pointerEvents:"none",zIndex:4
                  }}>▼</span>}
                  <div onClick={canClick?()=>onSelectPeriod(p.num):undefined}
                    style={{width:DOT,height:DOT,borderRadius:"50%",background:s.fill,border:`${bd}px solid ${s.border}`,
                      cursor:canClick?"pointer":"default",boxSizing:"border-box",
                      boxShadow:isSel?`0 0 0 2px ${C.charcoalBg}, 0 0 0 4px ${C.charcoalDk}`:st==="due_soon"?`0 0 0 2px ${C.redBg}`:"none",
                      transition:"box-shadow .15s"}}
                    title={`${periodLabel(p)} | ${p.isSupplementalMeasurement?"추가 측정일":p.isReference?"기준일":"도래일"}: ${p.isSupplementalMeasurement?(p.measured||"—"):(periodPointDate(p)||"—")}\n${periodMeasuredText(p)==="—"?"미측정":periodMeasuredText(p)} | ${s.label}`}/>
                </div>

                <div style={{minHeight:27,marginTop:7,display:"flex",flexDirection:"column",alignItems:"center",gap:1,width:"100%"}}>
                  <span style={{fontSize:dateFs,color:C.text4,whiteSpace:"nowrap"}}>{periodPointDate(p)?.slice(2,7)}</span>
                  {p.daysOver>0&&<span style={{fontSize:dateFs,color:C.red,whiteSpace:"nowrap"}}>+{p.daysOver}일</span>}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* 페이지 버튼은 주기별 이행 현황 영역 안에 항상 남아 있어 과거 구간으로 직접 이동할 수 있습니다. */}
      {totalPages>1&&<div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:5,marginTop:6,minHeight:26,flexWrap:"nowrap",overflowX:"auto"}}>
        {Array.from({length:totalPages},(_,idx)=>{
          const active=idx===activePage;
          return <button key={idx} type="button" onClick={()=>onPageChange&&onPageChange(idx)}
            aria-label={`${idx+1}번 페이지`} title={`${idx+1}번 페이지`}
            style={{minWidth:28,height:24,padding:"0 7px",border:`1px solid ${active?C.charcoalDk:C.bd2}`,borderRadius:6,
              background:active?C.charcoalDk:C.card,color:active?C.card:C.text3,fontSize:10,fontWeight:700,cursor:"pointer",flex:"0 0 auto"}}>{idx+1}</button>;
        })}
      </div>}
    </div>
  );
}
function PeriodTable({periods, selectedPeriodNum, onSelectPeriod, uploadTargetNum=null, uploadAllowed=false, uploadState=null, onUpload=null, uploadKey=""}){
  const SR={reference:{c:C.charcoal,bg:C.charcoalBg,l:"기준"},compliant:{c:C.charcoal,bg:C.charcoalBg,l:"이행"},measured:{c:C.charcoal,bg:C.charcoalBg,l:"측정 완료"},early:{c:C.charcoalMd,bg:C.charcoalBg,l:"조기측정"},late:{c:C.redDk,bg:C.redBg,l:"지연측정"},overdue:{c:C.red,bg:C.redBg,l:"미이행"},due_soon:{c:C.red,bg:C.redBg,l:"측정권장"},upcoming:{c:C.text4,bg:C.alt,l:"예정"}};
  const th={padding:"6px 9px",background:C.bg,border:`1px solid ${C.bd}`,fontWeight:600,color:C.text2,fontSize:10,letterSpacing:.2,textAlign:"center"};
  const td={padding:"7px 9px",borderBottom:`1px solid ${C.bd}`,fontSize:11,verticalAlign:"middle",textAlign:"center"};
  const inputOverlay={position:"absolute",inset:0,opacity:0,cursor:"pointer",width:"100%",height:"100%"};
  return(
    <div className="responsive-table-scroll table-scroll-medium" role="region" aria-label="주기별 이행 현황 표" tabIndex={0}>
      <table style={{width:"100%",minWidth:0,maxWidth:"100%",borderCollapse:"collapse",tableLayout:"fixed"}}>
        <colgroup>
          <col style={{width:"12%"}}/><col style={{width:"20%"}}/><col style={{width:"27%"}}/><col style={{width:"16%"}}/><col style={{width:"25%"}}/>
        </colgroup>
        <thead><tr>{["주기","기준·도래일","측정일","상태","XRF 결과 업로드"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
        <tbody>
          {periods?.map(p=>{
            const st=normStatus(p.status);
            const meta=SR[st]||SR.upcoming;
            const isSel=selectedPeriodNum===p.num;
            const canClick=!!onSelectPeriod;
            // R의 1차 ??도 같은 R에 Retest로 연결할 수 있어야 합니다.
            const isUploadTarget=Number(uploadTargetNum)===Number(p.num);
            const stateMatches=uploadState?.key && uploadState.key===uploadKey;
            const uploadTitle=stateMatches
              ? (uploadState.error || (uploadState.saved?`${uploadState.fileName} · 저장 완료`:uploadState.fileName))
              : "선택 주기의 XRF 결과 파일 업로드";
            return(
              <tr key={p.num} onClick={canClick?()=>onSelectPeriod(p.num):undefined}
                style={{background:isSel?C.charcoalBg:(["overdue","late"].includes(st)?C.redBg:"transparent"),cursor:canClick?"pointer":"default",outline:isSel?`1px solid ${C.red}`:"none",outlineOffset:-1}}>
                <td style={{...td,fontWeight:isSel?700:550,color:C.text1}}>{periodLabel(p)}{isSel&&<span style={{color:C.red,marginLeft:5,fontSize:8}}>●</span>}</td>
                <td style={{...td,color:C.text2,fontSize:11,fontWeight:400}}><AutoFitText value={periodPointDate(p)||"—"} baseFontSize={11} minFontSize={7} align="center" style={{color:C.text2,fontWeight:400}}/></td>
                <td style={{...td,fontWeight:p.measured?550:450,color:p.measured?C.text1:C.text3}}>
                  <AutoFitText value={`${periodMeasuredText(p)}${p.daysOver>0?` (+${p.daysOver}일)`:""}`} baseFontSize={11} minFontSize={7} align="center" style={{color:p.measured?C.text1:C.text3,fontWeight:p.measured?550:450}}/>
                </td>
                <td style={td}><span style={{fontSize:10,fontWeight:600,color:meta.c,background:meta.bg,padding:"3px 7px",borderRadius:6,border:`1px solid ${C.bd}`,whiteSpace:"nowrap"}}>{meta.l}</span></td>
                <td style={td} onClick={e=>e.stopPropagation()}>
                  {isUploadTarget && uploadAllowed && onUpload ? (
                    <label title={uploadTitle} style={{position:"relative",display:"inline-flex",alignItems:"center",justifyContent:"center",minHeight:28,padding:"0 12px",border:`1px solid ${C.bd2}`,borderRadius:UI.rs,background:C.bg,color:C.text1,fontSize:11,fontWeight:500,cursor:"pointer",whiteSpace:"nowrap",overflow:"hidden"}}>
                      {stateMatches ? (uploadState.saved?"저장 완료":uploadState.error?"오류 확인":"읽는 중") : "XRF 결과 업로드"}
                      <input type="file" accept=".xlsx,.xlsm,.xls,.csv,.json" onChange={onUpload} style={inputOverlay}/>
                    </label>
                  ) : null}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}

function VisualUatComplianceRenderMatrix(){
  const selectedPeriodNum=4;
  const statusRows=["reference","compliant","measured","early","late","overdue","due_soon","upcoming"];
  return(
    <div data-i18n-skip="true" data-uat-compliance-render-matrix="1" style={{padding:"16px 18px",marginBottom:16,border:`1px solid ${A.purple.ln}`,borderRadius:12,background:A.purple.bg,boxShadow:"0 8px 22px rgba(26,32,32,.075)"}}>
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:12,flexWrap:"wrap",marginBottom:14}}>
        <div>
          <div style={{fontSize:13,fontWeight:800,color:A.purple.tx}}>[UAT] 이행 현황 렌더 매트릭스</div>
          <div style={{fontSize:10.5,color:C.text3,lineHeight:1.6,marginTop:4}}>아래 데이터는 색상/도형 렌더링 확인용이며 업무 상태 산정 로직 검증에는 사용하지 않습니다.</div>
        </div>
        <span style={{fontSize:10,color:C.text4}}>선택 Ring: P4 · 기준일 {TODAY_STR}</span>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"minmax(180px,.65fr) minmax(0,1.35fr)",gap:12,marginBottom:12}}>
        <div style={{padding:"12px",background:C.card,border:`1px solid ${C.bd}`,borderRadius:8,minWidth:0}}>
          <div style={{fontSize:10.5,fontWeight:750,color:C.text2,marginBottom:10}}>부자재 리스트형 / DotCompact</div>
          <DotCompact periods={VISUAL_UAT_COMPLIANCE_RENDER_PERIODS} selectedPeriodNum={selectedPeriodNum}/>
        </div>
        <div style={{padding:"12px",background:C.card,border:`1px solid ${C.bd}`,borderRadius:8,minWidth:0}}>
          <div style={{fontSize:10.5,fontWeight:750,color:C.text2,marginBottom:4}}>XRF 상세형 / DotFull</div>
          <DotFull periods={VISUAL_UAT_COMPLIANCE_RENDER_PERIODS} windowDays={30} selectedPeriodNum={selectedPeriodNum}/>
        </div>
      </div>

      <div style={{padding:"12px",background:C.card,border:`1px solid ${C.bd}`,borderRadius:8,marginBottom:12}}>
        <div style={{fontSize:10.5,fontWeight:750,color:C.text2,marginBottom:9}}>상태 테이블형 / PeriodTable</div>
        <PeriodTable periods={VISUAL_UAT_COMPLIANCE_RENDER_PERIODS} selectedPeriodNum={selectedPeriodNum}/>
      </div>

      <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
        {statusRows.map(status=>{
          const meta=ST[status]||ST.upcoming;
          return <span key={status} style={{display:"inline-flex",alignItems:"center",gap:5,padding:"4px 7px",background:C.card,border:`1px solid ${C.bd}`,borderRadius:6,fontSize:9.5,color:C.text3}}>
            <span style={{width:11,height:11,borderRadius:"50%",background:meta.fill,border:`1.5px solid ${meta.border}`,boxSizing:"border-box"}}/>
            {status} · {meta.fill} / {meta.border}
          </span>;
        })}
      </div>
    </div>
  );
}

// 판정 라벨은 "같은 종류끼리" 동일 규격을 쓰되, 화면을 과하게 차지하지 않도록 최소 폭으로 유지합니다.
// XRF/정밀분석: OK · NG · 확인 필요 · 재측정 필요 → 62 × 20px
// 승인 상태: 승인 · 보류 · 반려 · 측정 진행중 · 의뢰 필요 · 확인 필요 → 64 × 20px
const RESULT_LABEL_WIDTH = 62;
const RESULT_LABEL_HEIGHT = 20;
const RESULT_LABEL_FONT_SIZE = 10;
const APPROVAL_LABEL_WIDTH = 64;
const APPROVAL_LABEL_HEIGHT = 20;
const APPROVAL_LABEL_FONT_SIZE = 10;
const FOLLOWUP_LABEL_WIDTH = 86;
const COMPACT_RESULT_VALUES = new Set(["OK","NG","확인 필요","재측정 필요","??","대기","측정 대기"]);

// 내부 상태값은 기존 호환성을 위해 `재측정 필요`를 유지하되, 화면 라벨은 간결하게 `재측정`으로 표시합니다.
function resultUiLabel(value){
  return String(value??"").trim()==="재측정 필요" ? "재측정" : value;
}

function Chip({v, small, result=false}){
  const M={"Not Allowed":{bg:A.rose.bg,c:A.rose.tx,bd:A.rose.ln},"H":{bg:A.rose.bg,c:A.rose.tx,bd:A.rose.ln},"M":{bg:A.amber.bg,c:A.amber.tx,bd:A.amber.ln},"L":{bg:A.green.bg,c:A.green.tx,bd:A.green.ln},"NG":{bg:A.rose.bg,c:A.rose.tx,bd:A.rose.ln},"Warning":{bg:A.amber.bg,c:A.amber.tx,bd:A.amber.ln},"OK":{bg:A.green.bg,c:A.green.tx,bd:A.green.ln},"??":{bg:A.amber.bg,c:A.amber.tx,bd:A.amber.ln},"확인 필요":{bg:A.amber.bg,c:A.amber.tx,bd:A.amber.ln},"재측정 필요":{bg:A.amber.bg,c:A.amber.tx,bd:A.amber.ln},"사용 가능":{bg:A.green.bg,c:A.green.tx,bd:A.green.ln},"정밀분석 인계":{bg:A.rose.bg,c:A.rose.tx,bd:A.rose.ln},"XRF 재측정":{bg:A.amber.bg,c:A.amber.tx,bd:A.amber.ln}};
  const normalized=String(v??"").trim();
  // XRF 판정값은 호출부에서 result prop을 빠뜨려도 항상 68 × 22px로 고정합니다.
  const compactResult=result || COMPACT_RESULT_VALUES.has(normalized);
  const chipSize=compactResult
    ? {width:RESULT_LABEL_WIDTH,minWidth:RESULT_LABEL_WIDTH,maxWidth:RESULT_LABEL_WIDTH,height:RESULT_LABEL_HEIGHT,padding:0,fontSize:RESULT_LABEL_FONT_SIZE,fontWeight:600,flex:"0 0 auto"}
    : {height:small?22:26,padding:small?"0 10px":"0 12px",fontSize:small?9.5:11,fontWeight:500};
  if(v==null || v==="—"){
    return <span style={{...chipSize,color:C.text3,background:C.card,borderRadius:UI.pill,display:"inline-flex",alignItems:"center",justifyContent:"center",border:`1px solid ${C.bd2}`,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",boxSizing:"border-box",lineHeight:1}}>—</span>;
  }
  const TONE={"Not Allowed":A.purple,"H":A.rose,"M":A.amber,"L":A.green,"NG":A.rose,"Warning":A.amber,"OK":A.green,"??":A.amber,"확인 필요":A.amber,"재측정 필요":A.amber,"사용 가능":A.green,"정밀분석 인계":A.rose,"XRF 재측정":A.amber};
  const tone=TONE[v];
  if(tone) return <StatusPill tone={tone} label={v==="Not Allowed"?"사용불허":resultUiLabel(v)} size={compactResult?"sm":(small?"sm":"md")} fixedWidth={compactResult?RESULT_LABEL_WIDTH:undefined}/>;
  const st=M[v]||{bg:C.card,c:C.text2,bd:C.bd2};
  return <span style={{...chipSize,fontWeight:compactResult?600:600,color:st.c,background:st.bg,borderRadius:UI.pill,display:"inline-flex",alignItems:"center",justifyContent:"center",border:`1px solid ${st.bd}`,lineHeight:1,overflow:"hidden",whiteSpace:"nowrap",boxSizing:"border-box"}}><AutoFitText value={resultUiLabel(v)} baseFontSize={compactResult?RESULT_LABEL_FONT_SIZE:(small?9.5:11)} minFontSize={7} align="center" style={{fontWeight:600,lineHeight:1}}/></span>;
}

// 부자재 리스트 탭의 ITEM NAME 열을 제외한 화면에서는 영문 모드일 때 영문 품목명을 우선 표시합니다.
function displayItemName(item, lang){
  if(lang!=="en") return item?.name || "";
  return String(item?.nameEn||ITEM_NAME_EN_FALLBACKS[String(item?.code||"").trim()]||"").trim() || item?.name || "";
}

function CatDot({cat, dark=false, lang="ko"}){
  const m=CAT_META[cat]||CAT_META.existing;
  const enLabel={existing:"Existing",new:"New",changed:"Change"}[cat] || "Existing";
  const label=lang==="en" ? enLabel : m.label;
  const labelColor=dark?"rgba(255,255,255,.88)":m.dot;
  return (
    <span style={{display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:11,color:labelColor,fontWeight:dark?700:500,whiteSpace:"nowrap",maxWidth:"100%",minWidth:0,overflow:"hidden"}}>
      <span style={{minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{label}</span>
    </span>
  );
}

function CRChip({crLevel, crType, compact}){
  if(!crLevel) return null;
  const col=crLevel==="H"?C.redDk:crLevel==="M"?C.charcoalMd:C.charcoal;
  const bg=crLevel==="H"?C.redBg:C.charcoalBg;
  return(
    <span style={{display:"inline-flex",alignItems:"center",gap:5,fontSize:11,fontWeight:600,color:col,background:bg,padding:"2px 8px",borderRadius:2}}>
      <span>{crLevel}</span>
      {!compact&&crType&&<span style={{fontWeight:400,fontSize:10,color:col,opacity:.8}}>· {crType}</span>}
    </span>
  );
}

function MetricChip({label, children, dark=false}){
  return(
    <div style={{display:"flex",flexDirection:"column",gap:4,minWidth:76}}>
      <div style={{fontSize:9,color:dark?"rgba(255,255,255,.70)":C.text3,fontWeight:600,letterSpacing:.25,whiteSpace:"nowrap"}}>{label}</div>
      <div style={{minHeight:20,display:"flex",alignItems:"center"}}>{children}</div>
    </div>
  );
}

function MiniMetric({label, children, hint, strong=false}){
  return(
    <div style={{background:C.card,border:`1px solid ${C.bd}`,borderRadius:10,padding:"9px 10px",minWidth:0,boxShadow:"0 1px 0 rgba(255,255,255,.9) inset"}}>
      <div style={{fontSize:9,color:C.text3,fontWeight:600,letterSpacing:.2,marginBottom:6,whiteSpace:"nowrap"}}>{label}</div>
      <div style={{minHeight:20,display:"flex",alignItems:"center",fontSize:strong?13:11,fontWeight:strong?700:600,color:C.text1,minWidth:0}}>{typeof children==="string" || typeof children==="number" ? <AutoFitText value={children} baseFontSize={strong?13:11} minFontSize={7} align="left" style={{fontWeight:strong?700:600,color:C.text1}}/> : children}</div>
      {hint&&<div style={{fontSize:9,color:C.text3,marginTop:4,lineHeight:1.35,fontWeight:400}}>{hint}</div>}
    </div>
  );
}

function XrfResultSummary({measurement, item}){
  if(!measurement){
    return(
      <div style={{border:`1px solid ${C.bd}`,background:C.alt,padding:"14px 16px",marginBottom:12,borderRadius:12}}>
        <div style={{...T.section,marginBottom:5}}>산정 요약</div>
        <div style={{fontSize:11,color:C.text3,fontWeight:400}}>선택한 주기에 연결된 측정 이력이 없어 XRF 결과를 표시하지 않습니다. 최초 등록 시 산정한 Final Risk는 유지됩니다.</div>
      </div>
    );
  }
  const xrfResult=measurementXrfResult(measurement);
  const xrfLevel=measurementXrfLevel(measurement);
  const final=itemFinalRisk(item)||"—";
  const basis=riskBasisMeasurementOf(item);
  const isReference=measurement?.id && basis?.id && measurement.id===basis.id;
  const followupInfo=measurementFollowupInfo(measurement);
  const {precisionElements,remeasureElements,flaggedElements,hasFollowup}=followupInfo;
  const visibleFollowupDetail=measurementFollowupVisibleDetail(followupInfo);
  const stepStyle={background:C.card,border:`1px solid ${C.bd}`,borderRadius:12,padding:"10px 12px",minWidth:0,textAlign:"center",boxShadow:"0 2px 10px rgba(26,32,32,.08)",boxSizing:"border-box"};
  const precisionSet=new Set(precisionElements);
  const remeasureSet=new Set(remeasureElements);
  const flaggedSet=new Set(flaggedElements);
  const tone=hasFollowup?A.rose:A.green;
  // R과 정기 측정 모두 원소별 후속조치를 같은 상태 레일로 표현합니다.
  // 정상 원소는 중립 셀, 후속 분석 대상만 강조합니다.
  // 후속 분석 대상이 없을 때는 "0"이나 빈 count badge를 따로 노출하지 않습니다.
  const followupRail=(
    <div style={{display:"grid",gridTemplateColumns:"1fr",justifyItems:"center",alignItems:"center",rowGap:12,
        padding:"18px 20px",marginBottom:12,borderRadius:UI.r,border:`1px solid ${tone.ln}`,
        background:tone.bg,textAlign:"center",boxSizing:"border-box"}}>

        {/* 원소 라벨 자체의 묶음 폭만 차지하게 해 카드의 정확한 중앙에 놓습니다. */}
        <div aria-label="원소별 후속조치 상태"
          style={{display:"flex",alignItems:"center",justifyContent:"center",gap:7,flexWrap:"wrap",
            width:"fit-content",maxWidth:"100%",margin:"0 auto"}}>
          {ELEMENTS.map(el=>{
            const flagged=flaggedSet.has(el);
            const action=precisionSet.has(el) && remeasureSet.has(el)
              ? "정밀분석 및 XRF 재측정"
              : precisionSet.has(el)
                ? "정밀분석 필요"
                : remeasureSet.has(el)
                  ? "XRF 재측정"
                  : "Internal Limit 이내";
            const title=el==="Cl+Br" && !flagged
              ? `${el} · Cl Content + Br Content 합계 / Internal Limit 1,050 ppm`
              : `${el} · ${action}`;
            return <span key={el} title={title}
              style={{display:"inline-flex",alignItems:"center",justifyContent:"center",minWidth:42,height:26,padding:"0 10px",
                borderRadius:UI.pill,border:`1px solid ${flagged?A.rose.solid:A.green.ln}`,
                background:flagged?C.card:"rgba(255,255,255,.78)",color:flagged?A.rose.tx:C.text1,
                fontSize:11,fontWeight:650,textAlign:"center",boxSizing:"border-box"}}>{el}</span>;
          })}
        </div>

        {/* 후속조치 값은 배지/버튼 없이 검정 텍스트로만 표시합니다. */}
        <div style={{display:"grid",justifyItems:"center",alignItems:"center",rowGap:5,width:"fit-content",maxWidth:"100%",margin:"0 auto"}}>
          <span style={{fontFamily:FONT_SANS,fontSize:11,fontWeight:500,color:C.text1,lineHeight:1.45,letterSpacing:0,textAlign:"center"}}>{followupInfo.label}</span>
          {hasFollowup
            ? visibleFollowupDetail&&<div style={{fontSize:11,color:C.text2,fontWeight:400,textAlign:"center",lineHeight:1.5,maxWidth:560}}>{visibleFollowupDetail}</div>
            : <div style={{fontSize:11,color:C.text2,fontWeight:400,textAlign:"center",lineHeight:1.5,maxWidth:560}}>원본 보고서 Judgment 기준 후속조치 대상이 없습니다.</div>}
        </div>
      </div>
  );
  if(!isReference) return followupRail;
  const summarySteps=[
    ["XRF Level", <Chip v={xrfLevel||"—"} small/>],
    ["C&R Level", <CRChip crLevel={item?.crLevel} crType={item?.crType} compact/>],
    ["Final Risk", <Chip v={final||"—"}/>],
    ["검사주기", <span style={{fontSize:11,fontWeight:700,color:C.text1}}>{CYCLE_KO[itemCycleFromRisk(item)]||itemCycleFromRisk(item)||"—"}</span>],
  ];
  return(<>
    <div style={{border:`1px solid ${C.bd}`,background:C.card,padding:"14px 16px",marginBottom:12,borderRadius:14,boxShadow:"0 8px 20px rgba(26,32,32,.08)"}}>
      <div style={{display:"flex",justifyContent:"space-between",gap:12,alignItems:"flex-start",marginBottom:12}}>
        <div>
          <div style={{...T.section,marginBottom:5}}>{isReference?"최초 위험도 평가 · R":"정기 XRF 결과"}</div>
          <div style={{fontSize:11,color:C.text3,lineHeight:1.55,fontWeight:400}}>{isReference?"R 단계에서 XRF Level과 C&R Level로 Final Risk와 검사주기를 확정하며, P1은 R 등록일에서 해당 주기만큼 지난 시점에 생성됩니다.":"Judgment와 후속조치만 갱신하며 Final Risk와 검사주기는 R 최초 평가 결과를 유지합니다."}</div>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(126px,1fr))",gap:8,alignItems:"stretch"}}>
        {summarySteps.map(([label,node],idx)=>(
          <div key={label} style={{...stepStyle,border:idx===summarySteps.length-1?`1px solid ${C.bd2}`:stepStyle.border,background:idx===summarySteps.length-1?C.alt:C.card}}>
            <div style={{fontSize:9,color:C.text3,marginBottom:7,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{label}</div>
            <div style={{minHeight:24,display:"flex",alignItems:"center",justifyContent:"center",minWidth:0,overflow:"hidden",textAlign:"center"}}>{node}</div>
          </div>
        ))}
      </div>
    </div>
    {followupRail}
  </>);
}

function PeriodMeasurementSummary({item, periods, selectedPeriodNum, currentPeriodNum, selectedMeasurementId=null, onSelectPeriod, onSelectMeasurement=null}){
  const currentNum=selectedPeriodNum ?? currentPeriodNum;
  const visiblePeriods=periods||[];
  const entries=visiblePeriods.flatMap(p=>{
    const measurements=periodMeasurementsOf(item,p);
    if(!measurements.length) return [{p,m:null,index:0,total:1}];
    return measurements.map((m,index)=>({p,m,index,total:measurements.length}));
  });
  return(
    <div style={{padding:"16px 22px",background:C.card}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end",gap:12,marginBottom:12}}>
        <div>
          <div style={{...T.section}}>주기별 산정 요약</div>
          <div style={{fontSize:10,color:C.text3,marginTop:3,fontWeight:400}}>같은 주기에 여러 측정이 있으면 같은 Pn으로 각각 표시되며, 측정 이력을 선택하면 상단 값과 우측 원소 분석도 해당 측정으로 전환됩니다.</div>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))",gap:10}}>
        {entries.map(({p,m,index,total})=>{
          const isLatestInPeriod=index===total-1;
          const isSel=selectedMeasurementId
            ? m?.id===selectedMeasurementId
            : (p.num===currentNum && (total===1 || isLatestInPeriod));
          const xrfResult=m?measurementXrfResult(m):"—";
          const xrfLevel=m?measurementXrfLevel(m):null;
          const final=itemFinalRisk(item)||"—";
          const followupInfo=m?measurementFollowupInfo(m):null;
          const hasFollowup=!!followupInfo?.hasFollowup;
          const followup=m ? followupInfo.label : "측정 없음";
          const followupDetail=m ? measurementFollowupVisibleDetail(followupInfo) : "";
          const st=ST[normStatus(p.status)]||ST.upcoming;
          const handleClick=()=>{
            if(m?.id && onSelectMeasurement) onSelectMeasurement(p.num,m.id);
            else if(onSelectPeriod) onSelectPeriod(p.num);
          };
          return(
            <button key={`${p.num}-${m?.id||"unmeasured"}-${index}`} onClick={handleClick}
              style={{textAlign:"left",padding:"12px 12px",background:isSel?C.card:C.charcoalBg,border:`1px solid ${isSel?"#FFFFFF":C.bd}`,borderRadius:16,cursor:"pointer",
                boxShadow:isSel?"0 10px 22px rgba(26,32,32,.18), 0 1px 0 rgba(255,255,255,.95) inset":"0 1px 0 rgba(26,32,32,.05) inset",
                transform:isSel?"translateY(-2px)":"none",transition:"all .16s ease",minWidth:0}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
                <span style={{fontSize:13,fontWeight:700,color:C.text1}}>{periodLabel(p)}</span>
                {total>1&&<span style={{fontSize:9,color:C.text3,fontWeight:600}}>측정 {index+1}/{total}</span>}
                <span style={{fontSize:10,fontWeight:600,color:st.c,background:st.bg,padding:"2px 7px",borderRadius:999,border:`1px solid ${C.bd}`}}>{st.l}</span>
                <span style={{marginLeft:"auto",fontSize:10,color:C.text3,fontWeight:400}}>{m?.date||"미측정"}</span>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(104px,1fr))",gap:6,marginBottom:9}}>
                {(p.isReference
                  ? [["XRF Result", <Chip v={displayXrfText(xrfResult)} small result/>], ["XRF Level", <Chip v={xrfLevel||"—"} small/>], ["Final Risk", <Chip v={final} small/>]]
                  : [["XRF Result", <Chip v={displayXrfText(xrfResult)} small result/>], ["적용 Risk", <Chip v={final} small/>], ["후속조치", <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:4,minWidth:0,width:"100%",textAlign:"center"}}>
                    <span style={{fontFamily:FONT_SANS,fontSize:10.5,fontWeight:500,color:C.text1,lineHeight:1.4,letterSpacing:0,textAlign:"center",whiteSpace:"normal"}}>{followup}</span>
                    {followupDetail&&<span style={{fontSize:9,fontWeight:400,color:C.text2,lineHeight:1.25,whiteSpace:"normal",textAlign:"center"}}>{followupDetail}</span>}
                  </div>]]
                ).map(([label,node])=>(
                  <div key={label} style={{background:isSel?C.alt:C.card,border:`1px solid ${C.bd}`,borderRadius:10,padding:"8px 6px",minWidth:0,overflow:"hidden",textAlign:"center",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
                    <div style={{fontSize:8,color:C.text3,marginBottom:6,fontWeight:500,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",width:"100%",textAlign:"center"}}>{label}</div>
                    <div style={{minHeight:24,display:"flex",alignItems:"center",justifyContent:"center",minWidth:0,width:"100%",overflow:"hidden",textAlign:"center"}}>{node}</div>
                  </div>
                ))}
              </div>
              <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap",fontSize:10,color:C.text3,lineHeight:1.45}}>
                <span>Risk 기준 <b style={{color:C.text2,fontWeight:600}}>R</b></span>
              </div>
              {m?.id&&<div style={{width:"100%",minWidth:0,marginTop:6}}><AutoFitText value={m.id} title={m.id} baseFontSize={9} minFontSize={6.5} align="left" style={{color:C.text4,fontWeight:400}}/></div>}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function CompText({item}){
  const cs=getCS(item);
  if(cs==="notAvailable" || cs==="replaced"){
    return <span style={{fontSize:11,color:C.text4,fontWeight:600}}>대상 아님</span>;
  }
  if(cs==="overdue") return <span style={{fontSize:11,fontWeight:700,color:C.red}}>미이행</span>;
  if(cs==="dueSoon") return <span style={{fontSize:11,fontWeight:700,color:C.red}}>측정권장</span>;
  if(cs==="upcoming") return <span style={{fontSize:11,color:C.text4,fontWeight:600}}>예정</span>;
  return <span style={{fontSize:11,color:C.charcoal,fontWeight:700}}>이행</span>;
}

function DDayText({dateStr}){
  const dd=daysUntil(dateStr);
  if(dd===null) return <span style={{color:C.text4,fontSize:12}}>—</span>;
  if(dd<0)   return <span style={{color:C.red,fontWeight:700,fontSize:12}}>D+{Math.abs(dd)}</span>;
  if(dd<=7)  return <span style={{color:C.red,fontWeight:700,fontSize:12}}>D-{dd}</span>;
  if(dd<=30) return <span style={{color:C.redDk,fontWeight:700,fontSize:12}}>D-{dd}</span>;
  if(dd<=90) return <span style={{color:C.charcoalMd,fontSize:12}}>D-{dd}</span>;
  return <span style={{color:C.text4,fontSize:12}}>D-{dd}</span>;
}

function listNextDueValue(item){
  const cycleMuted=isDiscontinueRequestItem(item) || isDiscontinuedItem(item) || !isComplianceTarget(item);
  return cycleMuted ? "—" : (nextDueOfItem(item)||"—");
}
function listDDayValue(item){
  if(item?.lifecycle==="ReplacedOld") return "—";
  const due=listNextDueValue(item);
  const dd=due==="—" ? null : daysUntil(due);
  if(dd===null) return "—";
  return dd<0 ? `D+${Math.abs(dd)}` : `D-${dd}`;
}
function listColumnValue(item,column,approvalStatus=""){
  if(!item) return "—";
  if(column==="category") return item.category||"—";
  if(column==="photo") return item.photoFileUrl?"사진 있음":"사진 없음";
  if(column==="code") return item.code||"—";
  if(column==="name") return [item.name,item.nameEn].filter(Boolean).join(" / ")||"—";
  if(column==="dept") return item.dept||"—";
  if(column==="cycle") return itemCycleFromRisk(item)||"Not Available";
  if(column==="firstDate") return item.firstDate||"—";
  if(column==="compliance") return getCS(item);
  if(column==="xrf") return itemXrfResult(item)||"—";
  if(column==="precision") return precisionAnalysisResultOf(item)||"—";
  if(column==="approval") return approvalFilterKey(approvalStatus||approvalStatusOfItemFallback(item));
  if(column==="nextDue") return listNextDueValue(item);
  if(column==="dDay") return listDDayValue(item);
  return "—";
}

function approvalStatusOfItemFallback(item){
  return item?.approvalStatus||deriveInitialApprovalStatus(latestMeasurementOf(item));
}

function itemMatchesListColumnFilter(item,column,selected,approvalStatus=""){
  const selectedValues=selected||[];
  if(!selectedValues.length) return true;
  if(column==="xrf"){
    const result=itemXrfResult(item);
    const level=riskBasisXrfLevelOf(item)||itemLatestXrfLevel(item);
    return selectedValues.includes(result)||selectedValues.includes(level);
  }
  return selectedValues.includes(listColumnValue(item,column,approvalStatus));
}

function itemMatchesListSearch(item,search,approvalStatus=""){
  const query=String(search||"").trim().toLocaleLowerCase("ko-KR");
  if(!query) return true;
  const searchableColumns=["category","photo","code","name","dept","cycle","firstDate","compliance","xrf","precision","approval","nextDue","dDay"];
  const searchableValues=searchableColumns.flatMap(column=>{
    const value=listColumnValue(item,column,approvalStatus);
    const label=FILTER_LABELS[column]?.[value];
    return [value,label];
  }).concat([
    item?.name,item?.nameEn,item?.manufacturer,item?.materialCategory,item?.materialState,
    item?.crLevel,itemFinalRisk(item),item?.lifecycle,LIFECYCLE_KO[item?.lifecycle],
    riskBasisXrfLevelOf(item),itemLatestXrfLevel(item),approvalStatus
  ]);
  return searchableValues.some(value=>String(value||"").toLocaleLowerCase("ko-KR").includes(query));
}

function isCancelledListItem(item,approvalStatus=""){
  return item?.lifecycle==="Cancelled"
    || approvalFilterKey(approvalStatus||approvalStatusOfItemFallback(item))==="cancelled"
    || (isDiscontinueRequestItem(item) && String(item?.processStatus||"").trim().toLowerCase()==="reverted");
}

function RiskBadge({risk}){
  const label=risk==="Not Allowed"?"사용불가":(risk||"—");
  const t=(label==="사용불가"||label==="H")?A.rose:(label==="M"?A.amber:(label==="L"?A.green:A.slate));
  return <StatusPill tone={t} label={label} minWidth={58}/>;
}
function RetestBadge({required}){
  return <StatusPill tone={required?A.rose:A.green} label={required?"재측정":"재측정 불필요"} size="sm" fixedWidth={RESULT_LABEL_WIDTH}/>;
}
function XrfFollowupBadge({measurement}){
  const label=measurementFollowupInfo(measurement).label;
  // 후속조치 필드는 상태와 관계없이 배지/버튼을 사용하지 않고 검정 텍스트로 통일합니다.
  return <AutoFitText value={label} title={label} baseFontSize={10.5} minFontSize={7} align="center" style={{fontFamily:FONT_SANS,fontWeight:500,color:C.text1,lineHeight:1.4,letterSpacing:0}}/>;
}
function firstResultValue(...values){
  for(const v of values){
    if(v!==undefined && v!==null && String(v).trim()!=="") return String(v).trim();
  }
  return "";
}
function normalizeAnalysisResult(value){
  const raw=String(value||"").trim();
  if(!raw) return "—";
  const u=raw.toUpperCase();
  if(["PASS","PASSED","OK","GOOD","COMPLIANT","적합","완료"].includes(u)) return "OK";
  if(["FAIL","FAILED","NG","NON-COMPLIANT","NONCOMPLIANT","부적합"].includes(u)) return "NG";
  if(["REQUIRED","PENDING","WAITING","IN PROGRESS","진행중","진행 중","대기","필요"].includes(u)) return "대기";
  if(["NOT REQUIRED","NOT_REQUIRED","N/A","NA","대상 아님","불필요"].includes(u)) return "대상 아님";
  return raw;
}
function precisionAnalysisResultOf(item){
  const wf=workflowForItem(item);
  // 정밀 "결과" 값은 OK/NG만 사용합니다. 미해결 Case가 있으면 결과 없음(—)으로 유지합니다.
  if(Array.isArray(wf?.precisionCases) && wf.precisionCases.length){
    const active=wf.precisionCases.find(c=>c.key===wf.activePrecisionCaseKey);
    if(active && !active.precision?.result) return "—";
    if(active?.precision?.result==="NG") return "NG";
    const latestResolved=wf.precisionCases.slice().reverse().find(c=>c.precision?.result==="OK" || c.precision?.result==="NG");
    if(latestResolved) return latestResolved.precision.result;
  }
  if(wf?.precision?.result==="OK" || wf?.precision?.result==="NG") return wf.precision.result;
  const m=latestMeasurementOf(item);
  const raw=normalizeAnalysisResult(firstResultValue(
    m?.precisionAnalysisResult,m?.precisionResult,m?.detailedAnalysisResult,m?.detailedResult,m?.Precision_Analysis_Result,
    item?.precisionAnalysisResult,item?.precisionResult,item?.detailedAnalysisResult,item?.detailedResult,item?.Precision_Analysis_Result
  ));
  return raw==="OK" || raw==="NG" ? raw : "—";
}
function AnalysisResultBadge({value}){
  const v=normalizeAnalysisResult(value);
  // 정밀분석 결과 라벨 전용 색상
  // OK: #1D5F99 / #BBD4EA / #DFF0FF
  // NG·후속조치 필요: #D97706 / #F3C998 / #FFEDDB
  if(v==="OK") return <StatusPill tone={A.green} label="OK" size="sm" fixedWidth={RESULT_LABEL_WIDTH}/>;
  if(v==="NG") return <StatusPill tone={A.rose} label="NG" size="sm" fixedWidth={RESULT_LABEL_WIDTH}/>;
  if(v==="대기" || v==="재측정 필요" || v==="확인 필요" || v==="측정 대기") return <StatusPill tone={A.amber} label={resultUiLabel(v)} size="sm" fixedWidth={RESULT_LABEL_WIDTH}/>;
  if(v==="대상 아님") return <StatusPill tone={A.slate} label="대상 아님" size="sm" fixedWidth={RESULT_LABEL_WIDTH}/>;
  if(v==="—") return <span title="결과 없음" style={{width:RESULT_LABEL_WIDTH,height:22,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:10.5,color:C.text4,fontWeight:500}}>—</span>;
  return <StatusPill tone={A.slate} label={v} size="sm" fixedWidth={RESULT_LABEL_WIDTH}/>;
}

function WorkflowStageText({text}){
  const display=resultUiLabel(text);
  return <div style={{width:"100%",minWidth:0,marginBottom:4}}><AutoFitText value={display||"—"} title={String(display||"—")} baseFontSize={10} minFontSize={7} align="center" style={{color:C.text3,fontWeight:550,lineHeight:1.3}}/></div>;
}
function XrfWorkflowCell({item}){
  const wf=workflowForItem(item);
  const result=itemXrfWorst(item);
  let status="";
  if(wf.stage==="XRF_REQUIRED") status="확인 필요";
  else if(wf.stage==="XRF_REMEASURE_REQUIRED") status="재측정 필요";
  else if(!result || result==="—") status="확인 필요";
  const showStatus=status && !(result && result!=="—" && status===result);
  const latestLevel=itemLatestXrfLevel(item)||"—";
  const rLevel=riskBasisXrfLevelOf(item)||"—";
  const rRisk=itemFinalRisk(item)||"—";
  return <div title={`최신 XRF ${result||"—"} · 최신 Level ${latestLevel} · R Level ${rLevel} · R Final Risk ${rRisk}`} style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minWidth:0}}>
    {showStatus&&<WorkflowStageText text={status}/>}
    {result&&result!=="—"?<Chip v={result} small result/>:displayXrfEmpty()}
  </div>;
}
function PrecisionWorkflowCell({item}){
  // 정밀분석 "결과" 열에는 OK/NG만 표시하고 진행상태 설명 텍스트는 표시하지 않습니다.
  return <div style={{display:"flex",alignItems:"center",justifyContent:"center",minWidth:0}}>
    <AnalysisResultBadge value={precisionAnalysisResultOf(item)}/>
  </div>;
}
function FollowUpCompact({item}){
  const p=workflowFollowUpProgress(item);
  if(!p || !p.total) return <span title="후속조치 대상 아님" style={{fontSize:12,color:C.text4}}>—</span>;
  const complete=p.done===p.total;
  return <span title={`B/S + 고객 사전 신고 ${p.done}/${p.total} 완료`} style={{fontSize:11,color:complete?C.charcoal:C.orange,fontWeight:700,whiteSpace:"nowrap"}}>{p.done}/{p.total} 완료</span>;
}
function WorkflowStrip({item, lang="ko"}){
  const wf=workflowForItem(item);
  const attempts=wf.xrfAttempts||[];
  const xrfRaw=attempts.length ? attempts[attempts.length-1].normalizedResult : itemXrfWorst(item);
  const xrfValue=xrfRaw==="REMEASURE"?"재측정 필요":(xrfRaw||"—");
  const precisionValue=wf.precision?.result
    || (wf.precision?.requestStatus==="REQUESTED"
      ? (wf.precision?.uploadStatus==="UPLOADED"?"결과 확인":"결과 대기")
      : wf.precision?.required
        ? "인계 필요"
        : (["PRECISION_FOLLOWUP_REQUIRED","PRECISION_REQUEST_REQUIRED"].includes(wf.stage)?"인계 필요"
          :wf.stage==="PRECISION_RESULT_WAITING"?"결과 대기":""));
  const approvalValue=wf.approval?.status || (wf.stage==="APPROVED"?"승인":"측정 진행중");
  const latest=latestMeasurementOf(item);
  const remeasureTargets=pendingXrfRemeasureElements(latest);

  const precisionSkipped=xrfValue==="OK" && !wf.precision?.result;
  const precisionStages=new Set(["PRECISION_FOLLOWUP_REQUIRED","PRECISION_REQUEST_REQUIRED","PRECISION_RESULT_WAITING","PRECISION_NG_REVIEW"]);
  const activeIdx=wf.stage==="APPROVED"?2:(precisionStages.has(wf.stage)?1:0);

  // Open Color Gray 7/8을 타임라인의 주 배경으로 사용하고,
  // 현재 단계는 별도 포인트 컬러 없이 흰색으로만 강조합니다.
  const GRAY7="#495057";
  const GRAY8="#343A40";
  const GRAY5="#ADB5BD";
  const GRAY4="#CED4DA";

  const statusTone=(v)=>{
    if(["OK","승인"].includes(v)) return A.green;
    if(["NG","반려"].includes(v)) return A.rose;
    if(["재측정 필요","인계 필요","결과 대기","보류"].includes(v)) return A.amber;
    if(v==="측정 진행중") return A.blue;
    return A.slate;
  };

  const nodes=[
    {
      label:"XRF",
      value:xrfValue,
      sub:remeasureTargets.length
        ? `재측정: ${remeasureTargets.join(", ")}`
        : attempts.length?`${attempts.length}차 · ${attempts[attempts.length-1]?.measuredDate||""}`:"",
      muted:false
    },
    {
      label:"정밀분석",
      value:precisionSkipped?"해당 없음":precisionValue,
      sub:precisionSkipped?"":(wf.precision?.measuredDate||wf.precision?.requestedAt||""),
      muted:precisionSkipped
    },
    {
      label:"승인",
      value:approvalValue,
      sub:wf.approval?.basis||"",
      muted:false
    }
  ];

  // 판정 결과는 화면에서 이미 쓰고 있는 StatusPill 디자인을 그대로 재사용합니다.
  const renderStatus=(nd,idx)=>{
    if(!nd.value) return <span style={{height:22}}/>;
    const tone=nd.muted?A.slate:statusTone(nd.value);
    if(idx===2){
      return <StatusPill tone={tone} label={nd.value} size="sm" variant="outlined" shape="rounded"
        fixedWidth={APPROVAL_LABEL_WIDTH} fixedHeight={APPROVAL_LABEL_HEIGHT} fixedFontSize={APPROVAL_LABEL_FONT_SIZE}/>;
    }
    if(nd.value==="재측정 필요"){
      const remeasureLabel=lang==="en" ? "Remeasurement Required" : nd.value;
      return <span data-i18n-skip="true" className="workflow-remeasure-pill" title={remeasureLabel} style={{display:"inline-flex",alignItems:"center",justifyContent:"center",width:"fit-content",minWidth:lang==="en"?"max-content":RESULT_LABEL_WIDTH,maxWidth:"none",height:RESULT_LABEL_HEIGHT,padding:"0 10px",background:tone.bg,border:`1px solid ${tone.ln}`,borderRadius:UI.pill,color:tone.tx,fontSize:RESULT_LABEL_FONT_SIZE,fontWeight:600,whiteSpace:"nowrap",boxSizing:"border-box",lineHeight:1,flex:"0 0 auto",position:"relative",zIndex:2}}>{remeasureLabel}</span>;
    }
    return <StatusPill tone={tone} label={nd.value} size="sm"
      fixedWidth={RESULT_LABEL_WIDTH} fixedHeight={RESULT_LABEL_HEIGHT} fixedFontSize={RESULT_LABEL_FONT_SIZE}/>;
  };

  return <div style={{padding:"18px clamp(18px, 2.8vw, 40px) 20px",background:GRAY8,borderBottom:`1px solid ${GRAY7}`,
    boxShadow:"inset 0 1px 0 rgba(255,255,255,.035)"}}>
    <div style={{fontSize:12,fontWeight:700,color:"#fff",letterSpacing:"-.05px",marginBottom:22}}>현재 단계</div>

    {/* 단계명을 별도 행으로 분리하고, 원/연결선 행의 정확한 50% 높이에 점선을 둡니다. */}
    <div style={{width:"100%",padding:"0 4px 1px"}}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",width:"100%",marginBottom:12}}>
        {nodes.map((nd,idx)=>{
          const current=idx===activeIdx && !nd.muted;
          const stageLabelColor=current?"#fff":(nd.muted?GRAY5:GRAY4);
          return <div key={`${nd.label}-label`} style={{minWidth:0,textAlign:"center",padding:"0 10px",
            fontSize:12.5,fontWeight:700,color:stageLabelColor,lineHeight:1.25,letterSpacing:"-.1px"}}>{nd.label}</div>;
        })}
      </div>

      <div style={{position:"relative",height:30,width:"100%"}}>
        <div aria-hidden="true" style={{position:"absolute",left:"16.666%",right:"16.666%",top:"50%",transform:"translateY(-50%)",
          height:0,borderTop:`1px dotted ${GRAY5}`,zIndex:0,opacity:.9}}/>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",width:"100%",height:"100%",position:"relative",zIndex:1}}>
          {nodes.map((nd,idx)=>{
            const done=idx<activeIdx && !nd.muted;
            const current=idx===activeIdx && !nd.muted;
            const circleBg=current?"#fff":(done?GRAY7:GRAY8);
            const circleBorder=current?"#fff":(done?GRAY5:GRAY7);
            const circleColor=current?GRAY8:(done?"#fff":GRAY5);
            return <div key={`${nd.label}-circle`} style={{display:"flex",justifyContent:"center",alignItems:"center",minWidth:0}}>
              <span style={{width:30,height:30,borderRadius:"50%",display:"inline-flex",alignItems:"center",justifyContent:"center",
                background:circleBg,color:circleColor,border:`1.5px solid ${circleBorder}`,fontSize:12,fontWeight:800,boxSizing:"border-box",
                boxShadow:current?`0 0 0 5px rgba(255,255,255,.13)`:`0 0 0 4px ${GRAY8}`}}>{idx+1}</span>
            </div>;
          })}
        </div>
      </div>

      <div className="workflow-stage-status-row" style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",width:"100%",marginTop:13}}>
        {nodes.map((nd,idx)=><div key={`${nd.label}-status`} style={{display:"flex",alignItems:"center",justifyContent:"center",minWidth:0,minHeight:22,padding:"0 10px"}}>
          {renderStatus(nd,idx)}
        </div>)}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",width:"100%",marginTop:6}}>
        {nodes.map((nd,idx)=>{
          const future=idx>activeIdx && !nd.muted;
          return <div key={`${nd.label}-sub`} style={{minWidth:0,padding:"0 10px",fontSize:10,
            color:nd.muted?GRAY5:(future?GRAY5:GRAY4),fontWeight:400,textAlign:"center",lineHeight:1.35,minHeight:14}}><AutoFitText value={nd.sub||""} title={String(nd.sub||"")} baseFontSize={10} minFontSize={7} align="center" style={{color:nd.muted?GRAY5:(future?GRAY5:GRAY4),fontWeight:400,lineHeight:1.35}}/></div>;
        })}
      </div>
    </div>
  </div>;
}

// 첨부 버튼 시스템 기준 라벨 2종.
//   tonal   … 채워진 톤 배경 + 테두리 없음  → XRF 판정처럼 "결과 값"에 사용
//   outlined… 흰 배경 + 색 테두리          → 승인 상태처럼 "처리 상태"에 사용
function StatusPill({tone, label, title, minWidth=0, fixedWidth, fixedHeight, fixedFontSize, size="md", variant="tonal", shape="pill"}){
  const sm=size==="sm";
  const hasFixed=Number.isFinite(fixedWidth) && fixedWidth>0;
  const resolvedHeight=hasFixed?(Number.isFinite(fixedHeight)?fixedHeight:RESULT_LABEL_HEIGHT):(sm?22:26);
  const resolvedFontSize=hasFixed?(Number.isFinite(fixedFontSize)?fixedFontSize:RESULT_LABEL_FONT_SIZE):(sm?10.5:11.5);
  const base={display:"inline-flex",alignItems:"center",justifyContent:"center",
    width:hasFixed?fixedWidth:undefined,minWidth:hasFixed?fixedWidth:minWidth,maxWidth:hasFixed?fixedWidth:undefined,
    height:resolvedHeight,padding:hasFixed?0:(sm?"0 10px":"0 12px"),
    borderRadius:shape==="rounded"?6:UI.pill,fontSize:resolvedFontSize,
    fontWeight:600,whiteSpace:"nowrap",boxSizing:"border-box",lineHeight:1,letterSpacing:"-.03px",
    flex:hasFixed?"0 0 auto":undefined,overflow:"hidden",textOverflow:"ellipsis"};
  const skin=variant==="outlined"
    ? {background:C.card,border:`1px solid ${tone.solid||tone.tx}`,color:tone.tx}
    : {background:tone.bg,border:`1px solid ${tone.ln}`,color:tone.tx};
  return <span title={title||label} style={{...base,...skin}}><AutoFitText value={label} baseFontSize={resolvedFontSize} minFontSize={7} align="center" style={{fontWeight:600,lineHeight:1}}/></span>;
}

function ApprovalBadge({status, context, detail=""}) {
  const key=approvalFilterKey(status);
  const isDiscontinue=context==="discontinue";
  const detailText=String(detail||"").trim();
  const base={
    approved:{t:A.green,  label:"승인"},
    measuring:{t:A.blue,  label:"측정 진행중"},
    request:{t:A.amber,   label:"의뢰 필요"},
    rejected:{t:A.rose,   label:"반려"},
    hold:{t:A.amber,      label:"보류"},
    cancelled:{t:A.slate, label:"처리 취소"},
    unknown:{t:A.slate,   label:"확인 필요"}
  };
  const meta=base[key] || {t:A.slate,label:status||"—"};
  // 승인/보류/반려는 앞에 사용·단종 같은 대상 단어를 붙이지 않고 상태만 표기합니다.
  // 대상 구분이 필요하면 툴팁으로만 알려줍니다.
  const tip=detailText || (isDiscontinue?`단종 처리 · ${meta.label}`:`사용 여부 · ${meta.label}`);
  if(key==="request"){
    return <span className="approval-request-pill" title={tip} style={{display:"inline-flex",alignItems:"center",justifyContent:"center",width:APPROVAL_LABEL_WIDTH,minWidth:APPROVAL_LABEL_WIDTH,height:APPROVAL_LABEL_HEIGHT,padding:0,background:C.card,border:`1px solid ${meta.t.solid||meta.t.tx}`,borderRadius:6,color:meta.t.tx,fontSize:APPROVAL_LABEL_FONT_SIZE,fontWeight:600,whiteSpace:"nowrap",textAlign:"center",lineHeight:1,boxSizing:"border-box",overflow:"visible"}}>{meta.label}</span>;
  }
  return <StatusPill tone={meta.t} variant="outlined" shape="rounded" label={meta.label} title={tip} size="sm"
    fixedWidth={APPROVAL_LABEL_WIDTH} fixedHeight={APPROVAL_LABEL_HEIGHT} fixedFontSize={APPROVAL_LABEL_FONT_SIZE}/>;
}

function DetailArrowButton({active=false, title="상세 보기"}){
  const [pressed,setPressed]=useState(false);
  return <button type="button" title={title} aria-pressed={active}
    onPointerDown={()=>setPressed(true)}
    onPointerUp={()=>setPressed(false)}
    onPointerCancel={()=>setPressed(false)}
    onPointerLeave={()=>setPressed(false)}
    style={{width:26,height:26,background:active?C.charcoalDk:(pressed?C.bg:C.card),border:`1px solid ${active?C.charcoalDk:C.bd2}`,borderRadius:UI.rs,fontSize:12,cursor:"pointer",color:active?"#fff":C.text3,fontWeight:500,boxShadow:"none",transform:active?"translateX(2px)":(pressed?"scale(.94)":"none"),transition:"background .16s ease, border-color .16s ease, box-shadow .16s ease, transform .12s ease"}}>→</button>;
}
function PrecisionDetailPanel({selected, lang, card, inp, updatePrecisionOverride, markPrecisionRequested, reloadSharePointDb, onPhotoPreview}){
  const isRequested=selected.precision.requestStatus==="REQUESTED";
  const isConfirmed=selected.finalConfirm==="확인";
  const targets=selected.targets||[];
  const elementResults=selected.elementResults||selected.precision?.elementResults||{};
  const [storedFiles,setStoredFiles]=useState({source:null,report:null});
  const [fileBusy,setFileBusy]=useState("");
  const [fileProgress,setFileProgress]=useState(0);
  const [fileError,setFileError]=useState("");
  const [fileMetadataLoading,setFileMetadataLoading]=useState(true);
  const [handoffAuthOpen,setHandoffAuthOpen]=useState(false);
  const [handoffPassword,setHandoffPassword]=useState("");
  const [handoffBusy,setHandoffBusy]=useState(false);
  const [handoffError,setHandoffError]=useState("");
  useEffect(()=>{
    let active=true;
    setStoredFiles({source:null,report:null});
    setFileMetadataLoading(true);
    setFileError("");
    fetchPrecisionStoredFiles(selected.key).then(files=>{
      if(active) setStoredFiles(files);
    }).catch(error=>{
      if(active) setFileError(error?.message||String(error));
    }).finally(()=>{if(active) setFileMetadataLoading(false);});
    return ()=>{active=false;};
  },[selected.key]);
  useEffect(()=>{
    setHandoffAuthOpen(false);
    setHandoffPassword("");
    setHandoffBusy(false);
    setHandoffError("");
  },[selected.key]);
  const resultFileName=storedFiles.report?.name||String(selected.precision?.resultFileId||"").trim();
  const sourceFileName=storedFiles.source?.name||String(selected.sourceFileName||"").trim();
  const hasPrecisionReportPdf=!!storedFiles.report?.id;
  const hasSourceFile=!!storedFiles.source?.id;
  const canUploadSourceFile=isRequested && !fileBusy && !fileMetadataLoading;
  const canUploadPrecisionPdf=isRequested && hasSourceFile && !fileBusy;
  const canEditPrecisionResult=isRequested && hasPrecisionReportPdf && !fileBusy;
  const derivedResult=derivePrecisionResult(targets,selected.precision,elementResults);
  const infoRows=[
    ["XRF 원본 파일",sourceFileName||"—"],
    ["XRF 측정 ID",selected.measurement?.id||"—"],
    ["XRF Judgment",measurementJudgementSummary(selected.measurement)],
    ["대상 원소",targets.join(", ")||"—"],
    ["인계 상태",isRequested?"인계 완료":"인계 필요"],
    ["인계일",selected.precision.requestedAt||"—"],
    ["담당자",selected.precision.requestedBy||"—"]
  ];
  const filePickStyle=(enabled=true)=>({marginTop:5,display:"grid",gridTemplateColumns:"24px minmax(0,1fr) 24px",alignItems:"center",width:"100%",maxWidth:"100%",minWidth:0,minHeight:36,padding:"0 8px",border:`1px solid ${enabled?A.blue.solid:C.bd2}`,background:enabled?A.blue.solid:C.alt,color:enabled?"#fff":C.text4,fontSize:11,fontWeight:650,cursor:enabled?"pointer":"not-allowed",boxSizing:"border-box",borderRadius:UI.rs,opacity:enabled?1:.62,overflow:"hidden",transition:"background .15s ease,border-color .15s ease,opacity .15s ease"});
  const UploadArrow=()=> <svg data-upload-arrow="true" aria-hidden="true" viewBox="0 0 24 24" width="20" height="20" style={{display:"block",justifySelf:"center"}}><path d="M12 19V5M7.5 9.5 12 5l4.5 4.5" fill="none" stroke="currentColor" strokeWidth="1.35" strokeLinecap="round" strokeLinejoin="round"/></svg>;
  const fileNameStyle=(has)=>({marginTop:5,display:"block",width:"100%",maxWidth:"100%",minWidth:0,fontSize:10,color:has?C.text2:C.text4,fontWeight:500,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",boxSizing:"border-box"});
  const updateElement=(el,patch)=>{
    if(!canEditPrecisionResult) return;
    const prevState=precisionElementState(selected.precision,elementResults,el);
    const nextElement={...prevState,...(elementResults?.[el]||{}),...patch};
    if(nextElement.presence==="미함유") nextElement.ppm="";
    const nextResults={...elementResults,[el]:nextElement};
    updatePrecisionOverride(selected.key,{elementResults:nextResults,finalConfirm:"",confirmedAt:"",confirmedBy:""});
  };
  const handleFileUpload=async(kind,file)=>{
    if(!file || !isRequested || fileBusy || selected.item?.__visualUat) return;
    setFileBusy(kind);
    setFileProgress(0);
    setFileError("");
    try{
      const metadata=await uploadPrecisionBinary(selected.key,kind,file,setFileProgress);
      setStoredFiles(prev=>({...prev,[kind]:metadata}));
      await reloadSharePointDb();
    }catch(error){
      setFileError(error?.message||String(error));
    }finally{
      setFileBusy("");
    }
  };
  const handlePrecisionReportUpload=(e)=>{
    const file=e.target.files?.[0];
    if(!file) return;
    if(!isRequested){
      window.alert("정밀분석 인계를 먼저 완료한 뒤 성적서 PDF를 업로드하세요.");
      e.target.value="";
      return;
    }
    if(!hasSourceFile){
      window.alert("XRF 원본 파일을 먼저 선택하세요.");
      e.target.value="";
      return;
    }
    const isPdf=/\.pdf$/i.test(String(file.name||"")) || String(file.type||"").toLowerCase()==="application/pdf";
    if(!isPdf){
      window.alert("정밀분석 성적서는 PDF 파일만 업로드할 수 있습니다.");
      e.target.value="";
      return;
    }
    void handleFileUpload("report",file);
    e.target.value="";
  };
  const submitPrecisionHandoff=async(event)=>{
    event.preventDefault();
    if(!handoffPassword || handoffBusy) return;
    setHandoffBusy(true);
    setHandoffError("");
    const completed=await markPrecisionRequested(selected.key,handoffPassword);
    setHandoffBusy(false);
    setHandoffPassword("");
    if(completed){
      setHandoffAuthOpen(false);
      return;
    }
    setHandoffError(lang==="en"?"Check the password and try again.":"비밀번호를 확인한 뒤 다시 시도하세요.");
  };
  const resultExplain=!isRequested
    ? "정밀분석 인계 버튼을 먼저 클릭해야 합니다. 인계 후 성적서 PDF를 업로드하면 결과 입력이 활성화됩니다."
    : !hasPrecisionReportPdf
      ? "정밀분석 성적서 PDF를 업로드해야 원소별 결과를 입력할 수 있습니다."
      : derivedResult==="OK"
        ? "모든 대상 원소가 미함유이거나 법적 기준치 이내입니다."
        : derivedResult==="NG"
          ? "법적 기준치를 초과한 대상 원소가 있습니다."
          : "성적서 내용을 기준으로 대상 원소의 함유 여부와 함유량을 입력하면 정밀 결과가 자동 산출됩니다.";
  return <><div style={{...card,overflow:"hidden"}}>
    <div className="xrf-hero" style={{background:C.charcoalDk}}>
      {selected.item?.photoFileUrl
        ? <button type="button" onClick={()=>onPhotoPreview?.({url:selected.item.photoFileUrl,name:selected.item.photoFileName||"품목 사진",code:selected.item.code,itemName:displayItemName(selected.item,lang)})} title={selected.item.photoFileName||"품목 사진"} aria-label={`${selected.item.code} 품목 사진 크게 보기`} style={{display:"inline-flex",alignSelf:"center",padding:0,border:0,borderRadius:10,background:"transparent",cursor:"zoom-in"}}><img className="xrf-hero-photo" src={selected.item.photoFileUrl} alt={selected.item.photoFileName||`${selected.item.code} 품목 사진`}/></button>
        : <div className="xrf-hero-photo" style={{display:"flex",alignItems:"center",justifyContent:"center",color:"rgba(255,255,255,.45)",fontSize:10,textAlign:"center",padding:6,boxSizing:"border-box"}}>사진 없음</div>}
      <div className="xrf-hero-title">
        <AutoFitText value={displayItemName(selected.item,lang)} title={displayItemName(selected.item,lang)} baseFontSize={16} minFontSize={8} align="left" style={{fontWeight:600,color:"#fff",letterSpacing:"-.2px"}}/>
        <AutoFitText value={`${selected.item.code} · ${selected.item.dept} · ${selected.measurement?.date||selected.measurement?.measuredDate||"측정일 미확인"}`} baseFontSize={12} minFontSize={7} align="left" style={{color:"rgba(255,255,255,.72)",marginTop:3,fontWeight:400}}/>
      </div>
      {/* XRF 분석 상단과 동일한 구분선형 정보표 디자인을 사용합니다. */}
      <div className="xrf-hero-meta">
        {[
          ["인계 사유", selected.triggerLabel||"—"],
          ["대상 원소", targets.join(", ")||"—"],
          ["인계 상태", isRequested?"인계 완료":"인계 필요"],
          ["성적서", hasPrecisionReportPdf?"PDF 등록":"미등록"],
          ["정밀 결과", derivedResult?<AnalysisResultBadge value={derivedResult}/>:"—"]
        ].map(([k,v],i,arr)=>(
          <div key={k} style={{padding:"8px 14px",minWidth:92,display:"flex",flexDirection:"column",
            alignItems:"center",justifyContent:"center",gap:5,
            borderRight:i<arr.length-1?"1px solid rgba(255,255,255,.14)":"none"}}>
            <div style={{fontSize:11,color:"rgba(255,255,255,.74)",fontWeight:500,whiteSpace:"nowrap"}}>{k}</div>
            <div style={{width:"100%",minWidth:0,fontSize:12,fontWeight:650,color:"#fff",whiteSpace:"nowrap",display:"flex",alignItems:"center",justifyContent:"center"}}>
              {typeof v==="string" || typeof v==="number"
                ? <AutoFitText value={v} baseFontSize={12} minFontSize={7} align="center" style={{fontWeight:650,color:"#fff"}}/>
                : v}
            </div>
          </div>
        ))}
      </div>
    </div>
    <div className="xrf-detail-grid">
      <div style={{padding:"16px 20px",borderRight:`1px solid ${C.bd}`}}>
        <div style={{...T.section,marginBottom:12}}>정밀분석 인계 관리</div>
        {infoRows.map(([k,v])=><div key={k} style={{display:"flex",justifyContent:"space-between",padding:"8px 0",borderBottom:`1px solid ${C.bd}`,fontSize:11,gap:12}}>
          <span style={{color:C.text3,flex:"0 0 auto"}}>{k}</span><span style={{fontWeight:600,color:C.text1,textAlign:"right",minWidth:0,flex:"1 1 auto"}}><AutoFitText value={v} title={String(v)} baseFontSize={11} minFontSize={7} align="right" style={{fontWeight:600,color:C.text1}}/></span>
        </div>)}
        <button disabled={isRequested} onClick={()=>{if(!isRequested){setHandoffPassword("");setHandoffError("");setHandoffAuthOpen(true);}}} className="pill-btn" style={{marginTop:12,width:"100%",padding:"10px",border:`1px solid ${isRequested?A.green.ln:A.blue.solid}`,background:isRequested?A.green.bg:A.blue.solid,color:isRequested?A.green.tx:"#fff",borderRadius:UI.rs,fontSize:12,fontWeight:600,cursor:isRequested?"default":"pointer"}}>
          {isRequested?"정밀분석 인계 완료":"정밀분석 인계"}
        </button>
      </div>
      <div style={{padding:"16px 20px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,marginBottom:12}}>
          <div>
            <div style={{...T.section}}>상세 확인</div>
            <div style={{fontSize:11,color:C.text3,marginTop:3}}>인계 완료 후 정밀분석 성적서 PDF를 등록해야 원소별 결과 입력이 활성화됩니다.</div>
          </div>
        </div>

        <div style={{marginBottom:12,padding:"10px 12px",border:`1px solid ${canEditPrecisionResult?A.green.ln:C.bd}`,background:canEditPrecisionResult?A.green.bg:C.alt,borderRadius:8,fontSize:11,color:canEditPrecisionResult?A.green.tx:C.text3,lineHeight:1.55}}>
          {!isRequested
            ? "1. 정밀분석 인계 버튼을 클릭하세요.  2. 성적서 PDF를 업로드하세요.  3. 원소별 결과를 입력하세요."
            : !hasPrecisionReportPdf
              ? "인계가 완료되었습니다. 정밀분석 성적서 PDF를 업로드하면 원소별 결과 입력이 가능합니다."
              : `성적서 등록 완료: ${resultFileName} · 원소별 결과를 입력할 수 있습니다.`}
        </div>

        <div style={{display:"grid",gridTemplateColumns:"minmax(0,1fr) minmax(0,1fr)",gap:10,marginBottom:12,minWidth:0}}>
          <div style={{fontSize:11,color:C.text2,fontWeight:600,minWidth:0,overflow:"hidden"}}>XRF 원본 파일<label style={filePickStyle(canUploadSourceFile)}><span aria-hidden="true"/><span style={{textAlign:"center"}}>파일 선택</span><UploadArrow/><input type="file" accept=".xlsx,.xlsm,.xls,.csv,.json,.pdf" disabled={!canUploadSourceFile} onChange={e=>{const file=e.target.files?.[0];if(file&&canUploadSourceFile) void handleFileUpload("source",file);e.target.value="";}} style={{display:"none"}}/></label><div title={sourceFileName||""} style={fileNameStyle(hasSourceFile)}>{hasSourceFile?<a href={storedFiles.source.webUrl||undefined} target="_blank" rel="noopener noreferrer" style={{color:C.text2}}>{sourceFileName}</a>:(sourceFileName?`${sourceFileName} · 원본 파일 재업로드 필요`:(!isRequested?"인계 완료 후 선택 가능":"업로드된 파일 없음"))}</div></div>
          <div style={{fontSize:11,color:C.text2,fontWeight:600,minWidth:0,overflow:"hidden"}}>정밀분석 성적서 PDF<label style={filePickStyle(canUploadPrecisionPdf)}><span aria-hidden="true"/><span style={{textAlign:"center"}}>PDF 선택</span><UploadArrow/><input type="file" accept=".pdf,application/pdf" disabled={!canUploadPrecisionPdf} onChange={handlePrecisionReportUpload} style={{display:"none"}}/></label><div title={resultFileName||""} style={fileNameStyle(hasPrecisionReportPdf)}>{hasPrecisionReportPdf?<a href={storedFiles.report.webUrl||undefined} target="_blank" rel="noopener noreferrer" style={{color:C.text2}}>{resultFileName}</a>:(resultFileName?`${resultFileName} · PDF 재업로드 필요`:(!isRequested?"인계 완료 후 업로드 가능":(!hasSourceFile?"XRF 원본 파일 선택 후 업로드 가능":"업로드된 PDF 없음")))}</div></div>
        </div>
        {fileBusy&&<div role="status" style={{marginBottom:12,fontSize:11,color:A.blue.solid}}>{fileBusy==="source"?"XRF 원본":"정밀분석 PDF"} 파일 업로드 중 · {fileProgress}%</div>}
        {fileError&&<div role="alert" style={{marginBottom:12,fontSize:11,color:C.redDk}}>{fileError}</div>}

        <div style={{border:`1px solid ${C.bd}`,borderRadius:8,overflow:"hidden",marginBottom:12,opacity:canEditPrecisionResult?1:.58}}>
          <div style={{padding:"9px 11px",background:C.bg,borderBottom:`1px solid ${C.bd}`,display:"flex",justifyContent:"space-between",gap:10,alignItems:"center",flexWrap:"wrap"}}>
            <div>
              <div style={{...T.section}}>원소 함유 결과</div>
              <div style={{fontSize:10,color:C.text3,marginTop:2}}>미함유 또는 함유량 ≤ 법적 기준치 → OK · 법적 기준치 초과 → NG</div>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:7}}>
              <span style={{fontSize:10,color:C.text3}}>정밀 결과</span>
              {derivedResult?<AnalysisResultBadge value={derivedResult}/>:<span style={{fontSize:12,color:C.text4}}>—</span>}
            </div>
          </div>
          <div className="responsive-table-scroll table-scroll-medium" role="region" aria-label="정밀분석 원소별 결과 표" tabIndex={0}>
            <table style={{width:"100%",minWidth:0,maxWidth:"100%",borderCollapse:"collapse",tableLayout:"fixed",fontSize:11}}>
              <colgroup>
                <col style={{width:"13%"}}/><col style={{width:"25%"}}/><col style={{width:"22%"}}/><col style={{width:"22%"}}/><col style={{width:"18%"}}/>
              </colgroup>
              <thead><tr>{["Element","원소 함유 결과","함유량 (ppm)","법적 기준치 (ppm)","정밀 결과"].map(h=><th key={h} style={{padding:"8px 5px",background:C.card,borderBottom:`1px solid ${C.bd}`,color:C.text2,fontWeight:600,textAlign:"center",whiteSpace:"normal",lineHeight:1.35}}>{h}</th>)}</tr></thead>
              <tbody>{targets.length?targets.map(el=>{
                const state=precisionElementState(selected.precision,elementResults,el);
                const decision=precisionElementDecision(state);
                return <tr key={el}>
                  <td style={{padding:"9px",borderTop:`1px solid ${C.bd}`,fontWeight:700,color:C.text1,textAlign:"center"}}>{el}</td>
                  <td style={{padding:"8px",borderTop:`1px solid ${C.bd}`,textAlign:"center"}}>
                    <select disabled={!canEditPrecisionResult} value={state.presence||""} onChange={e=>updateElement(el,{presence:e.target.value})} style={{...inp,width:"100%",minWidth:0,padding:"7px 6px",boxSizing:"border-box",opacity:canEditPrecisionResult?1:.65,cursor:canEditPrecisionResult?"pointer":"not-allowed"}}>
                      <option value="">선택</option><option value="미함유">미함유</option><option value="함유">함유</option>
                    </select>
                  </td>
                  <td style={{padding:"8px",borderTop:`1px solid ${C.bd}`,textAlign:"center"}}>
                    {state.presence==="미함유"
                      ? <span style={{fontSize:11,color:C.text3}}>미함유</span>
                      : <input type="number" min="0" step="0.1" disabled={!canEditPrecisionResult || state.presence!=="함유"} value={state.presence==="함유"?(state.ppm??""):""} onChange={e=>updateElement(el,{ppm:e.target.value})} placeholder={!canEditPrecisionResult?"성적서 필요":(state.presence==="함유"?"ppm 입력":"함유 선택")} style={{...inp,width:"100%",minWidth:0,padding:"7px 6px",boxSizing:"border-box",opacity:canEditPrecisionResult&&state.presence==="함유"?1:.55,cursor:canEditPrecisionResult&&state.presence==="함유"?"text":"not-allowed"}}/>}
                  </td>
                  <td style={{padding:"9px",borderTop:`1px solid ${C.bd}`,textAlign:"center",fontWeight:600,color:C.text2}}>{state.legalLimit==null?"—":Number(state.legalLimit).toLocaleString()}</td>
                  <td style={{padding:"9px",borderTop:`1px solid ${C.bd}`,textAlign:"center"}}>{decision?<AnalysisResultBadge value={decision}/>:<span style={{fontSize:12,color:C.text4}}>—</span>}</td>
                </tr>;
              }):<tr><td colSpan={5} style={{padding:20,textAlign:"center",color:C.text4}}>정밀분석 대상 원소가 없습니다.</td></tr>}</tbody>
            </table>
          </div>
          <div style={{padding:"9px 11px",borderTop:`1px solid ${C.bd}`,fontSize:11,color:derivedResult==="NG"?C.redDk:C.text3,background:derivedResult==="NG"?C.redBg:C.alt}}>{resultExplain}</div>
        </div>

        <button disabled={!canEditPrecisionResult || !derivedResult} onClick={()=>canEditPrecisionResult&&derivedResult&&updatePrecisionOverride(selected.key,{finalConfirm:"확인",confirmedAt:TODAY_STR,confirmedBy:"admin"})} className="pill-btn" style={{width:"100%",padding:"11px 12px",border:`1px solid ${isConfirmed?A.green.ln:(canEditPrecisionResult&&derivedResult?A.blue.solid:C.bd2)}`,background:isConfirmed?A.green.bg:(canEditPrecisionResult&&derivedResult?A.blue.solid:C.alt),color:!canEditPrecisionResult||!derivedResult?C.text4:(isConfirmed?A.green.tx:"#fff"),borderRadius:UI.rs,fontSize:12,fontWeight:600,opacity:canEditPrecisionResult&&derivedResult?1:.6,cursor:canEditPrecisionResult&&derivedResult?"pointer":"not-allowed"}}>
          {isConfirmed?"정밀분석 결과 확인 완료":"정밀분석 결과 확인"}
        </button>
      </div>
    </div>
  </div>
  {handoffAuthOpen&&<div role="presentation" onMouseDown={()=>{if(!handoffBusy)setHandoffAuthOpen(false);}} style={{position:"fixed",inset:0,zIndex:1200,display:"flex",alignItems:"center",justifyContent:"center",padding:20,background:"rgba(15,23,32,.58)",backdropFilter:"blur(2px)"}}>
    <form role="dialog" aria-modal="true" aria-labelledby="precision-handoff-auth-title" onSubmit={submitPrecisionHandoff} onMouseDown={event=>event.stopPropagation()} style={{width:"min(380px,100%)",padding:22,background:C.card,border:`1px solid ${C.bd}`,borderRadius:12,boxShadow:"0 18px 48px rgba(0,0,0,.28)"}}>
      <div id="precision-handoff-auth-title" style={{fontSize:16,fontWeight:750,color:C.text1}}>{lang==="en"?"Precision Transfer Authorization":"정밀분석 인계 인증"}</div>
      <div style={{marginTop:7,fontSize:11,color:C.text3,lineHeight:1.6}}>{lang==="en"?"Enter the administrator password to continue the transfer.":"정밀분석 인계를 진행하려면 관리자 비밀번호를 입력하세요."}</div>
      <input type="password" autoFocus autoComplete="off" value={handoffPassword} disabled={handoffBusy} onChange={event=>{setHandoffPassword(event.target.value);setHandoffError("");}} placeholder={lang==="en"?"Password":"비밀번호"} style={{...inp,width:"100%",marginTop:14,boxSizing:"border-box"}}/>
      {handoffError&&<div role="alert" style={{marginTop:8,fontSize:11,color:C.redDk}}>{handoffError}</div>}
      <div style={{display:"flex",justifyContent:"flex-end",gap:8,marginTop:16}}>
        <button type="button" disabled={handoffBusy} onClick={()=>{setHandoffAuthOpen(false);setHandoffPassword("");setHandoffError("");}} style={{padding:"8px 14px",border:`1px solid ${C.bd2}`,borderRadius:UI.rs,background:C.card,color:C.text2,fontWeight:650,cursor:handoffBusy?"wait":"pointer"}}>{lang==="en"?"Cancel":"취소"}</button>
        <button type="submit" disabled={!handoffPassword||handoffBusy} style={{padding:"8px 15px",border:`1px solid ${C.charcoalDk}`,borderRadius:UI.rs,background:C.charcoalDk,color:"#fff",fontWeight:650,opacity:(!handoffPassword||handoffBusy)?0.58:1,cursor:!handoffPassword||handoffBusy?"not-allowed":"pointer"}}>{handoffBusy?(lang==="en"?"Checking...":"확인 중..."):(lang==="en"?"Transfer":"인계")}</button>
      </div>
    </form>
  </div>}
  </>;
}

/* ══════════════════════════════════════════════════════════════════════════════
   원소별 XRF 추세 차트 (XRF Trend Chart)

   - 가로축: 실제 측정이 완료된 관리주기(R / P1 / P2 ... / F1) 전체.
     같은 주기의 복수 측정과 Retest는 해당 주기의 마지막 유효값으로 대표합니다.
   - 세로축: Content(ppm). Legal Limit(법적 규제)과 Internal Limit(법적 × 70%)
     두 기준선을 항상 함께 그립니다.
   - 상세 화면과 트렌드 모두 동일한 주기 대표 Measurement를 사용합니다.
   - N.D.는 0 ppm 위치에 속이 빈 마커, 결측/재측정(??)은 선을 끊어서 표시합니다.

   의존: ELEMENTS, XRF_ELEMENT_LIMITS, XRF_INTERNAL_LIMITS, XRF_INTERNAL_LIMIT_RATIO,
        C, T, A, UI, FONT_SANS, xrfNumber, sourcePpmValue, elementDataMissing,
        isXrfContentND, isXrfRemeasureRequired, xrfJudgementDecision,
        xrfLegalLimitFromData, xrfInternalLimitFromData, periodLabel,
        periodMeasuredDates, periodMeasurementIds, measurementById,
        measurementAttemptsForPeriod
   ══════════════════════════════════════════════════════════════════════════════ */

const XRF_TREND_ELEMENTS = ELEMENTS;                 // Pb, Hg, Cr, Cd, Cl, Br, Cl+Br
const XRF_TREND_EL_NAMES = {
  Pb:{ko:"납",en:"Lead"},Hg:{ko:"수은",en:"Mercury"},Cr:{ko:"크롬",en:"Chromium"},
  Cd:{ko:"카드뮴",en:"Cadmium"},Cl:{ko:"염소",en:"Chlorine"},Br:{ko:"브롬",en:"Bromine"},
  "Cl+Br":{ko:"염소+브롬",en:"Chlorine + Bromine"}
};

// 기준선 색: Legal = 실선 red(법적 상한), Internal = 파선 amber(사내 판정선)
const XRF_TREND_LEGAL_COLOR    = C.red;
const XRF_TREND_INTERNAL_COLOR = C.xWarn;
const XRF_TREND_LINE_COLOR     = C.blue;

function xrfTrendFmt(v){
  return v==null ? "—" : Number(v).toLocaleString(undefined,{maximumFractionDigits:1});
}

// 한 주기의 대표 측정 = 그 주기에 연결된 마지막 시도(Retest 포함).
// 상세 화면과 트렌드가 반드시 같은 Measurement를 읽도록 공용으로 사용합니다.
function xrfTrendMeasurement(item, period){
  if(!item || !period) return null;
  const attempts=measurementAttemptsForPeriod(item,period);
  const last=attempts[attempts.length-1]||null;
  const id=String(
    last?.id || last?.measurementId
    || period?.measurementId
    || periodMeasurementIds(period).slice(-1)[0]
    || ""
  ).trim();
  return id ? measurementById(item,id) : null;
}

// 트렌드 가로축은 Measurement 시도 횟수가 아니라 업무 관리주기(R/P1/P2...)입니다.
// 같은 주기의 재측정/복수 측정은 마지막 유효 Measurement 하나로 대표해 주기 라벨이
// R #1, R #2처럼 변형되지 않도록 합니다. 개별 시도 이력은 주기별 산정 요약에서 확인합니다.
function xrfTrendMeasurementEntries(item, periods){
  if(!item) return [];
  const rows=[];
  const safePeriods=periods||[];
  safePeriods.forEach(period=>{
    const attempts=measurementAttemptsForPeriod(item,period);
    const measurement=xrfTrendMeasurement(item,period);
    if(!measurement) return;
    rows.push({
      measurement,
      period,
      attempts,
      num:Number(period.num),
      baseLabel:periodLabel(period),
      label:periodLabel(period),
      total:attempts.length||1,
    });
  });

  rows.sort((a,b)=>{
    const periodA=Number.isFinite(Number(a.period?.displaySeq))?Number(a.period.displaySeq):Number(a.period?.num??999999);
    const periodB=Number.isFinite(Number(b.period?.displaySeq))?Number(b.period.displaySeq):Number(b.period?.num??999999);
    return periodA-periodB
      || String(a.measurement?.date||"").localeCompare(String(b.measurement?.date||""))
      || Number(a.measurement?.attemptNo||1)-Number(b.measurement?.attemptNo||1);
  });

  return rows;
}

// 실제 Measurement 하나 → 원소 하나의 추세 포인트
function xrfTrendPoint(entry, element){
  const measurement=entry?.measurement||null;
  const period=entry?.period||null;
  const d=measurement?.elements?.[element] || null;
  const rawPpm=d ? sourcePpmValue(d) : null;
  const remeasure=d ? isXrfRemeasureRequired(d) : false;
  const missing=!d || elementDataMissing(d);
  const isND=d ? (isXrfContentND(rawPpm) || !!d.isND) : false;
  // xrfNumber()는 N.D.를 0으로 환산합니다. 결측(----, 공백)은 null로 남깁니다.
  const ppm=missing ? null : xrfNumber(rawPpm);
  // EBar와 동일하게, 한도를 해석해 넣은 뒤 판정합니다(구형 데이터 호환).
  const legalLimit=(d ? xrfLegalLimitFromData(d) : null) ?? XRF_ELEMENT_LIMITS[element] ?? null;
  const internalLimit=(d ? xrfInternalLimitFromData(d) : null) ?? XRF_INTERNAL_LIMITS[element] ?? null;
  const judgement=d ? xrfJudgementDecision({...d,legalLimit,internalLimit}) : null;
  return {
    num:entry?.num,
    label:entry?.label||periodLabel(period),
    baseLabel:entry?.baseLabel||periodLabel(period),
    // 날짜와 값은 동일한 Measurement 객체에서만 읽어 서로 어긋날 수 없게 합니다.
    date:dateOnly(measurement?.date||measurement?.measuredDate)||null,
    measurementId:measurement?.id||measurement?.measurementId||null,
    role:measurement?.role || null,
    attemptNo:Number(measurement?.attemptNo||1),
    rawPpm,
    ppm,
    isND,
    missing,
    remeasure,
    judgement,
    legalLimit,
    internalLimit,
    hasValue:!missing && ppm!=null,
  };
}

function xrfTrendSeries(item, periods, element, entries=null){
  const source=entries||xrfTrendMeasurementEntries(item,periods);
  return source.map(entry=>xrfTrendPoint(entry,element));
}

// 선택한 주기보다 뒤의 측정 이력은 현재 선택 화면의 트렌드에 포함하지 않습니다.
// 선택 주기까지의 R/Pn 흐름만 사용해 타임라인 선택과 차트 범위를 일치시킵니다.
function xrfTrendPeriodsThroughSelection(periods, selectedPeriod){
  const rows=Array.isArray(periods)?periods:[];
  if(!selectedPeriod) return rows;
  const orderOf=period=>{
    const displaySeq=Number(period?.displaySeq);
    if(Number.isFinite(displaySeq)) return displaySeq;
    const num=Number(period?.num);
    return Number.isFinite(num)?num:Number.POSITIVE_INFINITY;
  };
  const selectedOrder=orderOf(selectedPeriod);
  return rows.filter(period=>orderOf(period)<=selectedOrder);
}

// 축 눈금을 1 / 2 / 5 × 10ⁿ 로 정돈합니다.
function xrfTrendNiceStep(range, target=4){
  if(!(range>0)) return 1;
  const rough=range/Math.max(1,target);
  const mag=Math.pow(10,Math.floor(Math.log10(rough)));
  const norm=rough/mag;
  const unit=norm<1.5?1:norm<3?2:norm<7?5:10;
  return unit*mag;
}
function xrfTrendTicks(max, target=4){
  const step=xrfTrendNiceStep(max,target);
  const out=[];
  for(let v=0; v<=max+step*0.001; v+=step) out.push(Math.round(v*1000)/1000);
  return out.length>1 ? out : [0,max];
}

// scaleMode:
//   "limit" → Legal Limit까지 모두 보이는 축 (기본값, 사용자가 요청한 "두 기준선이 모두 보이는" 축)
//   "data"  → 측정값 기준 확대. 기준선이 축을 벗어나면 상단에 ↑ 표시로 알립니다.
function xrfTrendAxisMax(series, element, scaleMode){
  const legal=series.find(p=>p.legalLimit!=null)?.legalLimit ?? XRF_ELEMENT_LIMITS[element] ?? 100;
  const internal=series.find(p=>p.internalLimit!=null)?.internalLimit ?? XRF_INTERNAL_LIMITS[element] ?? legal*XRF_INTERNAL_LIMIT_RATIO;
  const dataMax=series.reduce((mx,p)=>p.hasValue && p.ppm>mx ? p.ppm : mx, 0);
  if(scaleMode==="data"){
    // 전부 N.D.(0)이면 확대할 대상이 없으므로 사내 기준선 기준 축으로 되돌립니다.
    const base=dataMax>0 ? dataMax*1.35 : internal*1.15;
    return {max:Math.max(base,1), legal, internal};
  }
  return {max:Math.max(legal*1.14, dataMax*1.15, 1), legal, internal};
}

function xrfTrendPointColor(p){
  if(p.remeasure) return C.xWarn;
  if(p.judgement==="NG") return C.red;
  if(p.judgement==="OK") return C.xOk;
  return C.text4;
}

/* ── 플롯 본체 (소형 멀티플/확대 보기 공용) ───────────────────────────────── */
function XrfTrendPlot({
  element, series, scaleMode="limit", compact=false,
  selectedPeriodNum=null, onSelectPeriod=null, lang="ko",
}){
  const [hoverIdx,setHoverIdx]=useState(null);

  const n=series.length;
  const W=compact?330:780;
  const H=compact?196:366;
  const PAD=compact?{t:24,r:14,b:30,l:46}:{t:30,r:20,b:74,l:66};
  const inset=compact?10:18;

  const plotL=PAD.l, plotR=W-PAD.r, plotT=PAD.t, plotB=H-PAD.b;
  const plotW=plotR-plotL, plotH=plotB-plotT;

  const {max:axisMax,legal,internal}=useMemo(
    ()=>xrfTrendAxisMax(series,element,scaleMode),[series,element,scaleMode]);

  const xAt=i=>{
    if(n<=1) return plotL+plotW/2;
    const left=plotL+inset, right=plotR-inset;
    return left+(right-left)*(i/(n-1));
  };
  const yAt=v=>plotB-(Math.max(0,Math.min(v,axisMax))/axisMax)*plotH;
  const ticks=useMemo(()=>xrfTrendTicks(axisMax,compact?3:4),[axisMax,compact]);

  // 값이 있는 주기만 이어 하나의 선으로 그립니다.
  // 결측/재측정(??) 주기는 점을 찍지 않고 건너뛰되, 선은 끊지 않고 연결합니다.
  const linePoints=useMemo(()=>(
    series.map((p,i)=>p.hasValue?[xAt(i),yAt(p.ppm)]:null).filter(Boolean)
  ),[series,axisMax,W,H,compact]);

  const legalY=legal!=null?yAt(legal):null;
  const internalY=internal!=null?yAt(internal):null;
  const legalOff=legal!=null && legal>axisMax;       // 확대 보기에서 축을 벗어난 경우
  const internalOff=internal!=null && internal>axisMax;

  const measuredCount=series.filter(p=>p.hasValue).length;
  const hovered=hoverIdx!=null?series[hoverIdx]:null;

  const axisFont=compact?9:11;
  // 모든 측정점과 실선은 한 화면에 유지합니다. 주기가 화면 수용 범위를 넘으면
  // 가로축 텍스트를 일부만 남기지 않고 전부 숨겨 잘못 읽히는 것을 방지합니다.
  const denseAxis=n>(compact?9:14);
  const showAxisLabel=()=>!denseAxis;

  // 선 두께는 얇게 유지하고, 색 블록 대신 선과 점만으로 상태를 읽히게 합니다.
  const gridW=0.6;
  const limitW=compact?0.9:1;
  const lineW=compact?1.2:1.5;
  const dotR=compact?2.6:3.2;

  return (
    <div className="xrf-trend-view" style={{width:"100%",overflow:"visible",paddingBottom:4}} data-i18n-skip="true">
      <div style={{position:"relative",width:"100%",minWidth:0}}>
      <svg viewBox={`0 0 ${W} ${H}`} width="100%" height="auto"
        role="img" aria-label={lang==="en"?`${element} XRF trend by period`:`${element} 주기별 XRF 추세`}
        style={{display:"block",overflow:"visible",fontFamily:FONT_SANS}}>

        {/* 가로 그리드 */}
        {ticks.map(t=>(
          <g key={`t-${t}`}>
            <line x1={plotL} x2={plotR} y1={yAt(t)} y2={yAt(t)}
              stroke={C.bd} strokeWidth={gridW}/>
            <text x={plotL-7} y={yAt(t)+3} textAnchor="end"
              fontSize={axisFont} fontWeight={400} fill={C.text4}>{xrfTrendFmt(t)}</text>
          </g>
        ))}

        {/* 기준선 바닥 */}
        <line x1={plotL} x2={plotR} y1={plotB} y2={plotB} stroke={C.bd2} strokeWidth={gridW}/>

        {/* 선택된 주기 */}
        {series.map((p,i)=>(
          (selectedPeriodNum!=null && Number(p.num)===Number(selectedPeriodNum))
            ? <line key={`sel-${i}`} x1={xAt(i)} x2={xAt(i)} y1={plotT} y2={plotB}
                stroke={C.bd2} strokeWidth={0.8} strokeDasharray="2 3"/>
            : null
        ))}

        {/* Legal Limit */}
        {legalY!=null&&!legalOff&&(
          <g>
            <line x1={plotL} x2={plotR} y1={legalY} y2={legalY}
              stroke={XRF_TREND_LEGAL_COLOR} strokeWidth={limitW} opacity={.85}/>
            <text x={plotR} y={Math.max(plotT+9,legalY-5)} textAnchor="end"
              fontSize={axisFont} fontWeight={500} fill={XRF_TREND_LEGAL_COLOR}
              stroke={C.card} strokeWidth={3} paintOrder="stroke">
              Legal {xrfTrendFmt(legal)}
            </text>
          </g>
        )}
        {/* Internal Limit */}
        {internalY!=null&&!internalOff&&(
          <g>
            <line x1={plotL} x2={plotR} y1={internalY} y2={internalY}
              stroke={XRF_TREND_INTERNAL_COLOR} strokeWidth={limitW}
              strokeDasharray="4 4" opacity={.85}/>
            <text x={plotR} y={Math.max(plotT+9,internalY-5)} textAnchor="end"
              fontSize={axisFont} fontWeight={500} fill={XRF_TREND_INTERNAL_COLOR}
              stroke={C.card} strokeWidth={3} paintOrder="stroke">
              Internal {xrfTrendFmt(internal)}
            </text>
          </g>
        )}
        {/* 확대 보기에서 축을 벗어난 기준선 안내 */}
        {(legalOff||internalOff)&&(
          <text x={plotR} y={plotT-8} textAnchor="end" fontSize={axisFont} fill={C.text4}>
            {internalOff?`↑ Internal ${xrfTrendFmt(internal)}`:""}
            {internalOff&&legalOff?" · ":""}
            {legalOff?`↑ Legal ${xrfTrendFmt(legal)}`:""}
          </text>
        )}

        {/* 추세선 */}
        {linePoints.length>1&&(
          <polyline fill="none" stroke={XRF_TREND_LINE_COLOR}
            strokeWidth={lineW} strokeLinecap="round" strokeLinejoin="round"
            points={linePoints.map(([x,y])=>`${x},${y}`).join(" ")}/>
        )}

        {/* 데이터 마커 + X축 라벨 + 히트 영역 */}
        {series.map((p,i)=>{
          const x=xAt(i);
          const color=xrfTrendPointColor(p);
          const selected=selectedPeriodNum!=null&&Number(p.num)===Number(selectedPeriodNum);
          return (
            <g key={`p-${p.measurementId||i}`}>
              {p.hasValue&&selected&&(
                <circle cx={x} cy={yAt(p.ppm)} r={dotR+3} fill="none"
                  stroke={color} strokeWidth={0.9} opacity={.45}/>
              )}
              {p.hasValue&&(
                p.isND
                  ? <circle cx={x} cy={yAt(p.ppm)} r={dotR} fill={C.card} stroke={color} strokeWidth={1.2}/>
                  : <circle cx={x} cy={yAt(p.ppm)} r={dotR} fill={color} stroke={C.card} strokeWidth={1}/>
              )}
              {!p.hasValue&&p.remeasure&&(
                <text x={x} y={plotB-4} textAnchor="middle" fontSize={axisFont}
                  fontWeight={500} fill={C.xWarn}>??</text>
              )}
              {showAxisLabel(p,i)&&<text x={x} y={plotB+(compact?13:16)} textAnchor="middle" fontSize={axisFont}
                fontWeight={selected?600:400}
                fill={selected?C.text1:(p.hasValue?C.text3:C.text4)}>{p.label}</text>}
              {!compact&&p.date&&showAxisLabel(p,i)&&(
                <text x={x} y={plotB+30} textAnchor="middle" fontSize={9.5} fill={C.text4}>
                  {String(p.date)}
                </text>
              )}
              <rect x={x-(compact?9:13)} y={plotT} width={compact?18:26} height={plotH}
                fill="transparent" style={{cursor:onSelectPeriod?"pointer":"default"}}
                onMouseEnter={()=>setHoverIdx(i)}
                onMouseLeave={()=>setHoverIdx(null)}
                onClick={()=>{
                  if(onSelectPeriod && Number.isFinite(p.num)) onSelectPeriod(p.num);
                }}/>
            </g>
          );
        })}

        {/* 축 제목 */}
        {!compact&&(
          <>
            <text x={plotL-52} y={plotT-12} fontSize={10} fill={C.text4}>Content (ppm)</text>
            {!denseAxis&&<text x={plotR} y={plotB+62} textAnchor="end" fontSize={10} fill={C.text4}>{lang==="en"?"Management Period":"관리주기"}</text>}
          </>
        )}
      </svg>

      {/* 측정 이력이 아직 없을 때 */}
      {measuredCount===0&&(
        <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",
          justifyContent:"center",pointerEvents:"none"}}>
          <span style={{fontSize:compact?10:11,color:C.text4,background:"rgba(255,255,255,.86)",
            border:`1px solid ${C.bd}`,borderRadius:UI.pill,padding:"3px 10px"}}>
            {lang==="en"?"No Measurement History":"측정 이력 없음"}
          </span>
        </div>
      )}

      {/* 툴팁 */}
      {hovered&&(
        <div style={{
          position:"absolute",
          left:`${(xAt(hoverIdx)/W)*100}%`,
          top:`${((hovered.hasValue?yAt(hovered.ppm):plotT+plotH/2)/H)*100}%`,
          transform:`translate(${xAt(hoverIdx)>W*0.72?"-100%":"-50%"}, calc(-100% - 10px))`,
          background:C.charcoalDk,color:"#fff",borderRadius:8,padding:"7px 10px",
          fontSize:10.5,lineHeight:1.55,whiteSpace:"nowrap",pointerEvents:"none",zIndex:6,
          boxShadow:"0 6px 18px rgba(26,32,32,.28)"}}>
          <div style={{fontWeight:700,fontSize:11}}>
            {element} · {hovered.label}
          </div>
          <div style={{opacity:.78}}>{hovered.date||(lang==="en"?"Measurement Date —":"측정일 —")}</div>
          <div style={{marginTop:3,fontWeight:700}}>
            {hovered.missing
              ? (hovered.remeasure?(lang==="en"?"Remeasurement Required":"?? 재측정 대상"):(lang==="en"?"No Data":"데이터 없음"))
              : (hovered.isND?"N.D.":`${xrfTrendFmt(hovered.ppm)} ppm`)}
            {hovered.judgement?`  ·  ${hovered.judgement}`:""}
          </div>
          <div style={{opacity:.7,marginTop:2}}>
            Internal {xrfTrendFmt(hovered.internalLimit)} / Legal {xrfTrendFmt(hovered.legalLimit)} ppm
          </div>
        </div>
      )}
      </div>
    </div>
  );
}

/* ── 카드 (원소 선택 · 축 모드 · 범례 포함) ──────────────────────────────── */
function XrfTrendChart({
  item, periods, selectedPeriodNum=null, onSelectPeriod=null, lang="ko",
}){
  const [element,setElement]=useState("ALL");          // "ALL" = 전 원소 소형 멀티플
  const [scaleMode,setScaleMode]=useState("limit");

  const safePeriods=useMemo(()=>(periods||[]).filter(Boolean),[periods]);
  const measurementEntries=useMemo(()=>xrfTrendMeasurementEntries(item,safePeriods),[item,safePeriods]);

  const seriesByElement=useMemo(()=>{
    const out={};
    XRF_TREND_ELEMENTS.forEach(el=>{ out[el]=xrfTrendSeries(item,safePeriods,el,measurementEntries); });
    return out;
  },[item,safePeriods,measurementEntries]);

  const totalMeasured=useMemo(
    ()=>XRF_TREND_ELEMENTS.reduce((mx,el)=>Math.max(mx,seriesByElement[el].filter(p=>p.hasValue).length),0),
    [seriesByElement]
  );

  if(!measurementEntries.length){
    return (
      <div style={{background:C.card,border:`1px solid ${C.bd}`,borderRadius:14,padding:"28px 18px",
        textAlign:"center",fontSize:12,color:C.text3,marginTop:12}}>
        {lang==="en"?"No management-period measurements are available for the trend.":"추세를 그릴 관리주기가 아직 없습니다."}
      </div>
    );
  }

  const chipStyle=(active)=>({
    height:26,padding:"0 11px",borderRadius:UI.pill,fontSize:11,
    fontWeight:active?700:500,cursor:"pointer",whiteSpace:"nowrap",
    border:`1px solid ${active?C.charcoalDk:C.bd2}`,
    background:active?C.charcoalDk:C.card,
    color:active?"#fff":C.text2,
  });

  return (
    <div style={{background:C.card,border:`1px solid ${C.bd}`,borderRadius:14,
      overflow:"hidden",boxShadow:"0 10px 24px rgba(26,32,32,.10)",marginTop:12}}>

      {/* 헤더 */}
      <div style={{padding:"13px 14px",borderBottom:`1px solid ${C.bd}`,background:C.alt,
        display:"flex",flexWrap:"wrap",alignItems:"center",gap:10,justifyContent:"space-between"}}>
        <div style={{minWidth:0}}>
          <div style={{...T.section}}>{lang==="en"?"XRF Trend by Period":"주기별 XRF 추세"}</div>
          <div data-i18n-skip="true" style={{...T.micro,marginTop:3}}>
            {lang==="en"
              ? `${measurementEntries.length} measured periods · ${totalMeasured} plotted points`
              : `측정 완료 관리주기 ${measurementEntries.length}개 · 표시 측정점 ${totalMeasured}개`}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
          {[["limit",lang==="en"?"Full Scale":"기준 포함"],["data",lang==="en"?"Zoom to Measurements":"측정값 확대"]].map(([key,label])=>(
            <button key={key} type="button" className="pill-btn"
              onClick={()=>setScaleMode(key)} aria-pressed={scaleMode===key}
              style={chipStyle(scaleMode===key)}>{label}</button>
          ))}
        </div>
      </div>

      {/* 원소 선택 */}
      <div style={{padding:"10px 14px",borderBottom:`1px solid ${C.bd}`,
        display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
        <button type="button" className="pill-btn" onClick={()=>setElement("ALL")}
          aria-pressed={element==="ALL"} style={chipStyle(element==="ALL")}>{lang==="en"?"All":"전체"}</button>
        {XRF_TREND_ELEMENTS.map(el=>(
          <button key={el} type="button" className="pill-btn" data-i18n-skip="true"
            onClick={()=>setElement(el)} aria-pressed={element===el}
            style={chipStyle(element===el)}>{el}</button>
        ))}
      </div>

      {/* 범례 */}
      <div data-i18n-skip="true" style={{padding:"8px 14px",borderBottom:`1px solid ${C.bd}`,
        display:"flex",gap:14,flexWrap:"wrap",alignItems:"center",fontSize:10,color:C.text3}}>
        <span style={{display:"inline-flex",alignItems:"center",gap:5}}>
          <span style={{width:16,height:2,background:XRF_TREND_LEGAL_COLOR,display:"inline-block"}}/>
          Legal Limit
        </span>
        <span style={{display:"inline-flex",alignItems:"center",gap:5}}>
          <span style={{width:16,height:0,borderTop:`2px dashed ${XRF_TREND_INTERNAL_COLOR}`,display:"inline-block"}}/>
          Internal Limit (Legal × 70%)
        </span>
        <span style={{display:"inline-flex",alignItems:"center",gap:5}}>
          <span style={{width:9,height:9,borderRadius:"50%",background:C.xOk,display:"inline-block"}}/>OK
        </span>
        <span style={{display:"inline-flex",alignItems:"center",gap:5}}>
          <span style={{width:9,height:9,borderRadius:"50%",background:C.red,display:"inline-block"}}/>NG
        </span>
        <span style={{display:"inline-flex",alignItems:"center",gap:5}}>
          <span style={{width:9,height:9,borderRadius:"50%",background:C.card,
            border:`1.6px solid ${C.xOk}`,display:"inline-block"}}/>N.D.
        </span>
      </div>

      {/* 플롯 */}
      <div style={{padding:"14px 14px 16px"}}>
        {element==="ALL"
          ? <div className="xrf-trend-grid">
              {XRF_TREND_ELEMENTS.map(el=>(
                <div key={el} style={{border:`1px solid ${C.bd}`,borderRadius:10,
                  padding:"10px 10px 8px",background:C.card,minWidth:0}}>
                  <div style={{display:"flex",alignItems:"baseline",gap:6,marginBottom:4}}>
                    <span data-i18n-skip="true" style={{fontSize:12,fontWeight:700,color:C.text1}}>{el}</span>
                    <span style={{fontSize:9.5,color:C.text4}}>{XRF_TREND_EL_NAMES[el]?.[lang]||""}</span>
                  </div>
                  <XrfTrendPlot element={el} series={seriesByElement[el]} scaleMode={scaleMode}
                    compact selectedPeriodNum={selectedPeriodNum} onSelectPeriod={onSelectPeriod} lang={lang}/>
                </div>
              ))}
            </div>
          : <div className="xrf-trend-single">
              <XrfTrendPlot element={element} series={seriesByElement[element]} scaleMode={scaleMode}
                selectedPeriodNum={selectedPeriodNum} onSelectPeriod={onSelectPeriod} lang={lang}/>
            </div>
        }
      </div>
    </div>
  );
}

function EBar({el, d, measurement, last=false}){
  if(!d) return null;
  const EL_KO={Pb:"납",Hg:"수은",Cr:"크롬",Cd:"카드뮴",Cl:"염소",Br:"브롬","Cl+Br":"염소+브롬"};
  const rawPpm=sourcePpmValue(d);
  const rawSigma=sourceSigmaValue(d);
  const ppmVal=xrfNumber(rawPpm);
  const legalLimitVal=xrfLegalLimitFromData(d) ?? XRF_ELEMENT_LIMITS[el] ?? null;
  const internalLimitVal=xrfInternalLimitFromData(d) ?? XRF_INTERNAL_LIMITS[el] ?? null;
  const missing=elementDataMissing(d);
  const isND=isXrfContentND(rawPpm) || !!d?.isND;
  const level=xrfLevelFromJudgement({...d,legalLimit:legalLimitVal,internalLimit:internalLimitVal});
  const judgement=xrfJudgementDecision({...d,legalLimit:legalLimitVal,internalLimit:internalLimitVal});
  const barC=judgement==="NG"?C.red:(judgement==="OK"?C.xOk:C.text4);
  const remeasureRequired=isXrfRemeasureRequired(d);
  const fmt=v=>v==null?"—":Number(v).toLocaleString(undefined,{maximumFractionDigits:1});
  const sourceText=v=>v==null||String(v).trim()===""?"—":String(v).trim();
  const displayValue=missing
    ? sourceText(rawPpm)
    : (isND ? sourceText(rawPpm) : `${sourceText(rawPpm)} ppm`);
  const sigmaDisplay=missing ? sourceText(rawSigma) : `${sourceText(rawSigma)} ppm`;

  // 모든 원소 막대는 0~1,500 ppm 같은 축을 사용한다.
  // 판정선은 Legal Limit이 아니라 Legal Limit의 70%인 Internal Limit 위치에 표시합니다.
  const limitPct=internalLimitVal?axisPct(internalLimitVal):100;
  const fillPct=ppmVal==null?0:axisPct(ppmVal);
  const limitLabelTransform=limitPct<12?"translateX(0)":limitPct>88?"translateX(-100%)":"translateX(-50%)";
  const limitLabelTop=4;
  const barTop=44;
  const markerTop=36;
  const markerHeight=31;
  const followupInfo=measurementFollowupInfo(measurement);
  const followupText=followupInfo.precisionElements.includes(el)
    ? "정밀분석 필요"
    : followupInfo.remeasureElements.includes(el)
      ? "XRF 재측정"
      : "없음";
  const info=el==="Cl+Br"
    ? [
        ["Content", displayValue, "Cl Content + Br Content 합계"],
        ["산정 기준", "Cl + Br", "Cl과 Br의 원본 Content 값을 합산"],
      ]
    : [
        ["Content", displayValue, "XRF Report 원본 Content"],
        ["Std.Deviation", sigmaDisplay, "XRF Report 원본 Std.Deviation"],
      ];

  return(
    <div className="xrf-element-row" style={{
      display:"grid",
      gridTemplateColumns:"58px minmax(220px, 1fr) minmax(170px, 218px)",
      columnGap:14,
      rowGap:10,
      alignItems:"center",
      padding:"11px 14px",
      borderBottom:last?"none":`1px solid ${C.bd}`,
      background:C.card
    }}>
      <div>
        <div style={{fontFamily:FONT_SANS,fontSize:14,fontWeight:700,color:C.text1,lineHeight:1}}>{el}</div>
        <div style={{fontSize:10,color:C.text2,marginTop:4,fontWeight:400}}>{EL_KO[el]||""}</div>
      </div>

      <div style={{minWidth:0}}>
        <div style={{position:"relative",height:86,marginBottom:4,overflow:"visible"}}>
          <div style={{position:"absolute",left:`${limitPct}%`,top:limitLabelTop,transform:limitLabelTransform,fontSize:9,fontWeight:500,color:C.text1,background:C.card,border:`1px solid ${C.bd}`,borderRadius:5,padding:"1px 5px",whiteSpace:"nowrap",zIndex:4,pointerEvents:"none"}}>
            Internal Limit {fmt(internalLimitVal)} ppm
          </div>
          <div style={{position:"absolute",left:0,right:0,top:barTop,height:16,
            background:C.alt,border:`1px solid ${C.bd2}`,borderRadius:0,overflow:"hidden"}}>
            {!missing&&ppmVal!==null&&ppmVal>0&&<div title={`측정값 = ${fmt(ppmVal)} ppm`} style={{position:"absolute",left:0,top:0,bottom:0,width:`${fillPct}%`,background:barC,borderRight:`1px solid rgba(26,32,32,.22)`}}/>}
          </div>
          <div style={{position:"absolute",left:`${limitPct}%`,top:markerTop,width:0,height:markerHeight,borderLeft:`1px solid ${C.red}`,opacity:.9,zIndex:3,boxSizing:"content-box"}} title={`Internal Limit = ${fmt(internalLimitVal)} ppm · Legal Limit ${fmt(legalLimitVal)} ppm의 70%`}/>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",justifyContent:"center",gap:5,fontSize:10,color:C.text2,textAlign:"center"}}>
          {info.map(([k,v,h])=>(
            <span key={k} title={h} style={{display:"inline-flex",gap:4,alignItems:"center",padding:"2px 6px",background:C.alt,border:`1px solid ${C.bd}`,borderRadius:999,whiteSpace:"nowrap"}}>
              <span style={{fontWeight:400,color:C.text2}}>{k}</span>
              <span style={{display:"inline-block",maxWidth:150,minWidth:0}}><AutoFitText value={v} title={String(v)} baseFontSize={10} minFontSize={7} align="center" style={{fontWeight:600,color:C.text2}}/></span>
            </span>
          ))}
        </div>
      </div>

      <div style={{minWidth:0,borderLeft:`1px solid ${C.bd}`,paddingLeft:14,display:"flex",alignItems:"center",justifyContent:"center"}}>
        <div style={{display:"grid",gridTemplateColumns:"88px minmax(116px, 1fr)",rowGap:9,columnGap:10,alignItems:"center",justifyItems:"center",width:"100%",textAlign:"center"}}>
          <div style={{fontSize:10,color:C.text2,fontWeight:400,whiteSpace:"nowrap",textAlign:"center"}}>내부 판정</div>
          <StatusPill
            tone={judgement==="NG"?A.rose:(remeasureRequired?A.amber:(judgement==="OK"?A.green:A.slate))}
            label={judgementDisplayValue(d)}
            title={judgementDisplayValue(d)}
            size="sm"
            fixedWidth={RESULT_LABEL_WIDTH}
          />
          <div style={{fontSize:10,color:C.text2,fontWeight:400,whiteSpace:"nowrap",textAlign:"center"}}>후속조치</div>
          <AutoFitText value={followupText} title={followupText==="없음"?"추가 후속조치 없음":followupText} baseFontSize={10.5} minFontSize={7} align="center" style={{fontFamily:FONT_SANS,fontWeight:500,color:C.text1,lineHeight:1.4,letterSpacing:0}}/>
        </div>
      </div>
    </div>
  );
}

function UploadedXrfPreview({result, compact=false}){
  if(!result?.fileName && !result?.parsed && !result?.error) return null;
  if(result?.error){
    return <div style={{marginTop:10,padding:"10px 12px",background:C.redBg,border:`1px solid ${C.red}`,borderRadius:6,color:C.redDk,fontSize:11,lineHeight:1.6}}>{result.error}</div>;
  }
  if(!result?.parsed){
    return <div style={{marginTop:10,padding:"10px 12px",background:C.xWarnBg,border:`1px solid ${C.bd}`,borderRadius:6,color:C.text3,fontSize:11}}>파일을 읽는 중이거나 파싱 결과가 없습니다.</div>;
  }
  return <div style={{marginTop:10}}>
    <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center",marginBottom:8,fontSize:10,color:C.text3}}>
      <span style={{fontWeight:600,color:C.text1}}>파싱 완료</span>
      <span>{result.fileName}</span>
      {result.sourceSheet&&<span>시트 {result.sourceSheet}</span>}
      {result.partNumberCandidate&&<span>품번 {result.partNumberCandidate}</span>}
      {result.measuredDate&&<span>Meas.Date {result.measuredDate}</span>}
      <Chip v={result.xrfResult||result.xrfWorst||"—"} small result/>
      <ApprovalBadge status={result.approvalStatus||"측정 진행중"}/>
    </div>
    <div style={{display:"grid",gridTemplateColumns:compact?"repeat(auto-fit,minmax(118px,1fr))":"repeat(auto-fit,minmax(132px,1fr))",gap:6}}>
      {ELEMENTS.map(el=>{
        const d=result.elements?.[el]||{};
        const raw=sourcePpmValue(d);
        const sigma=sourceSigmaValue(d);
        const level=xrfLevelFromJudgement(d) || "—";
        const judgement=judgementDisplayValue(d);
        const followup=xrfJudgementDecision(d)==="NG"?"정밀분석 필요":isXrfRemeasureRequired(d)?"XRF 재측정":"후속조치 없음";
        const needsReview=!isXrfRemeasureRequired(d) && !xrfJudgementDecision(d);
        return <div key={el} style={{padding:"8px 9px",background:C.card,border:`1px solid ${(elementDataMissing(d)||needsReview)?C.red:C.bd}`,borderRadius:5,minWidth:0}}>
          <div style={{fontSize:10,fontWeight:600,color:C.text1,marginBottom:4}}>{el}</div>
          <AutoFitText value={`${raw===null||raw===undefined||raw===""?"—":String(raw)} ppm`} title={String(raw??"—")} baseFontSize={11} minFontSize={7} align="left" style={{fontWeight:500,color:C.text1}}/>
          <div style={{fontSize:10,color:C.text2,marginTop:2}}>Std.Deviation {sigma===null||sigma===undefined||sigma===""?"—":String(sigma)}</div>
          <div style={{fontSize:10,color:C.text2,marginTop:3}}>내부 판정 {judgement}</div>
          <div style={{fontSize:9,color:C.text4,marginTop:2}}>Internal Limit {(xrfInternalLimitFromData(d)??XRF_INTERNAL_LIMITS[el])?.toLocaleString?.()||"—"} ppm · Legal {(xrfLegalLimitFromData(d)??XRF_ELEMENT_LIMITS[el])?.toLocaleString?.()||"—"} ppm × 70%</div>
          <div style={{fontSize:10,color:C.text1,marginTop:3,fontWeight:500}}>{needsReview?"판정 확인 필요":elementDataMissing(d)?"데이터 확인 필요":followup}</div>
          {el==="Cl+Br"&&<div style={{fontSize:9,color:C.text4,marginTop:2}}>Cl + Br Content 합계 기준 · Internal Limit 1,050 ppm</div>}
        </div>;
      })}
    </div>
    {!!result.dataMissingElements?.length&&<div style={{marginTop:7,fontSize:10,color:C.redDk,fontWeight:700}}>데이터 부족: {result.dataMissingElements.join(", ")}</div>}
    {!!result.reviewRequiredElements?.length&&<div style={{marginTop:5,fontSize:10,color:C.redDk,fontWeight:700}}>Judgment 확인 필요: {result.reviewRequiredElements.join(", ")}</div>}
  </div>;
}

function RiskMatrix(){
  const MX={H:{H:"Not Allowed",M:"Not Allowed",L:"Not Allowed"},M:{H:"H",M:"M",L:"M"},L:{H:"M",M:"M",L:"L"}};
  const RS={"Not Allowed":{bg:C.charcoalDk,c:"#fff",l:"사용 불허"},"H":{bg:C.redBg,c:C.redDk,l:"사용 자제\n월 1회"},"M":{bg:C.charcoalBg,c:C.charcoalMd,l:"관리 필요\n반기 1회"},"L":{bg:C.charcoalBg,c:C.charcoal,l:"정상\n연 1회"}};
  const th={padding:"8px 12px",background:C.bg,border:`1px solid ${C.bd}`,fontSize:11,fontWeight:600,color:C.text3,textAlign:"center"};
  return(
    <div className="responsive-table-scroll table-scroll-compact" role="region" aria-label="위험성 평가 기준표" tabIndex={0}>
      <table style={{borderCollapse:"collapse",width:"100%",fontSize:12}}>
        <thead><tr><th style={{...th,textAlign:"center"}}>XRF ╲ C&R</th>{["H — 직접접촉+잔류","M — 직접접촉+비잔류","L — 비접촉"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
        <tbody>{["H","M","L"].map(x=>(
          <tr key={x}><td style={{padding:"8px 12px",border:`1px solid ${C.bd}`,fontWeight:500,color:C.text2,background:C.bg,fontSize:11}}>{x==="H"?"H · Exceeded":x==="M"?"M · Less than":"L · N.D."}</td>
            {["H","M","L"].map(c=>{const r=MX[x][c];const s=RS[r];return <td key={c} style={{padding:"10px 12px",border:`1px solid ${C.bd}`,background:s.bg,color:s.c,fontWeight:700,textAlign:"center",whiteSpace:"pre-line",fontSize:11,lineHeight:1.6}}>{s.l}</td>;})}
          </tr>
        ))}</tbody>
      </table>
    </div>
  );
}

/* ── 상단 통계 카드 (숫자 + 비중 배지) ── */
function StatCard({label, value, sub, color, share, onClick, active}){
  return(
    <div onClick={onClick}
      style={{background:C.card,border:`1px solid ${active?color:C.bd}`,borderRadius:UI.r,
        padding:"14px 16px",boxSizing:"border-box",cursor:onClick?"pointer":"default",
        boxShadow:active?UI.shHover:UI.sh,display:"flex",flexDirection:"column",
        justifyContent:"space-between",minHeight:96,transition:"box-shadow .15s, border-color .15s"}}
      onMouseEnter={e=>{if(onClick)e.currentTarget.style.boxShadow=UI.shHover;}}
      onMouseLeave={e=>{if(onClick&&!active)e.currentTarget.style.boxShadow=UI.sh;}}>
      <div style={{fontSize:12,color:C.text3,fontWeight:600}}>{label}</div>
      <div style={{marginTop:8}}>
        <div style={{display:"flex",alignItems:"center",gap:7}}>
          <span style={{fontSize:26,fontWeight:700,color:C.text1,lineHeight:1,letterSpacing:"-1px"}}>{value}</span>
          {share!==null&&share!==undefined&&(
            <span style={{fontSize:11,fontWeight:700,padding:"3px 8px",borderRadius:20,
              background:`${color}14`,color:color,whiteSpace:"nowrap"}}>{share}</span>
          )}
        </div>
        <div style={{width:"100%",minWidth:0,marginTop:6}}><AutoFitText value={sub} title={String(sub||"")} baseFontSize={11} minFontSize={7} align="left" style={{color:C.text4}}/></div>
      </div>
    </div>
  );
}

/* ── 차트 카드 껍데기 ── */
function ChartCard({title, hint, children}){
  return(
    <div style={{background:C.card,border:`1px solid ${C.bd}`,borderRadius:UI.r,
      padding:"16px 18px",boxShadow:UI.sh,boxSizing:"border-box",minWidth:0}}>
      <div style={{fontSize:13,fontWeight:700,color:C.text1}}>{title}</div>
      {hint&&<div style={{fontSize:11,color:C.text4,marginTop:3,marginBottom:12}}>{hint}</div>}
      {children}
    </div>
  );
}

/* ① 향후 6개월 측정 도래 예정 — 다음 마감일 기준 월별 부하 (세로 막대) */
function DueLoadChart({buckets, onPick}){
  const max=Math.max(1,...buckets.map(b=>b.total));
  const BAR_MAX_HEIGHT=88;
  return(
    <div style={{display:"flex",alignItems:"stretch",gap:10,height:132,paddingTop:4}}>
      {buckets.map(b=>{
        const clickable=!!onPick && b.total>0;
        // 비율을 % 높이에 맡기지 않고 고정 plot 높이(px) 안에서 직접 계산합니다.
        // 최대 건수 막대가 88px, 나머지는 건수 비율만큼 정확히 줄어 상대 차이가 즉시 보입니다.
        const barHeight=b.total>0 ? Math.max(6,Math.round((b.total/max)*BAR_MAX_HEIGHT)) : 3;
        return <div key={b.key} onClick={()=>clickable&&onPick(b)}
          style={{flex:1,height:"100%",display:"grid",gridTemplateRows:"16px 88px 16px",alignItems:"end",justifyItems:"center",rowGap:6,minWidth:0,cursor:clickable?"pointer":"default"}}
          title={`${b.label} · 총 ${b.total}건 (기한초과 ${b.overdue}건)`}>
          <span style={{fontSize:11,fontWeight:700,color:b.total?C.text2:C.text4,lineHeight:"16px"}}>{b.total}</span>
          <div style={{width:"100%",height:BAR_MAX_HEIGHT,display:"flex",alignItems:"flex-end",justifyContent:"center"}}>
            <div style={{width:"100%",maxWidth:34,height:barHeight,minHeight:barHeight,
              borderRadius:"8px 8px 4px 4px",overflow:"hidden",display:"flex",flexDirection:"column",justifyContent:"flex-end",
              background:b.total?C.charcoalBg:C.bd,transition:"height .18s ease"}}>
              {b.overdue>0&&<div style={{height:`${(b.overdue/Math.max(b.total,1))*100}%`,background:A.rose.solid,flex:"0 0 auto"}}/>}
              <div style={{flex:1,background:b.total?C.charcoalDk:C.bd}}/>
            </div>
          </div>
          <span style={{fontSize:10,color:C.text4,whiteSpace:"nowrap",lineHeight:"16px"}}>{b.label}</span>
        </div>;
      })}
    </div>
  );
}

/* ② 이행 상태 구성 — 100% 스택 바 */
function ComplianceBar({segments, total, onPick}){
  return(
    <div>
      <div style={{display:"flex",height:14,borderRadius:8,overflow:"hidden",background:C.bd,marginBottom:12}}>
        {segments.filter(s=>s.n>0).map(s=>(
          <div key={s.key} onClick={()=>onPick&&onPick(s)} title={`${s.label} ${s.n}건`}
            style={{width:`${(s.n/Math.max(total,1))*100}%`,background:s.color,cursor:onPick?"pointer":"default"}}/>
        ))}
      </div>
      <div style={{display:"flex",flexWrap:"wrap",gap:"6px 14px"}}>
        {segments.map(s=>(
          <div key={s.key} onClick={()=>s.n&&onPick&&onPick(s)}
            style={{display:"flex",alignItems:"center",gap:6,fontSize:11,cursor:s.n&&onPick?"pointer":"default",opacity:s.n?1:.45}}>
            <span style={{width:8,height:8,borderRadius:3,background:s.color,flex:"0 0 auto"}}/>
            <span style={{color:C.text3}}>{s.label}</span>
            <strong style={{color:C.text1,fontWeight:700}}>{s.n}</strong>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ③ 부서별 위험도 구성 — 가로 스택 바 */
function DeptRiskChart({rows, riskMeta}){
  const max=Math.max(1,...rows.map(r=>r.total));
  return(
    <div style={{display:"flex",flexDirection:"column",gap:9}}>
      {rows.map(r=>(
        <div key={r.dept} style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{width:66,flex:"0 0 auto",minWidth:0}}><AutoFitText value={r.dept} title={r.dept} baseFontSize={11} minFontSize={7} align="left" style={{color:C.text3}}/></span>
          <div style={{flex:1,display:"flex",height:14,borderRadius:7,overflow:"hidden",background:C.bg,minWidth:0}}>
            <div style={{display:"flex",width:`${(r.total/max)*100}%`,minWidth:r.total?6:0}}>
              {riskMeta.map(m=>r[m.key]>0&&(
                <div key={m.key} title={`${r.dept} · ${m.label} ${r[m.key]}건`}
                  style={{flex:r[m.key],background:m.color}}/>
              ))}
            </div>
          </div>
          <span style={{fontSize:11,fontWeight:700,color:C.text2,width:26,textAlign:"right",flex:"0 0 auto"}}>{r.total}</span>
        </div>
      ))}
      <div style={{display:"flex",flexWrap:"wrap",gap:"6px 14px",marginTop:4}}>
        {riskMeta.map(m=>(
          <div key={m.key} style={{display:"flex",alignItems:"center",gap:6,fontSize:11}}>
            <span style={{width:8,height:8,borderRadius:3,background:m.color}}/>
            <span style={{color:C.text3}}>{m.label}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// 첨부 시안에서 추출한 실제 배경색입니다. 위 #1F272F → 아래 #191F26 의 세로 그라데이션이며,
// C.charcoalDk(#1A2020)보다 약간 푸른 기가 도는 차콜입니다.
const SPLASH_BG_TOP = "#1F272F";
const SPLASH_BG_MID = "#1D242D";
const SPLASH_BG_BTM = "#191F26";
const SPLASH_BG     = `linear-gradient(180deg, ${SPLASH_BG_TOP} 0%, ${SPLASH_BG_MID} 46%, ${SPLASH_BG_BTM} 100%)`;
const SPLASH_DIM   = "rgba(255,255,255,.74)";
const SPLASH_DIM_2 = "rgba(255,255,255,.55)";
const SPLASH_LINE  = "rgba(255,255,255,.12)";
const APP_VERSION=typeof __APP_VERSION__==="string"?__APP_VERSION__:"1.0.0";

function SplashIcon({name}){
  const p={fill:"none",stroke:"currentColor",strokeWidth:1.5,
    strokeLinecap:"round",strokeLinejoin:"round"};
  return (
    <svg width="27" height="27" viewBox="0 0 24 24" aria-hidden="true">
      {name==="list"&&<g {...p}>
        <rect x="3.5" y="3.5" width="17" height="17" rx="2.5"/>
        <path d="M7.5 9h9M7.5 12.5h9M7.5 16h5.5"/></g>}
      {name==="xrf"&&<g {...p}>
        <path d="M3.5 18.5h17"/>
        <path d="M4 15.5l4.5-5 3.5 3.5 4-6.5 4 4.5"/>
        <circle cx="8.5" cy="10.5" r="1.3"/><circle cx="16" cy="7.5" r="1.3"/></g>}
      {name==="precision"&&<g {...p}>
        <path d="M6 3.5h7.5L18.5 8.5V20a.5.5 0 0 1-.5.5H6a.5.5 0 0 1-.5-.5V4a.5.5 0 0 1 .5-.5Z"/>
        <path d="M13 3.5V9h5.5"/><circle cx="11" cy="14.5" r="2.6"/><path d="M13 16.5l2 2"/></g>}
      {name==="risk"&&<g {...p}>
        <path d="M3.5 15.5a8.5 8.5 0 0 1 17 0"/><path d="M12 15.5l4-4.5"/>
        <circle cx="12" cy="15.5" r="1.1"/><path d="M3.5 18.5h17"/></g>}
      {name==="reg"&&<g {...p}>
        <circle cx="12" cy="12" r="3"/>
        <path d="M12 2.8v2.6M12 18.6v2.6M21.2 12h-2.6M5.4 12H2.8M18.5 5.5l-1.8 1.8M7.3 16.7l-1.8 1.8M18.5 18.5l-1.8-1.8M7.3 7.3 5.5 5.5"/></g>}
    </svg>
  );
}

function SplashScreen({onSelectTab, lang="ko", onLang, developer="SHLee", department="QEX"}){
  const t=(ko,en)=>lang==="en"?en:ko;

  // 궤도 = 앱이 자동으로 수행하는 4단계. 한 줄 소개문의 네 절과 그대로 대응합니다.
  //   판독 → 산정 → 추적 → 이관 순으로 한 바퀴 돌며, 이 과정이 주기마다 반복됩니다.
  const orbit=[
    {x:270,y:311,anchor:"end",  l1:"XRF SCREENING", l2:t("측정값 판독","Reads uploaded results")},
    {x:270,y:189,anchor:"end",  l1:"RISK & CYCLE",  l2:t("위험도와 주기 산정","Risk and cycle")},
    {x:730,y:189,anchor:"start",l1:"DUE TRACKING",  l2:t("도래일과 지연 추적","Due date and overdue")},
    {x:730,y:311,anchor:"start",l1:"ESCALATION",    l2:t("재측정과 정밀분석","Re-measure and precision")},
  ];

  // 하단 = 실제 TABS
  const menus=[
    {tab:"list",      icon:"list",      label:t("부자재 리스트","Materials")},
    {tab:"xrf",       icon:"xrf",       label:t("XRF 분석","XRF Analysis")},
    {tab:"precision", icon:"precision", label:t("정밀분석","Precision Analysis")},
    {tab:"risk",      icon:"risk",      label:t("위험도 현황","Risk Overview")},
    {tab:"reg",       icon:"reg",       label:t("품목 / 변경 관리","Registration")},
  ];

  return (
    <div className="mxk-splash" style={{height:"100dvh",minHeight:0,maxHeight:"100dvh",background:SPLASH_BG,fontFamily:FONT_SANS,
      display:"flex",flexDirection:"column",position:"relative",overflow:"hidden",boxSizing:"border-box"}}>

      {/* 배경 레이어
          ① 상단 광원 ② 중앙 하단의 붉은 잔광(그래픽의 레드 플레이트와 이어집니다)
          ③ 가장자리를 눌러 주는 비네트 — 시선이 가운데로 모이게 합니다 */}
      <div aria-hidden="true" style={{position:"absolute",inset:0,pointerEvents:"none",
        background:[
          "radial-gradient(1100px 520px at 50% 2%, rgba(255,255,255,.055), transparent 68%)",
          "radial-gradient(760px 340px at 50% 62%, rgba(230,10,50,.10), transparent 70%)",
          "radial-gradient(120% 90% at 50% 50%, transparent 52%, rgba(0,0,0,.38) 100%)",
        ].join(",")}}/>
      <style>{`
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable.css');
        @keyframes mxk-drift { to { stroke-dashoffset:-1000; } }
        @keyframes mxk-breathe { 0%,100%{opacity:.62} 50%{opacity:1} }
        .mxk-trace { animation: mxk-drift 14s linear infinite; }
        .mxk-breathe { animation: mxk-breathe 6s ease-in-out infinite; }
        .mxk-splash button { font-family:${FONT_SANS}; }
        .mxk-splash-header { flex:0 0 auto; }
        .mxk-splash-main { min-height:0; overflow:hidden; display:grid; grid-template-columns:minmax(0,1.16fr) minmax(410px,.84fr); grid-template-rows:auto auto; align-content:center; align-items:center; column-gap:clamp(28px,4vw,76px); row-gap:clamp(16px,2.7dvh,28px); }
        .mxk-splash-content { grid-column:2; grid-row:1; align-self:end; min-width:0; text-align:left; }
        .mxk-splash-title-row { display:flex; align-items:flex-end; gap:13px; flex-wrap:wrap; }
        .mxk-splash-version { margin-bottom:7px; color:rgba(255,255,255,.58); font-size:10px; font-weight:600; letter-spacing:.4px; white-space:nowrap; }
        .mxk-splash-hero-wrap { position:relative; grid-column:1; grid-row:1 / span 2; align-self:center; width:100%; min-width:0; display:flex; justify-content:center; }
        .mxk-splash-hero { width:100%; min-height:250px; }
        .mxk-splash-menu { grid-column:2; grid-row:2; align-self:start; display:grid; grid-template-columns:1fr; gap:8px; width:100%; max-width:530px; }
        .mxk-splash-orbit-mobile { display:none; }
        .mxk-splash-footer { flex:0 0 auto; }
        .mxk-splash-footer-meta { display:flex; align-items:center; gap:16px; flex-wrap:nowrap; padding:0 clamp(12px,4vw,52px) clamp(7px,1.5dvh,16px); font-size:10px; color:${SPLASH_DIM_2}; }
        .mxk-splash-menu-item { position:relative; appearance:none; -webkit-appearance:none; display:flex; align-items:center; gap:14px; width:100%; min-height:56px; padding:5px 14px; border:1px solid rgba(255,255,255,.16); border-radius:12px; background:rgba(255,255,255,.035); color:#fff; cursor:pointer; text-align:left; transition:color .18s ease, background .18s ease, border-color .18s ease, transform .18s ease; }
        .mxk-splash-menu-icon { display:grid; flex:0 0 auto; place-items:center; align-self:center; width:42px; height:42px; box-sizing:border-box; border:1px solid transparent; border-radius:12px; transition:background .18s ease, border-color .18s ease, box-shadow .18s ease; }
        .mxk-splash-menu-icon svg { display:block; }
        .mxk-splash-menu-label { min-width:0; color:#fff; font-size:16px; font-weight:650; line-height:1.35; }
        .mxk-splash-menu-arrow { margin-left:auto; color:rgba(255,255,255,.5); font-size:19px; line-height:1; }
        .mxk-splash-menu-item:hover .mxk-splash-menu-icon, .mxk-splash-menu-item:focus-visible .mxk-splash-menu-icon { background:${C.red}; border-color:${C.red}; box-shadow:0 0 0 4px rgba(230,10,50,.15), 0 8px 22px rgba(230,10,50,.3); }
        .mxk-splash-menu-item:focus-visible { outline:none; }
        .mxk-splash-menu-item:focus-visible .mxk-splash-menu-icon { outline:2px solid rgba(255,255,255,.75); outline-offset:3px; }
        .mxk-splash-menu-item:active .mxk-splash-menu-icon { background:#B90027; }
        @media (prefers-reduced-motion: reduce) {
          .mxk-trace, .mxk-breathe { animation: none; }
        }
        @media (max-width:980px) {
          .mxk-splash-main { display:flex; flex-direction:column; align-items:center; justify-content:center; }
          .mxk-splash-content { width:100%; max-width:780px; text-align:center; }
          .mxk-splash-title-row { justify-content:center; }
          .mxk-splash-title { font-size:clamp(40px,6vw,64px) !important; }
          .mxk-splash-version { margin-bottom:5px; }
          .mxk-splash-subtitle { font-size:clamp(15px,2.6vw,21px) !important; }
          .mxk-splash-rule { margin-left:auto !important; margin-right:auto !important; }
          .mxk-splash-copy { margin:0 auto; font-size:calc(clamp(12px,1.6vw,14px) - 1.5pt) !important; }
          .mxk-splash-hero-wrap { max-width:1000px; margin-top:clamp(9px,1.7dvh,18px); }
          .mxk-splash-hero { min-height:150px; height:clamp(165px,39dvh,430px) !important; }
          .mxk-splash-menu { width:min(710px,100%); max-width:none; grid-template-columns:repeat(5,minmax(0,1fr)); gap:0; margin-top:clamp(12px,2.3dvh,25px); }
          .mxk-splash-menu-item { flex-direction:column; justify-content:center; gap:clamp(4px,.8dvh,8px); min-height:76px; padding:2px 6px; border:0; border-radius:0; background:transparent; text-align:center; }
          .mxk-splash-menu-item:not(:first-child) { border-left:1px solid ${SPLASH_LINE}; }
          .mxk-splash-menu-icon { width:45px; height:45px; border-radius:14px; }
          .mxk-splash-menu-label { font-size:13px; white-space:normal; }
          .mxk-splash-menu-arrow { display:none; }
        }
        @media (max-width:620px) {
          .mxk-splash-header { padding-left:12px !important; padding-right:12px !important; gap:8px !important; }
          .mxk-splash-brand { gap:6px !important; }
          .mxk-splash-brand span:last-child { font-size:9px !important; }
          .mxk-splash-language button { padding:4px 8px !important; font-size:10px !important; }
          .mxk-splash-main { padding-left:8px !important; padding-right:8px !important; }
          .mxk-splash-menu-item { padding-left:2px !important; padding-right:2px !important; }
          .mxk-splash-menu-icon { width:38px; height:38px; }
          .mxk-splash-menu-item svg { width:23px; height:23px; }
          .mxk-splash-menu-label { font-size:11px !important; line-height:1.25 !important; white-space:normal !important; }
          .mxk-splash-orbit-text { display:none; }
          .mxk-splash-orbit-mobile { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:5px 12px; width:min(380px,100%); margin-top:4px; text-align:center; }
          .mxk-splash-orbit-mobile strong { display:block; color:#fff; font-size:10.5px; font-weight:600; letter-spacing:.3px; line-height:1.25; }
          .mxk-splash-orbit-mobile span { display:block; color:rgba(255,255,255,.67); font-size:10px; line-height:1.25; }
          .mxk-splash-footer-meta { gap:8px; font-size:9px; }
        }
        @media (max-height:650px) {
          .mxk-splash-main { row-gap:12px; }
          .mxk-splash-copy { font-size:calc(clamp(12px,1.15vw,15px) - 1.5pt) !important; line-height:1.4 !important; }
          .mxk-splash-hero { min-height:180px; height:46dvh !important; }
          .mxk-splash-menu { gap:5px; }
          .mxk-splash-menu-item { min-height:46px; }
          .mxk-splash-menu-icon { width:36px; height:36px; }
          .mxk-splash-menu-item svg { width:22px; height:22px; }
        }
        @media (max-width:980px) and (max-height:650px) {
          .mxk-splash-hero-wrap { margin-top:8px !important; }
          .mxk-splash-hero { min-height:130px; height:30dvh !important; }
          .mxk-splash-menu { gap:0; margin-top:7px !important; }
          .mxk-splash-menu-item { min-height:70px; }
          .mxk-splash-menu-label { font-size:11px !important; }
        }
      `}</style>

      {/* 상단 */}
      <div className="mxk-splash-header" style={{position:"relative",display:"flex",alignItems:"center",
        justifyContent:"space-between",gap:20,padding:"clamp(10px,2.5dvh,30px) clamp(20px,4vw,52px) 0",flexWrap:"nowrap"}}>
        <div className="mxk-splash-brand" data-i18n-skip="true" style={{display:"flex",alignItems:"baseline",gap:10,whiteSpace:"nowrap"}}>
          <span style={{fontFamily:"Arial Black,Arial,sans-serif",fontWeight:900,fontSize:21,
            color:C.red,letterSpacing:"-1px"}}>molex</span>
          <span style={{fontFamily:"Georgia,serif",fontStyle:"italic",fontSize:11,
            color:SPLASH_DIM}}>creating connections for life</span>
        </div>
        <div className="mxk-splash-language" style={{display:"flex",alignItems:"center",gap:2,border:`1px solid ${SPLASH_LINE}`,
          borderRadius:UI.pill,padding:3}}>
          {[["ko",t("한국어","Korean")],["en","English"]].map(([k,label])=>(
            <button key={k} type="button" onClick={()=>onLang&&onLang(k)} aria-pressed={lang===k}
              style={{padding:"5px 13px",borderRadius:UI.pill,border:"none",cursor:"pointer",
                fontSize:10.5,fontWeight:600,letterSpacing:".3px",
                background:lang===k?"rgba(255,255,255,.12)":"transparent",
                color:lang===k?"#fff":SPLASH_DIM}}>{label}</button>
          ))}
        </div>
      </div>

      {/* 본문 */}
      <div className="mxk-splash-main" style={{position:"relative",flex:"1 1 auto",width:"100%",maxWidth:1540,margin:"0 auto",
        boxSizing:"border-box",padding:"clamp(6px,1.5dvh,18px) clamp(20px,4vw,64px) clamp(8px,1.5dvh,18px)"}}>

        <div className="mxk-splash-content">
          <div className="mxk-splash-title-row" data-i18n-skip="true">
            <h1 className="mxk-splash-title" style={{margin:0,fontSize:"clamp(56px,min(5.7vw,10dvh),88px)",fontWeight:800,
              letterSpacing:"-2.5px",lineHeight:1,color:"#fff"}}>
              MXK<span style={{color:C.red}}>PCS</span>
            </h1>
            <span className="mxk-splash-version">Version {APP_VERSION}</span>
          </div>
          <div className="mxk-splash-subtitle" data-i18n-skip="true" style={{marginTop:"clamp(9px,1.5dvh,15px)",fontSize:"clamp(20px,2vw,27px)",
            fontWeight:500,letterSpacing:"-.2px",color:"rgba(255,255,255,.9)"}}>Product Compliance System</div>
          <div className="mxk-splash-rule" aria-hidden="true" style={{width:48,height:2,background:C.red,margin:"clamp(13px,2dvh,20px) 0 clamp(10px,1.5dvh,15px)"}}/>
          <div className="mxk-splash-copy" style={{maxWidth:600,fontSize:"calc(clamp(15px,1.25vw,18px) - 1.5pt)",lineHeight:1.6,
            color:"#fff"}}>
            {t("부자재의 신규 등록과 변경, 단종을 관리하고 업로드한 XRF 파일에서 위험도와 관리주기를 자동 산정합니다.",
               "Manage registration, change and discontinuation of auxiliary materials, and derive risk level and inspection cycle automatically from uploaded XRF files.")}
          </div>
        </div>

        {/* 히어로 그래픽 — 시안의 아이소메트릭 유리판 + 레드 글로우 + 궤도는 유지하고,
            방패·잎사귀 자리에 이 시스템의 실제 판정을 넣었습니다.
              · 유리판 위 7개 막대 = Pb · Hg · Cr · Cd · Cl · Br · Cl+Br 원소 측정값
              · 위쪽 두 평면      = Legal Limit(빨강) / Internal Limit = Legal × 70%(주황)
              · Cd 막대가 내부 기준 평면을 뚫고 올라가 초과 상태를 나타냅니다
              · 궤도             = 관리주기 R → P1 → P2 → P3 ··· 의 반복             */}
        <div className="mxk-splash-hero-wrap">
          <svg className="mxk-splash-hero" viewBox="85 60 830 340" width="100%" role="img" aria-hidden="true"
            style={{height:"clamp(290px,56dvh,500px)",margin:0,display:"block"}}>
          <defs>
            <linearGradient id="mxk-bar" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="rgba(255,255,255,.92)"/>
              <stop offset="55%" stopColor="rgba(255,255,255,.44)"/>
              <stop offset="100%" stopColor="rgba(255,255,255,.14)"/>
            </linearGradient>
            <linearGradient id="mxk-bar-ng" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="#FFD79B"/>
              <stop offset="45%" stopColor={A.amber.solid} stopOpacity=".85"/>
              <stop offset="100%" stopColor={A.amber.solid} stopOpacity=".16"/>
            </linearGradient>
            <linearGradient id="mxk-plate" x1="0" y1="0" x2="1" y2="1">
              <stop offset="0%" stopColor="rgba(255,255,255,.22)"/>
              <stop offset="48%" stopColor="rgba(255,255,255,.08)"/>
              <stop offset="100%" stopColor="rgba(255,255,255,.03)"/>
            </linearGradient>
            <linearGradient id="mxk-plate-red" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="#FF4D6E"/>
              <stop offset="100%" stopColor={C.red}/>
            </linearGradient>
            {/* 궤도 : 양 끝이 사라지는 선 */}
            <linearGradient id="mxk-ring" x1="0" y1="0" x2="1" y2="0">
              <stop offset="0%"   stopColor="rgba(255,255,255,0)"/>
              <stop offset="22%"  stopColor="rgba(255,255,255,.20)"/>
              <stop offset="50%"  stopColor="rgba(255,255,255,.07)"/>
              <stop offset="78%"  stopColor="rgba(255,255,255,.20)"/>
              <stop offset="100%" stopColor="rgba(255,255,255,0)"/>
            </linearGradient>
            <radialGradient id="mxk-glow" cx="50%" cy="50%" r="50%">
              <stop offset="0%" stopColor={C.red} stopOpacity=".62"/>
              <stop offset="55%" stopColor={C.red} stopOpacity=".18"/>
              <stop offset="100%" stopColor={C.red} stopOpacity="0"/>
            </radialGradient>
            <radialGradient id="mxk-orbit-dot-glow" cx="50%" cy="50%" r="50%">
              <stop offset="0%" stopColor="#FF536D" stopOpacity=".75"/>
              <stop offset="35%" stopColor={C.red} stopOpacity=".38"/>
              <stop offset="100%" stopColor={C.red} stopOpacity="0"/>
            </radialGradient>
            {/* 막대 아래로 비치는 반사 */}
            <linearGradient id="mxk-mirror" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="rgba(255,255,255,.16)"/>
              <stop offset="100%" stopColor="rgba(255,255,255,0)"/>
            </linearGradient>
            <filter id="mxk-soft" x="-60%" y="-60%" width="220%" height="220%">
              <feGaussianBlur stdDeviation="7"/>
            </filter>
            <filter id="mxk-tip" x="-120%" y="-120%" width="340%" height="340%">
              <feGaussianBlur stdDeviation="3.4"/>
            </filter>
            <marker id="mxk-arrow" markerWidth="9" markerHeight="9" refX="4.5" refY="4.5"
              orient="auto"><path d="M1 1 L8 4.5 L1 8 Z" fill="rgba(255,255,255,.34)"/></marker>
          </defs>

          <ellipse className="mxk-breathe" cx="500" cy="358" rx="228" ry="46" fill="url(#mxk-glow)"/>

          {/* 관리주기 궤도 + 궤도를 천천히 도는 빛줄기 */}
          <g transform="rotate(-4 500 250)">
            <ellipse cx="500" cy="250" rx="300" ry="95" fill="none"
              stroke="url(#mxk-ring)" strokeWidth="1.1"/>
            <ellipse className="mxk-trace" cx="500" cy="250" rx="300" ry="95" fill="none"
              stroke="rgba(255,255,255,.62)" strokeWidth="1.4" strokeLinecap="round"
              pathLength="1000" strokeDasharray="20 980"/>
          </g>
          <path d="M232 214 A 300 95 -4 0 1 292 170" fill="none"
            stroke="rgba(255,255,255,.26)" strokeWidth="1.2" markerEnd="url(#mxk-arrow)"/>
          <path d="M708 170 A 300 95 -4 0 1 768 214" fill="none"
            stroke="rgba(255,255,255,.26)" strokeWidth="1.2" markerEnd="url(#mxk-arrow)"/>

          {/* 기준 평면 — 뒤쪽 */}
          <path d="M348 140 L500 87 L652 140 L500 193 Z" fill={C.red} fillOpacity=".055"
            stroke={C.red} strokeOpacity=".50" strokeWidth="1"/>
          <path d="M348 188 L500 135 L652 188 L500 241 Z" fill={A.amber.solid} fillOpacity=".055"
            stroke={A.amber.solid} strokeOpacity=".50" strokeWidth="1" strokeDasharray="5 4"/>

          {/* 받침 : 레드 플레이트(번짐 포함) + 유리 플레이트 */}
          <path d="M500 292 L622 334 L500 376 L378 334 Z" fill={C.red} opacity=".55"
            filter="url(#mxk-soft)"/>
          <path d="M500 292 L622 334 L500 376 L378 334 Z" fill="url(#mxk-plate-red)" opacity=".92"/>
          <path d="M500 247 L652 300 L500 353 L348 300 Z" fill="url(#mxk-plate)"
            stroke="rgba(255,255,255,.26)" strokeWidth="1"/>
          <path d="M348 300 L500 247 L652 300" fill="none" stroke="rgba(255,255,255,.34)" strokeWidth="1"/>

          {/* 원소 측정 막대 (왼쪽부터 큰 값) */}
          {[["Cd",380,128,true],["Pb",420,92,false],["Cl+Br",460,86,false],["Cr",500,70,false],
            ["Cl",540,60,false],["Br",580,54,false],["Hg",620,46,false]].map(([el,x,hgt,over])=>{
            const top=300-hgt;
            const tint=over?A.amber.solid:"#fff";
            return (
              <g key={el}>
                <rect x={x-7} y={top+14} width="14" height={hgt-14} rx="3"
                  fill={tint} opacity={over?".34":".16"} filter="url(#mxk-soft)"/>
                <rect x={x-9} y={top} width="18" height={hgt+26} rx="3"
                  fill="url(#mxk-mirror)" opacity=".5"/>
                <rect x={x-9} y={top} width="18" height={hgt} rx="3"
                  fill={over?"url(#mxk-bar-ng)":"url(#mxk-bar)"}/>
                <ellipse cx={x} cy={top} rx="13" ry="4.5" fill={tint}
                  opacity={over?".55":".34"} filter="url(#mxk-tip)"/>
                <ellipse cx={x} cy={top} rx="9" ry="3" fill={over?"#FFD79B":"rgba(255,255,255,.96)"}/>
                <text x={x} y={top-11} textAnchor="middle" fontSize="9.5" fontWeight="600"
                  letterSpacing=".3" fill="#fff">{el}</text>
              </g>
            );
          })}

          {/* 기준 평면 — 앞쪽 모서리를 막대 위에 겹칩니다 */}
          <path d="M348 140 L500 193 L652 140" fill="none" stroke={C.red}
            strokeOpacity=".85" strokeWidth="1.5"/>
          <path d="M348 188 L500 241 L652 188" fill="none" stroke={A.amber.solid}
            strokeOpacity=".85" strokeWidth="1.5" strokeDasharray="5 4"/>
          <text x="648" y="128" textAnchor="end" fontSize="10.5" fontWeight="600"
            letterSpacing="1" fill="rgba(255,255,255,.78)">LEGAL LIMIT</text>
          <text x="648" y="176" textAnchor="end" fontSize="10.5" fontWeight="600"
            letterSpacing="1" fill="rgba(255,255,255,.78)">INTERNAL 70%</text>

          {orbit.map(o=>(
            <g key={o.l1}>
              <circle className="mxk-breathe" cx={o.x} cy={o.y} r="20" fill="url(#mxk-orbit-dot-glow)"/>
              <circle cx={o.x} cy={o.y} r="4.5" fill={C.red}/>
              <circle cx={o.x} cy={o.y} r="9.5" fill="none" stroke={C.red}
                strokeOpacity=".28" strokeWidth="1"/>
              <text className="mxk-splash-orbit-text" x={o.anchor==="end"?o.x-20:o.x+20} y={o.y-3} textAnchor={o.anchor}
                fontSize="16.5" fontWeight="600" letterSpacing=".7" fill="#fff">{o.l1}</text>
              <text className="mxk-splash-orbit-text" x={o.anchor==="end"?o.x-20:o.x+20} y={o.y+16} textAnchor={o.anchor}
                fontSize="14" fontWeight="400" fill="rgba(255,255,255,.68)">{o.l2}</text>
            </g>
          ))}
          </svg>
        </div>
        <div className="mxk-splash-orbit-mobile">
          {orbit.map(o=><div key={o.l1}><strong>{o.l1}</strong><span>{o.l2}</span></div>)}
        </div>

        {/* 메뉴 5개 */}
        <div className="mxk-splash-menu">
          {menus.map(m=>(
            <button type="button" className="mxk-splash-menu-item" data-splash-tab={m.tab} key={m.tab}
              onClick={()=>onSelectTab&&onSelectTab(m.tab)} aria-label={m.label}>
              <span className="mxk-splash-menu-icon"><SplashIcon name={m.icon}/></span>
              <span className="mxk-splash-menu-label">{m.label}</span>
              <span className="mxk-splash-menu-arrow" aria-hidden="true">→</span>
            </button>
          ))}
        </div>

      </div>

      {/* 푸터 */}
      <div className="mxk-splash-footer" style={{position:"relative"}}>
        <div className="mxk-splash-footer-meta">
          <span data-i18n-skip="true" style={{display:"inline-flex",alignItems:"center",gap:9}}>
            <span style={{letterSpacing:".6px"}}>Developer</span>
            <span style={{color:SPLASH_DIM,fontWeight:600}}>{developer}</span>
            <span aria-hidden="true" style={{width:1,height:11,background:SPLASH_LINE}}/>
            <span>{department}</span>
          </span>
        </div>
        <div aria-hidden="true" style={{height:4,background:C.red}}/>
      </div>
    </div>
  );
}



const MATERIAL_LIST_PAGE_SIZE=100;

function expandedPhotoDisplaySize(naturalWidth,naturalHeight){
  const width=Math.max(1,Number(naturalWidth)||1);
  const height=Math.max(1,Number(naturalHeight)||1);
  const viewportWidth=typeof window!=="undefined"?window.innerWidth:1200;
  const viewportHeight=typeof window!=="undefined"?window.innerHeight:800;
  const maxWidth=Math.max(240,Math.min(1100,viewportWidth-28));
  const maxHeight=Math.max(180,viewportHeight-150);
  const fitScale=Math.min(maxWidth/width,maxHeight/height);
  const minimumExpandedWidth=Math.min(720,maxWidth);
  const scale=Math.max(fitScale,minimumExpandedWidth/width);
  return {displayWidth:Math.round(width*scale),displayHeight:Math.round(height*scale)};
}

export default function App(){
  const appRootRef=useRef(null);
  const [entered,setEntered]=useState(false);
  const [lang,setLang]=useState(()=>{
    try{ return localStorage.getItem("xrf-ui-language")==="en"?"en":"ko"; }catch{return "ko";}
  });
  const [tab,setTab]=useState("list");
  const [fs,setFs]=useState({category:[],photo:[],code:[],name:[],dept:[],cycle:[],firstDate:[],xrf:[],compliance:[],precision:[],crLevel:[],risk:[],retest:[],approval:[],lifecycle:[],nextDue:[],dDay:[],dueMonth:[]});
  const [openFilter,setOpenFilter]=useState(null);
  const [fPos,setFPos]=useState({top:0,left:0});
  const [search,setSearch]=useState("");
  const [listPage,setListPage]=useState(0);
  const [visualUatOnly,setVisualUatOnly]=useState(false);
  const [selKey,setSelKey]=useState(null);
  const [rowSel,setRowSel]=useState(null);
  const [photoPreview,setPhotoPreview]=useState(null);
  const [selPeriodNum,setSelPeriodNum]=useState(null);
  const [selPeriodPage,setSelPeriodPage]=useState(null);
  const [precisionSelCaseId,setPrecisionSelCaseId]=useState(null);
  const [precisionDetailCaseId,setPrecisionDetailCaseId]=useState(null);
  const [precisionOverrides,setPrecisionOverrides]=useState({});
  const precisionSaveTimerRef=useRef(null);
  const [sharePointItems,setSharePointItems]=useState([]);
  const [dbLoading,setDbLoading]=useState(true);
  const [dbError,setDbError]=useState("");
  const [dbCounts,setDbCounts]=useState(null);
  const [selMeasurementId,setSelMeasurementId]=useState(null);
  const [xrfDetailView,setXrfDetailView]=useState("analysis");
  const [regType,setRegType]=useState(null);
  const [regStep,setRegStep]=useState(1);
  const [regCR,setRegCR]=useState("M");
  const [requestItems,setRequestItems]=useState([]);
  const [itemOverrides,setItemOverrides]=useState({});
  const [approvalOverrides,setApprovalOverrides]=useState({});
  const [itemEditOpen,setItemEditOpen]=useState(false);
  const [itemEditForm,setItemEditForm]=useState(null);
  const [itemEditAuth,setItemEditAuth]=useState({open:false,item:null,password:"",error:""});
  const [discontinueDecisionAuth,setDiscontinueDecisionAuth]=useState({open:false,item:null,password:"",error:"",verified:false,pending:false});
  const [regMode,setRegMode]=useState("requester");
  const [reqXrfMode,setReqXrfMode]=useState("request");
  const [reqUploadResult,setReqUploadResult]=useState(emptyUploadedXrfResult());
  const [reqUploadFileName,setReqUploadFileName]=useState("");
  const [reqPhotoFile,setReqPhotoFile]=useState(null);
  const [reqPhotoPreview,setReqPhotoPreview]=useState("");
  const [reqPhotoProgress,setReqPhotoProgress]=useState(0);
  const [adminUploadFileNameByCode,setAdminUploadFileNameByCode]=useState({});
  const [adminUploadResultByCode,setAdminUploadResultByCode]=useState({});
  const [periodUploadState,setPeriodUploadState]=useState({key:"",fileName:"",error:"",saved:false});
  const reqUploadFileRef=useRef(null);
  const measurementSavePendingRef=useRef(false);
  const [measurementSavePending,setMeasurementSavePending]=useState(false);
  const [reqForm,setReqForm]=useState({
    partNumber:"", itemName:"", itemNameEn:"", type:"Auxiliary Materials", handlingDept:"Assembly", materialCategory:REQUESTER_MATERIAL_CATEGORY_OPTIONS[0],
    materialState:"", manufacturer:"", useDate:TODAY_STR,
    crType:"직접접촉+비잔류",
    requestReason:"",
    replacementTargetCode:"",
    replacementReason:"기존 품목 단종",
    replacementReasonOther:"",
    oldItemDisposition:"신규 품목 승인 후 기존 재고 소진 후 단종",
    replacementDate:TODAY_STR,
    discontinueTargetCode:"",
    discontinueReason:"사용 중지",
    discontinueReasonOther:"",
    finalUseDate:TODAY_STR,
    stockDisposition:"재고 소진 후 사용 중지",
    note:""
  });
  const currentUserRole="admin"; // 실제 연동 시 로그인 사용자 또는 권한 그룹으로 치환
  const isAdminUser=currentUserRole==="admin";

  const reloadSharePointDb=useCallback(async()=>{
    setDbLoading(true);
    setDbError("");
    try{
      const payload=await fetchSharePointBootstrap();
      const runtime=buildRuntimeFromSharePoint(payload);
      setSharePointItems(runtime.items);
      setRequestItems(runtime.requestItems);
      setPrecisionOverrides(runtime.precisionOverrides);
      setApprovalOverrides({});
      setItemOverrides({});
      setDbCounts(runtime.counts);
    }catch(error){
      console.error(error);
      setDbError(error?.message||String(error));
    }finally{
      setDbLoading(false);
    }
  },[]);

  useEffect(()=>{
    reloadSharePointDb();
  },[reloadSharePointDb]);

  const persistSharePointAction=useCallback(async(action,payload,{reload=true,silent=false}={})=>{
    if(VISUAL_UAT_MUTATION_ACTIONS.has(action) && containsVisualUatReference(payload)){
      if(typeof window!=="undefined") window.alert(VISUAL_UAT_READ_ONLY_MESSAGE);
      return null;
    }
    try{
      const result=await mutateSharePoint(action,payload);
      if(reload) await reloadSharePointDb();
      return result;
    }catch(error){
      console.error(`[SharePoint ${action}]`,error);
      if(!silent && typeof window!=="undefined") window.alert(error?.message||String(error));
      if(reload) await reloadSharePointDb();
      return null;
    }
  },[reloadSharePointDb]);

  // UI 언어는 표시 계층에서만 변환합니다. 데이터값/필터값/Workflow 비교값은 한국어 원본을 유지합니다.
  useEffect(()=>{
    try{ localStorage.setItem("xrf-ui-language",lang); }catch{}
    if(typeof document!=="undefined") document.documentElement.lang=lang==="en"?"en":"ko";
    const root=appRootRef.current;
    if(!root) return;

    applyUiLanguage(root,lang);

    let applying=false;
    const observer=typeof MutationObserver!=="undefined" ? new MutationObserver(mutations=>{
      if(applying) return;
      applying=true;
      try{
        for(const mutation of mutations){
          if(mutation.type==="characterData") localizeUiTextNode(mutation.target,lang);
          if(mutation.type==="attributes") localizeUiElementAttributes(mutation.target,lang);
          mutation.addedNodes?.forEach(node=>{
            if(node.nodeType===3) localizeUiTextNode(node,lang);
            else if(node.nodeType===1) applyUiLanguage(node,lang);
          });
        }
      }finally{
        applying=false;
      }
    }) : null;
    observer?.observe(root,{subtree:true,childList:true,characterData:true,attributes:true,attributeFilter:UI_TRANSLATABLE_ATTRS});

    // 브라우저 기본 alert / confirm / prompt도 현재 표시 언어에 맞춥니다.
    const nativeAlert=typeof window!=="undefined"?window.alert:null;
    const nativeConfirm=typeof window!=="undefined"?window.confirm:null;
    const nativePrompt=typeof window!=="undefined"?window.prompt:null;
    if(typeof window!=="undefined"){
      if(nativeAlert) window.alert=(message)=>nativeAlert.call(window,lang==="en"?translateUiTextKoToEn(String(message??"")):message);
      if(nativeConfirm) window.confirm=(message)=>nativeConfirm.call(window,lang==="en"?translateUiTextKoToEn(String(message??"")):message);
      if(nativePrompt) window.prompt=(message,defaultValue)=>nativePrompt.call(window,lang==="en"?translateUiTextKoToEn(String(message??"")):message,defaultValue);
    }

    return ()=>{
      observer?.disconnect();
      if(typeof window!=="undefined"){
        if(nativeAlert) window.alert=nativeAlert;
        if(nativeConfirm) window.confirm=nativeConfirm;
        if(nativePrompt) window.prompt=nativePrompt;
      }
    };
  },[lang,entered]);

  // 품목이 바뀌면 이전 품목에서 수동으로 선택했던 주기 페이지를 이어받지 않습니다.
  useEffect(()=>{
    setSelPeriodPage(null);
    setSelPeriodNum(null);
    setSelMeasurementId(null);
    setXrfDetailView("analysis");
    setItemEditOpen(false);
    setItemEditForm(null);
    setItemEditAuth({open:false,item:null,password:"",error:""});
    setDiscontinueDecisionAuth({open:false,item:null,password:"",error:"",verified:false,pending:false});
  },[selKey]);
  // 탭별로 따로 계산하지 않고, 모든 탭이 같은 runtime workflow를 읽도록 단일 연결 구조를 구성합니다.
  // XRF 업로드(itemOverrides) → 정밀분석 입력(precisionOverrides) → Workflow/Approval 파생 → 목록/XRF/정밀분석/Risk 탭이 일괄 갱신됩니다.
  const realBaseItems=useMemo(()=>applyPeriodBasedCategories(deriveReplacementRelations([
    ...requestItems,
    ...sharePointItems.map(item=>{
      const key=itemStateKey(item);
      return itemOverrides[key] ? {...item, ...itemOverrides[key], _sourceCode:key} : item;
    }),
    ...(ENABLE_WORKFLOW_MOCK_DATA ? MOCK_WORKFLOW_ITEMS : [])
  ]).map(applyCurrentRiskPolicy)),[requestItems,sharePointItems,itemOverrides]);
  const visualUatBaseItems=useMemo(()=>VISUAL_UAT_MODE
    ? applyPeriodBasedCategories(deriveReplacementRelations(VISUAL_UAT_ALL_ITEMS).map(applyCurrentRiskPolicy))
    : [],[]);
  const baseItems=useMemo(()=>[...realBaseItems,...visualUatBaseItems],[realBaseItems,visualUatBaseItems]);
  const runtimePrecisionOverrides=useMemo(()=>VISUAL_UAT_MODE
    ? {...precisionOverrides,...VISUAL_UAT_PRECISION_OVERRIDES}
    : precisionOverrides,[precisionOverrides]);
  const allItems=useMemo(()=>baseItems.map(item=>({
    ...item,
    __workflow:deriveRuntimeWorkflow(item,runtimePrecisionOverrides,approvalOverrides)
  })),[baseItems,runtimePrecisionOverrides,approvalOverrides]);
  const realItems=useMemo(()=>allItems.filter(item=>!item.__visualUat),[allItems]);
  const currentPortfolioPeriodNo=useMemo(()=>derivePortfolioCurrentPeriodNo(realBaseItems),[realBaseItems]);
  const replacementChains=useMemo(()=>buildReplacementRelationGroups(allItems),[allItems]);
  const approvalStatusOf=useCallback((item)=>{
    const workflowStatus=workflowForItem(item)?.approval?.status;
    if(workflowStatus) return workflowStatus;
    const baseStatus=String(item?.approvalStatus||"").trim();
    const overrideStatus=String(approvalOverrides[item?.code]||"").trim();
    // SharePoint/마스터에서 관리자가 최종 Approval_Status를 입력한 경우,
    // 로컬에 남아 있는 과거 "측정 진행중" 임시값이 최종 상태를 덮어쓰지 않도록 한다.
    if(baseStatus && approvalFilterKey(baseStatus)!=="measuring"
      && overrideStatus && approvalFilterKey(overrideStatus)==="measuring"){
      return baseStatus;
    }
    return overrideStatus || baseStatus || "측정 진행중";
  },[approvalOverrides]);
  const selectableExistingItems=useMemo(
    ()=>realBaseItems.filter(item=>!item.__visualUat && !isDiscontinueRequestItem(item) && item.isCurrent!==false && item.lifecycle!=="ReplacedOld" && item.lifecycle!=="Discontinued"),
    [realBaseItems]
  );
  useEffect(()=>{
    if(!VISUAL_UAT_MODE || typeof window==="undefined") return;
    const byCode=new Map(allItems.filter(item=>item.__visualUat).map(item=>[item.code,item]));
    const rows=VISUAL_UAT_EXPECTATIONS.map(manifest=>{
      const item=byCode.get(manifest.code);
      return item
        ? visualUatSelfCheckRow(item,manifest,approvalStatusOf)
        : {code:manifest.code,group:manifest.group,PASS:false,failedFields:["fixture missing"],expected:manifest.expected,actual:null};
    });
    window.__VISUAL_UAT_SELF_CHECK__=rows;
    window.__VISUAL_UAT_FIXTURE_COUNT__=byCode.size;
    console.groupCollapsed(`[Visual UAT] ${rows.filter(row=>row.PASS).length}/${rows.length} expectations passed`);
    console.table(rows.map(row=>({
      code:row.code,group:row.group,expectedXrf:row.expectedXrf||"",actualXrf:row.actualXrf||"",
      expectedRisk:row.expectedRisk||"",actualRisk:row.actualRisk||"",expectedCycle:row.expectedCycle||"",
      actualCycle:row.actualCycle||"",expectedApproval:row.expectedApproval||"",actualApproval:row.actualApproval||"",
      expectedStage:row.expectedStage||"",actualStage:row.actualStage||"",PASS:row.PASS,
      failedFields:(row.failedFields||[]).join(", ")
    })));
    console.groupEnd();
    const complianceRows=VISUAL_UAT_COMPLIANCE_EXPECTATIONS.map(manifest=>{
      const item=byCode.get(manifest.code);
      return item
        ? visualUatComplianceSelfCheckRow(item,manifest)
        : {code:manifest.code,PASS:false,failedFields:["fixture missing"],expected:manifest.expected,actual:null};
    });
    window.__VISUAL_UAT_COMPLIANCE_SELF_CHECK__=complianceRows;
    window.__VISUAL_UAT_COMPLIANCE_COMPACT_WINDOW_MISMATCH__=complianceRows.some(row=>
      row.code==="VU-CMP-L12"
      && row.actual?.periodStatus==="upcoming"
      && row.actual?.compactPotential30DayOverride===true
    );
    console.groupCollapsed(`[Compliance Visual UAT] ${complianceRows.filter(row=>row.PASS).length}/${complianceRows.length} logic expectations passed`);
    console.table(complianceRows.map(row=>({
      code:row.code,cycle:row.cycle||"",p1Due:row.p1Due||"",p1Measured:row.p1Measured||"",
      expectedPeriodStatus:row.expectedPeriodStatus||"",actualPeriodStatus:row.actualPeriodStatus||"",
      expectedCurrentStatus:row.expectedCurrentStatus||"",actualCurrentStatus:row.actualCurrentStatus||"",
      compactLabels:row.compactLabels||"",compactCount:row.compactCount??"",
      businessStatus:row.businessStatus||"",compactVisualExpected:row.compactVisualExpected||"",
      compactVisualActual:row.compactVisualActual||"",
      compactPotential30DayOverride:row.compactPotential30DayOverride??false,
      PASS:row.PASS,failedFields:(row.failedFields||[]).join(", ")
    })));
    console.info("COMPACT_WINDOW_MISMATCH =",window.__VISUAL_UAT_COMPLIANCE_COMPACT_WINDOW_MISMATCH__?"FOUND":"NOT FOUND");
    console.groupEnd();
  },[allItems,approvalStatusOf]);
  const precisionRows=useMemo(()=>{
    const precisionStages=new Set([
      "PRECISION_TRANSFER_REQUEST_REQUIRED","PRECISION_TRANSFER_RESULT_WAITING",
      "PRECISION_FOLLOWUP_REQUIRED","PRECISION_REQUEST_REQUIRED","PRECISION_RESULT_WAITING","PRECISION_NG_REVIEW"
    ]);
    return allItems.flatMap(item=>{
      if(isDiscontinueRequestItem(item)) return [];
      const wf=workflowForItem(item);

      // 실제 품목은 현재 최신 XRF가 정밀분석 조건을 충족할 때 해당 최신 측정 건을 Case로 표시합니다.
      // R Final Risk가 H/Not Allowed여도 최신 XRF가 OK이면 과거 R NG만으로 정밀분석 대기 Case를 유지하지 않습니다.
      if(!wf.caseData && Array.isArray(wf.precisionCases) && wf.precisionCases.length){
        return wf.precisionCases.map(c=>({
          key:c.key,item,caseData:null,measurement:c.measurement,precision:c.precision,targets:c.targets,
          triggerLabel:c.triggerLabel,followUp:null,approval:wf.approval||null,
          sourceFileName:c.sourceFileName||"",finalConfirm:c.finalConfirm||"",
          elementResults:c.elementResults||c.precision?.elementResults||{}
        }));
      }

      // Workflow mock / 레거시 Case는 기존 단일 Case 표시를 유지합니다.
      const measurement=wf.latestXrf || latestMeasurementOf(item);
      const targets=precisionTargetsForWorkflow(item,wf);
      const needsPrecision=!!wf.precision || precisionStages.has(wf.stage) || targets.length>0;
      if(!needsPrecision) return [];
      const key=workflowOverrideKey(item,wf);
      if(!key) return [];
      const ov=runtimePrecisionOverrides[key]||{};
      const precision=wf.precision || {required:true,targetElements:targets,requestStatus:"NOT_REQUESTED",uploadStatus:"WAITING",result:""};
      const triggerLabel=precisionTriggerReason(measurement)!=="대상 아님"
        ? precisionTriggerReason(measurement)
        : (String(wf.latestXrf?.normalizedResult||"").toUpperCase()==="NG" || String(wf.latestXrf?.reportJudgement||"").toUpperCase()==="NG" ? "XRF NG 판정"
          : (String(wf.latestXrf?.normalizedResult||"").toUpperCase()==="REMEASURE" ? "재측정 후 ?? 판정" : "정밀분석 필요"));
      return [{
        key,item,caseData:wf.caseData||null,measurement,precision,targets,triggerLabel,
        followUp:wf.followUp||null,approval:wf.approval||null,
        sourceFileName:ov.sourceFileName||measurement?.sourceFileName||"",
        finalConfirm:ov.finalConfirm||precision.finalConfirm||"",
        elementResults:precision.elementResults||ov.elementResults||{},
      }];
    }).sort((a,b)=>String(b.measurement?.date||b.measurement?.measuredDate||b.precision?.requestedAt||"").localeCompare(String(a.measurement?.date||a.measurement?.measuredDate||a.precision?.requestedAt||"")));
  },[allItems,runtimePrecisionOverrides]);
  const updatePrecisionOverride=useCallback((key,patch)=>{
    const fixtureOwner=allItems.find(item=>item.__visualUat && !!measurementById(item,key));
    if(fixtureOwner){
      if(typeof window!=="undefined") window.alert(VISUAL_UAT_READ_ONLY_MESSAGE);
      return;
    }
    setPrecisionOverrides(prev=>{
      const nextCase={...(prev[key]||{}),...patch};
      const next={...prev,[key]:nextCase};
      if(precisionSaveTimerRef.current) clearTimeout(precisionSaveTimerRef.current);
      precisionSaveTimerRef.current=setTimeout(()=>{
        const owner=allItems.find(item=>!!measurementById(item,key)) || null;
        const wf=owner?workflowForItem(owner):null;
        const targets=owner?precisionTargetsForWorkflow(owner,wf):[];
        const precision=wf?.precision || {required:true,targetElements:targets,elements:{}};
        const finalResult=derivePrecisionResult(targets,precision,nextCase.elementResults||{});
        void persistSharePointAction("upsertPrecision",{
          triggerMeasurementId:key,
          precisionId:nextCase.precisionId||null,
          requestSpItemId:owner?._spRequestItemId||null,
          finalResult,
          actor:"admin",
          override:nextCase
        },{reload:nextCase.finalConfirm==="확인"});
      },500);
      return next;
    });
  },[allItems,persistSharePointAction]);
  const markPrecisionRequested=useCallback(async(key,password)=>{
    const fixtureOwner=allItems.find(item=>item.__visualUat && !!measurementById(item,key));
    if(fixtureOwner){
      if(typeof window!=="undefined") window.alert(VISUAL_UAT_READ_ONLY_MESSAGE);
      return false;
    }
    const owner=allItems.find(item=>!!measurementById(item,key)) || null;
    const wf=owner?workflowForItem(owner):null;
    const targets=owner?precisionTargetsForWorkflow(owner,wf):[];
    const precision=wf?.precision || {required:true,targetElements:targets,elements:{}};
    const existingOverride=runtimePrecisionOverrides[key]||{};
    const requestOverride={...existingOverride,requestStatus:"REQUESTED",requestedAt:TODAY_STR,requestedBy:"admin"};
    const result=await persistSharePointAction("upsertPrecision",{
      triggerMeasurementId:key,
      precisionId:existingOverride.precisionId||precision.precisionId||null,
      requestSpItemId:owner?._spRequestItemId||null,
      finalResult:derivePrecisionResult(targets,precision,requestOverride.elementResults||{}),
      actor:"admin",
      precisionHandoffPassword:password,
      override:requestOverride
    });
    return !!result;
  },[allItems,persistSharePointAction,runtimePrecisionOverrides]);
  useEffect(()=>()=>{ if(precisionSaveTimerRef.current) clearTimeout(precisionSaveTimerRef.current); },[]);
  useEffect(()=>{
    setApprovalOverrides(prev=>{
      let changed=false;
      const next={...prev};
      for(const item of allItems){
        const code=item?.code;
        if(!code || !next[code]) continue;
        const baseStatus=String(item?.approvalStatus||"").trim();
        if(baseStatus && approvalFilterKey(baseStatus)!=="measuring"
          && approvalFilterKey(next[code])==="measuring"){
          delete next[code];
          changed=true;
        }
      }
      return changed ? next : prev;
    });
  },[allItems]);
  const updateRequestField=useCallback((field,value)=>setReqForm(prev=>{
    const next={...prev,[field]:value};
    if(field==="replacementReason" && value!=="기타") next.replacementReasonOther="";
    if(field==="discontinueReason" && value!=="기타") next.discontinueReasonOther="";
    if(field==="materialCategory" && isChemicalMaterialCategory(value)){
      next.materialState=isValidChemicalMaterialState(prev.materialState) ? prev.materialState : "";
    }
    return next;
  }),[]);
  const goNextRequesterStep=useCallback(()=>{
    if(regStep===2 && regType!=="discontinue" && !reqPhotoFile){
      window.alert("품목 사진을 선택하세요.");
      return;
    }
    if(regStep===2 && regType!=="discontinue" && isChemicalMaterialCategory(reqForm.materialCategory)
      && !isValidChemicalMaterialState(reqForm.materialState)){
      window.alert("Material_Category가 Chemical인 경우 Material_State를 필수로 선택하세요.");
      return;
    }
    const maxStep=regType==="discontinue" ? 3 : 4;
    setRegStep(step=>Math.min(maxStep,step+1));
  },[regStep,regType,reqForm.materialCategory,reqForm.materialState,reqPhotoFile]);
  const setItemApproval=useCallback((code,status)=>setApprovalOverrides(prev=>({...prev,[code]:status})),[]);
  const handleRequesterXrfUpload=useCallback(async(e)=>{
    const file=e.target.files?.[0];
    if(!file) return;
    reqUploadFileRef.current=file;
    setReqUploadFileName(file.name);
    setReqUploadResult(emptyUploadedXrfResult(file.name,{pending:true}));
    const result=await parseUploadedXrfFile(file);
    setReqUploadResult(result);
    e.target.value="";
  },[]);
  const clearRequestPhoto=useCallback(()=>{
    setReqPhotoFile(null);
    setReqPhotoPreview("");
    setReqPhotoProgress(0);
  },[]);
  const handleRequestPhoto=useCallback((e)=>{
    const file=e.target.files?.[0]||null;
    e.target.value="";
    if(!file) return;
    if(file.size<1 || file.size>10*1024*1024){
      window.alert("품목 사진은 10MB 이하 파일만 선택하세요.");
      return;
    }
    if(!/^image\/(jpeg|png|webp)$/i.test(file.type||"") && !/\.(jpe?g|png|webp)$/i.test(file.name||"")){
      window.alert("품목 사진은 JPG, PNG, WEBP 형식만 선택하세요.");
      return;
    }
    setReqPhotoFile(file);
    setReqPhotoProgress(0);
    const reader=new FileReader();
    reader.onload=()=>setReqPhotoPreview(String(reader.result||""));
    reader.readAsDataURL(file);
  },[]);
  const commitParsedMeasurement=useCallback(async(targetItem,result,targetPeriod=null,sourceLabel="XRF 결과 업로드",sourceFile=null)=>{
    if(!targetItem || !result?.parsed) return null;
    if(targetItem.__readOnlyFixture){
      if(typeof window!=="undefined") window.alert(VISUAL_UAT_READ_ONLY_MESSAGE);
      return null;
    }
    const measuredDate=result.measuredDate || TODAY_STR;
    const historyRows=Array.isArray(targetItem.history)?targetItem.history:[];
    const sameDateCount=historyRows.filter(h=>dateOnly(h?.date)===measuredDate).length;
    const sequence=sameDateCount+1;
    const isInitial=!(targetItem.firstMeasurementId || historyRows.length || targetItem.lastMeasured);
    const allTargetPeriods=!isInitial ? itemAllPeriods(targetItem) : [];
    const fallbackPeriod=!isInitial ? (currentCompliancePeriod(targetItem) || allTargetPeriods.find(p=>!(p?.isReference||Number(p?.num)===0)&&!p?.measurementId&&!p?.measured)) : null;
    // 연 1회 품목은 업로드 대상 행보다 실제 측정연도를 우선합니다.
    // 같은 연도의 추가 측정은 항상 그 연도에 해당하는 동일 Pn으로 다시 묶입니다.
    const annualPeriodForMeasuredYear=!isInitial && cycleMonthsOf(targetItem)===12
      ? allTargetPeriods.find(p=>!(p?.isReference||Number(p?.num)===0) && String(p?.due||p?.targetDate||"").slice(0,4)===String(measuredDate).slice(0,4))
      : null;
    const resolvedPeriod=isInitial ? null : (annualPeriodForMeasuredYear || targetPeriod || fallbackPeriod);
    const periodNo=isInitial ? 0 : Number(resolvedPeriod?.num || 1);
    const isReferenceRetest=!isInitial && (resolvedPeriod?.isReference || periodNo===0);
    const periodAttempts=isInitial ? [] : measurementAttemptsForPeriod(targetItem,resolvedPeriod);
    const latestPeriodRow=periodAttempts[periodAttempts.length-1]||null;
    const latestPeriodId=latestPeriodRow?.id||latestPeriodRow?.measurementId||null;
    const latestPeriodMeasurement=latestPeriodId?measurementById(targetItem,latestPeriodId):null;
    const latestPeriodNeedsRemeasure=pendingXrfRemeasureElements(latestPeriodMeasurement).length>0;
    const latestPeriodRole=String(latestPeriodRow?.measurementRole||latestPeriodRow?.role||latestPeriodMeasurement?.role||"");
    let role="Initial";
    let attemptNo=1;
    let parentMeasurementId=null;
    if(!isInitial){
      const allowAnnualMultiple=cycleMonthsOf(targetItem)===12;
      if(!latestPeriodRow){
        role="Periodic";
        attemptNo=1;
      }else if(latestPeriodNeedsRemeasure){
        // 같은 측정건의 ?? 보완은 Retest 1회만 허용합니다.
        if(latestPeriodRole==="Retest" || Number(latestPeriodRow?.attemptNo||latestPeriodMeasurement?.attemptNo||1)>=2){
          if(typeof window!=="undefined") window.alert("XRF 재측정은 1회까지만 가능합니다. 이후에는 정밀분석 단계에서 처리하세요.");
          return null;
        }
        role="Retest";
        attemptNo=2;
        parentMeasurementId=latestPeriodId;
      }else if(allowAnnualMultiple){
        // 연 1회 품목은 같은 달력연도 안의 별도 측정을 개수 제한 없이 같은 Pn에 누적합니다.
        role="Periodic";
        attemptNo=1;
      }else{
        if(typeof window!=="undefined") window.alert("해당 주기에는 이미 XRF 측정값이 있습니다. 추가 등록은 재측정(??)이 필요한 경우에만 가능합니다.");
        return null;
      }
    }
    const measurementId=`${targetItem.code}_${measuredDate.replaceAll("-","")}_${String(sequence).padStart(2,"0")}`;
    const measurementBase={elements:result.elements||{},role,attemptNo};
    const xrfLevel=measurementXrfLevel(measurementBase);
    const xrfResult=measurementXrfResult(measurementBase);
    const exceededElements=xrfExceededElements(measurementBase);
    const precisionElements=precisionRequiredElements(measurementBase);
    const precisionRequired=precisionElements.length>0;
    const xrfRemeasureRequired=XRF_REPORT_ELEMENTS.some(el=>isXrfRemeasureRequired(result.elements?.[el]));
    let finalRisk=itemFinalRisk(targetItem) || "—";
    let cycle=itemCycleFromRisk(targetItem);
    let cycleMonths=cycleMonthsFromCycle(cycle);
    if(isInitial){
      const initialRisk=evaluateInitialRisk(measurementBase,targetItem.crLevel,isEquipmentItem(targetItem));
      finalRisk=initialRisk.finalRisk || "—";
      cycle=initialRisk.cycle;
      cycleMonths=initialRisk.cycleMonths;
    }
    const policyApproval=deriveInitialApprovalStatus({elements:result.elements||{},xrfResult,xrfRemeasureRequired,precisionRequired});
    const measurement={
      id:measurementId,
      date:measuredDate,
      sequence,
      role,
      attemptNo,
      parentMeasurementId,
      xrfWorst:xrfResult,
      xrfResult,
      xrfLevel,
      finalRisk,
      approvalStatus:policyApproval,
      elements:result.elements||{},
      exceededElements,
      retestRequired:xrfRemeasureRequired,
      retestRequiredElements:result.retestRequiredElements||"",
      xrfRemeasureRequired,
      precisionRequired,
      precisionRequiredElements:precisionElements,
      precisionStatus:precisionRequired ? "Required" : "Not Required",
      dataMissingElements:result.dataMissingElements||[],
      reviewRequiredElements:result.reviewRequiredElements||[],
      sourceFileName:result.fileName||"",
      sourceSheet:result.sourceSheet||"",
    };
    const measurementFollowup=measurementFollowupInfo(measurement);

    let periods=itemAllPeriods(targetItem).map(p=>({...p}));
    if(isInitial){
      periods=[{
        num:0,label:"R",displayLabel:"R",displaySeq:0,isReference:true,
        referenceDate:measuredDate,date:measuredDate,start:null,due:null,targetDate:null,windowStart:null,
        measured:measuredDate,measuredDates:[measuredDate],status:"reference",statusLabel:"등록완료",daysOver:0,
          measurementId,measurementIds:[measurementId],closeType:"registered",finalRisk,isDisplayTarget:false,
          sourceBasis:`${sourceLabel} · 최초 등록 기준 R`,cycleMonths:cycleMonths||null,windowDays:cycle==="Monthly"?15:30,
          prevDue:null,nextDue:null,cycleVersion:1,isCurrentCycle:true,measurementRole:role,
          xrfLevel,cycle,riskAssessment:true,riskAssessmentType:"Initial"
      }];
    }else if(isReferenceRetest){
      // R Retest는 새로운 P0를 만들거나 R 위험도 기준 ID를 교체하지 않습니다.
      // 최초 Initial ID는 measurementId로 유지하고, Retest는 measurementIds에 이력으로 연결합니다.
      periods=periods.map(p=>{
        if(!(p?.isReference || Number(p?.num)===0)) return p;
        const ids=periodMeasurementIds(p);
        if(!ids.includes(measurementId)) ids.push(measurementId);
        const dates=periodMeasuredDates(p);
        if(measuredDate && !dates.includes(measuredDate)) dates.push(measuredDate);
        dates.sort();
        return {
          ...p,
          measurementId:p.measurementId||targetItem.firstMeasurementId||latestPeriodId,
          measurementIds:ids,
          measured:measuredDate,
          measuredDates:dates,
          status:"reference",
          statusLabel:"R 재측정 완료",
          measurementRole:"Initial",
          sourceBasis:`${sourceLabel} · R Retest 연결`
        };
      });
    }else{
      const targetNum=Number(resolvedPeriod?.num||periodNo);
      let found=false;
      periods=periods.map(p=>{
        if(Number(p?.num)!==targetNum || p?.isReference){ return p; }
        found=true;
        const due=p.due||p.targetDate||null;
        const late=!!due && dateGt(measuredDate,due);
        const links=appendPeriodMeasurementLink(p,measurementId,measuredDate);
        return {
          ...p,
          ...links,
          status:late?"late":"compliant",
          statusLabel:late?"지연측정":"이행",
          daysOver:late?Math.max(0,diffDaysISO(measuredDate,due)):0,
          closeType:late?"measured_late":"measured",
          finalRisk,
          isDisplayTarget:true,
          measurementRole:role,
          sourceBasis:`${sourceLabel} · ${periodLabel(p)} 측정 연결`,
        };
      });
      if(!found){
        const due=resolvedPeriod?.due||resolvedPeriod?.targetDate||null;
        const late=!!due && dateGt(measuredDate,due);
        periods.push({
          ...(resolvedPeriod||{}),num:targetNum,label:`P${targetNum}`,displayLabel:`P${targetNum}`,displaySeq:targetNum,isReference:false,
          due,targetDate:due,measured:measuredDate,measuredDates:[measuredDate],measurementId,measurementIds:[measurementId],status:late?"late":"compliant",
          statusLabel:late?"지연측정":"이행",daysOver:late?Math.max(0,diffDaysISO(measuredDate,due)):0,
          closeType:late?"measured_late":"measured",finalRisk,isDisplayTarget:true,measurementRole:role,
          sourceBasis:`${sourceLabel} · P${targetNum} 측정 연결`
        });
        periods.sort((a,b)=>Number(a?.displaySeq??a?.num??0)-Number(b?.displaySeq??b?.num??0));
      }
    }

    const historyEntry={
      date:measuredDate,id:measurementId,periodNo,label:(isInitial||isReferenceRetest)?"R":`P${periodNo}`,measurementRole:role,attemptNo,
      xrf_worst:xrfResult,xrfWorst:xrfResult,xrfResult,xrfLevel,finalRisk,approvalStatus:policyApproval,
      retestRequired:xrfRemeasureRequired,retestRequiredElements:result.retestRequiredElements||"",
      exceededElements,precisionRequired,precisionRequiredElements:precisionElements,
      dataMissingElements:result.dataMissingElements||[],reviewRequiredElements:result.reviewRequiredElements||[],elements:result.elements||{},status:isInitial?"reference":"measured",
      sourceFileName:result.fileName||"",sourceSheet:result.sourceSheet||""
    };
    const nextOpen=periods
      .filter(p=>!(p?.isReference||Number(p?.num)===0)&&!p?.measured&&p?.due)
      .sort((a,b)=>String(a.due).localeCompare(String(b.due)))[0] || null;
    const updatedItem={
      ...targetItem,
      approvalStatus:policyApproval,
      approvalDetail:(result.reviewRequiredElements||[]).length
        ? `Judgment 확인 필요: ${(result.reviewRequiredElements||[]).join(", ")}`
        : (result.dataMissingElements||[]).length
          ? `데이터 확인 필요: ${(result.dataMissingElements||[]).join(", ")}`
          : measurementFollowup.hasFollowup
            ? measurementFollowup.detail
            : "",
      approvalCompletedAt:policyApproval==="승인"?measuredDate:null,
      approvalCompletedBy:policyApproval==="승인"?"admin":null,
      // Action_Note에는 사용자가 작성한 업무 메모만 유지하고 업로드 상태 문구는 저장하지 않습니다.
      actionNote:withoutGeneratedItemActionNote(targetItem.actionNote),
      latestXrfResult:xrfResult,
      latestXrfLevel:xrfLevel,
      xrfWorst:xrfResult,
      finalRisk,
      cycle,
      cycleMonths:cycleMonths||null,
      // 측정 의뢰가 이전 Pn에 생성되었더라도 최초 XRF(R)를 현재 Pn에서 등록하면
      // 그 현재 Pn의 신규 품목으로 판정되어야 하므로 R 등록 시점을 우선합니다.
      registrationPeriodNo:isInitial ? (currentPortfolioPeriodNo || targetItem.registrationPeriodNo || 1) : (targetItem.registrationPeriodNo || null),
      newCategoryUntil:null,
      nextDue:nextOpen?.due || (cycleMonths?addMonthsISO(measuredDate,cycleMonths):null),
      // 주기 이행은 분석 승인과 분리합니다. 측정 연결 결과는 periods/getCS가 판정합니다.
      complianceStatus:targetItem.complianceStatus || "ok",
      firstDate:targetItem.firstDate||measuredDate,
      lastMeasured:measuredDate,
      latestMeasurementId:measurementId,
      firstMeasurementId:targetItem.firstMeasurementId||measurementId,
      measurementCount:(targetItem.measurementCount||0)+1,
      riskBasisMeasurementId:isInitial ? measurementId : (targetItem.riskBasisMeasurementId||targetItem.firstMeasurementId||null),
      riskBasisXrfLevel:isInitial ? xrfLevel : riskBasisXrfLevelOf(targetItem),
      riskBasisCrLevel:isInitial ? targetItem.crLevel : (targetItem.riskBasisCrLevel||targetItem.crLevel||null),
      riskAssessmentType:isInitial ? "Initial" : (targetItem.riskAssessmentType||null),
      riskAssessmentVersion:isInitial ? 1 : (targetItem.riskAssessmentVersion||null),
      riskAssessedAt:isInitial ? measuredDate : (targetItem.riskAssessedAt||null),
      precisionRequired,
      retestRequired:xrfRemeasureRequired,
      retestRequiredElements:result.retestRequiredElements||"",
      periods,
      history:[...historyRows,historyEntry],
      latest:measurement,
    };

    // SharePoint가 source-of-truth가 되도록 Measurement header + 6개 원소 Raw 결과를 영구 저장합니다.
    // 저장 성공 전에는 화면의 Item/측정 이력을 변경하지 않아 DB 실패가 성공처럼 보이지 않게 합니다.
    if(measurementSavePendingRef.current) return null;
    measurementSavePendingRef.current=true;
    setMeasurementSavePending(true);
    let saved=null;
    try{
      saved=await persistSharePointAction("saveMeasurement",{
        itemId:targetItem.itemId,
        itemSpItemId:targetItem._spItemId||null,
        requestSpItemId:targetItem._spRequestItemId||null,
        requestStatus:requestStatusFromApproval(policyApproval),
        partNumber:targetItem.code,
        measuredDate,
        role,
        attemptNo,
        parentMeasurementId,
        sourceFileName:result.fileName||"",
        sourceSheet:result.sourceSheet||"",
        requesterName:"",
        elements:result.elements||{},
        actor:"admin",
        itemPatch:{
          actionNote:updatedItem.actionNote,
          registrationPeriodNo:updatedItem.registrationPeriodNo
        }
      });
    }finally{
      measurementSavePendingRef.current=false;
      setMeasurementSavePending(false);
    }
    if(!saved) return null;

    let sourceFileUploadError="";
    if(sourceFile){
      try{
        await uploadPrecisionBinary(saved.measurementId||measurementId,"source",sourceFile);
      }catch(error){
        sourceFileUploadError=error?.message||String(error);
        if(typeof window!=="undefined") window.alert(`XRF 측정값은 저장됐지만 원본 파일 저장에 실패했습니다.\n${sourceFileUploadError}`);
      }
    }

    const targetStateKey=itemStateKey(targetItem);
    setRequestItems(prev=>prev.map(item=>itemStateKey(item)===targetStateKey?{...updatedItem,_sourceCode:itemStateKey(item)}:item));
    setItemOverrides(prev=>({...prev,[targetStateKey]:{...(prev[targetStateKey]||{}),...updatedItem,_sourceCode:targetStateKey}}));
    setApprovalOverrides(prev=>({...prev,[targetItem.code]:policyApproval}));
    return {updatedItem,measurementId,periodNo,sourceFileUploadError};
  },[currentPortfolioPeriodNo,persistSharePointAction]);

  const buildRequesterPartNumber=useCallback((handlingDept)=>{
    const prefix=((handlingDept||"").trim().charAt(0)||"X").toUpperCase();
    const maxForPrefix=realBaseItems.reduce((max,item)=>{
      const m=String(item?.code||"").match(/^([A-Za-z])[-_]?(\d+)/);
      if(!m || m[1].toUpperCase()!==prefix) return max;
      const n=Number.parseInt(m[2],10);
      return Number.isFinite(n) ? Math.max(max,n) : max;
    },0);
    return `${prefix}-${String(maxForPrefix+1).padStart(2,"0")}`;
  },[realBaseItems]);
  const buildReplacementPartNumber=useCallback((targetCode)=>{
    const base=String(targetCode||"").trim();
    if(!base) return buildRequesterPartNumber(reqForm.handlingDept);
    const escapedBase=base.replace(/[.*+?^${}()|[\\]\\\\]/g,"\\$&");
    const rx=new RegExp(`^${escapedBase}-(\\d{2})$`);
    const maxSuffix=realBaseItems.reduce((max,item)=>{
      const m=String(item?.code||"").match(rx);
      if(!m) return max;
      const n=Number.parseInt(m[1],10);
      return Number.isFinite(n)?Math.max(max,n):max;
    },0);
    return `${base}-${String(maxSuffix+1).padStart(2,"0")}`;
  },[realBaseItems,buildRequesterPartNumber,reqForm.handlingDept]);
  const submitRequesterNewItem=useCallback(async()=>{
    const adminDirect=String(regType||"").startsWith("admin-");
    const requestType=regType==="admin-new"?"new":regType==="admin-replace"?"replacement":regType==="admin-discontinue"?"discontinue":(regType||"new");
    const effectiveXrfMode=adminDirect && requestType!=="discontinue" ? "upload" : reqXrfMode;
    const selectedReplacementItem=selectableExistingItems.find(i=>i.code===reqForm.replacementTargetCode);
    const selectedDiscontinueItem=selectableExistingItems.find(i=>i.code===reqForm.discontinueTargetCode);
    const replacementReasonFinal=reqForm.replacementReason==="기타"
      ? String(reqForm.replacementReasonOther||"").trim()
      : reqForm.replacementReason;
    const discontinueReasonFinal=reqForm.discontinueReason==="기타"
      ? String(reqForm.discontinueReasonOther||"").trim()
      : reqForm.discontinueReason;

    if(requestType==="discontinue" && !selectedDiscontinueItem){
      window.alert("단종 또는 사용 중지할 대상 품목을 선택하세요.");
      return;
    }
    if(requestType==="replacement" && !selectedReplacementItem){
      window.alert("대체할 기존 품목을 선택하세요.");
      return;
    }
    if(requestType!=="discontinue" && !String(reqForm.itemName||"").trim()){
      window.alert("Item_Name을 입력하세요.");
      return;
    }
    if(requestType!=="discontinue" && !String(reqForm.itemNameEn||"").trim()){
      window.alert("Item_Name_EN을 입력하세요.");
      return;
    }
    if(requestType!=="discontinue" && !String(reqForm.partNumber||"").trim()){
      window.alert("Part Number를 확인하세요.");
      return;
    }
    if(requestType!=="discontinue" && !reqPhotoFile){
      window.alert("품목 사진을 선택하세요.");
      return;
    }
    if(requestType==="discontinue" && reqForm.discontinueReason==="기타" && !discontinueReasonFinal){
      window.alert("단종 사유에서 기타를 선택한 경우 상세 설명을 입력하세요.");
      return;
    }
    if(requestType==="replacement" && reqForm.replacementReason==="기타" && !replacementReasonFinal){
      window.alert("대체 사유에서 기타를 선택한 경우 상세 설명을 입력하세요.");
      return;
    }
    if(requestType!=="discontinue" && isChemicalMaterialCategory(reqForm.materialCategory)
      && !isValidChemicalMaterialState(reqForm.materialState)){
      window.alert("Material_Category가 Chemical인 경우 Material_State를 필수로 선택하세요.");
      return;
    }

    if(adminDirect && requestType!=="discontinue" && !reqUploadResult?.parsed){
      window.alert("관리자 직접 등록은 XRF 원본 보고서를 먼저 선택하고 파싱을 완료해야 합니다.");
      return;
    }

    if(!adminDirect && requestType!=="discontinue" && effectiveXrfMode==="upload" && !reqUploadResult?.parsed){
      window.alert(reqUploadResult?.error || "XRF 원본 보고서 파싱을 완료한 뒤 등록 요청을 제출하세요.");
      return;
    }

    if(requestType==="discontinue"){
      const target=selectedDiscontinueItem || {};
      const code=`${target.code||"DISC"}-REQ-${TODAY_STR.replaceAll("-","")}`;
      const item={
        _isRequestRecord:true,
        code,
        name:`[단종 요청] ${target.name||"기존 품목"}`,
        nameEn:target.nameEn||"",
        dept:target.dept||reqForm.handlingDept,
        respDept:target.respDept||"PQE",
        type:target.type||"Auxiliary Materials",
        materialCategory:target.materialCategory||"부자재",
        materialState:target.materialState||"",
        manufacturer:target.manufacturer||"",
        useDate:reqForm.finalUseDate,
        crType:target.crType||"비접촉",
        crLevel:target.crLevel||"L",
        approvalStatus:adminDirect?"승인":"보류",
        actionNote:`의뢰자 등록 · 기존 품목 단종/사용 중지 요청\\n대상 품목: ${target.code||reqForm.discontinueTargetCode||"미선택"} · ${target.name||""}\\n사유: ${discontinueReasonFinal}\\n재고 처리: ${reqForm.stockDisposition}${reqForm.note?`\\n비고: ${reqForm.note}`:""}`,
        category:"changed",
        categoryLabel:"변경·대체",
        lifecycle:adminDirect?"Discontinued":(target.lifecycle||"ExistingActive"),
        isCurrent:!adminDirect,
        requestType:"discontinue",
        processStatus:adminDirect?"Applied":"Pending",
        processedAt:adminDirect?TODAY_STR:null,
        processedBy:adminDirect?"admin":null,
        discontinueTargetCode:target.code||reqForm.discontinueTargetCode,
        discontinueReason:discontinueReasonFinal,
        reasonCategory:reqForm.discontinueReason,
        reasonOtherDetail:reqForm.discontinueReason==="기타"?discontinueReasonFinal:"",
        reasonDetail:reqForm.note||"",
        reasonFinal:[discontinueReasonFinal,reqForm.note].filter(Boolean).join(" / "),
        finalUseDate:reqForm.finalUseDate,
        stockDisposition:reqForm.stockDisposition,
        replacementOf:null,
        replacedBy:null,
        replacementDate:reqForm.finalUseDate,
        firstDate:TODAY_STR,
        lastMeasured:null,
        latestMeasurementId:null,
        firstMeasurementId:null,
        measurementCount:0,
        xrfWorst:"—",
        finalRisk:"—",
        cycle:"Not Available",
        cycleMonths:null,
        nextDue:null,
        dDay:null,
        complianceStatus:"notAvailable",
        retestRequired:false,
        retestRequiredElements:"",
        periods:[],
        history:[],
        latest:null
      };
      if(!adminDirect){
        setRequestItems(prev=>[item,...prev.filter(x=>x.code!==code)]);
        setApprovalOverrides(prev=>({...prev,[code]:"보류"}));
      }
      if(adminDirect && target.code){
        const beforeTargetSnapshot={...target};
        const targetStateKey=itemStateKey(target);
        const terminalPatch={lifecycle:"Discontinued",isCurrent:false,nextDue:null,dDay:null,complianceStatus:"notAvailable",discontinuedByRequestCode:code,discontinueReason:discontinueReasonFinal,finalUseDate:reqForm.finalUseDate,replacementDate:reqForm.finalUseDate||TODAY_STR,beforeTargetSnapshot};
        setItemOverrides(prev=>({...prev,[targetStateKey]:{...(prev[targetStateKey]||{}),...terminalPatch,_sourceCode:targetStateKey}}));
        setRequestItems(prev=>prev.map(item=>itemStateKey(item)===targetStateKey?{...item,...terminalPatch}:item));
      }
      void persistSharePointAction("createRequest",{
        requestType:"Discontinue",
        status:adminDirect?"Applied":"Pending",
        targetItemId:target.itemId||null,
        targetSpItemId:target._spItemId||null,
        targetPartNumber:target.code||reqForm.discontinueTargetCode||"",
        requestedDate:TODAY_STR,
        actor:"admin",
        form:{...reqForm,itemName:target.name||reqForm.itemName,handlingDept:target.dept||reqForm.handlingDept,type:target.type||reqForm.type,materialCategory:target.materialCategory||reqForm.materialCategory,materialState:target.materialState||reqForm.materialState,crType:target.crType||reqForm.crType},
        beforeTargetSnapshot:adminDirect?{
          category:target.category,categoryLabel:target.categoryLabel,lifecycle:target.lifecycle,isCurrent:target.isCurrent,
          replacementDate:target.replacementDate||null,nextDue:target.nextDue||null,dDay:target.dDay??null,
          complianceStatus:target.complianceStatus||getCS(target),finalUseDate:target.finalUseDate||null
        }:null
      });
      setSelKey(null);
      setSelPeriodNum(null);
      setTab("list");
      setRegStep(1);
      setRegType(null);
      return;
    }

    const isReplacement=requestType==="replacement";
    const replacementTargetCode=selectedReplacementItem?.code||reqForm.replacementTargetCode||"";
    const code=(
      reqForm.partNumber
      || (isReplacement ? buildReplacementPartNumber(replacementTargetCode) : buildRequesterPartNumber(reqForm.handlingDept))
    ).trim();
    if(uploadedPartNumberFromFileName(reqUploadResult?.fileName) && effectiveXrfMode==="upload"
      && !samePartNumber(uploadedPartNumberFromFileName(reqUploadResult.fileName),code)){
      window.alert(`업로드 파일 품번 ${uploadedPartNumberFromFileName(reqUploadResult.fileName)}과 등록 품번 ${code}이 일치하지 않습니다.`);
      return;
    }
    const crLevel=crLevelFromType(reqForm.crType);
    const uploadParsed=effectiveXrfMode==="upload" && !!reqUploadResult?.parsed;
    const measurementBase={elements:reqUploadResult.elements||{}};
    const xrfLevel=uploadParsed ? measurementXrfLevel(measurementBase) : null;
    const xrfResult=uploadParsed ? measurementXrfResult(measurementBase) : "—";
    const xrfWorst=uploadParsed ? xrfResult : requestXrfWorstFromMode(effectiveXrfMode,reqUploadResult);
    const exceededElements=uploadParsed ? xrfExceededElements(measurementBase) : [];
    const precisionElements=uploadParsed ? precisionRequiredElements(measurementBase) : [];
    const precisionRequired=precisionElements.length>0;
    const retestRequired=uploadParsed ? XRF_REPORT_ELEMENTS.some(el=>isXrfRemeasureRequired(reqUploadResult.elements?.[el])) : false;
    const retestRequiredElements=uploadParsed ? (reqUploadResult.retestRequiredElements||"") : "";
    const approvalStatus=uploadParsed
      ? (reqUploadResult.approvalStatus || deriveInitialApprovalStatus({...measurementBase,xrfResult,xrfRemeasureRequired:retestRequired,precisionRequired}))
      : "보류";
    const isFacility=normalizeItemType(reqForm.type)==="Facility";
    const initialRisk=uploadParsed ? evaluateInitialRisk(measurementBase,crLevel,isFacility) : null;
    const finalRisk=initialRisk?.finalRisk || "—";
    const cycle=initialRisk?.cycle || "Not Available";
    const cycleMonths=initialRisk?.cycleMonths || null;
    const measuredDate=uploadParsed ? (reqUploadResult.measuredDate||TODAY_STR) : null;
    const measurementId=uploadParsed ? `${code}_${measuredDate.replaceAll("-","")}_01` : null;
    const initialHistory=uploadParsed ? [{
      date:measuredDate,
      id:measurementId,
      periodNo:0,
      label:"R",
      xrf_worst:xrfWorst,
      xrfResult,
      xrfLevel,
      finalRisk,
      retestRequired,
      retestRequiredElements,
      exceededElements,
      precisionRequired,
      precisionRequiredElements:precisionElements,
      dataMissingElements:reqUploadResult.dataMissingElements||[],
      reviewRequiredElements:reqUploadResult.reviewRequiredElements||[],
      elements:reqUploadResult.elements||{},
      approvalStatus,
      sourceFileName:reqUploadResult.fileName||"",
      sourceSheet:reqUploadResult.sourceSheet||"",
      status:"reference"
    }] : [];
    const initialPeriods=uploadParsed ? [{
      num:0,
      label:"R",
      displayLabel:"R",
      isReference:true,
      referenceDate:measuredDate,
      date:measuredDate,
      start:null,
      due:null,
      targetDate:null,
      windowStart:null,
      measured:measuredDate,
      status:"reference",
      statusLabel:"등록완료",
      daysOver:0,
      measurementId,
      closeType:"registered",
      finalRisk,
      xrfLevel,
      cycle,
      cycleMonths,
      riskAssessment:true,
      riskAssessmentType:"Initial",
      displaySeq:0,
      isDisplayTarget:false,
      sourceBasis:`${adminDirect?"Admin":"Requester"} uploaded XRF result registered as R baseline`
    }] : [];
    const item={
      _sourceCode:code,
      code, name:reqForm.itemName||"신규 등록 요청 품목", nameEn:String(reqForm.itemNameEn||"").trim(), dept:reqForm.handlingDept, respDept:"PQE",
      requesterName:"",
      photoFileName:reqPhotoFile?.name||"",photoFileUrl:reqPhotoPreview||"",
      type:reqForm.type,
      materialCategory:reqForm.materialCategory, materialState:reqForm.materialState||"", manufacturer:reqForm.manufacturer,
      useDate:reqForm.useDate, crType:reqForm.crType, crLevel, approvalStatus,
      // 신규 Item의 Action_Note는 시스템 상태 문구로 자동 채우지 않습니다.
      // 요청 사유와 사용자 Note는 XRF_Requests의 전용 필드에 별도로 저장됩니다.
      actionNote:isReplacement?[
        isReplacement?(adminDirect?"관리자 직접 등록 / 기존 품목 대체 등록":"의뢰자 등록 · 기존 품목 대체 등록"):(adminDirect?"관리자 직접 등록 / 신규 도입":"의뢰자 등록 · 신규 도입"),
        effectiveXrfMode==="request"?"XRF 측정 의뢰 상태":"XRF 원본 업로드 완료",
        isReplacement?`대체 대상: ${selectedReplacementItem?.code||reqForm.replacementTargetCode||"미선택"} · ${selectedReplacementItem?.name||""}`:"",
        isReplacement?`대체 사유: ${replacementReasonFinal}`:"",
        isReplacement?`기존 품목 처리: ${reqForm.oldItemDisposition}`:"",
        reqForm.requestReason?`요청 사유: ${reqForm.requestReason}`:"",
        reqForm.note?`비고: ${reqForm.note}`:""
      ].filter(Boolean).join("\n"):"",
      // 신규 대체품도 새로 등록되는 품목이므로 분류는 "신규 등록"입니다.
      // 기존 대체 대상 품목만 "변경·대체" 및 ReplacedOld로 전환합니다.
      originCategory:"new",
      originLifecycle:isReplacement?"NewReplacement":"NewItem",
      // 현재 Pn에서 요청/등록된 품목은 해당 Pn 동안 신규로 유지되고 다음 Pn부터 기존으로 전환됩니다.
      registrationPeriodNo:currentPortfolioPeriodNo || 1,
      newCategoryUntil:null,
      category:"new",
      categoryLabel:"신규 등록",
      lifecycle:isReplacement?"NewReplacement":"NewItem",
      isCurrent:true,
      requestType,
      replacementOf:isReplacement?(selectedReplacementItem?.code||reqForm.replacementTargetCode||null):null,
      replacementReason:isReplacement?replacementReasonFinal:"",
      reasonCategory:isReplacement?reqForm.replacementReason:"신규 도입",
      reasonOtherDetail:isReplacement&&reqForm.replacementReason==="기타"?replacementReasonFinal:"",
      reasonDetail:reqForm.requestReason||"",
      reasonFinal:[isReplacement?replacementReasonFinal:reqForm.requestReason,reqForm.note].filter(Boolean).join(" / "),
      oldItemDisposition:isReplacement?reqForm.oldItemDisposition:"",
      replacedBy:null,
      replacementDate:isReplacement?reqForm.replacementDate:null,
      firstDate:measuredDate||TODAY_STR, lastMeasured:uploadParsed?measuredDate:null,
      latestMeasurementId:measurementId, firstMeasurementId:measurementId, measurementCount:uploadParsed?1:0,
      latestXrfResult:xrfResult, latestXrfLevel:xrfLevel,
      xrfWorst, finalRisk, cycle,
      cycleMonths, nextDue:cycleMonths?addMonthsISO(measuredDate||TODAY_STR,cycleMonths):null, dDay:null,
      complianceStatus:approvalStatus==="승인"?"ok":"notAvailable",
      riskBasisMeasurementId:measurementId,
      riskBasisXrfLevel:xrfLevel,
      riskBasisCrLevel:crLevel,
      riskAssessmentType:uploadParsed?"Initial":null,
      riskAssessmentVersion:uploadParsed?1:null,
      riskAssessedAt:uploadParsed?measuredDate:null,
      precisionRequired, retestRequired, retestRequiredElements,
      periods:initialPeriods, history:initialHistory, latest:uploadParsed?{id:measurementId,date:measuredDate,sequence:1,role:"Initial",xrfWorst,xrfResult,xrfLevel,finalRisk,approvalStatus,elements:reqUploadResult.elements||{},exceededElements,precisionRequired,precisionRequiredElements:precisionElements,precisionStatus:precisionRequired?"Required":"Not Required",xrfRemeasureRequired:retestRequired,retestRequired,retestRequiredElements,dataMissingElements:reqUploadResult.dataMissingElements||[],reviewRequiredElements:reqUploadResult.reviewRequiredElements||[],sourceFileName:reqUploadResult.fileName||"",sourceSheet:reqUploadResult.sourceSheet||""}:null
    };
    setRequestItems(prev=>[item,...prev.filter(x=>x.code!==code)]);
    setApprovalOverrides(prev=>({...prev,[code]:approvalStatus}));

    if(isReplacement && selectedReplacementItem){
      const selectedStateKey=itemStateKey(selectedReplacementItem);
      const existingReplacedBy=String(selectedReplacementItem.replacedBy||"")
        .split(",").map(v=>v.trim()).filter(Boolean);
      const replacedBy=Array.from(new Set([...existingReplacedBy,code])).join(",");
      const replacementTerminalPatch={
        category:"changed",categoryLabel:"변경·대체",lifecycle:"ReplacedOld",isCurrent:false,replacedBy,
        replacementDate:reqForm.replacementDate||TODAY_STR,replacementReason:replacementReasonFinal,
        oldItemDisposition:reqForm.oldItemDisposition,nextDue:null,dDay:null,complianceStatus:"notAvailable"
      };
      setItemOverrides(prev=>({
        ...prev,
        [selectedStateKey]:{...(prev[selectedStateKey]||{}),...replacementTerminalPatch,_sourceCode:selectedStateKey}
      }));
      setRequestItems(prev=>prev.map(item=>itemStateKey(item)===selectedStateKey?{...item,...replacementTerminalPatch}:item));
    }

    const created=await persistSharePointAction("createRequest",{
      requestType:isReplacement?"Replacement":"New",
      status:uploadParsed?requestStatusFromApproval(approvalStatus):"Pending",
      targetItemId:selectedReplacementItem?.itemId||null,
      targetSpItemId:selectedReplacementItem?._spItemId||null,
      targetPartNumber:replacementTargetCode||null,
      partNumber:code,
      requestedDate:TODAY_STR,
      actor:"admin",
      xrfInputMode:effectiveXrfMode,
      form:{...reqForm,requesterName:"",partNumber:code},
      item:{
        name:item.name,nameEn:item.nameEn,type:item.type,dept:item.dept,materialCategory:item.materialCategory,
        materialState:item.materialState,crType:item.crType,lifecycle:item.lifecycle,isCurrent:item.isCurrent,
        registrationPeriodNo:item.registrationPeriodNo,actionNote:item.actionNote,replacementOf:item.replacementOf,replacementDate:item.replacementDate
      },
      initialMeasurement:uploadParsed?{
        measuredDate,sourceFileName:reqUploadResult.fileName||"",sourceSheet:reqUploadResult.sourceSheet||"",
        elements:reqUploadResult.elements||{},requestStatus:requestStatusFromApproval(approvalStatus)
      }:null
    });
    if(!created) return;
    let uploadedRemoteFile=false;
    if(reqPhotoFile && created.itemId){
      try{
        await uploadItemPhotoBinary(created.itemId,reqPhotoFile,setReqPhotoProgress,created.itemSpItemId);
        uploadedRemoteFile=true;
      }catch(error){
        window.alert(`품목은 저장됐지만 사진 원본 저장에 실패했습니다.\n${error?.message||error}`);
      }
    }
    if(uploadParsed && reqUploadFileRef.current && created.measurement?.measurementId){
      try{
        await uploadPrecisionBinary(created.measurement.measurementId,"source",reqUploadFileRef.current);
        uploadedRemoteFile=true;
      }catch(error){
        window.alert(`품목과 XRF 측정값은 저장됐지만 원본 파일 저장에 실패했습니다.\n${error?.message||error}`);
      }
    }
    if(uploadedRemoteFile) await reloadSharePointDb();
    // 업로드가 완료된 품목은 즉시 R 측정과 원소 분석을 확인할 수 있도록 XRF 분석 화면으로 이동합니다.
    // 측정 의뢰만 등록한 경우에는 리스트에 반영하고 관리자 업로드를 기다립니다.
    setSelKey(uploadParsed?code:null);
    setSelPeriodNum(uploadParsed?0:null);
    setTab(uploadParsed?"xrf":"list");
    setRegStep(1);
    setRegType(null);
    setReqXrfMode("request");
    reqUploadFileRef.current=null;
    setReqUploadFileName("");
    setReqUploadResult(emptyUploadedXrfResult());
    clearRequestPhoto();
  },[buildReplacementPartNumber,buildRequesterPartNumber,reqForm,reqXrfMode,reqUploadResult,reqPhotoFile,reqPhotoPreview,regType,selectableExistingItems,currentPortfolioPeriodNo,persistSharePointAction,reloadSharePointDb,clearRequestPhoto]);

  useEffect(()=>{
    const target=allItems.find(i=>i.code===reqForm.replacementTargetCode);
    const isReplacementMode=regType==="replacement" || regType==="admin-replace";
    const autoCode=isReplacementMode && target
      ? buildReplacementPartNumber(target.code)
      : buildRequesterPartNumber(reqForm.handlingDept);
    setReqForm(prev=>prev.partNumber===autoCode ? prev : {...prev,partNumber:autoCode});
  },[allItems,buildReplacementPartNumber,buildRequesterPartNumber,regType,reqForm.handlingDept,reqForm.replacementTargetCode]);

  useEffect(()=>{
    const h=(e)=>{
      if(!e.target.closest("[data-filter-panel]")&&!e.target.closest("[data-filter-header]"))
        setOpenFilter(null);
    };
    document.addEventListener("mousedown",h);
    return()=>document.removeEventListener("mousedown",h);
  },[openFilter]);

  const stats=useMemo(()=>{
    const c={existing:0,new:0,changed:0,overdue:0,late:0,dueSoon:0,active:0,history:0,replaced:0,retestRequired:0,precisionRequired:0,xrfRemeasureRequired:0,ng:0,measuring:0,pendingApproval:0,total:0,rawTotal:realItems.length};
    realItems.forEach(i=>{
      const isHistory=isHistoryRecordItem(i);
      const isOperational=isOperationalManagedItem(i);
      if(isHistory) c.history++;
      if(isOperational) c.total++;
      if(isOperational) c[i.category]=(c[i.category]||0)+1;
      if(isOperational) c.active++;
      if(i.lifecycle==="ReplacedOld") c.replaced++;
      if(isOperational && itemRetestRequired(i)) { c.retestRequired++; c.xrfRemeasureRequired++; }
      if(isOperational && itemHasPrecisionTrigger(i)) c.precisionRequired++;
      if(isOperational && itemXrfWorst(i)==="NG") c.ng++;
      const approvalKey=approvalFilterKey(approvalStatusOf(i));
      if(isOperational && approvalKey==="measuring") c.measuring++;
      if(isOperational && approvalKey!=="approved") c.pendingApproval++;
      const cs=getCS(i);
      if(isOperational && cs==="overdue") c.overdue++;
      if(isOperational && hasLatePeriod(i)) c.late++;
      if(isOperational && cs==="dueSoon") c.dueSoon++;
    });
    return c;
  },[realItems,approvalStatusOf]);

  // 대시보드 차트 데이터 — 전부 실제 품목에서 집계 (추정/더미값 없음)
  const dash=useMemo(()=>{
    const ops=realItems.filter(i=>isOperationalManagedItem(i));

    // ① 향후 6개월 측정 도래 예정 + 기한 초과
    const buckets=[];
    const base=new Date(TODAY.getFullYear(),TODAY.getMonth(),1);
    for(let k=0;k<6;k++){
      const d=new Date(base.getFullYear(),base.getMonth()+k,1);
      buckets.push({key:`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`,
        label:`${d.getMonth()+1}월`,total:0,overdue:0});
    }
    ops.forEach(i=>{
      const due=nextDueOfItem(i); if(!due) return;
      const b=buckets.find(x=>due.slice(0,7)===x.key); if(!b) return;
      b.total++;
      if(getCS(i)==="overdue") b.overdue++;
    });

    // ② 이행 상태 구성
    const csMeta=[
      {key:"overdue",  label:"미이행",   color:A.rose.solid},
      {key:"dueSoon",  label:"측정권장", color:A.amber.solid},
      {key:"ok",       label:"이행",     color:A.green.solid},
      {key:"upcoming", label:"예정",     color:A.blue.ln},
      {key:"notAvailable",label:"대상 아님",color:A.slate.ln},
    ];
    const csCount={};
    csMeta.forEach(m=>csCount[m.key]=0);
    ops.forEach(i=>{const k=getCS(i); if(csCount[k]!==undefined) csCount[k]++;});
    const segments=csMeta.map(m=>({...m,n:csCount[m.key]}));

    // ③ 부서별 Final Risk 구성
    const riskMeta=[
      {key:"NA",label:"사용불허",color:A.purple.solid},
      {key:"H", label:"H",       color:A.rose.solid},
      {key:"M", label:"M",       color:A.amber.solid},
      {key:"L", label:"L",       color:A.green.solid},
    ];
    const deptRows=DEPTS.map(d=>({dept:d,total:0,NA:0,H:0,M:0,L:0}));
    ops.forEach(i=>{
      const row=deptRows.find(r=>r.dept===i.dept); if(!row) return;
      const risk=itemFinalRisk(i);
      const k=risk==="Not Allowed"?"NA":(["H","M","L"].includes(risk)?risk:null);
      if(!k) return;
      row[k]++; row.total++;
    });

    return {buckets,segments,csTotal:ops.length,riskMeta,deptRows:deptRows.filter(r=>r.total>0)};
  },[realItems]);

  const listItems=useMemo(
    ()=>allItems.filter(item=>!isCancelledListItem(item,approvalStatusOf(item))),
    [allItems,approvalStatusOf]
  );

  const listFilterValues=useMemo(()=>{
    const uniqueSorted=values=>[...new Set(values.map(value=>String(value||"—")).filter(Boolean))]
      .sort((a,b)=>a.localeCompare(b,"ko",{numeric:true,sensitivity:"base"}));
    return {
      ...UNIQUE_VALS,
      photo:uniqueSorted(listItems.map(item=>listColumnValue(item,"photo",approvalStatusOf(item)))),
      code:uniqueSorted(listItems.map(item=>listColumnValue(item,"code",approvalStatusOf(item)))),
      name:uniqueSorted(listItems.map(item=>listColumnValue(item,"name",approvalStatusOf(item)))),
      dept:uniqueSorted(listItems.map(item=>listColumnValue(item,"dept",approvalStatusOf(item)))),
      cycle:uniqueSorted(listItems.map(item=>listColumnValue(item,"cycle",approvalStatusOf(item)))),
      firstDate:uniqueSorted(listItems.map(item=>listColumnValue(item,"firstDate",approvalStatusOf(item)))),
      precision:uniqueSorted(listItems.map(item=>listColumnValue(item,"precision",approvalStatusOf(item)))),
      nextDue:uniqueSorted(listItems.map(item=>listColumnValue(item,"nextDue",approvalStatusOf(item)))),
      dDay:uniqueSorted(listItems.map(item=>listColumnValue(item,"dDay",approvalStatusOf(item)))),
    };
  },[listItems,approvalStatusOf]);

  const filtered=useMemo(()=>listItems.filter(item=>{
    if(VISUAL_UAT_MODE && visualUatOnly && !item.__visualUat) return false;
    const approvalStatus=approvalStatusOf(item);
    for(const column of ["category","photo","code","name","dept","cycle","firstDate","compliance","xrf","precision","approval","nextDue","dDay"]){
      if(!itemMatchesListColumnFilter(item,column,fs[column],approvalStatus)) return false;
    }
    if(fs.crLevel.length>0&&!fs.crLevel.includes(item.crLevel)) return false;
    if(fs.risk.length>0&&!fs.risk.includes(itemFinalRisk(item))) return false;
    if(fs.retest.length>0&&!fs.retest.includes(itemRetestRequired(item)?"review":"trusted")) return false;
    if(fs.lifecycle.length>0&&!fs.lifecycle.includes(item.lifecycle)) return false;
    if((fs.dueMonth||[]).length>0){
      const due=nextDueOfItem(item);
      const dueMonth=due ? String(due).slice(0,7) : "";
      if(!fs.dueMonth.includes(dueMonth)) return false;
    }
    if(!itemMatchesListSearch(item,search,approvalStatus)) return false;
    return true;
  }).map((item,idx)=>({item,idx})).sort((a,b)=>{
    const closedRank=item=>{
      const approvalKey=approvalFilterKey(approvalStatusOf(item));
      const processKey=String(item.processStatus||"").trim().toLowerCase();
      const lifecycleKey=String(item.lifecycle||"").trim();
      if(approvalKey==="cancelled" || processKey==="reverted" || lifecycleKey==="Cancelled") return 2;
      if(isDiscontinueRequestItem(item) && (approvalKey==="approved" || approvalKey==="rejected" || processKey==="applied" || processKey==="rejected")) return 1;
      return 0;
    };
    const ra=closedRank(a.item), rb=closedRank(b.item);
    if(ra!==rb) return ra-rb;
    return a.idx-b.idx;
  }).map(({item})=>item),[listItems,fs,search,visualUatOnly,approvalStatusOf]);

  const listTotalPages=Math.max(1,Math.ceil(filtered.length/MATERIAL_LIST_PAGE_SIZE));
  useEffect(()=>setListPage(0),[fs,search,visualUatOnly]);
  useEffect(()=>setListPage(page=>Math.min(page,listTotalPages-1)),[listTotalPages]);
  const pagedFiltered=useMemo(()=>{
    const start=listPage*MATERIAL_LIST_PAGE_SIZE;
    return filtered.slice(start,start+MATERIAL_LIST_PAGE_SIZE);
  },[filtered,listPage]);

  const sel=allItems.find(i=>i.code===selKey);
  const selTimelinePeriods=useMemo(()=>itemTimelinePeriods(sel),[sel]);
  const selPeriodPageMeta=useMemo(
    ()=>xrfPeriodPageMeta(selTimelinePeriods,selPeriodNum,selPeriodPage),
    [selTimelinePeriods,selPeriodNum,selPeriodPage]
  );
  // XRF 분석 탭의 이행 상세와 주기별 산정 요약도 타임라인 현재 페이지의 최대 12개만 사용합니다.
  const selPeriods=selPeriodPageMeta.visiblePeriods;
  // XRF 분석은 주기 계산용 itemAllPeriods가 아니라 실제 측정 누락까지 보완한 표시용 기간을 사용합니다.
  // 따라서 위험도 현황의 최신 XRF와 XRF 분석 탭에서 선택 가능한 최신 측정이 항상 같은 Measurement ID를 가리킵니다.
  const selAllPeriods=useMemo(()=>itemXrfDisplayPeriods(sel),[sel]);

  const currentMeasurement=useMemo(()=>{
    if(!sel) return null;
    if(selMeasurementId){
      const byId=measurementById(sel,selMeasurementId);
      if(byId) return byId;
    }
    if(selPeriodNum !== null && selPeriodNum !== undefined){
      const p=selAllPeriods.find(pp=>Number(pp.num)===Number(selPeriodNum));
      return xrfTrendMeasurement(sel,p);
    }
    return latestMeasurementOf(sel) || null;
  },[sel,selPeriodNum,selAllPeriods,selMeasurementId]);

  const currentPeriodMeta=useMemo(()=>{
    if(!sel) return null;
    if(selPeriodNum === null || selPeriodNum === undefined) {
      const latestId=String(sel.latestMeasurementId||latestMeasurementOf(sel)?.id||"");
      const latestP=selAllPeriods.slice().reverse().find(p=>
        periodMeasurementIds(p).some(id=>String(id)===latestId)
        || measurementAttemptsForPeriod(sel,p).some(row=>String(row?.id||row?.measurementId||"")===latestId)
      );
      return latestP;
    }
    return selAllPeriods.find(p=>Number(p.num)===Number(selPeriodNum));
  },[sel,selPeriodNum,selAllPeriods]);
  const selectedTrendPeriods=useMemo(()=>{
    if(!currentMeasurement) return [];
    return xrfTrendPeriodsThroughSelection(selAllPeriods,currentPeriodMeta);
  },[currentMeasurement,currentPeriodMeta,selAllPeriods]);
  const selectedPeriodLabel=useMemo(()=>currentPeriodMeta?periodLabel(currentPeriodMeta):"선택 주기",[currentPeriodMeta]);
  const selectedMetrics=useMemo(()=>{
    if(!sel) return {xrf:"—", xrfLevel:null, judgement:"—", final:"—", measuredDate:"—", precisionRequired:false};
    if(currentMeasurement){
      return {
        xrf:displayXrfText(measurementXrfResult(currentMeasurement) || currentMeasurement.xrfResult || currentMeasurement.xrfWorst || "—"),
        xrfLevel:measurementXrfLevel(currentMeasurement),
        judgement:measurementJudgementSummary(currentMeasurement),
        final:itemFinalRisk(sel),
        measuredDate:currentMeasurement.date || "—",
        precisionRequired:precisionRequiredElements(currentMeasurement).length>0
      };
    }
    if(selPeriodNum !== null && selPeriodNum !== undefined) return {xrf:"---", xrfLevel:null, judgement:"---", final:itemFinalRisk(sel), measuredDate:"—", precisionRequired:false};
    const latest=latestMeasurementOf(sel);
    return {
      xrf:displayXrfText(itemXrfResult(sel)),
      xrfLevel:itemLatestXrfLevel(sel),
      judgement:measurementJudgementSummary(latest),
      final:itemFinalRisk(sel),
      measuredDate:latest?.date || sel.lastMeasured || "—",
      precisionRequired:latest ? precisionRequiredElements(latest).length>0 : false
    };
  },[sel,currentMeasurement,selPeriodNum]);

  const periodUploadTarget=useMemo(()=>{
    if(!sel || isDiscontinueRequestItem(sel) || isEquipmentItem(sel)) return null;
    const selected=selPeriodNum!==null && selPeriodNum!==undefined
      ? selAllPeriods.find(p=>Number(p?.num)===Number(selPeriodNum))
      : null;
    const pendingRemeasurePeriod=selAllPeriods.find(period=>{
      if(period?.isSupplementalMeasurement) return false;
      const attempts=measurementAttemptsForPeriod(sel,period);
      const latestAttempt=attempts[attempts.length-1]||null;
      const measurement=latestAttempt?measurementById(sel,latestAttempt.id||latestAttempt.measurementId):null;
      return pendingXrfRemeasureElements(measurement).length>0;
    }) || null;
    // 선택한 R/Pn에 1차 ??가 있으면 다른 주기보다 해당 측정의 Retest를 우선합니다.
    if(selected && pendingRemeasurePeriod && Number(selected.num)===Number(pendingRemeasurePeriod.num)) return selected;
    if(pendingRemeasurePeriod) return pendingRemeasurePeriod;
    if(selected && !selected?.isSupplementalMeasurement && !(selected?.isReference || Number(selected?.num)===0)) return selected;
    return currentCompliancePeriod(sel)
      || selAllPeriods.find(p=>!(p?.isReference||Number(p?.num)===0)&&!p?.measurementId&&!p?.measured)
      || null;
  },[sel,selAllPeriods,selPeriodNum]);
  const selectedPeriodAttempts=useMemo(()=>{
    if(!sel) return [];
    const periodNo=Number(periodUploadTarget?.num ?? selPeriodNum ?? currentPeriodMeta?.num ?? 0);
    const wf=workflowForItem(sel);
    if(wf.caseData && Number(wf.caseData.periodNo)===periodNo && wf.xrfAttempts.length){
      return wf.xrfAttempts.map(a=>({id:a.id,date:a.measuredDate,role:a.role,attemptNo:a.attemptNo,result:a.normalizedResult==="REMEASURE"?"재측정 필요":a.normalizedResult,reportJudgement:a.reportJudgement,mock:true}));
    }
    const resolved=periodUploadTarget || selAllPeriods.find(p=>Number(p?.num)===periodNo) || null;
    return measurementAttemptsForPeriod(sel,resolved).map((h,idx)=>{
      const id=h.id||h.measurementId;
      const linked=measurementById(sel,id);
      return {id,date:h.date,role:normalizeMeasurementRole(h.measurementRole||h.role||"Periodic"),attemptNo:Number(h.attemptNo||idx+1),result:linked?measurementXrfResult(linked):"—",reportJudgement:linked?measurementJudgementSummary(linked):(h.reportJudgement||"")};
    }).sort((a,b)=>String(a.date||"").localeCompare(String(b.date||"")) || Number(a.attemptNo)-Number(b.attemptNo));
  },[sel,periodUploadTarget,selPeriodNum,currentPeriodMeta,selAllPeriods]);
  const latestSelectedAttempt=selectedPeriodAttempts[selectedPeriodAttempts.length-1]||null;
  const latestSelectedMeasurement=latestSelectedAttempt && sel
    ? measurementById(sel,latestSelectedAttempt.id||latestSelectedAttempt.measurementId)
    : null;
  const selectedPendingRemeasure=pendingXrfRemeasureElements(latestSelectedMeasurement);
  const latestIsSecondRemeasure=latestSelectedAttempt?.result==="재측정 필요"
    && (String(latestSelectedAttempt?.role||"")==="Retest" || Number(latestSelectedAttempt?.attemptNo||1)>=2);
  const selectedIsAnnual=!!sel && cycleMonthsOf(sel)===12;
  const uploadTargetIsReference=!!periodUploadTarget && (periodUploadTarget?.isReference || Number(periodUploadTarget?.num)===0);
  // 연 1회는 같은 연도·동일 Pn의 별도 측정을 계속 추가할 수 있습니다.
  // 월/반기는 기존 규칙을 유지해 미측정 또는 1차 ??의 재측정일 때만 추가 업로드가 가능합니다.
  const periodUploadAllowed=!!periodUploadTarget && !latestIsSecondRemeasure
    && (selectedPendingRemeasure.length>0
      || (!uploadTargetIsReference && (selectedIsAnnual || selectedPeriodAttempts.length===0)));
  const handleSelectPeriod=useCallback((num)=>{ setSelPeriodNum(num); setSelMeasurementId(null); },[]);
  const handleSelectPeriodMeasurement=useCallback((num,measurementId)=>{ setSelPeriodNum(num); setSelMeasurementId(measurementId||null); },[]);

  const selectedApprovalKey = sel ? approvalFilterKey(approvalStatusOf(sel)) : "";
  const selectedIsDiscontinue = sel ? isDiscontinueRequestItem(sel) : false;
  // XRF 결과가 아직 없는 의뢰품은 Approval_Status가 보류여도 관리자 업로드 화면을 유지합니다.
  const selectedIsMeasuring = !selectedIsDiscontinue && !latestMeasurementOf(sel)
    && workflowForItem(sel).stage==="XRF_REQUIRED";

  const pendingMeasurementItems=useMemo(
    ()=>allItems.filter(item=>
      ["XRF_REQUIRED","XRF_REMEASURE_REQUIRED"].includes(workflowForItem(item).stage)
      && !isDiscontinueRequestItem(item)
      && String(item.processStatus||"").trim().toLowerCase()!=="reverted"
      && String(item.lifecycle||"").trim()!=="Cancelled"
    ),
    [allItems]
  );
  const discontinueRequestItems=useMemo(
    ()=>allItems.filter(item=>isDiscontinueRequestItem(item)),
    [allItems]
  );
  const pendingDiscontinueRequestItems=useMemo(
    ()=>discontinueRequestItems.filter(item=>{
      const approvalKey=approvalFilterKey(approvalStatusOf(item));
      const processKey=String(item.processStatus||"").trim().toLowerCase();
      const lifecycleKey=String(item.lifecycle||"").trim();
      const target=allItems.find(candidate=>candidate.code===item.discontinueTargetCode);
      const hasProcessedTrace=!!(
        item.isCurrent===false
        || item.processedAt
        || item.approvalCompletedAt
        || item.revertedAt
        || item.beforeTargetSnapshot
        || item.restoredTargetSnapshot
      );
      const targetAlreadyProcessed=!!target && (
        target.discontinuedByRequestCode===item.code
        || target.restoredFromDiscontinueRequestCode===item.code
        || target.lifecycle==="Discontinued"
      );
      return approvalKey==="hold"
        && !hasProcessedTrace
        && !targetAlreadyProcessed
        && processKey!=="applied"
        && processKey!=="rejected"
        && processKey!=="reverted"
        && lifecycleKey!=="Cancelled";
    }),
    [allItems,approvalStatusOf,discontinueRequestItems]
  );
  // 처리 완료된 단종 요청 이력은 대상 품목 상세에 표시하므로 XRF 품목 선택 목록에 중복 노출하지 않습니다.
  // 관리자 결정이 필요한 미처리 요청만 요청 행 자체를 선택할 수 있도록 유지합니다.
  const xrfSelectableItems=useMemo(()=>{
    const pendingCodes=new Set(pendingDiscontinueRequestItems.map(item=>item.code));
    return allItems.filter(item=>!isDiscontinueRequestItem(item) || pendingCodes.has(item.code));
  },[allItems,pendingDiscontinueRequestItems]);
  const canUndoDiscontinueRequest=useCallback((item)=>{
    if(!item || !hasDiscontinueRequestMetadata(item)) return false;
    const approvalKey=approvalFilterKey(approvalStatusOf(item));
    const processKey=String(item.processStatus||"").trim().toLowerCase();
    const lifecycleKey=String(item.lifecycle||"").trim();
    if(approvalKey==="cancelled" || approvalKey==="rejected") return false;
    if(processKey==="reverted" || lifecycleKey==="Cancelled") return false;
    return approvalKey==="approved" || processKey==="applied" || lifecycleKey==="Discontinued";
  },[approvalStatusOf]);
  const selectedDiscontinueRequestForTarget=useMemo(()=>{
    if(!sel || isDiscontinueRequestItem(sel)) return null;
    if(hasDiscontinueRequestMetadata(sel) && canUndoDiscontinueRequest(sel)) return sel;
    return discontinueRequestItems.find(item=>
      item.discontinueTargetCode===sel.code
      && canUndoDiscontinueRequest(item)
    ) || null;
  },[canUndoDiscontinueRequest,discontinueRequestItems,sel]);

  const handleAdminXrfUpload=useCallback(async(e)=>{
    const file=e.target.files?.[0];
    if(!file || !sel || !isAdminUser || measurementSavePendingRef.current) return;
    // 신규 품목은 XRF 결과가 없어 Approval이 보류여도 XRF_REQUIRED 단계에서 최초 측정을 업로드할 수 있습니다.
    if(!selectedIsMeasuring || isDiscontinueRequestItem(sel)) return;
    setAdminUploadFileNameByCode(prev=>({...prev,[sel.code]:file.name}));
    setAdminUploadResultByCode(prev=>({...prev,[sel.code]:emptyUploadedXrfResult(file.name,{pending:true})}));
    const result=await parseUploadedXrfFile(file);
    setAdminUploadResultByCode(prev=>({...prev,[sel.code]:result}));
    if(!result.parsed){
      e.target.value="";
      return;
    }
    if(result.partNumberCandidate && !samePartNumber(result.partNumberCandidate,sel.code)){
      const mismatch={...result,parsed:false,error:`파일 품번 ${result.partNumberCandidate}과 선택 품번 ${sel.code}이 일치하지 않습니다.`};
      setAdminUploadResultByCode(prev=>({...prev,[sel.code]:mismatch}));
      e.target.value="";
      return;
    }
    const committed=await commitParsedMeasurement(sel,result,null,"관리자 최초 XRF 업로드",file);
    if(committed){
      setSelPeriodNum(committed.periodNo);
      setSelMeasurementId(committed.measurementId);
      setTab("xrf");
    }else{
      setAdminUploadResultByCode(prev=>({...prev,[sel.code]:{...result,error:"SharePoint DB 저장에 실패했습니다. 측정 결과는 DB에 반영되지 않았습니다."}}));
    }
    e.target.value="";
  },[commitParsedMeasurement,isAdminUser,sel,selectedIsMeasuring]);

  const handlePeriodXrfUpload=useCallback(async(e)=>{
    const file=e.target.files?.[0];
    if(!file || !sel || !periodUploadTarget || !isAdminUser || !periodUploadAllowed || measurementSavePendingRef.current) return;
    const key=`${sel.code}|${periodUploadTarget.num}`;
    setPeriodUploadState({key,fileName:file.name,error:"",saved:false});
    const result=await parseUploadedXrfFile(file);
    if(!result.parsed){
      setPeriodUploadState({key,fileName:file.name,error:result.error||"XRF 보고서를 파싱하지 못했습니다.",saved:false});
      e.target.value="";
      return;
    }
    if(result.partNumberCandidate && !samePartNumber(result.partNumberCandidate,sel.code)){
      setPeriodUploadState({key,fileName:file.name,error:`파일 품번 ${result.partNumberCandidate}과 선택 품번 ${sel.code}이 일치하지 않습니다.`,saved:false});
      e.target.value="";
      return;
    }
    const committed=await commitParsedMeasurement(sel,result,periodUploadTarget,`관리자 주기 ${periodLabel(periodUploadTarget)} XRF 업로드`,file);
    if(committed){
      setPeriodUploadState({key,fileName:file.name,error:"",saved:true});
      setSelPeriodNum(committed.periodNo);
      setSelMeasurementId(committed.measurementId);
      setTab("xrf");
    }else{
      setPeriodUploadState({key,fileName:file.name,error:"SharePoint DB 저장에 실패했습니다. 측정 결과는 DB에 반영되지 않았습니다.",saved:false});
    }
    e.target.value="";
  },[commitParsedMeasurement,isAdminUser,periodUploadAllowed,periodUploadTarget,sel]);

  const handleAdminDirectXrfUpload=useCallback(async(e)=>{
    const file=e.target.files?.[0];
    if(!file || measurementSavePendingRef.current) return;
    reqUploadFileRef.current=file;
    setReqUploadFileName(file.name);
    setReqUploadResult(emptyUploadedXrfResult(file.name,{pending:true}));
    const result=await parseUploadedXrfFile(file);
    setReqUploadResult(result);

    if(regType!=="admin-period"){
      e.target.value="";
      return;
    }

    const target=allItems.find(item=>item.code===reqForm.replacementTargetCode);
    if(!target){
      window.alert("측정 추가 대상 품목을 먼저 선택하세요.");
      e.target.value="";
      return;
    }
    if(!result.parsed){
      window.alert(result.error || "XRF 원본 보고서를 파싱하지 못했습니다.");
      e.target.value="";
      return;
    }
    if(result.partNumberCandidate && !samePartNumber(result.partNumberCandidate,target.code)){
      window.alert(`파일 품번 ${result.partNumberCandidate}과 선택 품번 ${target.code}이 일치하지 않습니다.`);
      e.target.value="";
      return;
    }

    const targetPeriod=currentCompliancePeriod(target)
      || itemAllPeriods(target).find(p=>!(p?.isReference||Number(p?.num)===0)&&!p?.measurementId&&!p?.measured);
    if(!targetPeriod){
      window.alert("측정 결과를 연결할 주기를 찾지 못했습니다.");
      e.target.value="";
      return;
    }

    const committed=await commitParsedMeasurement(target,result,targetPeriod,`관리자 직접 주기 ${periodLabel(targetPeriod)} 등록`,file);
    if(committed){
      setSelKey(target.code);
      setSelPeriodNum(committed.periodNo);
      setTab("xrf");
      setRegType(null);
      reqUploadFileRef.current=null;
      setReqUploadResult(emptyUploadedXrfResult());
      setReqUploadFileName("");
    }else{
      setReqUploadResult({...result,error:"SharePoint DB 저장에 실패했습니다. 측정 결과는 DB에 반영되지 않았습니다."});
    }
    e.target.value="";
  },[allItems,commitParsedMeasurement,regType,reqForm.replacementTargetCode]);

  const handleAdminDraftSave=useCallback(()=>{
    try{
      localStorage.setItem("xrf-admin-registration-draft",JSON.stringify({regType,reqForm,reqUploadResult,savedAt:new Date().toISOString()}));
      window.alert("관리자 입력 내용과 파싱 결과를 브라우저 임시 저장소에 저장했습니다.");
    }catch(error){
      window.alert(`임시 저장 실패: ${error?.message||error}`);
    }
  },[regType,reqForm,reqUploadResult]);

  const handleAdminDirectSubmit=useCallback(async()=>{
    if(regType==="admin-period"){
      const target=allItems.find(item=>item.code===reqForm.replacementTargetCode);
      if(!target){
        window.alert("측정 추가 대상 품목을 선택하세요.");
        return;
      }
      if(!reqUploadResult?.parsed){
        window.alert("XRF 원본 보고서를 선택하고 파싱을 완료하세요.");
        return;
      }
      if(reqUploadResult.partNumberCandidate && !samePartNumber(reqUploadResult.partNumberCandidate,target.code)){
        window.alert(`파일 품번 ${reqUploadResult.partNumberCandidate}과 선택 품번 ${target.code}이 일치하지 않습니다.`);
        return;
      }
      const targetPeriod=currentCompliancePeriod(target)
        || itemAllPeriods(target).find(p=>!(p?.isReference||Number(p?.num)===0)&&!p?.measurementId&&!p?.measured);
      if(!targetPeriod){
        window.alert("측정 결과를 연결할 주기를 찾지 못했습니다.");
        return;
      }
      const committed=await commitParsedMeasurement(target,reqUploadResult,targetPeriod,`관리자 직접 주기 ${periodLabel(targetPeriod)} 등록`,reqUploadFileRef.current);
      if(committed){
        setSelKey(target.code);
        setSelPeriodNum(committed.periodNo);
        setTab("xrf");
        setRegType(null);
        reqUploadFileRef.current=null;
        setReqUploadResult(emptyUploadedXrfResult());
        setReqUploadFileName("");
      }else{
        setReqUploadResult({...reqUploadResult,error:"SharePoint DB 저장에 실패했습니다. 측정 결과는 DB에 반영되지 않았습니다."});
      }
      return;
    }
    submitRequesterNewItem();
  },[allItems,commitParsedMeasurement,regType,reqForm.replacementTargetCode,reqUploadResult,submitRequesterNewItem]);

  const openDiscontinueDecisionAuth=useCallback(()=>{
    if(!sel || !isAdminUser || !isDiscontinueRequestItem(sel)) return;
    if(sel.__readOnlyFixture){
      window.alert(VISUAL_UAT_READ_ONLY_MESSAGE);
      return;
    }
    setDiscontinueDecisionAuth({open:true,item:sel,password:"",error:"",verified:false,pending:false});
  },[isAdminUser,sel]);

  const closeDiscontinueDecisionAuth=useCallback(()=>{
    setDiscontinueDecisionAuth({open:false,item:null,password:"",error:"",verified:false,pending:false});
  },[]);

  const submitDiscontinueDecisionAuth=useCallback((event)=>{
    event?.preventDefault?.();
    if(!discontinueDecisionAuth.item) return closeDiscontinueDecisionAuth();
    if(discontinueDecisionAuth.password!==ADMIN_ITEM_EDIT_PASSWORD){
      setDiscontinueDecisionAuth(prev=>({...prev,error:"관리자 비밀번호가 올바르지 않습니다."}));
      return;
    }
    setDiscontinueDecisionAuth(prev=>({...prev,password:"",error:"",verified:true}));
  },[closeDiscontinueDecisionAuth,discontinueDecisionAuth.item,discontinueDecisionAuth.password]);

  const handleDiscontinueApproval=useCallback(async(status,candidate)=>{
    const requestItem=candidate || sel;
    const authenticatedItem=discontinueDecisionAuth.item;
    const authenticated=discontinueDecisionAuth.open
      && discontinueDecisionAuth.verified
      && authenticatedItem
      && requestItem
      && authenticatedItem.code===requestItem.code;
    if(!authenticated || !requestItem || !isAdminUser || !isDiscontinueRequestItem(requestItem)) return;
    if(requestItem.__readOnlyFixture){
      window.alert(VISUAL_UAT_READ_ONLY_MESSAGE);
      return;
    }
    const targetCode=requestItem.discontinueTargetCode;
    const targetBefore=allItems.find(item=>item.code===targetCode) || null;

    if(status==="승인"){
      const ok=window.confirm([
        "단종 처리를 승인하시겠습니까?",
        targetBefore ? `${targetBefore.code} · ${targetBefore.name}` : `대상 품목: ${targetCode||"미선택"}`,
        "승인 후에도 '단종 처리 원복'으로 복구할 수 있으며 처리 이력이 남습니다."
      ].join("\n"));
      if(!ok) return;
    }

    const beforeTargetSnapshot=targetBefore ? {
      category:targetBefore.category,
      categoryLabel:targetBefore.categoryLabel,
      lifecycle:targetBefore.lifecycle,
      isCurrent:targetBefore.isCurrent,
      replacedBy:targetBefore.replacedBy||null,
      replacementDate:targetBefore.replacementDate||null,
      replacementReason:targetBefore.replacementReason||"",
      oldItemDisposition:targetBefore.oldItemDisposition||"",
      nextDue:targetBefore.nextDue||null,
      dDay:targetBefore.dDay??null,
      complianceStatus:targetBefore.complianceStatus||"notAvailable",
      discontinuedByRequestCode:targetBefore.discontinuedByRequestCode||null,
      discontinueReason:targetBefore.discontinueReason||"",
      finalUseDate:targetBefore.finalUseDate||null,
      discontinueApprovalNote:targetBefore.discontinueApprovalNote||""
    } : null;

    setDiscontinueDecisionAuth(prev=>({...prev,pending:true,error:""}));
    const saved=await persistSharePointAction("processDiscontinue",{
      requestSpItemId:requestItem._spRequestItemId||null,
      targetSpItemId:targetBefore?._spItemId||null,
      status:status==="승인"?"Applied":"Rejected",
      processedAt:new Date().toISOString(),
      actor:"admin",
      finalUseDate:requestItem.finalUseDate||null,
      beforeTargetSnapshot:status==="승인"?(requestItem.beforeTargetSnapshot||beforeTargetSnapshot):null
    });
    if(!saved){
      setDiscontinueDecisionAuth(prev=>({...prev,pending:false,error:"처리 결과를 SharePoint에 저장하지 못했습니다. 다시 시도하기 전에 현재 DB 상태를 확인하세요."}));
      return;
    }

    setRequestItems(prev=>prev.map(item=>{
      if(item.code!==requestItem.code) return item;
      return {
        ...item,
        approvalStatus:status,
        approvalCompletedAt:TODAY_STR,
        approvalCompletedBy:"admin",
        processStatus:status==="승인"?"Applied":"Rejected",
        processedAt:TODAY_STR,
        processedBy:"admin",
        beforeTargetSnapshot:status==="승인" ? (item.beforeTargetSnapshot||beforeTargetSnapshot) : item.beforeTargetSnapshot||null,
        lifecycle:status==="승인" ? "Discontinued" : "Cancelled",
        isCurrent:false,
        xrfWorst:"—",
        finalRisk:"—",
        cycle:"Not Available",
        cycleMonths:null,
        nextDue:null,
        dDay:null,
        complianceStatus:"notAvailable",
        retestRequired:false,
        retestRequiredElements:"",
        periods:[],
        history:[],
        latest:null,
        actionNote:`${item.actionNote||""}\n관리자 처리: ${status} (${TODAY_STR})`
      };
    }));
    setApprovalOverrides(prev=>({...prev,[requestItem.code]:status}));
    if(status==="승인" && targetCode){
      const targetItem=allItems.find(item=>item.code===targetCode);
      const targetStateKey=itemStateKey(targetItem)||targetCode;
      const terminalPatch={
        lifecycle:"Discontinued",isCurrent:false,nextDue:null,dDay:null,complianceStatus:"notAvailable",
        discontinuedByRequestCode:requestItem.code,discontinueReason:requestItem.discontinueReason,finalUseDate:requestItem.finalUseDate,
        replacementDate:requestItem.finalUseDate||TODAY_STR,
        discontinueApprovalNote:[`관리자 단종 승인: ${TODAY_STR}`,`단종 요청 번호: ${requestItem.code}`].filter(Boolean).join("\n")
      };
      setItemOverrides(prev=>({
        ...prev,
        [targetStateKey]:{...(prev[targetStateKey]||{}),...terminalPatch,_sourceCode:targetStateKey}
      }));
      setRequestItems(prev=>prev.map(item=>itemStateKey(item)===targetStateKey?{...item,...terminalPatch}:item));
    }
    // 승인/반려 후에는 요청용 가상 행이 아니라 실제 대상 품목으로 이동합니다.
    // 반려된 경우 대상 Master는 수정되지 않았으므로 기존 상태와 모든 XRF 이력이 그대로 보입니다.
    closeDiscontinueDecisionAuth();
    if(targetCode){
      setSelKey(targetCode);
      setSelPeriodNum(null);
      setSelMeasurementId(null);
      setTab("xrf");
    }
  },[allItems,closeDiscontinueDecisionAuth,discontinueDecisionAuth,isAdminUser,sel,persistSharePointAction]);

  const handleUndoDiscontinue=useCallback((candidate)=>{
    const requestItem=hasDiscontinueRequestMetadata(candidate)
      ? candidate
      : (hasDiscontinueRequestMetadata(sel)?sel:null);
    if(!requestItem || !isAdminUser || !canUndoDiscontinueRequest(requestItem)) return;
    if(requestItem.__readOnlyFixture || candidate?.__readOnlyFixture || sel?.__readOnlyFixture){
      window.alert(VISUAL_UAT_READ_ONLY_MESSAGE);
      return;
    }
    const targetCode=requestItem.discontinueTargetCode
      || (!isDiscontinueRequestItem(requestItem)?requestItem.code:"");
    const baseTarget=allItems.find(item=>item.code===targetCode) || sharePointItems.find(item=>item.code===targetCode) || null;
    const snapshot=requestItem.beforeTargetSnapshot || (baseTarget ? {
      category:baseTarget.originCategory || baseTarget.category || "existing",
      categoryLabel:baseTarget.originCategory==="new" ? "신규 등록" : (baseTarget.categoryLabel || "기존 품목"),
      lifecycle:baseTarget.originLifecycle || (baseTarget.lifecycle==="Discontinued" ? "ExistingActive" : baseTarget.lifecycle) || "ExistingActive",
      isCurrent:true,
      replacedBy:baseTarget.replacedBy||null,
      replacementDate:baseTarget.replacementDate||null,
      replacementReason:baseTarget.replacementReason||"",
      oldItemDisposition:baseTarget.oldItemDisposition||"",
      nextDue:baseTarget.nextDue||null,
      dDay:baseTarget.dDay??null,
      complianceStatus:baseTarget.complianceStatus||getCS(baseTarget),
      discontinuedByRequestCode:null,
      discontinueReason:"",
      finalUseDate:baseTarget.finalUseDate||null,
      discontinueApprovalNote:""
    } : null);
    if(!snapshot || !targetCode){
      window.alert("원복할 대상 품목을 찾을 수 없습니다. 단종 요청의 대상 품번을 확인하세요.");
      return;
    }
    const ok=window.confirm([
      "승인된 단종 처리를 원복하시겠습니까?",
      `대상 품목: ${targetCode}`,
      "기존 XRF 측정 이력은 그대로 유지되고, 품목의 사용 상태와 이행주기 판정만 승인 전 상태로 복구됩니다."
    ].join("\n"));
    if(!ok) return;
    const reason=String(window.prompt("원복 사유를 입력하세요.", "관리자 오선택")||"").trim();
    if(!reason){
      window.alert("감사 추적을 위해 원복 사유를 입력해야 합니다.");
      return;
    }
    const restoredLifecycle=snapshot.lifecycle==="Discontinued" || snapshot.lifecycle==="Cancelled"
      ? "ExistingActive"
      : (snapshot.lifecycle||"ExistingActive");
    const restoredCategory=normalizedClassification(snapshot.category)==="changed"
      ? "existing"
      : (snapshot.category||"existing");
    const restoredState={
      ...snapshot,
      category:restoredCategory,
      categoryLabel:restoredCategory==="new" ? "신규 등록" : "기존 품목",
      lifecycle:restoredLifecycle,
      isCurrent:true,
      discontinuedByRequestCode:null,
      discontinueReason:"",
      finalUseDate:snapshot.finalUseDate||null,
      discontinueApprovalNote:"",
      restoredFromDiscontinueRequestCode:discontinueRequestCodeOf(requestItem),
      restoredAt:TODAY_STR,
      restoredBy:"admin",
      restoreReason:reason,
      restoreAuditNote:[
        `단종 처리 원복: ${TODAY_STR}`,
        `단종 요청 번호: ${requestItem.code}`,
        `원복 사유: ${reason}`
      ].join("\n"),
      discontinuationHistory:[
        ...(Array.isArray(baseTarget?.discontinuationHistory)?baseTarget.discontinuationHistory:[]),
        {
          requestCode:discontinueRequestCodeOf(requestItem),
          approvedAt:requestItem.processedAt||requestItem.approvalCompletedAt||TODAY_STR,
          revertedAt:TODAY_STR,
          revertedBy:"admin",
          revertReason:reason,
          restoredLifecycle,
          restoredCategory
        }
      ]
    };

    // 대상이 정적 마스터 품목이든 세션에서 생성한 품목이든 동일하게 복구한다.
    const restoreStateKey=itemStateKey(baseTarget)||targetCode;
    setItemOverrides(prev=>({
      ...prev,
      [restoreStateKey]:{
        ...(prev[restoreStateKey]||{}),
        ...restoredState,
        _sourceCode:restoreStateKey
      }
    }));
    setRequestItems(prev=>prev.map(item=>{
      if(item.code===targetCode) return {...item,...restoredState};
      if(item.code!==requestItem.code) return item;
      return {
        ...item,
        approvalStatus:"처리 취소",
        lifecycle:"Cancelled",
        isCurrent:false,
        processStatus:"Reverted",
        revertedAt:TODAY_STR,
        revertedBy:"admin",
        revertReason:reason,
        restoredTargetCode:targetCode,
        restoredTargetSnapshot:restoredState,
        originalDiscontinueApprovalStatus:"승인",
        nextDue:null,
        dDay:null,
        complianceStatus:"notAvailable",
        actionNote:`${item.actionNote||""}\n단종 처리 원복: ${TODAY_STR}\n원복 사유: ${reason}`
      };
    }));
    setApprovalOverrides(prev=>({...prev,[requestItem.code]:"처리 취소"}));
    void persistSharePointAction("revertDiscontinue",{
      requestSpItemId:requestItem._spRequestItemId||null,
      targetSpItemId:baseTarget?._spItemId||null,
      revertedAt:TODAY_STR,
      actor:"admin",
      reason,
      restored:restoredState
    });
    setSelKey(targetCode);
    setSelPeriodNum(null);
    setTab("list");
  },[allItems,sharePointItems,canUndoDiscontinueRequest,isAdminUser,sel,persistSharePointAction]);

  const handleToggle=useCallback((col,val)=>{
    setFs(prev=>{const cur=prev[col]||[];return{...prev,[col]:cur.includes(val)?cur.filter(v=>v!==val):[...cur,val]};});
  },[]);
  const handleClear=useCallback((col)=>setFs(prev=>({...prev,[col]:[]})),[]);
  const openAdminItemEditor=useCallback((item)=>{
    if(!item || isDiscontinueRequestItem(item)) return;
    if(item.__readOnlyFixture){
      window.alert(VISUAL_UAT_READ_ONLY_MESSAGE);
      return;
    }
    if(!isAdminUser){
      window.alert("품목 정보 편집은 관리자만 가능합니다.");
      return;
    }
    setItemEditAuth({open:true,item,password:"",error:""});
  },[isAdminUser]);

  const closeItemEditAuth=useCallback(()=>{
    setItemEditAuth({open:false,item:null,password:"",error:""});
  },[]);

  const submitItemEditAuth=useCallback((event)=>{
    event?.preventDefault?.();
    const item=itemEditAuth.item;
    if(!item) return closeItemEditAuth();
    if(itemEditAuth.password!==ADMIN_ITEM_EDIT_PASSWORD){
      setItemEditAuth(prev=>({...prev,error:"관리자 비밀번호가 올바르지 않습니다."}));
      return;
    }
    setItemEditForm({
      sourceKey:itemStateKey(item),
      originalCode:item.code,
      code:item.code||"",
      name:item.name||"",
      nameEn:item.nameEn||"",
      dept:item.dept||"",
    });
    setItemEditOpen(true);
    closeItemEditAuth();
  },[closeItemEditAuth,itemEditAuth]);

  const updateItemEditField=useCallback((field,value)=>{
    setItemEditForm(prev=>prev?{...prev,[field]:value}:prev);
  },[]);

  const saveAdminItemEditor=useCallback(()=>{
    if(!isAdminUser || !itemEditForm) return;
    const fixtureSource=allItems.find(item=>itemStateKey(item)===itemEditForm.sourceKey || item.code===itemEditForm.originalCode);
    if(fixtureSource?.__readOnlyFixture){
      window.alert(VISUAL_UAT_READ_ONLY_MESSAGE);
      return;
    }
    const oldCode=String(itemEditForm.originalCode||"").trim();
    const newCode=String(itemEditForm.code||"").trim();
    const newName=String(itemEditForm.name||"").trim();
    const newNameEn=String(itemEditForm.nameEn||"").trim();
    const newDept=String(itemEditForm.dept||"").trim();
    const sourceKey=String(itemEditForm.sourceKey||oldCode).trim();
    if(!newCode || !newName || !newDept){
      window.alert("품번, 한글 품목명, 공정명은 필수 입력입니다.");
      return;
    }
    const duplicate=allItems.some(item=>item.code!==oldCode && samePartNumber(item.code,newCode));
    if(duplicate){
      window.alert(`이미 사용 중인 품번입니다: ${newCode}`);
      return;
    }

    const replaceCsvCode=(value)=>Array.from(new Set(
      String(value||"").split(",").map(v=>v.trim()).filter(Boolean).map(v=>v===oldCode?newCode:v)
    )).join(",");
    const rewriteRefs=(item)=>{
      const patch={};
      if(item?.replacementOf===oldCode) patch.replacementOf=newCode;
      if(String(item?.replacedBy||"").split(",").map(v=>v.trim()).includes(oldCode)) patch.replacedBy=replaceCsvCode(item.replacedBy);
      if(item?.replacementTargetCode===oldCode) patch.replacementTargetCode=newCode;
      if(item?.discontinueTargetCode===oldCode) patch.discontinueTargetCode=newCode;
      if(item?.restoredTargetCode===oldCode) patch.restoredTargetCode=newCode;
      return patch;
    };
    const selectedPatch={
      code:newCode,name:newName,nameEn:newNameEn,dept:newDept,
      editedAt:TODAY_STR,editedBy:"admin",_sourceCode:sourceKey
    };

    // 세션에서 새로 생성된 품목과 해당 품목을 참조하는 요청/대체 관계를 함께 갱신합니다.
    setRequestItems(prev=>prev.map(item=>{
      const refs=rewriteRefs(item);
      if(itemStateKey(item)===sourceKey || item.code===oldCode) return {...item,...refs,...selectedPatch};
      return Object.keys(refs).length?{...item,...refs}:item;
    }));

    // 정적 마스터 품목은 원본 배열을 직접 변경하지 않고 sourceKey 기준 override에 저장합니다.
    // 동시에 다른 정적 품목의 Replacement_Of / Replaced_By 참조도 새 품번으로 연결합니다.
    setItemOverrides(prev=>{
      const next={...prev};
      sharePointItems.forEach(base=>{
        const key=itemStateKey(base);
        const current={...base,...(prev[key]||{}),_sourceCode:key};
        const refs=rewriteRefs(current);
        const isSelected=key===sourceKey || current.code===oldCode;
        if(!isSelected && !Object.keys(refs).length) return;
        next[key]={...(prev[key]||{}),...refs,...(isSelected?selectedPatch:{}),_sourceCode:key};
      });
      return next;
    });

    if(oldCode!==newCode){
      const moveKey=(setter)=>setter(prev=>{
        if(!Object.prototype.hasOwnProperty.call(prev,oldCode)) return prev;
        const next={...prev,[newCode]:prev[oldCode]};
        delete next[oldCode];
        return next;
      });
      moveKey(setApprovalOverrides);
      moveKey(setAdminUploadFileNameByCode);
      moveKey(setAdminUploadResultByCode);
      setReqForm(prev=>({
        ...prev,
        replacementTargetCode:prev.replacementTargetCode===oldCode?newCode:prev.replacementTargetCode,
        discontinueTargetCode:prev.discontinueTargetCode===oldCode?newCode:prev.discontinueTargetCode,
      }));
      if(rowSel===oldCode) setRowSel(newCode);
      setSelKey(newCode);
    }

    const persistedSource=sharePointItems.find(item=>itemStateKey(item)===sourceKey || item.code===oldCode)
      || allItems.find(item=>itemStateKey(item)===sourceKey || item.code===oldCode);
    if(persistedSource?._spItemId){
      void persistSharePointAction("updateItem",{
        itemSpItemId:persistedSource._spItemId,oldCode,newCode,name:newName,nameEn:newNameEn,dept:newDept,actor:"admin"
      });
    }else{
      window.alert("SharePoint Item ID를 찾지 못해 화면만 수정되었습니다. 새로고침 후 다시 시도하세요.");
    }
    setItemEditOpen(false);
    setItemEditForm(null);
  },[allItems,sharePointItems,isAdminUser,itemEditForm,rowSel,persistSharePointAction]);

  const handleClearAll=()=>setFs({category:[],photo:[],code:[],name:[],dept:[],cycle:[],firstDate:[],xrf:[],compliance:[],precision:[],crLevel:[],risk:[],retest:[],approval:[],lifecycle:[],nextDue:[],dDay:[],dueMonth:[]});
  const positionFilterPanel=useCallback((column)=>{
    if(!column || typeof window==="undefined") return;
    const anchor=appRootRef.current?.querySelector(`[data-filter-header="${column}"]`);
    if(!anchor) return;
    const rect=anchor.getBoundingClientRect();
    if(rect.bottom<=0 || rect.top>=window.innerHeight){
      setOpenFilter(current=>current===column?null:current);
      return;
    }
    const left=Math.max(8,Math.min(rect.left,window.innerWidth-288));
    setFPos({top:rect.bottom+4,left});
  },[]);
  const handleOpen=(col,e)=>{
    const next=openFilter===col?null:col;
    if(next){
      const rect=e.currentTarget.getBoundingClientRect();
      setFPos({top:rect.bottom+4,left:Math.max(8,Math.min(rect.left,window.innerWidth-288))});
    }
    setOpenFilter(next);
  };
  useLayoutEffect(()=>{
    if(!openFilter) return undefined;
    let frame=0;
    const syncPosition=()=>{
      window.cancelAnimationFrame(frame);
      frame=window.requestAnimationFrame(()=>positionFilterPanel(openFilter));
    };
    syncPosition();
    document.addEventListener("scroll",syncPosition,true);
    window.addEventListener("scroll",syncPosition,{passive:true});
    window.addEventListener("resize",syncPosition);
    return ()=>{
      window.cancelAnimationFrame(frame);
      document.removeEventListener("scroll",syncPosition,true);
      window.removeEventListener("scroll",syncPosition);
      window.removeEventListener("resize",syncPosition);
    };
  },[openFilter,positionFilterPanel]);
  useEffect(()=>{
    if(!openFilter) return undefined;
    const closeOnOutsidePointer=event=>{
      const target=event.target;
      if(target instanceof Element && (
        target.closest('[data-filter-panel="1"]')
        || target.closest(".material-list-table")
        || target.closest(`[data-filter-header="${openFilter}"]`)
      )) return;
      setOpenFilter(null);
    };
    document.addEventListener("pointerdown",closeOnOutsidePointer,true);
    return ()=>document.removeEventListener("pointerdown",closeOnOutsidePointer,true);
  },[openFilter]);
  const gotoXrf=(code)=>{ setSelKey(code); setSelPeriodNum(null); setSelMeasurementId(null); setTab("xrf"); };
  const gotoPrecisionDetail=(item)=>{
    if(!item) return;
    const wf=workflowForItem(item);
    const key=wf.caseData?.id || wf.activePrecisionCaseKey || wf.precisionCases?.[0]?.key || latestMeasurementOf(item)?.id || null;
    if(!key){
      gotoXrf(item.code);
      return;
    }
    setSelKey(item.code);
    setPrecisionSelCaseId(key);
    setPrecisionDetailCaseId(key);
    setTab("precision-detail");
  };
  const gotoItemDetail=(item)=>{
    if(itemHasPrecisionTrigger(item) || workflowForItem(item).precision || workflowForItem(item).precisionCases?.length) gotoPrecisionDetail(item);
    else gotoXrf(item.code);
  };
  const gotoListWithFilter=(newFs)=>{ setFs({...fs, ...newFs}); setTab("list"); };
  const gotoDueMonthList=(monthKey)=>{
    if(!monthKey) return;
    setSearch("");
    setFs({category:[],photo:[],code:[],name:[],dept:[],cycle:[],firstDate:[],xrf:[],compliance:[],precision:[],crLevel:[],risk:[],retest:[],approval:[],lifecycle:[],nextDue:[],dDay:[],dueMonth:[monthKey]});
    setTab("list");
  };
  const hasFilters=Object.values(fs).some(v=>v.length>0);

  const root={fontFamily:FONT_SANS,minHeight:"100vh",background:C.bg,color:C.text2,fontSize:13};
  const card={background:C.card,border:`1px solid ${C.bd}`,borderRadius:12,boxShadow:"0 8px 22px rgba(26,32,32,.075)"};
  const inp={padding:"8px 12px",border:`1px solid ${C.bd2}`,borderRadius:UI.rs,fontSize:12,outline:"none",background:C.card,color:C.text2,fontFamily:FONT_SANS};
  const fileInputOverlay={position:"absolute",inset:0,opacity:0,cursor:"pointer",width:"100%",height:"100%"};
  const uploadBoxStyle={position:"relative",display:"flex",alignItems:"center",justifyContent:"center",minHeight:52,border:`1px dashed ${C.bd2}`,background:C.card,color:C.text3,fontSize:12,fontWeight:600,cursor:"pointer",overflow:"hidden",textAlign:"center",padding:"0 12px",boxSizing:"border-box"};
  const TABS=[{id:"list",l:"부자재 리스트"},{id:"xrf",l:"XRF 분석"},{id:"precision",l:"정밀분석"},{id:"risk",l:"위험도 현황"},{id:"reg",l:"품목 / 변경 관리"}];
  const fhp={filterState:fs,openFilter,onOpen:handleOpen};
  // 세로 방향은 페이지 자체가 스크롤하고, 15개 열의 최소 가독 폭보다 화면이 좁을 때만
  // 표 컨테이너가 가로로 스크롤됩니다. 품목명은 축소하지 않고 셀 안에서 줄바꿈합니다.
  const listColumns=[
    {key:"cat", el:<FH label="분류" col="category" {...fhp}/>, w:80},
    {key:"photo",el:<FH label="사진" col="photo" {...fhp}/>, w:72},
    {key:"code",el:<FH label="품번" col="code" {...fhp}/>, w:86},
    {key:"name",el:<FH label="품목명" col="name" {...fhp}/>, w:180},
    {key:"dept",el:<FH label="부서" col="dept" {...fhp}/>, w:84},
    {key:"cyc",el:<FH label="주기" col="cycle" {...fhp}/>, w:62},
    {key:"first",el:<FH label="최초 등록" col="firstDate" {...fhp}/>, w:84},
    {key:"dot",el:<div style={{display:"flex",alignItems:"center",justifyContent:"center",width:"100%"}}><PlainH label="이행 현황"/></div>, w:118},
    {key:"comp",el:<FH label="현재 상태" col="compliance" {...fhp}/>, w:78},
    {key:"xrf",el:<FH label="XRF" col="xrf" {...fhp}/>, w:84},
    {key:"precision",el:<FH label="정밀분석" col="precision" {...fhp}/>, w:100},
    {key:"approval",el:<FH label="승인 상태" col="approval" {...fhp}/>, w:92},
    {key:"nd",el:<FH label="다음 마감" col="nextDue" {...fhp}/>, w:90},
    {key:"dd",el:<FH label="D-DAY" col="dDay" {...fhp}/>, w:60},
    {key:"btn",el:null, w:34},
  ];
  const listTableWeight=listColumns.reduce((sum,c)=>sum+c.w,0);
  const requesterStepLabels = regType==="discontinue"
    ? ["요청 유형 선택","단종 정보 입력","요청 완료"]
    : ["요청 유형 선택","기본정보 입력","XRF 방식 선택","요청 완료"];
  const requesterMaxStep = requesterStepLabels.length;
  const registrationButtonStyle=(variant="secondary",disabled=false)=>({
    minWidth:86,height:36,padding:"0 16px",display:"inline-flex",alignItems:"center",justifyContent:"center",
    border:`1px solid ${variant==="primary"?C.charcoalDk:C.bd2}`,borderRadius:UI.rs,
    background:disabled?C.alt:(variant==="primary"?C.charcoalDk:C.card),
    color:disabled?C.text4:(variant==="primary"?"#fff":C.text2),fontSize:12,fontWeight:650,
    cursor:disabled?"default":"pointer",opacity:disabled?.65:1,boxSizing:"border-box",boxShadow:"none"
  });
  const selectedReplacementTarget = selectableExistingItems.find(i=>i.code===reqForm.replacementTargetCode);
  const selectedDiscontinueTarget = selectableExistingItems.find(i=>i.code===reqForm.discontinueTargetCode);

  if(!entered){
    return <SplashScreen onSelectTab={(nextTab)=>{setTab(nextTab);setEntered(true);}} lang={lang} onLang={setLang}/>;
  }
  if(dbLoading){
    return <div style={{fontFamily:FONT_SANS,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,color:C.text2}}>
      <div style={{padding:"22px 26px",background:C.card,border:`1px solid ${C.bd}`,borderRadius:UI.r,boxShadow:UI.sh,fontSize:13,fontWeight:600}}>
        SharePoint DB를 불러오는 중입니다.
      </div>
    </div>;
  }
  if(dbError){
    return <div style={{fontFamily:FONT_SANS,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:C.bg,color:C.text2,padding:24,boxSizing:"border-box"}}>
      <div style={{width:"min(680px,100%)",padding:"22px 26px",background:C.card,border:`1px solid ${C.red}`,borderRadius:UI.r,boxShadow:UI.sh}}>
        <div style={{fontSize:14,fontWeight:700,color:C.redDk}}>SharePoint DB 연결 오류</div>
        <div style={{marginTop:8,fontSize:12,lineHeight:1.6,color:C.text3,wordBreak:"break-word"}}>{dbError}</div>
        <button type="button" onClick={reloadSharePointDb} style={{marginTop:14,padding:"8px 14px",border:`1px solid ${C.bd2}`,borderRadius:UI.rs,background:C.charcoalDk,color:"#fff",cursor:"pointer",fontWeight:600}}>다시 시도</button>
      </div>
    </div>;
  }

  return(
    <div ref={appRootRef} className="app-root" data-lang={lang} data-db-items={dbCounts?.items??0} data-visual-uat-mode={VISUAL_UAT_MODE?"1":"0"} style={root}>
      <style>{`
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable.css');

        /* 영문·한글·숫자를 Pretendard 한 서체로 통일합니다.
           input·select·button·table까지 상속이 끊기지 않게 명시합니다. */
        .app-root, .app-root * {
          font-family:${FONT_SANS};
          /* antialiased / grayscale 강제는 Windows Chrome에서 획을 얇게 만들어
             글자가 뿌옇게 보이는 원인이 됩니다. 브라우저 기본값을 씁니다. */
          -webkit-font-smoothing:auto;
          -moz-osx-font-smoothing:auto;
          text-rendering:optimizeSpeed;
        }
        .app-root input, .app-root select, .app-root textarea, .app-root button {
          font-family:inherit;
          font-size:inherit;
          color:inherit;
        }
        /* 의뢰 필요는 한국어에서 다른 승인 상태와 동일 규격을 유지하고,
           영문 Request Required일 때만 글자 크기는 유지한 채 세로폭을 늘립니다. */
        .app-root[data-lang="en"] .approval-request-pill {
          height:32px !important;
          padding:3px 5px !important;
          white-space:normal !important;
          line-height:1.15 !important;
        }
        /* XRF 분석 > Current Stage 전용: 영문 라벨은 현재 10px 글자 크기를 유지하고,
           고정 높이 대신 필요한 만큼 세로로 늘어나며 줄바꿈되도록 합니다. */
        .app-root[data-lang="en"] .workflow-stage-status-row > div > span {
          height:auto !important;
          min-height:${RESULT_LABEL_HEIGHT}px !important;
          padding:4px 5px !important;
          white-space:normal !important;
          line-height:1.15 !important;
          overflow:visible !important;
          text-overflow:clip !important;
          text-align:center !important;
        }
        .app-root[data-lang="en"] .workflow-stage-status-row > div:nth-child(-n+2) > span:not(.workflow-remeasure-pill) {
          width:${RESULT_LABEL_WIDTH}px !important;
          min-width:${RESULT_LABEL_WIDTH}px !important;
          max-width:${RESULT_LABEL_WIDTH}px !important;
        }
        .app-root[data-lang="en"] .workflow-stage-status-row > div:nth-child(3) > span {
          width:${APPROVAL_LABEL_WIDTH}px !important;
          min-width:${APPROVAL_LABEL_WIDTH}px !important;
          max-width:${APPROVAL_LABEL_WIDTH}px !important;
        }
        /* Remeasurement Required는 고정 폭을 해제해 배경이 영문 문구 전체를 감싸도록 합니다. */
        .app-root[data-lang="en"] .workflow-stage-status-row > div > .workflow-remeasure-pill {
          display:inline-flex !important;
          width:fit-content !important;
          min-width:max-content !important;
          max-width:none !important;
          height:${RESULT_LABEL_HEIGHT}px !important;
          min-height:${RESULT_LABEL_HEIGHT}px !important;
          padding:0 10px !important;
          white-space:nowrap !important;
          line-height:1 !important;
          overflow:visible !important;
          flex:0 0 auto !important;
        }
        .app-root[data-lang="en"] .workflow-stage-status-row > div > span > span {
          font-size:${RESULT_LABEL_FONT_SIZE}px !important;
          white-space:normal !important;
          overflow:visible !important;
          text-overflow:clip !important;
          line-height:1.15 !important;
          overflow-wrap:anywhere !important;
          word-break:normal !important;
        }
        .app-root table { font-variant-numeric:tabular-nums; }
        .responsive-table-scroll{
          width:100%;
          max-width:100%;
          overflow-x:auto !important;
          overflow-y:visible;
          -webkit-overflow-scrolling:touch;
          overscroll-behavior-x:contain;
          scrollbar-width:thin;
          box-sizing:border-box;
        }
        .responsive-table-scroll > table{
          max-width:none !important;
        }
        .app-root ::placeholder { color:${C.text4}; opacity:1; }
        .num { font-family:${FONT_NUM}; font-variant-numeric:tabular-nums; letter-spacing:-.01em; }

        /* 첨부 레퍼런스의 pill 버튼 규칙 */
        .xrf-trend-grid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(250px,1fr)); gap:10px; }
        .xrf-trend-single{ width:50%; min-width:280px; margin:0 auto; }
        @media (max-width: 560px){
          .xrf-trend-grid{ grid-template-columns:1fr; }
          .xrf-trend-single{ width:100%; min-width:0; }
        }
        .pill-btn{
          display:inline-flex; align-items:center; justify-content:center; gap:6px;
          border-radius:${UI.pill}px; cursor:pointer; white-space:nowrap;
          transition:background .15s ease, border-color .15s ease, box-shadow .15s ease, transform .1s ease;
        }
        .pill-btn:hover{ filter:brightness(.98); }
        .pill-btn:active{ filter:brightness(.95); }
        .pill-btn:disabled{ opacity:.45; cursor:not-allowed; filter:none; }

        .app-shell{
          width:min(100%, 1400px);
          margin:0 auto;
          padding-left:clamp(12px, 2vw, 24px);
          padding-right:clamp(12px, 2vw, 24px);
          box-sizing:border-box;
        }
        .app-topbar{
          min-height:50px;
          display:flex;
          align-items:center;
          gap:14px;
          flex-wrap:nowrap;
          padding-top:6px;
          padding-bottom:6px;
        }
        .app-brand{
          display:flex;
          align-items:baseline;
          gap:10px;
          flex:0 0 auto;
          min-width:0;
          overflow:hidden;
          white-space:nowrap;
        }
        .app-brand span:first-child{
          flex:0 0 auto;
        }
        .app-brand span:last-child{
          min-width:0;
          overflow:hidden;
          text-overflow:ellipsis;
        }
        .app-topbar-alerts{
          margin-left:auto;
          display:flex;
          gap:14px;
          align-items:center;
          justify-content:flex-end;
          flex:1 1 auto;
          min-width:140px;
          overflow:hidden;
          flex-wrap:wrap;
          white-space:nowrap;
        }
        .app-topbar-alerts > span{
          flex:0 0 auto;
        }
        .language-toggle{
          display:inline-flex;
          align-items:center;
          gap:2px;
          padding:2px;
          border:1px solid ${C.bd2};
          border-radius:999px;
          background:${C.bg};
          flex:0 0 auto;
        }
        .language-toggle button{
          height:26px;
          min-width:58px;
          padding:0 10px;
          border:0;
          border-radius:999px;
          background:transparent;
          font-size:10.5px;
          font-weight:600;
          cursor:pointer;
          color:${C.text3};
          white-space:nowrap;
        }
        .language-toggle button.active{
          background:${C.charcoalDk};
          color:#fff;
        }
        .app-tabs{
          display:flex;
          overflow-x:auto;
          scrollbar-width:thin;
        }
        .app-tabs button{
          flex:0 0 auto;
          white-space:nowrap;
        }
        .app-main{
          padding-top:clamp(12px, 1.8vw, 20px);
          padding-bottom:clamp(12px, 1.8vw, 20px);
        }
        .xrf-hero{
          padding:14px clamp(14px, 2.2vw, 24px);
          display:flex;
          align-items:stretch;
          gap:12px;
          flex-wrap:nowrap;
          overflow-x:auto;
          overflow-y:hidden;
        }
        .xrf-hero-title{
          flex:1 1 260px;
          min-width:220px;
          display:flex;
          flex-direction:column;
          justify-content:center;
          align-self:stretch;
          padding:2px 0;
        }
        .xrf-hero-photo{
          width:78px;
          height:78px;
          flex:0 0 78px;
          border-radius:10px;
          border:1px solid rgba(255,255,255,.22);
          background:rgba(255,255,255,.08);
          object-fit:cover;
          align-self:center;
        }
        .xrf-hero-title > div{
          overflow:hidden;
          text-overflow:ellipsis;
          white-space:nowrap;
          overflow-wrap:anywhere;
        }
        .xrf-hero-metrics{
          display:flex;
          gap:10px;
          align-items:flex-start;
          flex-wrap:nowrap;
          flex:0 1 auto;
          min-width:0;
        }
        .xrf-hero-meta{
          margin-left:auto;
          display:flex;
          align-items:stretch;
          flex-wrap:nowrap;
          flex:0 0 auto;
          min-width:max-content;
          border:1px solid rgba(255,255,255,.18);
          border-radius:10px;
          background:rgba(255,255,255,.06);
          overflow:hidden;
        }
        .xrf-hero-meta > div{
          min-width:64px;
        }
        .xrf-detail-grid{
          display:grid;
          grid-template-columns:minmax(360px,.9fr) minmax(420px,1.1fr);
        }
        .risk-attention-row{
          cursor:pointer;
          transition:background .15s ease;
        }
        .risk-attention-row td{
          transition:background .15s ease, border-color .15s ease;
        }
        .risk-attention-row:hover td{
          background:#F6F8F9 !important;
        }
        .risk-attention-row:focus{
          outline:none;
        }
        .risk-attention-row:focus-visible{
          outline:2px solid ${A.purple.solid};
          outline-offset:-2px;
        }
        .risk-attention-row:hover td:first-child,
        .risk-attention-row:focus-visible td:first-child{
          color:${A.purple.tx};
        }
        @media (max-width: 760px){
          .xrf-element-row{grid-template-columns:1fr !important;}
          .xrf-element-row > div{border-left:0 !important;padding-left:0 !important;}
          .responsive-table-scroll > table{
            width:100% !important;
            min-width:720px !important;
          }
          .responsive-table-scroll.table-scroll-wide > table{
            min-width:1280px !important;
          }
          .responsive-table-scroll.table-scroll-medium > table{
            min-width:820px !important;
          }
          .responsive-table-scroll.table-scroll-compact > table{
            min-width:520px !important;
          }
          .responsive-table-scroll th,
          .responsive-table-scroll td{
            overflow-wrap:break-word;
            word-break:keep-all;
          }
        }
        @media (max-width: 1180px){
          .stat-grid{grid-template-columns:repeat(3,1fr) !important;}
          .chart-grid{grid-template-columns:1fr !important;}
        }
        @media (max-width: 640px){
          .stat-grid{grid-template-columns:repeat(2,1fr) !important;}
        }
        @media (max-width: 980px){
          .xrf-hero{
            align-items:stretch;
          }
          .xrf-hero-title > div{
            white-space:normal;
          }
          .xrf-hero-meta{
            margin-left:0;
            justify-content:flex-start;
          }
          .xrf-detail-grid{
            grid-template-columns:1fr;
          }
          .xrf-detail-grid > div:first-child{
            border-right:0 !important;
            border-bottom:1px solid ${C.bd};
          }
        }
        @media (max-width: 560px){
          .app-shell{
            padding-left:10px;
            padding-right:10px;
          }
          .app-topbar{
            flex-wrap:wrap;
          }
          .app-brand{
            width:100%;
          }
          .app-topbar-alerts{
            margin-left:0;
            flex-basis:100%;
            justify-content:flex-start;
            overflow-x:auto;
            flex-wrap:nowrap;
            scrollbar-width:thin;
          }
          .app-tabs button{
            padding-left:12px !important;
            padding-right:12px !important;
          }
        }
      `}</style>
      {openFilter&&<FilterPanel column={openFilter} pos={fPos} selected={fs[openFilter]||[]} values={listFilterValues[openFilter]||[]} onToggle={handleToggle} onClear={handleClear} onClose={()=>setOpenFilter(null)}/>}
      {photoPreview&&(
        <div role="presentation" onMouseDown={()=>setPhotoPreview(null)} onKeyDown={event=>{if(event.key==="Escape")setPhotoPreview(null);}} style={{position:"fixed",inset:0,zIndex:10020,display:"flex",alignItems:"center",justifyContent:"center",padding:"clamp(14px,3vw,34px)",background:"rgba(8,12,14,.82)",boxSizing:"border-box"}}>
          <div role="dialog" aria-modal="true" aria-label={lang==="en"?"Item photo preview":"품목 사진 크게 보기"} onMouseDown={event=>event.stopPropagation()} style={{position:"relative",width:photoPreview.displayWidth?`${photoPreview.displayWidth}px`:"min(720px,calc(100vw - 28px))",maxWidth:"calc(100vw - 28px)",maxHeight:"calc(100dvh - 28px)",display:"flex",flexDirection:"column",background:C.card,border:`1px solid ${C.bd2}`,borderRadius:14,overflow:"hidden",boxShadow:"0 24px 70px rgba(0,0,0,.42)"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:14,padding:"12px 14px",borderBottom:`1px solid ${C.bd}`,background:C.card}}>
              <div style={{minWidth:0}}>
                <div style={{fontSize:13,fontWeight:750,color:C.text1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{photoPreview.code} · {photoPreview.itemName}</div>
                <div style={{marginTop:2,fontSize:10.5,color:C.text4,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{photoPreview.name}</div>
              </div>
              <button type="button" autoFocus onClick={()=>setPhotoPreview(null)} style={{flex:"0 0 auto",display:"inline-flex",alignItems:"center",gap:7,padding:"7px 11px",border:`1px solid ${C.bd2}`,borderRadius:7,background:C.card,color:C.text2,fontSize:11,fontWeight:650,cursor:"pointer"}}>{lang==="en"?"Close":"닫기"}<span aria-hidden="true" style={{fontSize:16,lineHeight:1}}>×</span></button>
            </div>
            <div style={{minHeight:0,width:"100%",alignSelf:"center",display:"block",padding:0,background:"#151b1e",overflow:"auto",boxSizing:"border-box"}}>
              <img src={photoPreview.url} alt={photoPreview.name||`${photoPreview.code} 품목 사진`} onLoad={event=>{const size=expandedPhotoDisplaySize(event.currentTarget.naturalWidth,event.currentTarget.naturalHeight);setPhotoPreview(prev=>prev&&prev.url===photoPreview.url?{...prev,...size}:prev);}} style={{display:"block",width:photoPreview.displayWidth?`${photoPreview.displayWidth}px`:"100%",height:photoPreview.displayHeight?`${photoPreview.displayHeight}px`:"auto",maxWidth:"100%",objectFit:"fill",background:"#fff"}}/>
            </div>
          </div>
        </div>
      )}
      {discontinueDecisionAuth.open&&(
        <div role="presentation" onMouseDown={()=>{if(!discontinueDecisionAuth.pending) closeDiscontinueDecisionAuth();}} style={{position:"fixed",inset:0,zIndex:10001,display:"flex",alignItems:"center",justifyContent:"center",padding:18,background:"rgba(10,16,16,.58)",boxSizing:"border-box"}}>
          <form role="dialog" aria-modal="true" aria-labelledby="admin-discontinue-auth-title" onSubmit={submitDiscontinueDecisionAuth} onMouseDown={event=>event.stopPropagation()} onKeyDown={event=>{if(event.key==="Escape"&&!discontinueDecisionAuth.pending) closeDiscontinueDecisionAuth();}} style={{width:"min(430px,100%)",padding:"20px",background:C.card,border:`1px solid ${C.bd}`,borderRadius:12,boxShadow:"0 18px 48px rgba(0,0,0,.28)"}}>
            <div id="admin-discontinue-auth-title" style={{fontSize:15,fontWeight:750,color:C.text1}}>{discontinueDecisionAuth.verified?"단종 요청 승인 여부 결정":"관리자 인증"}</div>
            <div style={{marginTop:6,fontSize:11,color:C.text3,lineHeight:1.65}}>
              요청 번호 {discontinueDecisionAuth.item?.code||"—"}<br/>
              대상 품목 {discontinueDecisionAuth.item?.discontinueTargetCode||"—"}
            </div>
            {!discontinueDecisionAuth.verified?(
              <>
                <div style={{marginTop:10,fontSize:11,color:C.text3,lineHeight:1.55}}>단종 요청을 승인하거나 반려하려면 관리자 비밀번호를 입력하세요.</div>
                <input type="password" value={discontinueDecisionAuth.password} onChange={event=>setDiscontinueDecisionAuth(prev=>({...prev,password:event.target.value,error:""}))} autoFocus autoComplete="current-password" aria-label="관리자 비밀번호" style={{...inp,width:"100%",marginTop:14,boxSizing:"border-box"}}/>
                {discontinueDecisionAuth.error&&<div role="alert" style={{marginTop:7,fontSize:11,color:C.redDk}}>{discontinueDecisionAuth.error}</div>}
                <div style={{display:"flex",justifyContent:"flex-end",gap:8,marginTop:16}}>
                  <button type="button" onClick={closeDiscontinueDecisionAuth} style={{padding:"7px 13px",border:`1px solid ${C.bd2}`,borderRadius:7,background:C.card,color:C.text2,fontSize:11,fontWeight:600,cursor:"pointer"}}>취소</button>
                  <button type="submit" style={{padding:"7px 15px",border:`1px solid ${C.charcoal}`,borderRadius:7,background:C.charcoal,color:"#fff",fontSize:11,fontWeight:650,cursor:"pointer"}}>인증</button>
                </div>
              </>
            ):(
              <>
                <div style={{marginTop:12,padding:"11px 12px",background:C.alt,border:`1px solid ${C.bd}`,borderRadius:7,fontSize:11,color:C.text2,lineHeight:1.65}}>관리자 인증이 완료되었습니다. 승인하면 대상 품목이 단종 상태로 전환되고, 반려하면 대상 품목의 현재 상태와 기존 XRF 이력이 그대로 유지됩니다.</div>
                {discontinueDecisionAuth.error&&<div role="alert" style={{marginTop:9,fontSize:11,color:C.redDk,lineHeight:1.55}}>{discontinueDecisionAuth.error}</div>}
                <div style={{display:"flex",justifyContent:"flex-end",gap:8,marginTop:16,flexWrap:"wrap"}}>
                  <button type="button" disabled={discontinueDecisionAuth.pending} onClick={closeDiscontinueDecisionAuth} style={{padding:"7px 13px",border:`1px solid ${C.bd2}`,borderRadius:7,background:C.card,color:C.text2,fontSize:11,fontWeight:600,cursor:discontinueDecisionAuth.pending?"not-allowed":"pointer",opacity:discontinueDecisionAuth.pending?0.6:1}}>취소</button>
                  <button type="button" disabled={discontinueDecisionAuth.pending} onClick={()=>handleDiscontinueApproval("반려",discontinueDecisionAuth.item)} style={{padding:"7px 15px",border:`1px solid ${A.rose.ln}`,borderRadius:7,background:A.rose.bg,color:A.rose.tx,fontSize:11,fontWeight:650,cursor:discontinueDecisionAuth.pending?"not-allowed":"pointer",opacity:discontinueDecisionAuth.pending?0.6:1}}>반려</button>
                  <button type="button" disabled={discontinueDecisionAuth.pending} onClick={()=>handleDiscontinueApproval("승인",discontinueDecisionAuth.item)} style={{padding:"7px 15px",border:`1px solid ${C.charcoal}`,borderRadius:7,background:C.charcoal,color:"#fff",fontSize:11,fontWeight:650,cursor:discontinueDecisionAuth.pending?"not-allowed":"pointer",opacity:discontinueDecisionAuth.pending?0.6:1}}>{discontinueDecisionAuth.pending?"SharePoint 처리 중...":"단종 승인"}</button>
                </div>
              </>
            )}
          </form>
        </div>
      )}
      {itemEditAuth.open&&(
        <div role="presentation" onMouseDown={closeItemEditAuth} style={{position:"fixed",inset:0,zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:18,background:"rgba(10,16,16,.58)",boxSizing:"border-box"}}>
          <form role="dialog" aria-modal="true" aria-labelledby="admin-item-auth-title" onSubmit={submitItemEditAuth} onMouseDown={event=>event.stopPropagation()} onKeyDown={event=>{if(event.key==="Escape") closeItemEditAuth();}} style={{width:"min(380px,100%)",padding:"20px",background:C.card,border:`1px solid ${C.bd}`,borderRadius:12,boxShadow:"0 18px 48px rgba(0,0,0,.28)"}}>
            <div id="admin-item-auth-title" style={{fontSize:15,fontWeight:750,color:C.text1}}>관리자 인증</div>
            <div style={{marginTop:6,fontSize:11,color:C.text3,lineHeight:1.55}}>품목 정보를 수정하려면 관리자 비밀번호를 입력하세요.</div>
            <input type="password" value={itemEditAuth.password} onChange={event=>setItemEditAuth(prev=>({...prev,password:event.target.value,error:""}))} autoFocus autoComplete="current-password" aria-label="관리자 비밀번호" style={{...inp,width:"100%",marginTop:14,boxSizing:"border-box"}}/>
            {itemEditAuth.error&&<div role="alert" style={{marginTop:7,fontSize:11,color:C.redDk}}>{itemEditAuth.error}</div>}
            <div style={{display:"flex",justifyContent:"flex-end",gap:8,marginTop:16}}>
              <button type="button" onClick={closeItemEditAuth} style={{padding:"7px 13px",border:`1px solid ${C.bd2}`,borderRadius:7,background:C.card,color:C.text2,fontSize:11,fontWeight:600,cursor:"pointer"}}>취소</button>
              <button type="submit" style={{padding:"7px 15px",border:`1px solid ${C.charcoal}`,borderRadius:7,background:C.charcoal,color:"#fff",fontSize:11,fontWeight:650,cursor:"pointer"}}>확인</button>
            </div>
          </form>
        </div>
      )}

      <div style={{background:C.card,borderBottom:`1px solid ${C.bd}`}}>
        <div className="app-shell app-topbar">
          <button type="button" className="app-brand" onClick={()=>setEntered(false)} aria-label={lang==="en"?"Return to start screen":"시작 화면으로 돌아가기"} title={lang==="en"?"Return to start screen":"시작 화면으로 돌아가기"} style={{padding:0,border:0,background:"transparent",cursor:"pointer",fontFamily:FONT_SANS}}>
            <span style={{fontFamily:"Arial Black,Arial,sans-serif",fontWeight:900,fontSize:21,color:C.red,letterSpacing:"-1px"}}>molex</span>
            <span style={{fontFamily:"Georgia,serif",fontStyle:"italic",fontSize:11,color:C.text4}}>creating connections for life</span>
          </button>
          <div className="app-topbar-alerts">
            <span style={{fontSize:11,color:C.text4}}>기준일 {TODAY_STR}</span>
            <div className="language-toggle" data-i18n-skip="true" role="group" aria-label="Language">
              <button type="button" className={lang==="ko"?"active":""} aria-pressed={lang==="ko"} onClick={()=>setLang("ko")}>{lang==="en"?"Korean":"한국어"}</button>
              <button type="button" className={lang==="en"?"active":""} aria-pressed={lang==="en"} onClick={()=>setLang("en")}>English</button>
            </div>
          </div>
        </div>
        <div style={{background:C.charcoalDk}}>
          <div className="app-shell app-tabs">
            {TABS.map(t=>{const activeTab=tab===t.id || (tab==="precision-detail"&&t.id==="precision");return <button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"9px 18px",background:"none",border:"none",cursor:"pointer",fontSize:12,fontWeight:activeTab?600:400,color:activeTab?"#fff":"rgba(255,255,255,.55)",borderBottom:activeTab?`2px solid ${C.red}`:"2px solid transparent",transition:"color .15s ease, border-color .15s ease"}}>{t.l}</button>})}
          </div>
        </div>
      </div>

      {VISUAL_UAT_MODE&&(
        <div data-i18n-skip="true" data-visual-uat-banner="1" style={{background:A.purple.bg,borderBottom:`1px solid ${A.purple.ln}`}}>
          <div className="app-shell" style={{paddingTop:9,paddingBottom:9,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <strong style={{fontSize:11,color:A.purple.tx,letterSpacing:".35px"}}>VISUAL UAT MODE</strong>
            <span style={{fontSize:11,color:C.text2}}>로컬 테스트 데이터가 실제 SharePoint 데이터와 함께 표시됩니다.</span>
            <span style={{fontSize:11,color:A.purple.tx,fontWeight:650}}>테스트 데이터는 SharePoint에 저장되지 않습니다.</span>
            <span style={{marginLeft:"auto",fontSize:10,color:C.text4}}>Fixtures {VISUAL_UAT_ALL_ITEMS.length} · Compliance {VISUAL_UAT_COMPLIANCE_ITEMS.length}</span>
          </div>
        </div>
      )}

      <div className="app-shell app-main">

        {tab==="list"&&(
          <div>
            <div className="stat-grid" style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:12,marginBottom:12}}>
              {[
                ["전체 품목",stats.total,A.slate.solid,"운영 관리 대상",null,false],
                ["신규 등록",stats.new,A.purple.solid,"첫 관리주기 내 등록",{category:["new"]},true],
                ["측정 의뢰",stats.measuring,A.blue.solid,"XRF 측정 대기",{approval:["measuring"]},true],
                ["미이행",stats.overdue,A.amber.solid,"기한 내 미측정",{compliance:["overdue"]},true],
                ["지연측정",stats.late,A.rose.solid,"과거 주기의 지연 측정 이력",null,true],
                ["이력",stats.history,A.slate.ln,"단종·취소·요청 기록",{lifecycle:["Discontinued","Cancelled","ReplacedOld"]},false],
              ].map(([l,n,c,s,filter,showShare])=>(
                <StatCard key={l} label={l} value={n} sub={s} color={c}
                  share={showShare&&stats.total?`${Math.round((n/stats.total)*100)}%`:null}
                  onClick={filter?()=>gotoListWithFilter(filter):undefined}/>
              ))}
            </div>

            <div className="chart-grid" style={{display:"grid",gridTemplateColumns:"1.15fr 1fr 1fr",gap:12,marginBottom:16}}>
              <ChartCard title="측정 도래 예정" hint={`${TODAY_STR} 기준 향후 6개월`}>
                <DueLoadChart buckets={dash.buckets} onPick={b=>b.total>0&&gotoDueMonthList(b.key)}/>
              </ChartCard>
              <ChartCard title="이행 상태 구성" hint={`운영 관리 대상 ${dash.csTotal}건`}>
                <ComplianceBar segments={dash.segments} total={dash.csTotal}
                  onPick={s=>s.key!=="notAvailable"&&gotoListWithFilter({compliance:[s.key]})}/>
              </ChartCard>
              <ChartCard title="부서별 위험도 구성" hint="R 단계에서 확정된 Final Risk 기준">
                <DeptRiskChart rows={dash.deptRows} riskMeta={dash.riskMeta}/>
              </ChartCard>
            </div>

            {VISUAL_UAT_MODE&&<VisualUatComplianceRenderMatrix/>}

            <div style={{position:"relative",zIndex:2,padding:"0 0 12px"}}>
              <div style={{minHeight:34,display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:hasFilters?8:0}}>
                <div style={{position:"relative",flex:"0 0 auto"}}>
                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke={C.text4} strokeWidth="2.2" strokeLinecap="round"
                    style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",pointerEvents:"none"}}>
                    <circle cx="11" cy="11" r="7"/><path d="M20 20l-3.6-3.6"/>
                  </svg>
                  <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="전체 열 검색"
                    style={{height:34,width:220,boxSizing:"border-box",padding:"0 12px 0 32px",border:`1px solid ${C.bd}`,borderRadius:UI.rs,fontSize:12,outline:"none",background:C.card,color:C.text1}}/>
                </div>
                {VISUAL_UAT_MODE&&<button type="button" data-i18n-skip="true" aria-pressed={visualUatOnly} onClick={()=>setVisualUatOnly(value=>!value)} className="pill-btn" style={{height:30,padding:"0 12px",border:`1px solid ${visualUatOnly?A.purple.solid:A.purple.ln}`,background:visualUatOnly?A.purple.solid:A.purple.bg,color:visualUatOnly?"#fff":A.purple.tx,fontSize:11,fontWeight:650,cursor:"pointer"}}>UAT 테스트만 보기</button>}
                <span style={{fontSize:11,color:C.text4}}>결과 <strong style={{color:C.text2,fontWeight:700}}>{filtered.length}</strong>건 · 운영 {stats.total}건 · 이력 {stats.history}건</span>
                {hasFilters&&<button onClick={handleClearAll} className="pill-btn" style={{height:26,padding:"0 13px",background:C.card,border:`1px solid ${C.bd2}`,fontSize:11,color:C.text3,fontWeight:500}}>필터 초기화</button>}
                <div style={{marginLeft:"auto"}}/>
                <button onClick={()=>{setTab("reg");setRegMode("requester");setRegStep(1);setRegType(null);setReqXrfMode("request");}}
                  style={{height:34,padding:"0 16px",background:C.charcoalDk,color:"#fff",border:"none",borderRadius:UI.rs,fontSize:12,fontWeight:600,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:6}}>
                  <span style={{fontSize:15,lineHeight:1}}>+</span> 품목 / 변경 관리
                </button>
              </div>
              <FilterChips filterState={fs} onClear={handleClear} onClearAll={handleClearAll}/>
            </div>

            <div className="responsive-table-scroll table-scroll-wide material-list-table" role="region" aria-label="부자재 목록 표" tabIndex={0} style={{background:C.card,border:`1px solid ${C.bd}`,borderRadius:UI.r,boxShadow:UI.sh,width:"100%",maxWidth:"100%"}}>
              <table style={{width:"100%",minWidth:listTableWeight,maxWidth:"none",borderCollapse:"collapse",tableLayout:"fixed"}}>
                <colgroup>
                  {listColumns.map(c=><col key={c.key} style={{width:`${(c.w/listTableWeight)*100}%`}}/>)}
                </colgroup>
                <thead>
                  <tr style={{height:40}}>
                    {listColumns.map(({key,el})=>(
                      <th key={key} style={{height:40,padding:"4px 6px",borderBottom:`1px solid ${C.bd}`,textAlign:"center",verticalAlign:"middle",boxSizing:"border-box",background:C.card,whiteSpace:"normal",overflow:"visible",lineHeight:1.15}}>{el}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {filtered.length===0?(
                    <tr><td colSpan={listColumns.length} style={{padding:"40px",textAlign:"center",color:C.text4,borderBottom:`1px solid ${C.bd}`}}>
                      조건에 맞는 품목이 없습니다.
                      <button onClick={handleClearAll} style={{marginLeft:10,padding:"4px 10px",background:"none",border:`1px solid ${C.bd}`,borderRadius:0,fontSize:11,cursor:"pointer",color:C.text3}}>초기화</button>
                    </td></tr>
                  ):pagedFiltered.map((item,idx)=>{
                    const isReplaced=item.lifecycle==="ReplacedOld";
                    const isDiscontinueItem=isDiscontinueRequestItem(item);
                    const isDiscontinuedExisting=isDiscontinuedItem(item) && !isDiscontinueItem;
                    const itemPeriods=itemListPeriods(item);
                    const nextDue=nextDueOfItem(item);
                    const approvalKey=approvalFilterKey(approvalStatusOf(item));
                    const isMeasuringItem=approvalKey==="measuring";
                    const isDiscontinueProcessed=isDiscontinueItem && (approvalKey==="approved" || approvalKey==="rejected");
                    const isCycleMuted=isDiscontinueItem || isDiscontinuedExisting || !isComplianceTarget(item);
                    const hideTimeline=isDiscontinueItem || isDiscontinuedExisting || isReplaced || item.category==="changed";
                    const showTimeline=!hideTimeline && itemPeriods.length>0;
                    const rowOpacity=(isReplaced || isDiscontinueProcessed || isDiscontinuedExisting) ? 0.55 : 1;
                    const tdS={height:UI.rowH,padding:"0 6px",borderBottom:`1px solid ${C.bd}`,verticalAlign:"middle",opacity:rowOpacity,boxSizing:"border-box",textAlign:"center"};
                    const mutedCell=<span style={{fontSize:11,color:C.text4}}>—</span>;
                    const notTargetCell=<span title="대상 아님" style={{fontSize:12,color:C.text4,fontWeight:500}}>—</span>;
                    const isRowActive=rowSel===item.code;
                    const primaryItemName=displayItemName(item,lang)||"—";
                    const secondaryItemName=lang==="en"?(item.name||"—"):(item.nameEn||ITEM_NAME_EN_FALLBACKS[item.code]||"—");
                    return(
                      <tr key={item.code}
                        onClick={()=>setRowSel(prev=>prev===item.code?null:item.code)}
                        style={{background:isRowActive?UI.selRow:C.card,cursor:"pointer",transition:"background .12s"}}
                        onMouseEnter={e=>{if(!isRowActive)e.currentTarget.style.background=UI.hoverRow;}}
                        onMouseLeave={e=>{if(!isRowActive)e.currentTarget.style.background=C.card;}}>
                        <td style={{...tdS,overflow:"hidden"}}>
                          <div style={{width:"100%",minWidth:0,overflow:"hidden",display:"flex",alignItems:"center",justifyContent:"center"}}>
                            <CatDot cat={item.category} lang={lang}/>
                          </div>
                        </td>
                        <td style={{...tdS,padding:"5px 6px",textAlign:"center"}}>
                          {item.photoFileUrl?(
                            <button type="button" onClick={e=>{e.stopPropagation();setPhotoPreview({url:item.photoFileUrl,name:item.photoFileName||"품목 사진",code:item.code,itemName:displayItemName(item,lang)});}} title={item.photoFileName||"품목 사진"} aria-label={`${item.code} 품목 사진 크게 보기`} style={{display:"inline-flex",padding:0,border:0,borderRadius:7,background:"transparent",cursor:"zoom-in"}}>
                              <img src={item.photoFileUrl} alt={item.photoFileName||`${item.code} 품목 사진`} loading="lazy" style={{width:44,height:44,objectFit:"cover",borderRadius:7,border:`1px solid ${C.bd}`,background:C.alt}}/>
                            </button>
                          ):<span style={{fontSize:9.5,color:C.text4,lineHeight:1.2}}>사진 없음</span>}
                        </td>
                        <td style={{...tdS,color:C.text3,fontSize:11,fontWeight:600,overflow:"hidden",textAlign:"center"}}>
                          <AutoFitText value={item.code} title={item.code} baseFontSize={11} minFontSize={7} align="center" style={{fontWeight:600,color:C.text3}}/>
                        </td>
                        <td data-i18n-skip="true" style={{...tdS,textAlign:"left",overflow:"visible",padding:"7px 6px"}}>
                          <div title={primaryItemName} style={{fontSize:12,fontWeight:600,color:C.text1,textAlign:"left",lineHeight:1.35,whiteSpace:"normal",wordBreak:"keep-all",overflowWrap:"anywhere"}}>
                            {primaryItemName}
                          </div>
                          <div title={secondaryItemName} style={{fontSize:10,color:C.text2,marginTop:3,textAlign:"left",lineHeight:1.35,whiteSpace:"normal",wordBreak:"keep-all",overflowWrap:"anywhere"}}>
                            {secondaryItemName}
                          </div>
                        </td>
                        <td style={{...tdS,color:C.text3,fontSize:11}}><AutoFitText value={item.dept} baseFontSize={11} minFontSize={7} align="center" style={{color:C.text3}}/></td>
                        <td style={{...tdS,color:C.text3,fontSize:11,textAlign:"center"}}><AutoFitText value={isCycleMuted||isEquipmentItem(item)?"—":(CYCLE_KO[itemCycleFromRisk(item)]||"—")} baseFontSize={11} minFontSize={7} align="center" style={{color:C.text3}}/></td>
                        <td style={{...tdS,color:C.text3,fontSize:11,textAlign:"center"}}><AutoFitText value={item.firstDate||"—"} baseFontSize={11} minFontSize={7} align="center" style={{color:C.text3}}/></td>
                        <td style={{...tdS,padding:"0 6px",textAlign:"center"}}>{showTimeline?<DotCompact periods={itemPeriods}/>:mutedCell}</td>
                        <td style={{...tdS,padding:"0 6px",textAlign:"center"}}>{isCycleMuted?mutedCell:<CompText item={item}/>}</td>
                        <td style={{...tdS,padding:"0 6px",cursor:"pointer",textAlign:"center"}} onClick={e=>{e.stopPropagation();gotoXrf(item.code);}} title="XRF 분석으로 이동">
                          {isDiscontinueItem?notTargetCell:<XrfWorkflowCell item={item}/>}
                        </td>
                        <td style={{...tdS,padding:"0 6px",textAlign:"center",cursor:itemHasPrecisionTrigger(item)||workflowForItem(item).precision||workflowForItem(item).precisionCases?.length?"pointer":"default"}} onClick={e=>{e.stopPropagation();const wf=workflowForItem(item);if(itemHasPrecisionTrigger(item)||wf.precision||wf.precisionCases?.length){setPrecisionSelCaseId(wf.caseData?.id||wf.activePrecisionCaseKey||wf.precisionCases?.[0]?.key||latestMeasurementOf(item)?.id||null);setSelKey(item.code);setTab("precision");}}}>{isDiscontinueItem?notTargetCell:<PrecisionWorkflowCell item={item}/>}</td>
                        <td style={{...tdS,padding:"0 6px",textAlign:"center"}}><ApprovalBadge status={approvalStatusOf(item)} context={isDiscontinueItem?"discontinue":"use"} detail={item.approvalDetail||item.retestRequiredElements||latestMeasurementOf(item)?.retestRequiredElements}/></td>
                        <td style={{...tdS,color:C.text3,fontSize:11,textAlign:"center"}}><AutoFitText value={isCycleMuted?"—":(nextDue||"—")} baseFontSize={11} minFontSize={7} align="center" style={{color:C.text3}}/></td>
                        <td style={{...tdS,padding:"0 6px",textAlign:"center"}}>{isCycleMuted||isReplaced?<span style={{color:C.text4,fontSize:11}}>—</span>:<DDayText dateStr={nextDue}/>}</td>
                        <td style={{...tdS,padding:"0 6px",whiteSpace:"nowrap",textAlign:"center"}}>
                          <span onClick={e=>{e.stopPropagation();gotoItemDetail(item);}}><DetailArrowButton active={selKey===item.code && ["xrf","precision-detail"].includes(tab)} title={itemHasPrecisionTrigger(item)||workflowForItem(item).precision||workflowForItem(item).precisionCases?.length?"정밀분석 상세 열기":"XRF 분석 열기"}/></span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
            {listTotalPages>1&&(
              <div data-material-list-pagination="1" style={{display:"flex",alignItems:"center",justifyContent:"center",gap:6,marginTop:12,flexWrap:"wrap"}}>
                <button type="button" disabled={listPage===0} onClick={()=>setListPage(page=>Math.max(0,page-1))}
                  style={{padding:"6px 10px",border:`1px solid ${C.bd}`,borderRadius:6,background:C.card,color:C.text2,fontSize:11,cursor:listPage===0?"default":"pointer",opacity:listPage===0?.45:1}}>이전</button>
                {Array.from({length:listTotalPages},(_,page)=><button key={page} type="button" aria-current={page===listPage?"page":undefined} onClick={()=>setListPage(page)}
                  style={{minWidth:30,padding:"6px 8px",border:`1px solid ${page===listPage?C.charcoalDk:C.bd}`,borderRadius:6,background:page===listPage?C.charcoalDk:C.card,color:page===listPage?"#fff":C.text2,fontSize:11,fontWeight:page===listPage?750:500,cursor:"pointer"}}>{page+1}</button>)}
                <button type="button" disabled={listPage>=listTotalPages-1} onClick={()=>setListPage(page=>Math.min(listTotalPages-1,page+1))}
                  style={{padding:"6px 10px",border:`1px solid ${C.bd}`,borderRadius:6,background:C.card,color:C.text2,fontSize:11,cursor:listPage>=listTotalPages-1?"default":"pointer",opacity:listPage>=listTotalPages-1?.45:1}}>다음</button>
                <span style={{marginLeft:4,fontSize:10.5,color:C.text4}}>{listPage*MATERIAL_LIST_PAGE_SIZE+1}–{Math.min(filtered.length,(listPage+1)*MATERIAL_LIST_PAGE_SIZE)} / {filtered.length}건</span>
              </div>
            )}
          </div>
        )}

        {tab==="xrf"&&(
          <div>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:18}}>
              <span style={{...T.section}}>품목 선택</span>
              <select data-i18n-skip="true" value={selKey||""} onChange={e=>{setSelKey(e.target.value);setSelPeriodNum(null);}} style={{...inp,minWidth:220,width:"min(100%, 420px)",maxWidth:"100%"}}>
                <option value="">{lang==="en"?"-- Select --":"-- 선택하세요 --"}</option>
                {xrfSelectableItems.map(i=><option key={i.code} value={i.code}>{i.code}  |  {i.dept}  |  {displayItemName(i,lang)}</option>)}
              </select>
            </div>
            {pendingDiscontinueRequestItems.length>0&&<div style={{...card,padding:"12px 14px",marginBottom:14,background:C.alt}}>
              <div style={{fontSize:11,fontWeight:700,color:C.charcoal,marginBottom:8}}>단종 처리 요청 리스트</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {pendingDiscontinueRequestItems.map(item=>{
                  const key=approvalFilterKey(approvalStatusOf(item));
                  const processed=key==="approved" || key==="rejected";
                  return <button key={item.code} onClick={()=>{setSelKey(item.code);setSelPeriodNum(null);}}
                    style={{padding:"6px 10px",border:`1px solid ${selKey===item.code?C.charcoal:C.bd}`,background:selKey===item.code?C.charcoalBg:C.card,color:C.text2,fontSize:11,cursor:"pointer",opacity:processed?0.55:1}}>
                    {item.code} · {displayItemName(item,lang)} · <ApprovalBadge status={approvalStatusOf(item)} context="discontinue"/>
                  </button>;
                })}
              </div>
            </div>}
            {!sel?(
              <div style={{...card,padding:60,textAlign:"center",color:C.text4}}>품목을 선택하면 XRF 분석 결과가 표시됩니다.</div>
            ):(
              <div style={{display:"flex",flexDirection:"column",...card,overflow:"hidden"}}>
                <div className="xrf-hero" style={{background:C.charcoalDk}}>
                  {sel.photoFileUrl
                    ? <button type="button" onClick={()=>setPhotoPreview({url:sel.photoFileUrl,name:sel.photoFileName||"품목 사진",code:sel.code,itemName:displayItemName(sel,lang)})} title={sel.photoFileName||"품목 사진"} aria-label={`${sel.code} 품목 사진 크게 보기`} style={{display:"inline-flex",alignSelf:"center",padding:0,border:0,borderRadius:10,background:"transparent",cursor:"zoom-in"}}><img className="xrf-hero-photo" src={sel.photoFileUrl} alt={sel.photoFileName||`${sel.code} 품목 사진`}/></button>
                    : <div className="xrf-hero-photo" style={{display:"flex",alignItems:"center",justifyContent:"center",color:"rgba(255,255,255,.45)",fontSize:10,textAlign:"center",padding:6,boxSizing:"border-box"}}>사진 없음</div>}
                  <div className="xrf-hero-title">
                    <div style={{fontSize:16,fontWeight:600,color:"#fff",marginBottom:3,letterSpacing:"-.2px"}}>{displayItemName(sel,lang)}</div>
                    <div style={{fontSize:12,color:"rgba(255,255,255,.72)",fontWeight:400}}>{sel.code} · {sel.dept}</div>
                    {isAdminUser&&!selectedIsDiscontinue&&<button type="button" onClick={()=>openAdminItemEditor(sel)} style={{alignSelf:"flex-start",marginTop:8,padding:"5px 9px",border:"1px solid rgba(255,255,255,.38)",borderRadius:6,background:"rgba(255,255,255,.08)",color:"#fff",fontSize:10.5,fontWeight:650,cursor:"pointer"}}>품목정보 수정</button>}
                  </div>
                  {selectedIsDiscontinue&&(
                    <div className="xrf-hero-metrics">
                      <>
                        <MetricChip label="요청 유형" dark><span style={{fontSize:11,fontWeight:800,color:"#fff"}}>단종 / 사용 중지</span></MetricChip>
                        <MetricChip label="Approval_Status" dark><ApprovalBadge status={approvalStatusOf(sel)} context="discontinue"/></MetricChip>
                        <MetricChip label="XRF" dark><span style={{fontSize:11,fontWeight:800,color:"rgba(255,255,255,.86)"}}>대상 아님</span></MetricChip>
                      </>
                    </div>
                  )}
                  {/* 우측 정보표 — 변동값(선택 주기 XRF 결과 · 현재 단계)과
                      고정값(최초 등록 · R Final Risk · R 검사주기 · C&R)을 한 표에 칸으로 나눠 담습니다. */}
                  <div className="xrf-hero-meta">
                    {(selectedIsDiscontinue
                      ? [["요청일",sel.firstDate],["최종 사용 예정일",sel.finalUseDate||"—"]]
                      : [
                          [`${selectedPeriodLabel} XRF 결과`, <Chip v={displayXrfText(selectedMetrics.xrf)} small result/>],
                          ["현재 단계", workflowStageLabel(workflowForItem(sel).stage)],
                          ["최초 등록", sel.firstDate],
                          ["R Final Risk", itemFinalRisk(sel)||"—"],
                          ["R 검사주기", CYCLE_KO[itemCycleFromRisk(sel)]||itemCycleFromRisk(sel)||"—"],
                          ["C&R", `${sel.crLevel||"—"} · ${sel.crType||"—"}`]
                        ]
                    ).map(([k,v],i,arr)=>(
                      <div key={k} style={{padding:"8px 14px",minWidth:92,display:"flex",flexDirection:"column",
                        alignItems:"center",justifyContent:"center",gap:5,
                        borderRight:i<arr.length-1?"1px solid rgba(255,255,255,.14)":"none"}}>
                        <div style={{fontSize:11,color:"rgba(255,255,255,.74)",fontWeight:500,whiteSpace:"nowrap"}}>{k}</div>
                        <div style={{width:"100%",minWidth:0,fontSize:12,fontWeight:650,color:"#fff",whiteSpace:"nowrap",display:"flex",alignItems:"center",justifyContent:"center"}}>{typeof v==="string" || typeof v==="number" ? <AutoFitText value={v} baseFontSize={12} minFontSize={7} align="center" style={{fontWeight:650,color:"#fff"}}/> : v}</div>
                      </div>
                    ))}
                  </div>
                </div>
                {itemEditOpen&&itemEditForm&&itemEditForm.sourceKey===itemStateKey(sel)&&!selectedIsDiscontinue&&(
                  <div style={{padding:"14px clamp(14px, 2.4vw, 32px)",borderBottom:`1px solid ${C.bd}`,background:C.alt}}>
                    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,marginBottom:12,flexWrap:"wrap"}}>
                      <div>
                        <div style={{fontSize:12,fontWeight:750,color:C.text1}}>관리자 품목정보 수정</div>
                        <div style={{fontSize:10.5,color:C.text3,marginTop:3}}>저장 즉시 부자재 리스트와 XRF 분석 화면에 같은 값이 반영됩니다.</div>
                      </div>
                      <span style={{fontSize:10,color:C.text4}}>관리자 인증 완료</span>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(150px,1fr))",gap:10}}>
                      {[
                        ["한글 품목명","name","text"],
                        ["영문 품목명","nameEn","text"],
                        ["공정명","dept","text"],
                        ["품번","code","text"],
                      ].map(([label,key,type])=>(
                        <label key={key} style={{display:"flex",flexDirection:"column",gap:5,minWidth:0}}>
                          <span style={{fontSize:10.5,fontWeight:650,color:C.text3}}>{label}</span>
                          <input type={type} value={itemEditForm[key]??""} onChange={e=>updateItemEditField(key,e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}/>
                        </label>
                      ))}
                    </div>
                    <div style={{display:"flex",justifyContent:"flex-end",gap:8,marginTop:12}}>
                      <button type="button" onClick={()=>{setItemEditOpen(false);setItemEditForm(null);}} style={{padding:"7px 12px",border:`1px solid ${C.bd2}`,borderRadius:7,background:C.card,color:C.text2,fontSize:11,fontWeight:600,cursor:"pointer"}}>취소</button>
                      <button type="button" onClick={saveAdminItemEditor} style={{padding:"7px 14px",border:`1px solid ${C.charcoal}`,borderRadius:7,background:C.charcoal,color:"#fff",fontSize:11,fontWeight:650,cursor:"pointer"}}>변경사항 저장</button>
                    </div>
                  </div>
                )}
                {!selectedIsDiscontinue&&<WorkflowStrip item={sel} lang={lang}/>}
                {selectedDiscontinueRequestForTarget&&(
                   <div style={{padding:"12px clamp(14px, 2.4vw, 32px)",borderBottom:`1px solid ${C.bd}`,background:C.charcoalBg,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                    <span style={{fontSize:11,color:C.text2,fontWeight:700}}>이 품목은 단종 승인 처리됨</span>
                    <span style={{fontSize:10,color:C.text3}}>요청 번호 {discontinueRequestCodeOf(selectedDiscontinueRequestForTarget)}</span>
                    {isAdminUser&&<button onClick={()=>handleUndoDiscontinue(selectedDiscontinueRequestForTarget)} className="pill-btn" style={{marginLeft:"auto",padding:"7px 14px",border:`1px solid ${A.amber.ln}`,background:A.amber.bg,color:A.amber.tx,fontSize:11,fontWeight:600}}>단종 처리 원복</button>}
                  </div>
                )}
                {!selectedDiscontinueRequestForTarget && sel.restoredFromDiscontinueRequestCode&&(
                  <div style={{padding:"12px clamp(14px, 2.4vw, 32px)",borderBottom:`1px solid ${C.bd}`,background:C.alt,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                    <span style={{fontSize:11,color:C.charcoal,fontWeight:800}}>단종 원복 완료</span>
                    <span style={{fontSize:10,color:C.text3}}>요청 번호 {sel.restoredFromDiscontinueRequestCode}</span>
                    <span style={{fontSize:10,color:C.text3}}>원복일 {sel.restoredAt||"—"}</span>
                    <span style={{fontSize:10,color:C.text3}}>사유 {sel.restoreReason||"—"}</span>
                  </div>
                )}
                {selectedIsDiscontinue?(
                  <div style={{padding:"16px clamp(14px, 2.4vw, 32px)",borderBottom:`1px solid ${C.bd}`,background:C.alt}}>
                    <div style={{padding:"14px 16px",background:C.card,border:`1px solid ${C.bd}`,borderRadius:4}}>
                      <div style={{fontSize:12,fontWeight:800,color:C.text1,marginBottom:5}}>단종 / 사용 중지 처리 요청</div>
                      <div style={{fontSize:11,color:C.text3,lineHeight:1.7,marginBottom:10}}>
                        단종 요청은 XRF 측정 대상이 아니므로 원소 분석, Final Risk, 주기별 이행현황을 계산하지 않습니다. 관리자는 요청 사유만 검토하여 승인 또는 반려를 적용합니다.
                      </div>
                      {isAdminUser && selectedApprovalKey==="hold"?(
                        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                          <button onClick={openDiscontinueDecisionAuth} style={{padding:"8px 15px",border:"none",borderRadius:UI.rs,background:C.charcoalDk,color:"#fff",fontSize:11,fontWeight:600,cursor:"pointer"}}>관리자 인증 후 승인 여부 결정</button>
                        </div>
                      ):(
                        <div style={{display:"flex",alignItems:"center",gap:8,fontSize:11,color:C.text3,flexWrap:"wrap"}}>
                          <span>현재 처리 상태</span><ApprovalBadge status={approvalStatusOf(sel)} context="discontinue"/>
                          {isAdminUser && canUndoDiscontinueRequest(sel)&&(
                            <button onClick={handleUndoDiscontinue} className="pill-btn" style={{padding:"7px 14px",border:`1px solid ${A.amber.ln}`,background:A.amber.bg,color:A.amber.tx,fontSize:11,fontWeight:600}}>단종 처리 원복</button>
                          )}
                          {selectedApprovalKey==="cancelled"&&<span style={{fontSize:10,color:C.text4}}>원복 사유: {sel.revertReason||"—"}</span>}
                        </div>
                      )}
                    </div>
                  </div>
                ):selectedIsMeasuring?(
                  <div style={{padding:"16px clamp(14px, 2.4vw, 32px)",borderBottom:`1px solid ${C.bd}`,background:C.card}}>
                    <div style={{padding:"14px 16px",background:C.xWarnBg,border:`1px solid ${C.bd}`,borderRadius:4}}>
                      <div style={{fontSize:12,fontWeight:800,color:C.text1,marginBottom:5}}>현재 상태: 보류 · XRF 결과 대기</div>
                      <div style={{fontSize:11,color:C.text3,lineHeight:1.6,marginBottom:10}}>의뢰자가 XRF 측정을 요청했지만 아직 판정 지표가 없어 보류 상태입니다. XRF 결과 업로드 후 R 위험도 평가와 주기별 이행현황이 생성됩니다.</div>
                      {isAdminUser?<div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                        <label style={{position:"relative",display:"inline-flex",alignItems:"center",justifyContent:"center",minHeight:34,padding:"0 14px",border:`1px dashed ${C.bd2}`,background:C.card,color:C.text2,fontSize:11,fontWeight:700,cursor:"pointer",overflow:"hidden"}}>
                          {measurementSavePending?"SharePoint 저장 중...":"관리자 XRF 결과 업로드 (.xlsx)"}
                          <input type="file" accept=".xlsx,.xlsm,.xls,.csv,.json" onChange={handleAdminXrfUpload} disabled={measurementSavePending} style={fileInputOverlay}/>
                        </label>
                        <span style={{fontSize:10,color:C.text3}}>
                          {adminUploadFileNameByCode[sel.code] ? `최근 업로드: ${adminUploadFileNameByCode[sel.code]}` : "업로드 후 XRF 결과, 재측정 여부, Approval_Status가 자동 반영됩니다."}
                        </span>
                      </div>:<div style={{fontSize:10,color:C.text3}}>관리자만 XRF 결과를 업로드할 수 있습니다.</div>}
                      {isAdminUser&&adminUploadResultByCode[sel.code]&&<UploadedXrfPreview result={adminUploadResultByCode[sel.code]} compact/>}
                    </div>
                  </div>
                ):(
                  <div style={{padding:"16px clamp(14px, 2.4vw, 32px) 12px",borderBottom:`1px solid ${C.bd}`,background:C.card}}>
                    <div style={{position:"relative",display:"flex",alignItems:"center",justifyContent:"flex-start",minHeight:26,marginBottom:12,width:"100%"}}>
                      <div style={{...T.section,textAlign:"left",margin:0}}>주기별 이행 현황</div>
                      {selPeriodNum!==null&&<div style={{position:"absolute",right:0,top:"50%",transform:"translateY(-50%)"}}>
                        <button onClick={()=>{setSelPeriodNum(null);setSelPeriodPage(null);}} className="pill-btn" style={{padding:"4px 12px",background:C.card,border:`1px solid ${C.bd2}`,fontSize:11,color:C.text3}}>최근 측정으로</button>
                      </div>}
                    </div>
                    <DotFull periods={selTimelinePeriods} windowDays={30} selectedPeriodNum={selPeriodNum} onSelectPeriod={handleSelectPeriod} pageMeta={selPeriodPageMeta} onPageChange={setSelPeriodPage}/>
                    <div style={{marginTop:14,paddingTop:14,borderTop:`1px solid ${C.bd}`}}>
                      <div style={{...T.section,marginBottom:10}}>이행 상세</div>
                      <PeriodTable
                        periods={selPeriods}
                        selectedPeriodNum={selPeriodNum}
                        onSelectPeriod={handleSelectPeriod}
                        uploadTargetNum={isAdminUser?periodUploadTarget?.num:null}
                        uploadAllowed={isAdminUser&&periodUploadAllowed&&!measurementSavePending}
                        uploadState={periodUploadState}
                        onUpload={handlePeriodXrfUpload}
                        uploadKey={periodUploadTarget?`${sel.code}|${periodUploadTarget.num}`:""}
                      />
                    </div>
                  </div>
                )}
                {selectedIsDiscontinue?(
                  <div style={{padding:"22px clamp(14px, 2.4vw, 32px)",background:C.card}}>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:10,marginBottom:12}}>
                      {[
                        ["단종 대상 품목", sel.discontinueTargetCode || "—"],
                        ["단종 사유", sel.discontinueReason || "—"],
                        ["최종 사용 예정일", sel.finalUseDate || "—"],
                        ["잔여 재고 처리", sel.stockDisposition || "—"],
                      ].map(([k,v])=>(
                        <div key={k} style={{padding:"10px 12px",background:C.bg,border:`1px solid ${C.bd}`,borderRadius:3}}>
                          <div style={{fontSize:10,color:C.text3,fontWeight:700,marginBottom:4}}>{k}</div>
                          <AutoFitText value={v} title={String(v)} baseFontSize={12} minFontSize={7} align="left" style={{color:C.text1,fontWeight:700}}/>
                        </div>
                      ))}
                    </div>
                    <div style={{padding:"16px 18px",textAlign:"left",border:`1px dashed ${C.bd2}`,borderRadius:8,background:C.alt,color:C.text3,fontSize:12,lineHeight:1.75,whiteSpace:"pre-line"}}>
                      {sel.actionNote || "단종 요청 상세 정보가 없습니다."}
                    </div>
                  </div>
                ):selectedIsMeasuring?(
                  <div style={{padding:"22px clamp(14px, 2.4vw, 32px)",background:C.card}}>
                    <div style={{padding:"34px 20px",textAlign:"center",border:`1px dashed ${C.bd2}`,borderRadius:8,background:C.bg,color:C.text3,fontSize:12,lineHeight:1.7}}>
                      XRF 측정 결과가 아직 업로드되지 않았습니다.<br/>
                      관리자가 결과 파일을 업로드하면 원소별 분석, Final Risk, 재측정 여부가 자동 표시됩니다.
                    </div>
                  </div>
                ):xrfDetailView==="trend"?(
                  <div style={{padding:"18px clamp(14px, 2.4vw, 32px) 26px",background:C.bg}}>
                    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,flexWrap:"wrap",marginBottom:12,padding:"13px 14px",background:"#343A40",borderRadius:UI.rs,boxShadow:"0 8px 20px rgba(24,28,32,.16)"}}>
                      <div style={{display:"flex",alignItems:"center",gap:10,minWidth:0}}>
                        <button type="button" onClick={()=>setXrfDetailView("analysis")} className="pill-btn"
                          aria-label={lang==="en"?"Back to XRF Analysis":"XRF 분석으로 돌아가기"}
                          style={{width:32,height:32,display:"inline-flex",alignItems:"center",justifyContent:"center",border:"1px solid rgba(255,255,255,.32)",background:"rgba(255,255,255,.08)",color:"#fff",fontSize:17,fontWeight:700,cursor:"pointer",flex:"0 0 auto"}}>←</button>
                        <div style={{minWidth:0}}>
                          <div style={{...T.cardTitle,color:"#fff"}}>{lang==="en"?"XRF Trend by Element":"원소별 XRF 트렌드"}</div>
                          <div data-i18n-skip="true" style={{...T.micro,marginTop:3,color:"rgba(255,255,255,.72)"}}>{sel.code} · {displayItemName(sel,lang)}</div>
                        </div>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",fontSize:10.5,color:"rgba(255,255,255,.86)"}}>
                        <span style={{padding:"5px 9px",background:"rgba(255,255,255,.08)",border:"1px solid rgba(255,255,255,.22)",borderRadius:UI.rs}}>{selectedPeriodLabel}</span>
                        <span style={{padding:"5px 9px",background:"rgba(255,255,255,.08)",border:"1px solid rgba(255,255,255,.22)",borderRadius:UI.rs}}>Meas.Date {currentMeasurement?.date||"—"}</span>
                      </div>
                    </div>
                    {currentMeasurement
                      ? <XrfTrendChart
                          item={sel}
                          periods={selectedTrendPeriods}
                          selectedPeriodNum={currentPeriodMeta?.num ?? null}
                          onSelectPeriod={handleSelectPeriod}
                          lang={lang}
                        />
                      : <div style={{background:C.card,border:`1px dashed ${C.bd2}`,borderRadius:14,
                          padding:"38px 18px",textAlign:"center",fontSize:12,color:C.text3,marginTop:12}}>
                          {lang==="en"
                            ? "No XRF measurement is linked to the selected period."
                            : "선택한 주기에 연결된 XRF 측정 이력이 없습니다."}
                        </div>}
                  </div>
                ):(
                  <div className="xrf-detail-grid">
                  <div style={{borderRight:`1px solid ${C.bd}`,display:"flex",flexDirection:"column"}}>
                    <PeriodMeasurementSummary item={sel} periods={selPeriods} selectedPeriodNum={selPeriodNum} currentPeriodNum={currentPeriodMeta?.num} selectedMeasurementId={selMeasurementId} onSelectPeriod={handleSelectPeriod} onSelectMeasurement={handleSelectPeriodMeasurement}/>
                  </div>
                  <div style={{padding:"16px clamp(14px, 2vw, 24px)"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12,gap:16,flexWrap:"wrap"}}>
                      <div>
                        <div style={{...T.section,marginBottom:4}}>
                          원소별 XRF 분석
                          {currentPeriodMeta&&<span style={{...T.caption,marginLeft:8}}>
                            {periodLabel(currentPeriodMeta)}
                          </span>}
                        </div>
                      </div>
                    </div>

                    {/* 측정 요약 */}
                    <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(118px,1fr))",gap:8,marginBottom:12}}>
                      {[
                        ["Meas.Date",    currentMeasurement?.date||"—"],
                        ["Sample Name",  sel.code],
                        ["Measurement Type", currentMeasurement?.role ? `${normalizeMeasurementRole(currentMeasurement.role)}${currentMeasurement?.attemptNo?` #${currentMeasurement.attemptNo}`:""}` : "—"],
                        ["측정 방법",    "by ED-XRF"],
                      ].map(([k,v])=>(
                        <div key={k} style={{padding:"10px 10px",background:C.alt,border:`1px solid ${C.bd}`,borderRadius:8,minWidth:0,textAlign:"center",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
                          <div style={{...T.label,marginBottom:5,whiteSpace:"nowrap",textAlign:"center",width:"100%"}}>{k}</div>
                          <AutoFitText value={v} title={String(v)} baseFontSize={12} minFontSize={7} align="center" style={{fontWeight:600,color:C.text1}}/>
                        </div>
                      ))}
                    </div>

                    <XrfResultSummary measurement={currentMeasurement} item={sel}/>

                    <details style={{padding:"10px 12px",background:C.bg,
                      border:`1px solid ${C.bd}`,borderRadius:UI.rs,marginBottom:12,color:C.text3}}>
                      <summary style={{cursor:"pointer",listStyle:"none",display:"flex",justifyContent:"space-between",gap:12,alignItems:"center"}}>
                        <span style={{...T.section}}>XRF 내부 판정 기준</span>
                        <span data-i18n-skip="true" style={{fontSize:10,color:C.text2,fontWeight:400}}>{lang==="en"?"Internal Limit = 70% of Legal Limit":"Internal Limit = Legal Limit의 70%"}</span>
                      </summary>
                      <div className="responsive-table-scroll table-scroll-compact" role="region" aria-label="XRF 내부 판정 기준표" tabIndex={0}>
                        <table style={{width:"100%",borderCollapse:"collapse",tableLayout:"fixed",fontSize:11,background:C.card,marginTop:8}}>
                          <thead>
                            <tr>
                              {["내부 판정 조건", "처리 기준"].map(h=><th key={h} style={{padding:"7px 8px",border:`1px solid ${C.bd}`,background:C.alt,color:C.text3,fontWeight:800,textAlign:"center"}}>{h}</th>)}
                            </tr>
                          </thead>
                          <tbody>
                            {[
                              [<div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:4}}><Chip v="OK" small result/><span style={{fontSize:10,color:C.text3}}>Content ≤ Internal Limit 또는 N.D.</span></div>, "사용 가능"],
                              [<div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:4}}><Chip v="NG" small result/><span style={{fontSize:10,color:C.text3}}>Content &gt; Internal Limit</span></div>, "정밀분석 인계"],
                              [<Chip v="??" small result/>, lang==="en"?"One XRF remeasurement\n* If the remeasurement result is ?? / NG, proceed to Precision Analysis":"XRF 1회 재측정\n* 재측정 결과가 ?? / NG이면 정밀분석"],
                            ].map(([judgement,action],idx)=>(
                              <tr key={idx}>
                                <td style={{padding:"7px 8px",border:`1px solid ${C.bd}`,textAlign:"center",fontWeight:700,color:C.text1}}>{judgement}</td>
                                <td style={{padding:"7px 8px",border:`1px solid ${C.bd}`,textAlign:"center",color:C.text2,fontWeight:400,whiteSpace:"pre-line",lineHeight:1.7}}>{action}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </details>

                    {/* 원소 바 */}
                    <div style={{background:C.card,border:`1px solid ${C.bd}`,borderRadius:14,overflow:"hidden",boxShadow:"0 10px 24px rgba(26,32,32,.10)"}}>
                      <div style={{padding:"10px 14px",borderBottom:`1px solid ${C.bd}`,background:C.alt,display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,flexWrap:"wrap"}}>
                        <div>
                          <div style={{fontSize:11,fontWeight:700,color:C.text1}}>측정값</div>
                          <div style={{fontSize:9.5,color:C.text3,marginTop:2}}>현재 선택 Measurement 기준</div>
                        </div>
                        <button type="button" onClick={()=>setXrfDetailView("trend")}
                          style={{alignSelf:"flex-start",padding:"7px 10px",display:"inline-flex",alignItems:"center",justifyContent:"center",gap:8,border:"1px solid #343A40",borderRadius:6,background:"#343A40",color:"#fff",fontSize:10.5,fontWeight:650,cursor:"pointer",boxShadow:"none"}}>
                          <span>원소별 트렌드 확인</span><span data-i18n-skip="true" aria-hidden="true" style={{fontSize:13,lineHeight:1,fontWeight:400}}>→</span>
                        </button>
                      </div>
                      {currentMeasurement
                        ? ELEMENTS.map((el,idx)=><EBar key={el} el={el} d={currentMeasurement.elements?.[el]} measurement={currentMeasurement} last={idx===ELEMENTS.length-1}/>)
                        : <div style={{padding:"38px 18px",textAlign:"center",fontSize:12,color:C.text3,fontWeight:600}}>선택한 주기에 연결된 XRF 측정 이력이 없습니다.</div>
                      }
                    </div>

                  </div>
                  </div>
                )}
              </div>
            )}
          </div>
        )}

        {tab==="precision"&&(()=>{
          const selected=precisionRows.find(r=>r.key===precisionSelCaseId) || precisionRows[0] || null;
          const requestNeed=precisionRows.filter(r=>r.precision.requestStatus!=="REQUESTED").length;
          const waiting=precisionRows.filter(r=>r.precision.requestStatus==="REQUESTED"&&r.precision.uploadStatus!=="UPLOADED").length;
          const completed=precisionRows.filter(r=>r.finalConfirm==="확인").length;
          const sent=precisionRows.filter(r=>r.precision.requestStatus==="REQUESTED").length;
          return <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:10}}>{[["정밀분석 대상",precisionRows.length],["인계 필요",requestNeed],["인계 완료",sent],["최종 컨펌",completed]].map(([l,n])=><div key={l} style={{...card,padding:"14px 16px"}}><div style={{...T.label,marginBottom:8}}>{l}</div><div style={{fontSize:24,fontWeight:700,color:C.text1,lineHeight:1,letterSpacing:"-.5px"}}>{n}</div></div>)}</div>
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              <div style={{display:"flex",alignItems:"baseline",gap:10,flexWrap:"wrap"}}>
                <div style={{...T.cardTitle}}>정밀분석 관리</div>
              </div>
              {precisionRows.length===0
                ? <div style={{...card,padding:"44px 20px",textAlign:"center",color:C.text4,fontSize:12}}>정밀분석 대상이 없습니다.</div>
                : precisionRows.map(r=>{
                  const active=selected?.key===r.key;
                  const isRequested=r.precision.requestStatus==="REQUESTED";
                  const hasReport=!!(r.precision.resultFileId||r.precision.reportNo);
                  const isConfirmed=r.finalConfirm==="확인";
                  const result=r.precision.result;   // OK / NG / 미판정
                  const periodTxt=r.caseData?`P${r.caseData.periodNo}`:`P${((itemAllPeriods(r.item)||[]).find(p=>p.measurementId===r.measurement?.id)?.num||"—")}`;

                  // 3단계 진행 상태. 각 단계는 완료 / 진행 / 대기 셋 중 하나입니다.
                  //   완료 = 검정 채움, 진행 = 검정 테두리, 대기 = 회색 테두리
                  const steps=[
                    {label:"인계",   state:isRequested?"done":"active"},
                    {label:"성적서", state:hasReport?"done":(isRequested?"active":"wait")},
                    {label:"컨펌",   state:isConfirmed?"done":(hasReport?"active":"wait")}
                  ];

                  return <div key={r.key} onClick={()=>{setPrecisionSelCaseId(r.key);setSelKey(r.item.code);}}
                    style={{...card,borderColor:active?C.charcoalDk:C.bd,padding:"16px 18px",cursor:"pointer",
                      display:"flex",alignItems:"center",gap:20,flexWrap:"wrap",
                      boxShadow:active?UI.shHover:UI.sh,transition:"box-shadow .15s, border-color .15s"}}>

                    {/* 좌측 · 품목과 인계 사유 */}
                    <div style={{minWidth:0,flex:"1 1 280px"}}>
                      <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",marginBottom:6}}>
                        <span style={{...T.cardTitle}}>{displayItemName(r.item,lang)}</span>
                        <span style={{...T.caption,color:C.text4}}>{r.item.code} · {r.item.dept} · {periodTxt}</span>
                      </div>
                      <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
                        <span style={{...T.caption}}>인계 사유 <strong style={{color:C.text2,fontWeight:600}}>{r.triggerLabel}</strong></span>
                        <span style={{display:"flex",alignItems:"center",gap:5,flexWrap:"wrap"}}>
                          <span style={{...T.caption}}>대상 원소</span>
                          {(r.targets||[]).length
                            ? r.targets.map(el=><span key={el} style={{display:"inline-flex",alignItems:"center",justifyContent:"center",
                                minWidth:28,height:20,padding:"0 8px",borderRadius:UI.pill,background:C.bg,
                                border:`1px solid ${C.bd2}`,color:C.text1,fontSize:11,fontWeight:500}}>{el}</span>)
                            : <span style={{...T.caption,color:C.text4}}>—</span>}
                        </span>
                      </div>
                    </div>

                    {/* 우측 · 진행 단계와 결과 */}
                    <div style={{display:"flex",alignItems:"center",gap:18,marginLeft:"auto",flex:"0 0 auto",flexWrap:"wrap"}}>
                      <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:7}}>
                        <div style={{display:"flex",alignItems:"center"}}>
                          {steps.map((st,i)=>{
                            const done=st.state==="done", act=st.state==="active";
                            return <div key={st.label} style={{display:"flex",alignItems:"center"}}>
                              <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:5,width:52}}>
                                <span style={{width:20,height:20,borderRadius:"50%",display:"inline-flex",alignItems:"center",justifyContent:"center",
                                  fontSize:10,fontWeight:600,flex:"0 0 auto",
                                  background:done?C.charcoalDk:C.card,color:done?"#fff":(act?C.text1:C.text4),
                                  border:`1px solid ${done?C.charcoalDk:(act?C.charcoalDk:C.bd2)}`}}>{done?"✓":i+1}</span>
                                <span style={{fontSize:10,fontWeight:done||act?500:400,color:done||act?C.text2:C.text4,whiteSpace:"nowrap"}}>{st.label}</span>
                              </div>
                              {i<steps.length-1&&<span style={{width:18,height:1,marginBottom:16,background:steps[i+1].state==="wait"?C.bd2:C.charcoalDk}}/>}
                            </div>;
                          })}
                        </div>
                      </div>

                      <div style={{width:1,height:44,background:C.bd}}/>

                      <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:6,minWidth:74}}>
                        <span style={{...T.label}}>정밀 결과</span>
                        {result
                          ? <AnalysisResultBadge value={result}/>
                          : <span title="정밀분석 결과 없음" style={{width:RESULT_LABEL_WIDTH,height:RESULT_LABEL_HEIGHT,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:RESULT_LABEL_FONT_SIZE,color:C.text4,fontWeight:500}}>—</span>}
                      </div>

                      <span onClick={e=>{e.stopPropagation();setPrecisionSelCaseId(r.key);setPrecisionDetailCaseId(r.key);setSelKey(r.item.code);setTab("precision-detail");}}><DetailArrowButton active={precisionDetailCaseId===r.key}/></span>
                    </div>
                  </div>;
                })}
            </div>
          </div>;
        })()}

        {tab==="precision-detail"&&(()=>{
          const selected=precisionRows.find(r=>r.key===precisionDetailCaseId) || precisionRows.find(r=>r.key===precisionSelCaseId) || null;
          return <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,flexWrap:"wrap"}}>
              <div>
                <div style={{...T.cardTitle}}>정밀분석 상세</div>
                <div style={{fontSize:11,color:C.text4,marginTop:3}}>정밀분석 인계, 원본/결과 파일, 원소 함유 확인을 항목별로 관리합니다.</div>
              </div>
              <button onClick={()=>setTab("precision")} className="pill-btn" style={{padding:"8px 15px",background:C.card,border:`1px solid ${C.bd2}`,fontSize:11,fontWeight:500,color:C.text3}}>목록으로</button>
            </div>
            {selected
              ? <PrecisionDetailPanel key={selected.key} selected={selected} lang={lang} card={card} inp={inp} updatePrecisionOverride={updatePrecisionOverride} markPrecisionRequested={markPrecisionRequested} reloadSharePointDb={reloadSharePointDb} onPhotoPreview={setPhotoPreview}/>
              : <div style={{...card,padding:28,textAlign:"center",color:C.text4,fontSize:12}}>선택된 정밀분석 항목이 없습니다.</div>}
          </div>;
        })()}

        {tab==="risk"&&(()=>{
          const active=allItems.filter(i=>i.isCurrent && !isDiscontinueRequestItem(i));
          const matrix={H:{},M:{},L:{}};
          active.forEach(i=>{const x=riskBasisXrfLevelOf(i);const c=riskBasisCrLevelOf(i);if(x&&c)matrix[x][c]=(matrix[x][c]||0)+1;});
          const crCounts={H:active.filter(i=>riskBasisCrLevelOf(i)==="H").length,M:active.filter(i=>riskBasisCrLevelOf(i)==="M").length,L:active.filter(i=>riskBasisCrLevelOf(i)==="L").length};
          const cellColors={H:{bg:C.redBg,c:C.red,l:"사용 불허 Risk"},M:{H:{bg:C.xWarnBg,c:C.xWarn,l:"월 1회"},M:{bg:C.xWarnBg,c:C.xWarn,l:"반기"},L:{bg:C.xWarnBg,c:C.xWarn,l:"반기"}},L:{H:{bg:C.xOkBg,c:C.xOk,l:"반기"},M:{bg:C.xOkBg,c:C.xOk,l:"반기"},L:{bg:C.xOkBg,c:C.xOk,l:"연 1회"}}};
          const highRisk=active.filter(i=>["H","Not Allowed"].includes(itemFinalRisk(i)));
          const workflowCounts=analysisWorkflowSummary(active);
          const cycleCounts={Monthly:active.filter(i=>itemCycleFromRisk(i)==="Monthly").length,Semiannual:active.filter(i=>itemCycleFromRisk(i)==="Semiannual").length,Annual:active.filter(i=>itemCycleFromRisk(i)==="Annual").length};
          const lifecycleCounts={};allItems.forEach(i=>{lifecycleCounts[i.lifecycle]=(lifecycleCounts[i.lifecycle]||0)+1;});
          const precisionActionStages=new Set([
            "PRECISION_TRANSFER_REQUEST_REQUIRED",
            "PRECISION_TRANSFER_RESULT_WAITING",
            "PRECISION_FOLLOWUP_REQUIRED",
            "PRECISION_REQUEST_REQUIRED",
            "PRECISION_RESULT_WAITING",
            "PRECISION_NG_REVIEW",
          ]);
          const attention=active.filter(i=>{
            const finalRisk=itemFinalRisk(i);
            const latestXrf=itemXrfWorst(i);
            const complianceKey=getCS(i);
            // 고위험/후속분석 표의 대상은 다음 세 조건으로만 구성합니다.
            // 1) R Final Risk H/사용불허, 2) 최신 XRF NG, 3) 주기 이행상태 미이행.
            return ["H","Not Allowed"].includes(finalRisk) || latestXrf==="NG" || complianceKey==="overdue";
          }).slice(0,30);
          return <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:10}}>{[["활성 품목",active.length,C.charcoal],["High Risk",highRisk.length,C.red],["주기 미이행",stats.overdue,C.redDk],["후속분석 진행중",workflowCounts.total,C.orange]].map(([l,n,c])=><div key={l} style={{...card,padding:"14px 16px"}}><div style={{fontSize:25,fontWeight:700,color:c,lineHeight:1}}>{n}</div><div style={{fontSize:11,fontWeight:700,color:C.text1,marginTop:5}}>{l}</div></div>)}</div>
            <div className="chart-grid" style={{display:"grid",gridTemplateColumns:"1.15fr 1fr 1fr",gap:14}}>
              <ChartCard title="측정 도래 예정" hint={`${TODAY_STR} 기준 향후 6개월`}>
                <DueLoadChart buckets={dash.buckets} onPick={b=>b.total>0&&gotoDueMonthList(b.key)}/>
              </ChartCard>
              <ChartCard title="이행 상태 구성" hint={`운영 관리 대상 ${dash.csTotal}건`}>
                <ComplianceBar segments={dash.segments} total={dash.csTotal}
                  onPick={s=>s.key!=="notAvailable"&&gotoListWithFilter({compliance:[s.key]})}/>
              </ChartCard>
              <ChartCard title="부서별 위험도 구성" hint="R 단계에서 확정된 Final Risk 기준">
                <DeptRiskChart rows={dash.deptRows} riskMeta={dash.riskMeta}/>
              </ChartCard>
            </div>
            <div className="responsive-table-scroll table-scroll-compact" role="region" aria-label="위험성 평가 현황표" tabIndex={0} style={{...card,padding:"18px 20px"}}><div style={{...T.cardTitle,marginBottom:5}}>위험성 평가 매트릭스 · R 기준</div><div style={{fontSize:11,color:C.text3,marginBottom:14}}>최초 등록 시 R XRF Level과 C&R 등급으로 산정한 위험성 평가 결과입니다. 후속 정밀분석 승인 결과와는 별도로 유지합니다.</div><table style={{borderCollapse:"collapse",width:"100%",fontSize:12}}><thead><tr><th style={{padding:"8px 12px",background:C.bg,border:`1px solid ${C.bd}`,textAlign:"center",fontSize:11,color:C.text3}}>XRF Level ╲ C&R</th>{[["H",`${crCounts.H}건 · 직접+잔류`],["M",`${crCounts.M}건 · 직접+비잔류`],["L",`${crCounts.L}건 · 비접촉`]].map(([c,d])=><th key={c} style={{padding:"8px 12px",background:C.bg,border:`1px solid ${C.bd}`,textAlign:"center",fontSize:11,color:C.text3}}>{c}<br/><span style={{fontSize:9,color:C.text4,fontWeight:400}}>{d}</span></th>)}</tr></thead><tbody>{["H","M","L"].map(x=><tr key={x}><td style={{padding:"10px 12px",border:`1px solid ${C.bd}`,fontWeight:600,color:C.text2,background:C.bg,fontSize:11}}><span style={{color:x==="H"?C.red:x==="M"?C.xWarn:C.xOk}}>●</span> {x==="H"?"H · Exceeded":x==="M"?"M · Less than":"L · N.D."}</td>{["H","M","L"].map(c=>{const cnt=matrix[x]?.[c]||0;const cs=x==="H"?cellColors.H:cellColors[x][c];return <td key={c} onClick={()=>cnt>0&&gotoListWithFilter({xrf:[x],crLevel:[c]})} style={{padding:"12px",border:`1px solid ${C.bd}`,background:cnt?cs.bg:C.card,textAlign:"center",cursor:cnt?"pointer":"default"}}><div style={{fontSize:20,fontWeight:700,color:cnt?cs.c:C.text4}}>{cnt}</div><div style={{fontSize:9,color:cnt?cs.c:C.text4,marginTop:3}}>{cs.l}</div></td>})}</tr>)}</tbody></table></div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <div style={{...card,padding:"18px 20px"}}><div style={{...T.cardTitle,marginBottom:12}}>주기 관리 현황</div>{[["월 1회",cycleCounts.Monthly],["반기 1회",cycleCounts.Semiannual],["연 1회",cycleCounts.Annual],["미이행",stats.overdue],["측정권장 (월 D-15 / 반기·연 D-30)",stats.dueSoon],["현재 이행",active.filter(i=>getCS(i)==="ok").length]].map(([l,n])=><div key={l} style={{display:"flex",justifyContent:"space-between",padding:"8px 10px",borderBottom:`1px solid ${C.bd}`,fontSize:11}}><span style={{color:C.text2}}>{l}</span><strong style={{color:l==="미이행"?C.red:C.charcoal}}>{n}건</strong></div>)}</div>
              <div style={{...card,padding:"18px 20px"}}><div style={{...T.cardTitle,marginBottom:12}}>분석 Workflow 현황</div>{[["XRF 재측정",workflowCounts.xrfRemeasure,"xrf"],["정밀분석 인계 필요",workflowCounts.precisionTransferRequired,"precision"],["정밀분석 결과 대기",workflowCounts.precisionResultWaiting,"precision"],["정밀분석 후속조치 필요",workflowCounts.precisionFollowupRequired,"precision"],["정밀분석 NG · 반려",workflowCounts.precisionNgReview,"precision"]].map(([l,n,target])=><div key={l} onClick={()=>{if(n){setTab(target);}}} style={{display:"flex",justifyContent:"space-between",padding:"8px 10px",borderBottom:`1px solid ${C.bd}`,fontSize:11,cursor:n?"pointer":"default",background:n?C.alt:"transparent"}}><span style={{color:n?C.text2:C.text4}}>{l}</span><strong style={{color:n?C.orange:C.text4}}>{n}건</strong></div>)}</div>
            </div>
            <div style={{...card,padding:"18px 20px"}}>
              <div style={{display:"flex",alignItems:"flex-end",justifyContent:"space-between",gap:12,marginBottom:5,flexWrap:"wrap"}}>
                <div style={{...T.cardTitle}}>고위험 / 후속분석 대상 품목</div>
              </div>
              <div className="responsive-table-scroll table-scroll-medium" role="region" aria-label="고위험 및 후속분석 대상 품목 표" tabIndex={0}>
                <table style={{width:"100%",minWidth:0,maxWidth:"100%",borderCollapse:"collapse",tableLayout:"fixed"}}>
                  <colgroup>
                    <col style={{width:"11%"}}/><col style={{width:"18%"}}/><col style={{width:"11%"}}/><col style={{width:"11%"}}/><col style={{width:"10%"}}/><col style={{width:"11%"}}/><col style={{width:"18%"}}/><col style={{width:"10%"}}/>
                  </colgroup>
                  <thead>
                    <tr>{["품번","품목명","R Final Risk","주기 이행 상태","최신 XRF","정밀분석 결과","현재 조치","승인 상태"].map(h=><th key={h} style={{padding:"8px 5px",background:C.bg,border:`1px solid ${C.bd}`,fontSize:10.5,fontWeight:600,color:C.text2,textAlign:"center",whiteSpace:"normal",lineHeight:1.35}}>{h}</th>)}</tr>
                  </thead>
                  <tbody>
                    {attention.length?attention.map((it,idx)=>{
                      const wf=workflowForItem(it);
                      const complianceKey=getCS(it);
                      const precisionLinked=!!(itemHasPrecisionTrigger(it) || wf.precision || wf.precisionCases?.length || precisionActionStages.has(wf.stage));
                      const currentAction=(()=>{
                        // 분석 Workflow가 아직 끝나지 않았다면 그 조치를 가장 먼저 보여 줍니다.
                        if(wf.stage && wf.stage!=="APPROVED"){
                          return {
                            label:workflowStageLabel(wf.stage),
                            target:precisionActionStages.has(wf.stage)?"precision":"xrf"
                          };
                        }
                        // 분석이 완료된 뒤에도 Pn 측정 의무는 별도입니다.
                        // 따라서 미이행/측정권장 상태를 "처리 완료"로 덮어쓰지 않습니다.
                        if(complianceKey==="overdue") return {label:"주기 측정 필요",target:"xrf"};
                        if(complianceKey==="dueSoon") return {label:"주기 측정 권장",target:"xrf"};
                        return {label:"후속조치 없음",target:"xrf"};
                      })();
                      const openPrecision=(e)=>{e.stopPropagation(); precisionLinked?gotoPrecisionDetail(it):gotoXrf(it.code);};
                      const openCurrentAction=(e)=>{
                        e.stopPropagation();
                        currentAction.target==="precision" && precisionLinked ? gotoPrecisionDetail(it) : gotoXrf(it.code);
                      };
                      const openXrf=(e)=>{e.stopPropagation(); gotoXrf(it.code);};
                      return <tr key={it.code} className="risk-attention-row" tabIndex={0} role="button"
                        title="클릭하여 품목 상세 열기"
                        onClick={()=>gotoItemDetail(it)}
                        onKeyDown={e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();gotoItemDetail(it);}}}
                        style={{background:idx%2?C.alt:C.card}}>
                        <td onClick={openXrf} title="XRF 분석 상세로 이동" style={{padding:"9px 5px",border:`1px solid ${C.bd}`,fontSize:10.5,fontWeight:700,color:C.text1,overflow:"hidden"}}><AutoFitText value={it.code} title={it.code} baseFontSize={10.5} minFontSize={7} align="center" style={{fontWeight:700,color:C.text1}}/></td>
                        <td onClick={openXrf} title="XRF 분석 상세로 이동" style={{padding:"9px 5px",border:`1px solid ${C.bd}`,fontSize:10.5,fontWeight:500,color:C.text1,overflow:"hidden"}}><AutoFitText value={displayItemName(it,lang)} title={displayItemName(it,lang)} baseFontSize={10.5} minFontSize={7} align="center" style={{fontWeight:500,color:C.text1}}/></td>
                        <td onClick={openXrf} title="R 기준 위험도와 XRF 상세로 이동" style={{padding:"9px",border:`1px solid ${C.bd}`,textAlign:"center"}}><RiskBadge risk={itemFinalRisk(it)}/></td>
                        <td onClick={openXrf} title="해당 품목의 이행/XRF 상세로 이동" style={{padding:"9px",border:`1px solid ${C.bd}`,fontSize:11,textAlign:"center"}}><CompText item={it}/></td>
                        <td onClick={openXrf} title="XRF 결과 상세로 이동" style={{padding:"9px",border:`1px solid ${C.bd}`,textAlign:"center"}}><Chip v={itemXrfWorst(it)} small result/></td>
                        <td onClick={openPrecision} title={precisionLinked?"정밀분석 상세로 이동":"정밀분석 대상이 없어 XRF 상세로 이동"} style={{padding:"9px",border:`1px solid ${C.bd}`,textAlign:"center"}}><AnalysisResultBadge value={precisionAnalysisResultOf(it)}/></td>
                        <td onClick={openCurrentAction} title={currentAction.target==="precision"?"현재 조치의 정밀분석 상세로 이동":"현재 조치의 주기/XRF 상세로 이동"} style={{padding:"9px 5px",border:`1px solid ${C.bd}`,fontSize:10.5,fontWeight:500,color:C.text2,textAlign:"center",overflow:"hidden",lineHeight:1.35}}><AutoFitText value={currentAction.label} title={currentAction.label} baseFontSize={10.5} minFontSize={7} align="center" style={{fontWeight:500,color:C.text2,lineHeight:1.35}}/></td>
                        <td onClick={e=>{e.stopPropagation();gotoItemDetail(it);}} title="현재 Workflow 상세로 이동" style={{padding:"9px",border:`1px solid ${C.bd}`,textAlign:"center"}}><ApprovalBadge status={approvalStatusOf(it)}/></td>
                      </tr>;
                    }):<tr><td colSpan={8} style={{padding:24,textAlign:"center",color:C.text4}}>현재 조치 대상이 없습니다.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
            {replacementChains.length>0&&<div style={{...card,padding:"16px 20px"}}>
              <div data-i18n-skip="true" style={{...T.cardTitle,fontFamily:FONT_SANS,marginBottom:4}}>{lang==="en"?`Replacement Chains — ${replacementChains.length} items`:`대체 관계 체인 — ${replacementChains.length}건`}</div>
              <div data-i18n-skip="true" style={{fontFamily:FONT_SANS,fontSize:11,fontWeight:400,lineHeight:1.45,color:C.text3,marginBottom:14}}>{lang==="en"?"History of existing items replaced by new items":"기존 품목이 신규 항목으로 대체된 이력"}</div>
              {replacementChains.map((chain,i)=>(
                <div key={`${chain.old.code}-${i}`} style={{display:"flex",alignItems:"center",gap:8,padding:"10px 0",borderTop:i>0?`1px solid ${C.bd}`:"none",flexWrap:"wrap"}}>
                  <div onClick={()=>gotoXrf(chain.old.code)} title={`${chain.old.code} XRF 상세로 이동`}
                    style={{fontFamily:FONT_SANS,padding:"8px 12px",background:C.alt,border:`1px dashed ${C.bd2}`,borderRadius:6,minWidth:180,cursor:"pointer",lineHeight:1.4}}>
                    <div data-i18n-skip="true" style={{fontSize:10,color:C.text4}}>{lang==="en"?`Existing Measurements: ${chain.old.measurementCount||chain.old.history?.length||0} preserved · Transition ${chain.old.replacementDate||chain.old.finalUseDate||"—"}`:`기존 측정이력 ${chain.old.measurementCount||chain.old.history?.length||0}건 보존 · 전환 ${chain.old.replacementDate||chain.old.finalUseDate||"—"}`}</div>
                    <div style={{fontSize:12,fontWeight:500,color:C.text2,marginTop:2,fontFamily:FONT_SANS}}>{chain.old.code}</div>
                    <div data-i18n-skip="true" style={{fontSize:12,color:C.text2}}>{displayItemName(chain.old,lang)}</div>
                    <div data-i18n-skip="true" style={{fontSize:10,color:chain.old.lifecycle==="Discontinued"?C.redDk:C.text4,marginTop:3}}>{lang==="en"?translateUiTextKoToEn(LIFECYCLE_KO[chain.old.lifecycle]||chain.old.lifecycle||"—"):(LIFECYCLE_KO[chain.old.lifecycle]||chain.old.lifecycle||"—")}</div>
                  </div>
                  <span aria-hidden="true" style={{fontSize:16,color:C.text4}}>→</span>
                  <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                    {chain.successors.map(s=>(
                      <div key={s.code} onClick={()=>gotoXrf(s.code)} title={`${s.code} XRF 상세로 이동`}
                        style={{fontFamily:FONT_SANS,padding:"8px 12px",background:C.charcoalBg,border:`1px solid ${C.bd}`,borderRadius:6,cursor:"pointer",minWidth:170,lineHeight:1.4}}>
                        <div data-i18n-skip="true" style={{fontSize:10,color:C.charcoalMd}}>{lang==="en"?`Replacement Item · ${translateUiTextKoToEn(LIFECYCLE_KO[s.lifecycle]||"신규 대체")} · ${s.replacementDate||s.firstDate||"—"}`:`대체품 · ${LIFECYCLE_KO[s.lifecycle]||"신규 대체"} · ${s.replacementDate||s.firstDate||"—"}`}</div>
                        <div style={{fontSize:12,fontWeight:500,color:C.charcoal,marginTop:2,fontFamily:FONT_SANS}}>{s.code}</div>
                        <div data-i18n-skip="true" style={{fontSize:12,color:C.text1}}>{displayItemName(s,lang)}</div>
                        <div data-i18n-skip="true" style={{fontSize:10,color:C.text4,marginTop:3}}>{lang==="en"?`Replaces Existing Item ${s.replacementOf||chain.old.code}`:`기존 품목 ${s.replacementOf||chain.old.code} 대체`}</div>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>}
            <div style={{...card,padding:"16px 20px"}}><div style={{fontSize:12,fontWeight:700,color:C.text1,marginBottom:10}}>품목 라이프사이클 참고</div><div style={{display:"flex",gap:18,flexWrap:"wrap"}}>{[["기존 운영","ExistingActive"],["신규 대체","NewReplacement"],["대체됨","ReplacedOld"],["단종","Discontinued"]].map(([l,k])=><span key={k} style={{fontSize:11,color:C.text3}}>{l} <strong style={{color:C.text1}}>{lifecycleCounts[k]||0}건</strong></span>)}</div></div>
          </div>;
        })()}

        {tab==="reg"&&(
          <div style={{maxWidth:960,margin:"0 auto"}}>
            <div style={{...card,padding:0,overflow:"hidden"}}>
              <div style={{padding:"18px 22px",borderBottom:`1px solid ${C.bd}`,display:"flex",justifyContent:"space-between",gap:12,alignItems:"center",flexWrap:"wrap"}}>
                <div>
                  <div style={{...T.pageTitle,marginBottom:5}}>품목 / 변경 관리</div>
                </div>
                <div style={{display:"flex",gap:6,background:C.bg,border:`1px solid ${C.bd}`,padding:3}}>
                  {[['requester','의뢰자 등록'],['admin','관리자 직접 등록']].map(([m,l])=>(
                    <button key={m} onClick={()=>{setRegMode(m);setRegStep(1);setRegType(null);setReqUploadResult(emptyUploadedXrfResult());setReqUploadFileName("");clearRequestPhoto();}} style={{padding:"6px 12px",background:regMode===m?C.charcoalDk:C.card,color:regMode===m?'#fff':C.text2,border:"none",fontSize:12,fontWeight:600,cursor:"pointer"}}>{l}</button>
                  ))}
                </div>
              </div>

              {regMode==="requester"?(
                <div style={{padding:22}}>
                  <div style={{display:"flex",marginBottom:20}}>
                    {requesterStepLabels.map((s,i)=>(
                      <div key={s} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center"}}>
                        <div style={{display:"flex",width:"100%",alignItems:"center"}}>
                          <div style={{flex:1,height:1,background:i===0?"transparent":regStep>i?C.charcoal:C.bd}}/>
                          <div style={{width:26,height:26,borderRadius:"50%",border:`2px solid ${regStep>i?C.charcoal:regStep===i+1?C.red:C.bd}`,background:regStep>i?C.charcoal:C.card,color:regStep>i?"#fff":regStep===i+1?C.red:C.text4,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:700,fontSize:12}}>{regStep>i?"✓":i+1}</div>
                          <div style={{flex:1,height:1,background:i===requesterStepLabels.length-1?"transparent":regStep>i+1?C.charcoal:C.bd}}/>
                        </div>
                        <div style={{fontSize:11,color:regStep===i+1?C.text1:C.text4,marginTop:6,fontWeight:regStep===i+1?700:500,textAlign:"center"}}>{s}</div>
                      </div>
                    ))}
                  </div>

                  {regStep===1&&<div>
                    <div style={{fontSize:10,fontWeight:700,color:C.charcoalLt,letterSpacing:.6,textTransform:"uppercase",marginBottom:12}}>요청 유형 선택</div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(3,minmax(0,1fr))",gap:12}}>
                      {[
                        ["new","신규 도입","기존 관리 품목과 관계없이 새로운 부자재/공정부자재를 등록합니다."],
                        ["replacement","기존 품목 대체 등록","현재 사용 중인 품목을 새로운 품목으로 변경하거나 대체합니다."],
                        ["discontinue","기존 품목 단종 / 사용 중지","신규 품목 추가 없이 기존 품목의 사용 중지 또는 단종을 요청합니다."]
                      ].map(([key,title,desc])=>(
                        <button key={key} onClick={()=>{setRegType(key);setRegStep(2);setReqUploadResult(emptyUploadedXrfResult());setReqUploadFileName("");clearRequestPhoto();}}
                          style={{textAlign:"left",padding:"18px 18px",border:`2px solid ${regType===key?C.red:C.bd}`,background:regType===key?C.redBg:C.card,borderRadius:10,cursor:"pointer",minHeight:150}}>
                          <div style={{fontSize:14,fontWeight:800,color:C.text1,marginBottom:8}}>{title}</div>
                          <div style={{fontSize:11,color:C.text3,lineHeight:1.7}}>{desc}</div>
                        </button>
                      ))}
                    </div>
                  </div>}

                  {regStep===2&&regType!=="discontinue"&&<div>
                    <div style={{fontSize:10,fontWeight:700,color:C.charcoalLt,letterSpacing:.6,textTransform:"uppercase",marginBottom:12}}>
                      {regType==="replacement"?"대체 대상 및 신규 대체품 기본정보":"신규 도입 기본정보"}
                    </div>

                    {regType==="replacement"&&<div style={{padding:"12px",background:C.bg,border:`1px solid ${C.bd}`,borderRadius:8,marginBottom:14}}>
                      <div style={{fontSize:12,fontWeight:700,color:C.text1,marginBottom:8}}>대체 대상 품목 선택 <span style={{color:C.red}}>*</span></div>
                      <select data-i18n-skip="true" value={reqForm.replacementTargetCode} onChange={e=>updateRequestField("replacementTargetCode",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box",marginBottom:10}}>
                        <option value="">대체할 기존 품목을 선택하세요</option>
                        {selectableExistingItems.map(item=><option key={item.code} value={item.code}>{item.code} | {item.dept} | {displayItemName(item,lang)}</option>)}
                      </select>
                      {selectedReplacementTarget&&<div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:8}}>
                        {[
                          ["대상 품목",`${selectedReplacementTarget.code} · ${displayItemName(selectedReplacementTarget,lang)}`],
                          ["부서",selectedReplacementTarget.dept],
                          ["현재 XRF",itemXrfWorst(selectedReplacementTarget)],
                          ["현재 주기",CYCLE_KO[itemCycleFromRisk(selectedReplacementTarget)]||"—"]
                        ].map(([k,v])=><div key={k} style={{background:C.card,border:`1px solid ${C.bd}`,padding:"8px 9px",fontSize:11,minWidth:0}}><div style={{color:C.text4,marginBottom:4}}>{k}</div><AutoFitText value={v} title={String(v)} baseFontSize={11} minFontSize={7} align="left" style={{fontWeight:600,color:C.text1}}/></div>)}
                      </div>}
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginTop:10}}>
                        <div>
                          <div style={{fontSize:12,fontWeight:600,marginBottom:5}}>대체 사유</div>
                          <select value={reqForm.replacementReason} onChange={e=>updateRequestField("replacementReason",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}><option>기존 품목 단종</option><option>공급업체 변경</option><option>제조사 변경</option><option>사양 변경</option><option>품질 개선</option><option>원가 절감</option><option>고객 요구</option><option>기타</option></select>
                          {reqForm.replacementReason==="기타"&&<textarea value={reqForm.replacementReasonOther} onChange={e=>updateRequestField("replacementReasonOther",e.target.value)} placeholder="기타 대체 사유를 구체적으로 입력하세요" style={{...inp,width:"100%",boxSizing:"border-box",minHeight:64,resize:"vertical",marginTop:8,borderColor:reqForm.replacementReasonOther.trim()?C.bd:C.red}}/>}
                        </div>
                        <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>기존 품목 처리 방식</div><select value={reqForm.oldItemDisposition} onChange={e=>updateRequestField("oldItemDisposition",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}><option>신규 품목 승인 후 기존 재고 소진 후 단종</option><option>승인 후 즉시 단종</option><option>신규 품목 승인 후 병행 사용</option><option>기존 품목 유지, 신규 품목 추가 등록만 진행</option></select></div>
                        <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>전환 예정일</div><input type="date" value={reqForm.replacementDate} onChange={e=>updateRequestField("replacementDate",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}/></div>
                        <div style={{gridColumn:"1 / -1",padding:"10px 12px",background:C.charcoalBg,border:`1px solid ${C.bd}`,fontSize:11,color:C.text3,lineHeight:1.65}}>
                          신규 대체품 품번은 선택한 기존 품번 뒤에 순번을 붙여 자동 생성합니다. 예: 기존 품번 뒤에 -01, -02 순번을 추가합니다. 기존 품목은 변경·대체 이력으로 전환되고, 신규 대체품은 신규 등록 품목으로 주기 관리를 시작합니다.
                        </div>
                      </div>
                    </div>}

                    <div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:12}}>
                      {[
                        ["partNumber","Part Number","자동 생성"],
                        ["itemName",regType==="replacement"?"신규 대체품 Item_Name":"Item_Name","예: 제품 포장용 엠보스"],
                        ["itemNameEn","Item_Name_EN","e.g. Packaging Emboss"],
                        ["manufacturer","제조사","예: ABC Chemical"],
                        ["useDate","사용 예정 일자",TODAY_STR]
                      ].map(([k,l,ph])=>(
                        <div key={k}><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>{l} <span style={{color:C.red}}>*</span></div><input type={k==="useDate"?"date":"text"} value={reqForm[k]} onChange={e=>updateRequestField(k,e.target.value)} placeholder={ph} readOnly={k==="partNumber"} style={{...inp,width:"100%",boxSizing:"border-box",background:k==="partNumber"?C.bg:C.card,color:k==="partNumber"?C.text3:C.text1}}/></div>
                      ))}
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Type <span style={{color:C.red}}>*</span></div><select value={reqForm.type} onChange={e=>updateRequestField("type",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}>{ITEM_TYPE_OPTIONS.map(v=><option key={v} value={v}>{ITEM_TYPE_KO[v]} ({v})</option>)}</select></div>
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Handling_Dept <span style={{color:C.red}}>*</span></div><select value={reqForm.handlingDept} onChange={e=>updateRequestField("handlingDept",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}>{DEPTS.map(d=><option key={d}>{d}</option>)}</select></div>
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Material_Category <span style={{color:C.red}}>*</span></div><select value={reqForm.materialCategory} onChange={e=>updateRequestField("materialCategory",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}>{REQUESTER_MATERIAL_CATEGORY_OPTIONS.map(v=><option key={v}>{v}</option>)}</select></div>
                      <div>
                        <div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Material_State {isChemicalMaterialCategory(reqForm.materialCategory)&&<span style={{color:C.red}}>*</span>}</div>
                        {isChemicalMaterialCategory(reqForm.materialCategory)?(
                          <select value={reqForm.materialState} onChange={e=>updateRequestField("materialState",e.target.value)} required style={{...inp,width:"100%",boxSizing:"border-box",borderColor:reqForm.materialState?C.bd:C.red}}>
                            <option value="">Material_State를 선택하세요</option>
                            {CHEMICAL_MATERIAL_STATE_OPTIONS.map(v=><option key={v} value={v}>{v}</option>)}
                          </select>
                        ):(
                          <input value={reqForm.materialState} onChange={e=>updateRequestField("materialState",e.target.value)} placeholder="예: Solid / Powder" style={{...inp,width:"100%",boxSizing:"border-box"}}/>
                        )}
                      </div>
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>CR_Type <span style={{color:C.red}}>*</span></div><select value={reqForm.crType} onChange={e=>updateRequestField("crType",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}><option>직접접촉+잔류</option><option>직접접촉+비잔류</option><option>비접촉</option></select></div>
                      <div style={{gridColumn:"1 / -1"}}>
                        <div style={{fontSize:12,fontWeight:600,marginBottom:5}}>품목 사진 <span style={{color:C.red}}>*</span></div>
                        <div style={{display:"grid",gridTemplateColumns:reqPhotoPreview?"120px 1fr":"1fr",gap:12,alignItems:"stretch"}}>
                          {reqPhotoPreview&&<img src={reqPhotoPreview} alt="선택한 품목 사진" style={{width:120,height:96,objectFit:"cover",borderRadius:8,border:`1px solid ${C.bd}`}}/>}
                          <label style={{...uploadBoxStyle,minHeight:96,borderRadius:8}}>
                            <span data-i18n-skip="true">{reqPhotoFile?.name||(lang==="en"?"Choose Photo (JPG, PNG, WEBP / max. 10 MB)":"사진 선택 (JPG, PNG, WEBP / 최대 10MB)")}</span>
                            <input type="file" accept="image/jpeg,image/png,image/webp,.jpg,.jpeg,.png,.webp" onChange={handleRequestPhoto} style={fileInputOverlay}/>
                          </label>
                        </div>
                      </div>
                      <div style={{gridColumn:"1 / -1"}}><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>요청 사유 / 비고</div><textarea value={reqForm.requestReason} onChange={e=>updateRequestField("requestReason",e.target.value)} placeholder="신규 도입 또는 대체 등록 사유를 입력하세요" style={{...inp,width:"100%",boxSizing:"border-box",minHeight:74,resize:"vertical"}}/></div>
                    </div>
                  </div>}

                  {regStep===2&&regType==="discontinue"&&<div>
                    <div style={{fontSize:10,fontWeight:700,color:C.charcoalLt,letterSpacing:.6,textTransform:"uppercase",marginBottom:12}}>단종 / 사용 중지 정보 입력</div>
                    <div style={{padding:"12px",background:C.bg,border:`1px solid ${C.bd}`,borderRadius:8,marginBottom:14}}>
                      <div style={{fontSize:12,fontWeight:700,color:C.text1,marginBottom:8}}>단종 대상 품목 선택 <span style={{color:C.red}}>*</span></div>
                      <select data-i18n-skip="true" value={reqForm.discontinueTargetCode} onChange={e=>updateRequestField("discontinueTargetCode",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box",marginBottom:10}}>
                        <option value="">단종 또는 사용 중지할 기존 품목을 선택하세요</option>
                        {selectableExistingItems.map(item=><option key={item.code} value={item.code}>{item.code} | {item.dept} | {displayItemName(item,lang)}</option>)}
                      </select>
                      {selectedDiscontinueTarget&&<div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:8}}>
                        {[
                          ["대상 품목",`${selectedDiscontinueTarget.code} · ${displayItemName(selectedDiscontinueTarget,lang)}`],
                          ["부서",selectedDiscontinueTarget.dept],
                          ["Lifecycle",LIFECYCLE_KO[selectedDiscontinueTarget.lifecycle]||selectedDiscontinueTarget.lifecycle],
                          ["Approval",approvalStatusOf(selectedDiscontinueTarget)]
                        ].map(([k,v])=><div key={k} style={{background:C.card,border:`1px solid ${C.bd}`,padding:"8px 9px",fontSize:11,minWidth:0}}><div style={{color:C.text4,marginBottom:4}}>{k}</div><AutoFitText value={v} title={String(v)} baseFontSize={11} minFontSize={7} align="left" style={{fontWeight:600,color:C.text1}}/></div>)}
                      </div>}
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:12}}>
                      <div>
                        <div style={{fontSize:12,fontWeight:600,marginBottom:5}}>단종 / 사용 중지 사유</div>
                        <select value={reqForm.discontinueReason} onChange={e=>updateRequestField("discontinueReason",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}><option>사용 중지</option><option>공정 변경</option><option>고객 요구</option><option>공급 중단</option><option>대체품 적용 완료</option><option>기타</option></select>
                        {reqForm.discontinueReason==="기타"&&<textarea value={reqForm.discontinueReasonOther} onChange={e=>updateRequestField("discontinueReasonOther",e.target.value)} placeholder="기타 단종 사유를 구체적으로 입력하세요" style={{...inp,width:"100%",boxSizing:"border-box",minHeight:64,resize:"vertical",marginTop:8,borderColor:reqForm.discontinueReasonOther.trim()?C.bd:C.red}}/>}
                      </div>
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>최종 사용 예정일</div><input type="date" value={reqForm.finalUseDate} onChange={e=>updateRequestField("finalUseDate",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}/></div>
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>잔여 재고 처리 방식</div><select value={reqForm.stockDisposition} onChange={e=>updateRequestField("stockDisposition",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}><option>재고 소진 후 사용 중지</option><option>폐기</option><option>반품</option><option>별도 보관</option><option>기타</option></select></div>
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>요청 부서</div><select value={reqForm.handlingDept} onChange={e=>updateRequestField("handlingDept",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}>{DEPTS.map(d=><option key={d}>{d}</option>)}</select></div>
                      <div style={{gridColumn:"1 / -1"}}><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>비고</div><textarea value={reqForm.note} onChange={e=>updateRequestField("note",e.target.value)} placeholder="단종/사용 중지 관련 특이사항을 입력하세요" style={{...inp,width:"100%",boxSizing:"border-box",minHeight:74,resize:"vertical"}}/></div>
                    </div>
                  </div>}

                  {regStep===3&&regType!=="discontinue"&&<div>
                    <div style={{fontSize:10,fontWeight:700,color:C.charcoalLt,letterSpacing:.6,textTransform:"uppercase",marginBottom:12}}>XRF 방식 선택</div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:14}}>
                      <div onClick={()=>setReqXrfMode("upload")} style={{border:`2px solid ${reqXrfMode==="upload"?C.red:C.bd}`,borderRadius:8,padding:"18px 20px",background:reqXrfMode==="upload"?C.redBg:C.card,cursor:"pointer"}}>
                        <div style={{fontSize:14,fontWeight:700,marginBottom:8}}>방식 A — XRF 사전 완료</div>
                        <div style={{fontSize:12,color:C.text3,lineHeight:1.75}}>XRF 원본 보고서가 이미 있을 때<br/>파일 업로드 → 값 추출 → 즉시 Risk 판정</div>
                        <div style={{marginTop:12,border:`1px dashed ${C.bd2}`,padding:"16px",textAlign:"center",background:C.bg,fontSize:12,color:C.text4}}>아래 업로드 영역에서 파일 선택</div>
                      </div>
                      <div onClick={()=>setReqXrfMode("request")} style={{border:`2px solid ${reqXrfMode==="request"?C.red:C.bd}`,borderRadius:8,padding:"18px 20px",background:reqXrfMode==="request"?C.redBg:C.card,cursor:"pointer"}}>
                        <div style={{fontSize:14,fontWeight:700,marginBottom:8}}>방식 B — XRF 측정 의뢰</div>
                        <div style={{fontSize:12,color:C.text3,lineHeight:1.75}}>보고서가 없을 때<br/>의뢰서 작성 → 관리자 알림 → 측정 완료 후 판정</div>
                        <div style={{marginTop:12,padding:"16px",textAlign:"center",background:C.bg,fontSize:12,color:C.text4}}>Approval_Status = 보류</div>
                      </div>
                    </div>
                    {reqXrfMode==="upload"&&<div style={{display:"grid",gridTemplateColumns:"1.2fr .8fr",gap:12,padding:"12px",background:C.bg,border:`1px solid ${C.bd}`}}>
                      <div>
                        <div style={{fontSize:11,fontWeight:600,marginBottom:6}}>XRF 원본 Excel 업로드</div>
                        <label style={uploadBoxStyle}>
                          {reqUploadFileName || "XRF 파일 선택 (.xlsx 또는 파싱 결과 .json)"}
                          <input type="file" accept=".xlsx,.xlsm,.xls,.csv,.json" onChange={handleRequesterXrfUpload} style={fileInputOverlay}/>
                        </label>
                        <div style={{fontSize:10,color:C.text4,lineHeight:1.55,marginTop:6}}>업로드한 XRF Report의 Content(ppm), Std.Deviation(ppm), Judgment를 자동 추출합니다. XRF OK/NG는 Content를 원소별 Internal Limit(법적 기준의 70%)과 비교해 산정하고, 원본 Judgment는 ??/Cannot Judge 재측정 판단과 추적용으로 보존합니다.</div>
                      </div>
                      <div style={{display:"grid",gridTemplateColumns:"1fr",gap:7}}>
                        <div style={{padding:"8px 10px",background:C.card,border:`1px solid ${C.bd}`}}>
                          <div style={{fontSize:10,color:C.text4,marginBottom:4}}>자동 산출 XRF</div>
                          {reqUploadResult.parsed?<Chip v={reqUploadResult.xrfWorst||"—"} small result/>:<span style={{fontSize:11,color:C.text4}}>{reqUploadResult.pending?"읽는 중":"파싱 대기"}</span>}
                        </div>
                        <div style={{padding:"8px 10px",background:C.card,border:`1px solid ${C.bd}`}}>
                          <div style={{fontSize:10,color:C.text4,marginBottom:4}}>후속조치</div>
                          {reqUploadResult.parsed?<XrfFollowupBadge measurement={{elements:reqUploadResult.elements||{}}}/>:<span style={{fontSize:11,color:C.text4}}>—</span>}
                        </div>
                        <div style={{padding:"8px 10px",background:C.card,border:`1px solid ${C.bd}`}}>
                          <div style={{fontSize:10,color:C.text4,marginBottom:4}}>예상 Approval_Status</div>
                          <ApprovalBadge status={reqUploadResult.parsed?(reqUploadResult.approvalStatus||deriveXrfApprovalStatus(reqUploadResult.xrfResult||reqUploadResult.xrfWorst)):"보류"}/>
                        </div>
                      </div>
                      <div style={{gridColumn:"1 / -1"}}><UploadedXrfPreview result={reqUploadResult}/></div>
                      <div style={{gridColumn:"1 / -1",fontSize:11,color:C.text4,lineHeight:1.7}}>브라우저에서 원본 보고서의 O5 측정일과 Element / Content / Std.Deviation / Judgment 행을 직접 읽습니다. 등록 후에는 같은 결과를 SharePoint XRF_Measurements와 XRF_ElementResults에 저장하도록 연결합니다.</div>
                    </div>}
                  </div>}

                  {((regType==="discontinue"&&regStep===3)||(regType!=="discontinue"&&regStep===4))&&<div>
                    <div style={{fontSize:10,fontWeight:700,color:C.charcoalLt,letterSpacing:.6,textTransform:"uppercase",marginBottom:12}}>요청 요약</div>
                    <div style={{background:C.bg,border:`1px solid ${C.bd}`,borderRadius:4,padding:"12px 14px",marginBottom:14}}>
                      {(regType==="discontinue"?[
                        ["요청 유형","기존 품목 단종 / 사용 중지"],
                        ["단종 대상",selectedDiscontinueTarget?`${selectedDiscontinueTarget.code} · ${displayItemName(selectedDiscontinueTarget,lang)}`:"미선택"],
                        ["단종 사유",reqForm.discontinueReason==="기타"?(reqForm.discontinueReasonOther||"기타 상세 미입력"):reqForm.discontinueReason],
                        ["최종 사용 예정일",reqForm.finalUseDate],
                        ["잔여 재고 처리",reqForm.stockDisposition],
                        ["예상 Approval_Status","보류"]
                      ]:[
                        ["요청 유형",regType==="replacement"?"기존 품목 대체 등록":"신규 도입"],
                        ...(regType==="replacement"?[["대체 대상",selectedReplacementTarget?`${selectedReplacementTarget.code} · ${displayItemName(selectedReplacementTarget,lang)}`:"미선택"],["대체 사유",reqForm.replacementReason==="기타"?(reqForm.replacementReasonOther||"기타 상세 미입력"):reqForm.replacementReason],["기존 품목 처리",reqForm.oldItemDisposition],["전환 예정일",reqForm.replacementDate]]:[]),
                        ["Part Number",reqForm.partNumber||"자동 발번"],["품목 사진",reqPhotoFile?.name||"미선택"],["Item_Name",reqForm.itemName||"—"],["Item_Name_EN",reqForm.itemNameEn||"—"],["Type",`${ITEM_TYPE_KO[reqForm.type]||reqForm.type} (${reqForm.type})`],["Handling_Dept",reqForm.handlingDept],["Material_Category",reqForm.materialCategory],["Material_State",reqForm.materialState||"—"],["제조사",reqForm.manufacturer||"—"],["사용 예정 일자",reqForm.useDate],["CR_Type",reqForm.crType],["XRF 방식",reqXrfMode==="upload"?"XRF 원본 업로드":"XRF 측정 의뢰"],["업로드 파일",reqXrfMode==="upload"?(reqUploadFileName||"미선택"):"—"],["XRF Result",reqXrfMode==="upload"?(reqUploadResult.parsed?(reqUploadResult.xrfResult||reqUploadResult.xrfWorst):"파싱 대기"):"측정 후 산출"],["Judgment",reqXrfMode==="upload"?(reqUploadResult.parsed?measurementJudgementSummary({elements:reqUploadResult.elements||{}}):"파싱 대기"):"측정 후 산출"],["후속조치",reqXrfMode==="upload"?(reqUploadResult.parsed?measurementFollowupInfo({elements:reqUploadResult.elements||{},role:"Initial",attemptNo:1}).label:"파싱 대기"):"측정 후 산출"],["예상 Approval_Status",reqXrfMode==="request"||!reqUploadResult.parsed?"보류":(reqUploadResult.approvalStatus||deriveXrfApprovalStatus(reqUploadResult.xrfResult||reqUploadResult.xrfWorst))]
                      ]).map(([k,v])=><div key={k} style={{display:"flex",justifyContent:"space-between",gap:10,padding:"7px 0",borderBottom:`1px solid ${C.bd}`,fontSize:12}}><span style={{color:C.text4}}>{k}</span><span style={{fontWeight:600,color:C.text1,textAlign:"right"}}>{v}</span></div>)}
                    </div>
                    <button onClick={submitRequesterNewItem} style={{width:"100%",padding:"12px 0",background:C.charcoalDk,color:"#fff",border:"none",borderRadius:UI.rs,fontSize:13,fontWeight:600,cursor:"pointer"}}>
                      {regType==="discontinue"?"단종 / 사용 중지 요청 제출":"신규 등록 요청 제출"}
                    </button>
                  </div>}

                  <div style={{display:"flex",justifyContent:"space-between",marginTop:18,paddingTop:14,borderTop:`1px solid ${C.bd}`}}>
                    <button disabled={regStep===1} onClick={()=>setRegStep(s=>Math.max(1,s-1))} style={registrationButtonStyle("secondary",regStep===1)}>이전</button>
                    {regStep<requesterMaxStep&&<button disabled={regStep===1&&!regType} onClick={goNextRequesterStep} style={registrationButtonStyle("primary",regStep===1&&!regType)}>다음</button>}
                  </div>
                </div>
              ):(
                <div style={{padding:22}}>
                  {!regType&&<div>
                    <div style={{fontSize:14,fontWeight:700,marginBottom:8}}>관리자 직접 등록</div>
                    <div style={{fontSize:11,color:C.text4,lineHeight:1.7,marginBottom:14}}>관리자가 수동으로 신규 도입, 기존 주기 측정 추가, 변경·대체, 단종 처리를 직접 선택합니다.</div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(4,minmax(0,1fr))",gap:12}}>
                      {[
                        ["admin-new","신규 도입","새 품목을 즉시 등록하고 XRF 결과까지 반영"],
                        ["admin-period","기존 주기 도래 / 측정 추가","기존 품목의 정기 XRF 결과를 추가 등록"],
                        ["admin-replace","변경 / 대체 등록","기존 품목과 신규 대체품을 연결"],
                        ["admin-discontinue","단종 처리","기존 품목의 사용 중지 또는 이력 전환"]
                      ].map(([key,title,desc])=>(
                        <button key={key} onClick={()=>{setRegType(key);setRegStep(1);setReqUploadResult(emptyUploadedXrfResult());setReqUploadFileName("");clearRequestPhoto();}}
                          style={{textAlign:"left",padding:"16px",border:`1px solid ${C.bd}`,background:C.card,borderRadius:8,cursor:"pointer",minHeight:130}}>
                          <div style={{fontSize:13,fontWeight:800,color:C.text1,marginBottom:8}}>{title}</div>
                          <div style={{fontSize:11,color:C.text3,lineHeight:1.65}}>{desc}</div>
                        </button>
                      ))}
                    </div>
                  </div>}

                  {regType&&<div>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,marginBottom:14}}>
                      <div>
                        <div style={{fontSize:14,fontWeight:800,color:C.text1}}>
                          {regType==="admin-new"?"관리자 직접 등록 / 신규 도입":regType==="admin-period"?"관리자 직접 등록 / 기존 주기 측정 추가":regType==="admin-replace"?"관리자 직접 등록 / 변경/대체 등록":"관리자 직접 등록 / 단종 처리"}
                        </div>
                        <div style={{fontSize:11,color:C.text4,lineHeight:1.6,marginTop:3}}>기본 정보, C&R 등급, XRF 결과, 위험도 산정, 승인 상태를 관리자가 직접 확인합니다.</div>
                      </div>
                      <button onClick={()=>{setRegType(null);setReqUploadResult(emptyUploadedXrfResult());setReqUploadFileName("");clearRequestPhoto();}} className="pill-btn" style={{padding:"7px 14px",background:C.card,border:`1px solid ${C.bd2}`,fontSize:11,color:C.text3}}>유형 다시 선택</button>
                    </div>

                    {(regType==="admin-period"||regType==="admin-replace"||regType==="admin-discontinue")&&<div style={{padding:"12px",background:C.bg,border:`1px solid ${C.bd}`,borderRadius:8,marginBottom:14}}>
                      <div style={{fontSize:12,fontWeight:700,color:C.text1,marginBottom:8}}>
                        {regType==="admin-period"?"측정 추가 대상 품목":regType==="admin-replace"?"대체 대상 기존 품목":"단종 대상 기존 품목"}
                      </div>
                      <select value={regType==="admin-discontinue"?reqForm.discontinueTargetCode:reqForm.replacementTargetCode}
                        onChange={e=>updateRequestField(regType==="admin-discontinue"?"discontinueTargetCode":"replacementTargetCode",e.target.value)}
                        style={{...inp,width:"100%",boxSizing:"border-box"}}>
                        <option value="">기존 품목을 선택하세요</option>
                        {selectableExistingItems.map(item=><option key={item.code} value={item.code}>{item.code} | {item.dept} | {displayItemName(item,lang)}</option>)}
                      </select>
                    </div>}

                    {regType!=="admin-period"&&regType!=="admin-discontinue"&&<div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:12,marginBottom:14}}>
                      {[
                        ["partNumber","Part Number","자동 생성"],
                        ["itemName","Item_Name","예: 제품 포장용 엠보스"],
                        ["itemNameEn","Item_Name_EN","e.g. Packaging Emboss"],
                        ["manufacturer","제조사","예: ABC Chemical"],
                        ["useDate","사용 일자",TODAY_STR]
                      ].map(([k,l,ph])=>(
                        <div key={k}><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>{l}</div><input type={k==="useDate"?"date":"text"} value={reqForm[k]} onChange={e=>updateRequestField(k,e.target.value)} placeholder={ph} readOnly={k==="partNumber"} style={{...inp,width:"100%",boxSizing:"border-box",background:k==="partNumber"?C.bg:C.card}}/></div>
                      ))}
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Type</div><select value={reqForm.type} onChange={e=>updateRequestField("type",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}>{ITEM_TYPE_OPTIONS.map(v=><option key={v} value={v}>{ITEM_TYPE_KO[v]} ({v})</option>)}</select></div>
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Handling_Dept</div><select value={reqForm.handlingDept} onChange={e=>updateRequestField("handlingDept",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}>{DEPTS.map(d=><option key={d}>{d}</option>)}</select></div>
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Material_Category</div><select value={reqForm.materialCategory} onChange={e=>updateRequestField("materialCategory",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}>{REQUESTER_MATERIAL_CATEGORY_OPTIONS.map(v=><option key={v}>{v}</option>)}</select></div>
                      <div>
                        <div style={{fontSize:12,fontWeight:600,marginBottom:5}}>Material_State {isChemicalMaterialCategory(reqForm.materialCategory)&&<span style={{color:C.red}}>*</span>}</div>
                        {isChemicalMaterialCategory(reqForm.materialCategory)?(
                          <select value={reqForm.materialState} onChange={e=>updateRequestField("materialState",e.target.value)} required style={{...inp,width:"100%",boxSizing:"border-box",borderColor:reqForm.materialState?C.bd:C.red}}>
                            <option value="">Material_State를 선택하세요</option>
                            {CHEMICAL_MATERIAL_STATE_OPTIONS.map(v=><option key={v} value={v}>{v}</option>)}
                          </select>
                        ):(
                          <input value={reqForm.materialState} onChange={e=>updateRequestField("materialState",e.target.value)} placeholder="예: Solid / Powder" style={{...inp,width:"100%",boxSizing:"border-box"}}/>
                        )}
                      </div>
                      <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>CR_Type</div><select value={reqForm.crType} onChange={e=>updateRequestField("crType",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}><option>직접접촉+잔류</option><option>직접접촉+비잔류</option><option>비접촉</option></select></div>
                      <div style={{gridColumn:"1 / -1"}}>
                        <div style={{fontSize:12,fontWeight:600,marginBottom:5}}>품목 사진 <span style={{color:C.red}}>*</span></div>
                        <div style={{display:"grid",gridTemplateColumns:reqPhotoPreview?"120px 1fr":"1fr",gap:12,alignItems:"stretch"}}>
                          {reqPhotoPreview&&<img src={reqPhotoPreview} alt="선택한 품목 사진" style={{width:120,height:96,objectFit:"cover",borderRadius:8,border:`1px solid ${C.bd}`}}/>}
                          <label style={{...uploadBoxStyle,minHeight:96,borderRadius:8}}>
                            <span data-i18n-skip="true">{reqPhotoFile?.name||(lang==="en"?"Choose Photo (JPG, PNG, WEBP / max. 10 MB)":"사진 선택 (JPG, PNG, WEBP / 최대 10MB)")}</span>
                            <input type="file" accept="image/jpeg,image/png,image/webp,.jpg,.jpeg,.png,.webp" onChange={handleRequestPhoto} style={fileInputOverlay}/>
                          </label>
                        </div>
                      </div>
                    </div>}

                    {regType==="admin-discontinue"?(
                      <div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:12,marginBottom:14}}>
                        <div>
                          <div style={{fontSize:12,fontWeight:600,marginBottom:5}}>단종 사유</div>
                          <select value={reqForm.discontinueReason} onChange={e=>updateRequestField("discontinueReason",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}><option>사용 중지</option><option>공정 변경</option><option>고객 요구</option><option>공급 중단</option><option>대체품 적용 완료</option><option>기타</option></select>
                          {reqForm.discontinueReason==="기타"&&<textarea value={reqForm.discontinueReasonOther} onChange={e=>updateRequestField("discontinueReasonOther",e.target.value)} placeholder="기타 단종 사유를 구체적으로 입력하세요" style={{...inp,width:"100%",boxSizing:"border-box",minHeight:64,resize:"vertical",marginTop:8,borderColor:reqForm.discontinueReasonOther.trim()?C.bd:C.red}}/>}
                        </div>
                        <div><div style={{fontSize:12,fontWeight:600,marginBottom:5}}>최종 사용일</div><input type="date" value={reqForm.finalUseDate} onChange={e=>updateRequestField("finalUseDate",e.target.value)} style={{...inp,width:"100%",boxSizing:"border-box"}}/></div>
                        <div style={{gridColumn:"1 / -1",padding:"14px 16px",background:C.bg,border:`1px dashed ${C.bd2}`,fontSize:12,color:C.text3,lineHeight:1.8}}>승인 시 선택 품목은 Discontinued로 전환되고, 리스트에서는 이력 상태로 관리됩니다.</div>
                      </div>
                    ):(
                      <div style={{display:"grid",gridTemplateColumns:"1.1fr .9fr",gap:12,padding:"12px",background:C.bg,border:`1px solid ${C.bd}`,borderRadius:8}}>
                        <div>
                          <div style={{fontSize:11,fontWeight:600,marginBottom:6}}>XRF 결과 Excel 업로드</div>
                          <label style={uploadBoxStyle}>
                            {reqUploadFileName || "XRF 결과 파일 선택 (.xlsx)"}
                            <input type="file" accept=".xlsx,.xlsm,.xls,.csv,.json" onChange={regType==="admin-period"?handleAdminDirectXrfUpload:handleRequesterXrfUpload} disabled={measurementSavePending} style={fileInputOverlay}/>
                          </label>
                          <div style={{fontSize:10,color:C.text4,lineHeight:1.55,marginTop:6}}>
                            {regType==="admin-period"
                              ? "측정 추가 대상 품목을 선택한 뒤 파일을 넣으면 해당 주기에 바로 저장되고 XRF 분석 화면으로 이동합니다."
                              : "XRF Report Content와 Internal Limit(법적 기준의 70%)을 비교해 XRF Result와 R 위험도 평가를 자동 산정합니다."}
                          </div>
                        </div>
                        <div style={{display:"grid",gridTemplateColumns:"1fr",gap:7}}>
                          <div style={{padding:"8px 10px",background:C.card,border:`1px solid ${C.bd}`}}><div style={{fontSize:10,color:C.text4,marginBottom:4}}>XRF 결과</div>{reqUploadResult.parsed?<Chip v={reqUploadResult.xrfWorst||"—"} small result/>:<span style={{fontSize:11,color:C.text4}}>{reqUploadResult.pending?"읽는 중":"파싱 대기"}</span>}</div>
                          <div style={{padding:"8px 10px",background:C.card,border:`1px solid ${C.bd}`}}><div style={{fontSize:10,color:C.text4,marginBottom:4}}>후속조치</div>{reqUploadResult.parsed?<XrfFollowupBadge measurement={{elements:reqUploadResult.elements||{}}}/>:<span style={{fontSize:11,color:C.text4}}>—</span>}</div>
                          <div style={{padding:"8px 10px",background:C.card,border:`1px solid ${C.bd}`}}><div style={{fontSize:10,color:C.text4,marginBottom:4}}>Approval_Status</div><ApprovalBadge status={reqUploadResult.parsed?(reqUploadResult.approvalStatus||deriveXrfApprovalStatus(reqUploadResult.xrfResult||reqUploadResult.xrfWorst)):"보류"}/></div>
                        </div>
                        <div style={{gridColumn:"1 / -1"}}><UploadedXrfPreview result={reqUploadResult}/></div>
                      </div>
                    )}

                    <div style={{display:"flex",justifyContent:"flex-end",gap:8,marginTop:16,paddingTop:14,borderTop:`1px solid ${C.bd}`}}>
                      <button onClick={handleAdminDraftSave} style={registrationButtonStyle("secondary")}>임시 저장</button>
                      <button onClick={handleAdminDirectSubmit} disabled={measurementSavePending} style={{...registrationButtonStyle("primary",measurementSavePending),cursor:measurementSavePending?"wait":"pointer",minWidth:96}}>{measurementSavePending?"SharePoint 저장 중...":"등록"}</button>
                    </div>
                  </div>}
                </div>
              )}
            </div>
          </div>
        )}

      </div>
      <div style={{borderTop:`1px solid ${C.bd}`,padding:"12px 24px",textAlign:"center",fontSize:10,color:C.text4}}>
        Molex | 부자재 XRF 관리 시스템
      </div>
    </div>
  );
}
