import assert from "node:assert/strict";
import sharePointHandler from "../api/xrf-sharepoint.js";

const originalFetch=globalThis.fetch;
const originalSecret=process.env.MXKPCS_CLIENT_SECRET;
let createdColumn=null;
let measurementFields=null;
let elementSequence=0;

async function invoke(body){
  let output="";
  const response={
    statusCode:200,
    setHeader(){},
    end(value=""){ output=value; },
  };
  await sharePointHandler({method:"POST",body,url:"/api/xrf-sharepoint",headers:{}},response);
  return {status:response.statusCode,body:output?JSON.parse(output):null};
}

try{
  process.env.MXKPCS_CLIENT_SECRET="test-only-secret";
  globalThis.fetch=async(input,options={})=>{
    const url=String(input);
    const method=String(options.method||"GET").toUpperCase();
    if(url.includes("/oauth2/v2.0/token")){
      return new Response(JSON.stringify({access_token:"fake-token"}),{status:200});
    }
    if(url.includes("/lists/dd0bf876-6066-4877-be6d-eb10eb5ce35c/columns")&&method==="GET"){
      return new Response(JSON.stringify({value:[{name:"field_7",displayName:"Source_Sheet"}]}),{status:200});
    }
    if(url.includes("/lists/dd0bf876-6066-4877-be6d-eb10eb5ce35c/items")){
      if(method==="GET") return new Response(JSON.stringify({value:[]}),{status:200});
      if(method==="POST"){
        measurementFields=JSON.parse(options.body)?.fields||null;
        return new Response(JSON.stringify({id:"measurement-row",fields:measurementFields}),{status:201});
      }
    }
    if(url.includes("/lists/b0a877ee-5334-4237-9b8f-e94235b118a4/items")&&method==="POST"){
      elementSequence+=1;
      return new Response(JSON.stringify({id:`element-${elementSequence}`,fields:JSON.parse(options.body)?.fields||{}}),{status:201});
    }
    throw new Error(`Unexpected mock Graph request: ${method} ${url}`);
  };

  const result=await invoke({action:"saveMeasurement",payload:{
    itemId:"ITM_DB_PERIOD",
    partNumber:"DB-PERIOD",
    measuredDate:"2026-10-21",
    role:"Periodic",
    periodNo:3,
    attemptNo:1,
    elements:{Pb:{rawPpm:"10",rawSigma:"1",rawJudgement:"OK"}},
  }});

  assert.equal(result.status,200);
  assert.equal(result.body.ok,true);
  assert.equal(createdColumn,null);
  assert.equal(measurementFields.field_7,"[[MXKPCS:Period_No=3]]");
  assert.equal(measurementFields.field_2,undefined);
  assert.equal(measurementFields.field_3,"Periodic");
  assert.equal(elementSequence,6);
  console.log("measurement-period-db: PASS");
}finally{
  globalThis.fetch=originalFetch;
  if(originalSecret===undefined) delete process.env.MXKPCS_CLIENT_SECRET;
  else process.env.MXKPCS_CLIENT_SECRET=originalSecret;
}
