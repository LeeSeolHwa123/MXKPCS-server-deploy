import { spawn } from "node:child_process";
import { createServer } from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot=path.resolve(path.dirname(fileURLToPath(import.meta.url)),"..");
const host="127.0.0.1";
const uiPort=Number(process.env.MXKPCS_LOCAL_UI_PORT||5176);
const bridgePort=Number(process.env.MXKPCS_LOCAL_API_PORT||3100);
const deployment=String(
  process.env.MXKPCS_VERCEL_DEPLOYMENT
  || "https://mxkpcs-pb05lfser-yah2685-2386.vercel.app"
).trim();
const viteCli=path.join(projectRoot,"node_modules","vite","bin","vite.js");
const vercelCli=path.join(path.dirname(process.execPath),"node_modules","vercel","dist","vc.js");

function sendJson(response,status,body){
  response.writeHead(status,{
    "Content-Type":"application/json; charset=utf-8",
    "Cache-Control":"no-store, max-age=0",
  });
  response.end(JSON.stringify(body));
}

function runVercelCurl(request,response){
  const requestUrl=new URL(request.url||"/",`http://${host}:${bridgePort}`);
  if(requestUrl.pathname!=="/api/xrf-sharepoint"){
    sendJson(response,404,{ok:false,error:"Local API route not found"});
    return;
  }

  const method=String(request.method||"GET").toUpperCase();
  const args=[
    vercelCli,"curl",`${requestUrl.pathname}${requestUrl.search}`,
    "--deployment",deployment,"--yes","--",
    "--silent","--show-error","--request",method,
    "--header","Accept: application/json",
  ];
  if(!["GET","HEAD"].includes(method)){
    args.push(
      "--header",`Content-Type: ${request.headers["content-type"]||"application/json"}`,
      "--data-binary","@-"
    );
  }

  const child=spawn(process.execPath,args,{
    cwd:projectRoot,
    env:process.env,
    windowsHide:true,
    stdio:["pipe","pipe","pipe"],
  });
  const stdout=[];
  const stderr=[];
  child.stdout.on("data",chunk=>stdout.push(chunk));
  child.stderr.on("data",chunk=>stderr.push(chunk));
  child.on("error",error=>{
    if(!response.headersSent) sendJson(response,502,{ok:false,error:`Local API bridge failed: ${error.message}`});
  });
  child.on("close",code=>{
    if(response.writableEnded) return;
    const body=Buffer.concat(stdout).toString("utf8").trim();
    if(code!==0){
      const detail=Buffer.concat(stderr).toString("utf8").trim().split(/\r?\n/).slice(-1)[0]||`exit ${code}`;
      sendJson(response,502,{ok:false,error:`Local API bridge failed: ${detail}`});
      return;
    }
    try{
      JSON.parse(body);
      response.writeHead(200,{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store, max-age=0"});
      response.end(body);
    }catch{
      sendJson(response,502,{ok:false,error:"Local API bridge received a non-JSON response"});
    }
  });

  if(["GET","HEAD"].includes(method)) child.stdin.end();
  else request.pipe(child.stdin);
}

const bridge=createServer(runVercelCurl);
bridge.listen(bridgePort,host,()=>{
  const vite=spawn(process.execPath,[viteCli,"--host",host,"--port",String(uiPort),"--strictPort"],{
    cwd:projectRoot,
    env:{...process.env,MXKPCS_LOCAL_API_TARGET:`http://${host}:${bridgePort}`},
    windowsHide:true,
    stdio:"inherit",
  });

  console.log(`MXKPCS local API bridge: http://${host}:${bridgePort}`);
  console.log(`MXKPCS local UI: http://${host}:${uiPort}`);

  const stop=()=>{
    if(!vite.killed) vite.kill("SIGTERM");
    bridge.close(()=>process.exit(0));
  };
  process.once("SIGINT",stop);
  process.once("SIGTERM",stop);
  vite.once("exit",code=>{
    bridge.close(()=>process.exit(code??0));
  });
});

bridge.on("error",error=>{
  console.error(`MXKPCS local API bridge error: ${error.message}`);
  process.exit(1);
});
