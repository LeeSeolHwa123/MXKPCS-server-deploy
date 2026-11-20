import {readFileSync} from "node:fs";
import {defineConfig} from "vite";

const packageJson=JSON.parse(readFileSync(new URL("./package.json",import.meta.url),"utf8"));
const localApiTarget=String(process.env.MXKPCS_LOCAL_API_TARGET||"").trim();

export default defineConfig({
  publicDir:"outputs/sharepoint_db_schema_2026-11-03",
  server:localApiTarget?{
    proxy:{
      "/api/xrf-sharepoint":{
        target:localApiTarget,
        changeOrigin:true,
      }
    }
  }:undefined,
  define:{
    __APP_VERSION__:JSON.stringify(String(packageJson.version||"0.0.0"))
  }
});
