import {readFileSync} from "node:fs";
import {defineConfig} from "vite";

const packageJson=JSON.parse(readFileSync(new URL("./package.json",import.meta.url),"utf8"));

export default defineConfig({
  publicDir:"outputs/sharepoint_db_schema_2026-11-03",
  define:{
    __APP_VERSION__:JSON.stringify(String(packageJson.version||"0.0.0"))
  }
});
