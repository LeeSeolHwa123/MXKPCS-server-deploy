import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createRequire } from "node:module";
import { transformWithOxc } from "vite";
import React from "react";
import { renderToStaticMarkup } from "react-dom/server";

const require = createRequire(import.meta.url);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourcePath = path.join(projectRoot, "app.jsx");
let source = await fs.readFile(sourcePath, "utf8");
const originalSource = source;
source = source
  .replace('from "react"', `from "${pathToFileURL(require.resolve("react")).href}"`)
  .replace('from "xlsx"', `from "${pathToFileURL(require.resolve("xlsx")).href}"`);
source += "\nexport { SplashScreen };\n";
const transformed = await transformWithOxc(source, sourcePath, { lang: "jsx", jsx: { runtime: "automatic" } });
const executable = transformed.code.replace(
  'from "react/jsx-runtime"',
  `from "${pathToFileURL(require.resolve("react/jsx-runtime")).href}"`
);
const moduleUrl = `data:text/javascript;base64,${Buffer.from(executable).toString("base64")}`;
const { default: App, SplashScreen } = await import(moduleUrl);

const originalConsoleError = console.error;
let markup;
try {
  // The existing App uses client-only layout effects; React warns during this SSR-only assertion.
  console.error = (...args) => {
    if (String(args[0] || "").includes("useLayoutEffect does nothing on the server")) return;
    originalConsoleError(...args);
  };
  markup = renderToStaticMarkup(React.createElement(App));
} finally {
  console.error = originalConsoleError;
}
assert.match(markup, /MXK/);
assert.doesNotMatch(markup, /대시보드 열기/);
assert.match(markup, /부자재 리스트/);
assert.match(markup, /품목 \/ 변경 관리/);
assert.match(markup, /Version 1\.0\.0/);
assert.doesNotMatch(markup, /Build local/);
assert.doesNotMatch(markup, /SharePoint DB를 불러오는 중입니다/);

let selectedTab = "";
let chosenLanguage = "";
const splash = SplashScreen({
  onSelectTab: value => { selectedTab = value; },
  lang: "ko",
  onLang: value => { chosenLanguage = value; },
});
function find(node, predicate) {
  if (!node || typeof node !== "object") return null;
  if (Array.isArray(node)) return node.map(child => find(child, predicate)).find(Boolean) || null;
  if (predicate(node)) return node;
  return find(node.props?.children, predicate);
}
for (const tab of ["list", "xrf", "precision", "risk", "reg"]) {
  const menu = find(splash, node => node.type === "button" && node.props?.["data-splash-tab"] === tab);
  assert.ok(menu, `${tab} splash menu is missing`);
  menu.props.onClick();
  assert.equal(selectedTab, tab);
}
const english = find(splash, node => node.type === "button" && node.props?.children === "English");
assert.ok(english);
english.props.onClick();
assert.equal(chosenLanguage, "en");

assert.match(originalSource, /if\(!entered\)\{\s*return <SplashScreen/);
assert.match(originalSource, /className="app-brand" onClick=\{\(\)=>setEntered\(false\)\}/);
assert.match(originalSource, /\},\[lang,entered\]\);/);
assert.match(originalSource, /className="mxk-splash"[^>]*height:"100dvh"/);
assert.match(originalSource, /className="mxk-splash-hero"/);
console.log("splash-screen: PASS");
