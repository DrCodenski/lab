/* Headless sanity tests for the lookup logic. No browser needed.
   Reuses the sample cutsheet and re-implements the same matching pipeline
   the browser app uses, asserting on known expected results. */
import { createRequire } from "node:module";
import path from "node:path";
import assert from "node:assert";

const require = createRequire(import.meta.url);
const XLSX = require(path.join(process.cwd(), "vendor", "xlsx.full.min.js"));

// --- Import the pure functions from app.js via a tiny shim.
// app.js guards a module.exports for exactly this purpose.
const appPath = path.join(process.cwd(), "app.js");
const src = require("node:fs").readFileSync(appPath, "utf8");
const moduleShim = { exports: {} };
// Provide the globals app.js touches at import time.
const sandbox = {
  module: moduleShim,
  document: { addEventListener() {} },
  window: { matchMedia: () => ({ matches: false }) },
  localStorage: { getItem: () => null, setItem() {} },
  navigator: {},
  XLSX,
};
const fn = new Function(...Object.keys(sandbox), src);
fn(...Object.values(sandbox));
const { parseLocation, locFromParts, locMatches, locLabel } = moduleShim.exports;

let pass = 0;
function check(name, cond) {
  assert.ok(cond, "FAIL: " + name);
  pass++;
  console.log("  ✓ " + name);
}

console.log("parseLocation / matching:");
const q = parseLocation("s1:05:30");
check("parses s1:05:30 -> S1 R5 U30", q.sector === 1 && q.rack === 5 && q.ru === 30);
check("05 == 5 (leading zeros)", parseLocation("s1:5:30").rack === parseLocation("s1:05:30").rack);
check("dash format s1-05-30 parses same", locMatches(q, parseLocation("s1-05-30")));
check("labelled r05 ru30 parses", (() => { const l = parseLocation("S1 R05 RU30"); return l.rack === 5 && l.ru === 30; })());
check("partial query s1:05 matches RU 30 and 31", locMatches(parseLocation("s1:05"), parseLocation("s1:05:31")));
check("s1:05:30 does NOT match s1:06:30", !locMatches(q, parseLocation("s1:06:30")));
check("s1:05:30 does NOT match s2:05:30", !locMatches(q, parseLocation("s2:05:30")));
check("split cols compose location", (() => { const l = locFromParts("1", "05", "30"); return l.sector === 1 && l.rack === 5 && l.ru === 30; })());
check("locLabel formats canonically", locLabel(q) === "S1:05:30");

console.log("\nend-to-end against sample cutsheet:");
const buf = require("node:fs").readFileSync(path.join("sample", "sample-cutsheet.xlsx"));
const wb = XLSX.read(buf, { type: "buffer" });
const ws = wb.Sheets[wb.SheetNames[0]];
const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
const headers = aoa[0];
const rows = aoa.slice(1).filter((r) => r.some((c) => String(c).trim() !== ""));

const idx = (name) => headers.indexOf(name);
const A_LOC = idx("A-End Location"), Z_LOC = idx("Z-End Location");
const A_PN = idx("A-End Optic PN"), Z_PN = idx("Z-End Optic PN");

const query = parseLocation("s1:05:30");
const matched = rows.filter((r) =>
  locMatches(query, parseLocation(r[A_LOC])) || locMatches(query, parseLocation(r[Z_LOC]))
);
check("6 connections reference s1:05:30", matched.length === 6);

// Optics "at searched location" = the optic on whichever end matched.
const atLoc = {};
for (const r of matched) {
  if (locMatches(query, parseLocation(r[A_LOC]))) atLoc[r[A_PN]] = (atLoc[r[A_PN]] || 0) + 1;
  if (locMatches(query, parseLocation(r[Z_LOC]))) atLoc[r[Z_PN]] = (atLoc[r[Z_PN]] || 0) + 1;
}
const totalAtLoc = Object.values(atLoc).reduce((a, b) => a + b, 0);
console.log("    optics at s1:05:30:", JSON.stringify(atLoc));
check("6 optics needed at the searched location", totalAtLoc === 6);
check("400G DR4 x2 at location", atLoc["OPT-QSFPDD-400G-DR4"] === 2);
check("100G SR4 x1 at location", atLoc["OPT-QSFP28-100G-SR4"] === 1);
check("400G FR4 x1 at location", atLoc["OPT-QSFPDD-400G-FR4"] === 1);
check("25G SR x1 at location", atLoc["OPT-SFP28-25G-SR"] === 1);
check("800G 2FR4 x1 at location", atLoc["OPT-OSFP-800G-2FR4"] === 1);

console.log(`\nAll ${pass} checks passed.`);
