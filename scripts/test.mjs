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
const {
  parseLocation, locFromParts, locMatches, locLabel,
  sideToken, matchRowsPure, aggregatePure,
} = moduleShim.exports;

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

// ---------------------------------------------------------------------------
// Regression for the real MASTER-US-EAST cutsheet layout: several location
// columns per side (LOCODE / LOC:CAB:RU / BREAKOUT / PATCH-PANEL). Optics must
// bind to a SIDE, not the first location column, or "at searched location"
// wrongly returns 0.
// ---------------------------------------------------------------------------
console.log("\nreal-cutsheet layout (side-based optic binding):");
const H2 = [
  "A-SIDE LOCODE", "A-LOC:CAB:RU", "A-BREAKOUT LOC:CAB:RU", "A-OPTIC", "A-PATCH-PANEL LOC:CAB:RU:PORT",
  "Z-SIDE LOCODE", "Z-LOC:CAB:RU", "Z-BREAKOUT LOC:CAB:RU", "Z-OPTIC", "Z-PATCH-PANEL LOC:CAB:RU:PORT",
];
check("sideToken(Z-SIDE LOCODE) === z", sideToken("Z-SIDE LOCODE") === "z");
check("sideToken(Z-LOC:CAB:RU) === z", sideToken("Z-LOC:CAB:RU") === "z");
check("sideToken(A-OPTIC) === a && sideToken(Z-OPTIC) === z", sideToken("A-OPTIC") === "a" && sideToken("Z-OPTIC") === "z");

// Mapping as auto-detect would build it: every "loc"-ish column is an endpoint,
// each carrying its side; the two optic columns carry their side.
const ep = (i) => ({ mode: "combined", combined: i, sector: null, rack: null, ru: null, name: H2[i], side: sideToken(H2[i]) });
const endpoints2 = [0, 1, 2, 4, 5, 6, 7, 9].map(ep);
const optic = (i) => ({ pn: i, qty: null, desc: null, endpoint: null, side: sideToken(H2[i]), name: H2[i] });
const optics2 = [optic(3), optic(8)];

const DR = "OSFP-800G-2DR4";
// Z-LOC = s4:021:39 on the matched rows; A-LOC is a different rack.
const R = (aLoc, aOpt, zLoc, zOpt) =>
  ["US-CDZ01", aLoc, "", aOpt, "", "US-CDZ01", zLoc, "", zOpt, ""];
const rows2 = [
  R("s4:010:11", "",  "s4:021:39", DR),   // A-optic blank
  R("s4:010:12", "",  "s4:021:39", DR),   // A-optic blank
  R("s4:010:13", DR,  "s4:021:39", DR),
  R("s4:010:14", DR,  "s4:021:39", DR),
  R("s4:010:15", DR,  "s4:021:39", DR),
  R("s4:099:01", DR,  "s4:099:02", DR),   // noise: does not touch s4:021:39
];

const q2 = [{ text: "s4:021:39", loc: parseLocation("s4:021:39") }];
const matched2 = matchRowsPure(rows2, endpoints2, q2);
check("5 connections reference s4:021:39", matched2.length === 5);

const atSide = aggregatePure(matched2, optics2, endpoints2, "endpoint");
const totalSide = atSide.list.reduce((s, r) => s + r.qty, 0);
console.log("    at searched location:", JSON.stringify(atSide.list));
check("optics at location is NOT zero (the reported bug)", totalSide > 0);
check("5 optics at searched location (Z-OPTIC only)", totalSide === 5);
check("exactly one distinct PN at location", atSide.list.length === 1 && atSide.list[0].pn === DR);

const both = aggregatePure(matched2, optics2, endpoints2, "all");
const totalBoth = both.list.reduce((s, r) => s + r.qty, 0);
check("whole-connection counts both ends (5 Z + 3 A = 8)", totalBoth === 8);

console.log(`\nAll ${pass} checks passed.`);
