/* Builds a fixture cutsheet that mimics the real MASTER-US-EAST layout
   (multiple location columns per side) to verify auto-detect + aggregation
   end-to-end in the browser. Invented data only. */
import { writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
const require = createRequire(import.meta.url);
const XLSX = require(path.join(process.cwd(), "vendor", "xlsx.full.min.js"));

const headers = [
  "A-SIDE LOCODE", "A-LOC:CAB:RU", "A-BREAKOUT LOC:CAB:RU", "A-OPTIC", "A-PATCH-PANEL LOC:CAB:RU:PORT",
  "Z-SIDE LOCODE", "Z-LOC:CAB:RU", "Z-BREAKOUT LOC:CAB:RU", "Z-OPTIC", "Z-PATCH-PANEL LOC:CAB:RU:PORT",
];
const DR = "OSFP-800G-2DR4";
const R = (aLoc, aOpt, zLoc, zOpt, pp) =>
  ["US-CDZ01", aLoc, "", aOpt, pp || "", "US-CDZ01", zLoc, "", zOpt, pp || ""];

const rows = [];
// 78 connections whose Z end is s4:021:39; some A-optics blank (breakouts).
for (let i = 1; i <= 78; i++) {
  const aBlank = i <= 24; // first chunk are breakout children with no A-optic
  rows.push(R(`s4:0${10 + (i % 9)}:${(i % 40) + 1}`, aBlank ? "" : DR, "s4:021:39", DR));
}
// noise rows that must not match
rows.push(R("s4:099:01", DR, "s4:099:02", DR));
rows.push(R("s5:021:39", DR, "s6:010:10", DR));

const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Cutsheet");
writeFileSync(path.join("sample", "fixture-real-layout.xlsx"), XLSX.write(wb, { bookType: "xlsx", type: "buffer" }));
console.log("Wrote sample/fixture-real-layout.xlsx (78 matching + 2 noise rows)");
