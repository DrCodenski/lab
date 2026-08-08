/* Generates a dummy CoreWeave-style cutsheet (xlsx + csv) for demoing the tool.
   No real data — invented racks, ports and part numbers. */
import { writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";

const require = createRequire(import.meta.url);
const XLSX = require(path.join(process.cwd(), "vendor", "xlsx.full.min.js"));

const headers = [
  "Cable ID", "Media Type", "Length (m)",
  "A-End Location", "A-End Device", "A-End Port", "A-End Optic PN", "A-End Optic Description",
  "Z-End Location", "Z-End Device", "Z-End Port", "Z-End Optic PN", "Z-End Optic Description",
];

// Optic catalog: PN -> description
const OPT = {
  "OPT-QSFP28-100G-SR4": "100G QSFP28 SR4 MMF MPO-12",
  "OPT-QSFP28-100G-LR4": "100G QSFP28 LR4 SMF LC",
  "OPT-QSFPDD-400G-DR4": "400G QSFP-DD DR4 SMF MPO-12",
  "OPT-QSFPDD-400G-FR4": "400G QSFP-DD FR4 SMF LC",
  "OPT-SFP28-25G-SR": "25G SFP28 SR MMF LC",
  "OPT-OSFP-800G-2FR4": "800G OSFP 2xFR4 SMF",
};

function row(id, media, len, aLoc, aDev, aPort, aPN, zLoc, zDev, zPort, zPN) {
  return [id, media, len, aLoc, aDev, aPort, aPN, OPT[aPN], zLoc, zDev, zPort, zPN, OPT[zPN]];
}

// Several connections deliberately reference S1:05:30 on one end or the other.
const rows = [
  row("C-0001", "SMF", 12, "s1:05:30", "leaf-s1-05a", "et-0/0/1", "OPT-QSFPDD-400G-DR4", "s1:01:40", "spine-s1-01", "et-0/0/12", "OPT-QSFPDD-400G-DR4"),
  row("C-0002", "SMF", 12, "s1:05:30", "leaf-s1-05a", "et-0/0/2", "OPT-QSFPDD-400G-DR4", "s1:02:40", "spine-s1-02", "et-0/0/12", "OPT-QSFPDD-400G-DR4"),
  row("C-0003", "MMF", 3,  "s1:05:30", "leaf-s1-05a", "et-0/0/3", "OPT-QSFP28-100G-SR4", "s1:05:31", "srv-s1-05-31", "eth0", "OPT-QSFP28-100G-SR4"),
  row("C-0004", "SMF", 15, "s1:03:22", "spine-s1-03", "et-0/0/5", "OPT-QSFPDD-400G-FR4", "s1:05:30", "leaf-s1-05b", "et-0/0/4", "OPT-QSFPDD-400G-FR4"),
  row("C-0005", "MMF", 2,  "s1:05:31", "srv-s1-05-31", "eth1", "OPT-SFP28-25G-SR", "s1:05:30", "leaf-s1-05a", "et-0/0/5", "OPT-SFP28-25G-SR"),
  row("C-0006", "SMF", 30, "s1:05:30", "leaf-s1-05a", "et-0/0/6", "OPT-OSFP-800G-2FR4", "s2:10:12", "dci-s2-10", "et-0/1/0", "OPT-OSFP-800G-2FR4"),
  // noise rows that should NOT match s1:05:30
  row("C-0007", "SMF", 12, "s1:06:30", "leaf-s1-06a", "et-0/0/1", "OPT-QSFPDD-400G-DR4", "s1:01:41", "spine-s1-01", "et-0/0/13", "OPT-QSFPDD-400G-DR4"),
  row("C-0008", "MMF", 5,  "s2:05:30", "leaf-s2-05a", "et-0/0/1", "OPT-QSFP28-100G-SR4", "s2:01:40", "spine-s2-01", "et-0/0/1", "OPT-QSFP28-100G-SR4"),
  row("C-0009", "SMF", 10, "s1:05:29", "leaf-s1-05a", "et-0/0/9", "OPT-QSFP28-100G-LR4", "s1:02:38", "spine-s1-02", "et-0/0/9", "OPT-QSFP28-100G-LR4"),
];

const aoa = [headers, ...rows];
const ws = XLSX.utils.aoa_to_sheet(aoa);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Cutsheet");

const xlsxBuf = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
writeFileSync(path.join("sample", "sample-cutsheet.xlsx"), xlsxBuf);
const csvStr = XLSX.utils.sheet_to_csv(ws);
writeFileSync(path.join("sample", "sample-cutsheet.csv"), csvStr);
console.log("Wrote sample/sample-cutsheet.xlsx and sample/sample-cutsheet.csv");
