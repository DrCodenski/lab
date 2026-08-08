/* Drives the built app in Chromium against the real-layout fixture and asserts
   the optics table populates (the reported bug) with one-side counts. */
import { chromium } from "playwright";
import path from "node:path";
import { pathToFileURL } from "node:url";
import assert from "node:assert";

const appUrl = pathToFileURL(path.resolve("dist/coreweave-cutsheet-lookup.html")).href;
const fixture = path.resolve("sample/fixture-real-layout.xlsx");

const browser = await chromium.launch({
  executablePath: "/opt/pw-browsers/chromium", headless: true, args: ["--headless=new", "--no-sandbox"],
});
try {
  const page = await browser.newPage({ viewport: { width: 1300, height: 1700 }, deviceScaleFactor: 2 });
  page.on("pageerror", (e) => console.log("[pageerror]", e.message));
  await page.goto(appUrl);
  await page.setInputFiles("#fileInput", fixture);
  await page.waitForSelector("#searchCard:not(.hidden)");
  await page.fill("#queryInput", "s4:021:39");
  await page.click("#searchBtn");
  await page.waitForSelector("#results.show");
  await page.waitForTimeout(300);

  const conns = parseInt(await page.textContent("#statConns"), 10);
  const optics = parseInt(await page.textContent("#statOptics"), 10);
  const pns = parseInt(await page.textContent("#statPNs"), 10);
  const rowsInOpticTable = await page.$$eval("#opticsBody tr", (t) => t.length);
  console.log(`connections=${conns} opticsAtLocation=${optics} distinctPNs=${pns} opticRows=${rowsInOpticTable}`);

  assert.strictEqual(conns, 78, "should match 78 connections");
  assert.ok(optics > 0, "BUG: optics at location must not be 0");
  assert.strictEqual(optics, 78, "Z-OPTIC only: 78 at the searched location");
  assert.strictEqual(pns, 1, "one distinct PN");
  assert.strictEqual(rowsInOpticTable, 1, "optics table should have a row");

  await page.screenshot({ path: "scratch-fixture.png", fullPage: true });
  console.log("OK — optics table populated correctly (screenshot: scratch-fixture.png)");
} finally {
  await browser.close();
}
