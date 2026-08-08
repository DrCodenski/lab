/* Drives the built single-file app in Chromium: loads the sample cutsheet,
   searches s1:05:30, and screenshots the result (light + dark). */
import { chromium } from "playwright";
import path from "node:path";
import { pathToFileURL } from "node:url";

const appUrl = pathToFileURL(path.resolve("dist/coreweave-cutsheet-lookup.html")).href;
const samplePath = path.resolve("sample/sample-cutsheet.xlsx");

const browser = await chromium.launch({
  executablePath: "/opt/pw-browsers/chromium",
  headless: true,
  args: ["--headless=new", "--no-sandbox"],
});
try {
  for (const theme of ["light", "dark"]) {
    const page = await browser.newPage({ viewport: { width: 1200, height: 1600 }, deviceScaleFactor: 2 });
    page.on("console", (m) => console.log(`  [console:${m.type()}]`, m.text()));
    page.on("pageerror", (e) => console.log("  [pageerror]", e.message));
    await page.goto(appUrl);
    await page.evaluate((t) => { document.documentElement.setAttribute("data-theme", t); }, theme);
    await page.setInputFiles("#fileInput", samplePath);
    await page.waitForSelector("#searchCard:not(.hidden)", { timeout: 5000 });
    await page.fill("#queryInput", "s1:05:30");
    await page.click("#searchBtn");
    await page.waitForSelector("#results.show");
    await page.waitForTimeout(400);
    const out = `scratch-screenshot-${theme}.png`;
    await page.screenshot({ path: out, fullPage: true });
    // Report a couple of on-screen numbers so we know it truly rendered.
    const conns = await page.textContent("#statConns");
    const optics = await page.textContent("#statOptics");
    const pns = await page.textContent("#statPNs");
    console.log(`[${theme}] connections=${conns} optics=${optics} distinctPNs=${pns} -> ${out}`);
    await page.close();
  }
} finally {
  await browser.close();
}
