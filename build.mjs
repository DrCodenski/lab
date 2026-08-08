/* Inlines styles.css, vendor/xlsx.full.min.js and app.js into index.html to
   produce a single self-contained file the team can double-click.
   Output: dist/coreweave-cutsheet-lookup.html */
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";

const html = readFileSync("index.html", "utf8");
const css = readFileSync("styles.css", "utf8");
const xlsx = readFileSync("vendor/xlsx.full.min.js", "utf8");
const app = readFileSync("app.js", "utf8");

// Use function replacements so `$` sequences in the payloads are inserted
// verbatim (a string replacement would treat $&, $`, $', $n specially).
let out = html
  .replace(
    '<link rel="stylesheet" href="styles.css" />',
    () => "<style>\n" + css + "\n</style>"
  )
  .replace(
    '<script src="vendor/xlsx.full.min.js"></script>',
    () => "<script>\n" + xlsx + "\n</script>"
  )
  .replace(
    '<script src="app.js"></script>',
    () => "<script>\n" + app + "\n</script>"
  );

mkdirSync("dist", { recursive: true });
writeFileSync("dist/coreweave-cutsheet-lookup.html", out);
const kb = (Buffer.byteLength(out) / 1024).toFixed(0);
console.log(`Wrote dist/coreweave-cutsheet-lookup.html (${kb} KB, self-contained)`);
