// Copies the embed page and the powerbi-client browser bundle into media/
// so both the embed host (Playwright) and the VS Code webview can serve them.
const fs = require("fs");
const path = require("path");

const pkgRoot = path.join(__dirname, "..");
const mediaDir = path.join(pkgRoot, "media");
fs.mkdirSync(mediaDir, { recursive: true });

fs.copyFileSync(
  path.join(pkgRoot, "src", "embed.html"),
  path.join(mediaDir, "embed.html")
);

const pbiDist = path.dirname(require.resolve("powerbi-client/dist/powerbi.min.js"));
fs.copyFileSync(
  path.join(pbiDist, "powerbi.min.js"),
  path.join(mediaDir, "powerbi.min.js")
);

// Report-authoring bundle (field-binding readback). Some releases ship only
// the unminified dist; either way it lands under the stable .min.js name
// that embed.html, the embed host and the webview rewrite all expect.
let authoringSrc;
try {
  authoringSrc = require.resolve("powerbi-report-authoring/dist/powerbi-report-authoring.min.js");
} catch {
  authoringSrc = require.resolve("powerbi-report-authoring/dist/powerbi-report-authoring.js");
}
fs.copyFileSync(authoringSrc, path.join(mediaDir, "powerbi-report-authoring.min.js"));

console.log("assets copied to", mediaDir);
