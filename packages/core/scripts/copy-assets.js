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

console.log("assets copied to", mediaDir);
