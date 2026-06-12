// Builds a self-contained .vsix for the PBI Lens extension.
// The extension code (+ @pbi-lens/core, msal, express) is bundled into one
// file with esbuild; playwright-core stays a real node_modules dependency
// because it loads its own files at runtime. Browsers are NOT shipped —
// capture drives the locally installed Edge/Chrome.
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const repoRoot = path.join(__dirname, "..");
const staging = path.join(repoRoot, "dist-vsix");
const corePkg = require(path.join(repoRoot, "packages", "core", "package.json"));
const extPkg = require(path.join(repoRoot, "packages", "vscode", "package.json"));

const run = (cmd, cwd) => execSync(cmd, { cwd, stdio: "inherit" });

fs.rmSync(staging, { recursive: true, force: true });
fs.mkdirSync(path.join(staging, "dist"), { recursive: true });

// 1. Bundle extension + workspace deps into a single CJS file.
run(
  [
    "npx esbuild packages/vscode/src/extension.ts",
    "--bundle",
    "--platform=node",
    "--target=node18",
    "--format=cjs",
    "--external:vscode",
    "--external:playwright-core",
    `--outfile=${path.join(staging, "dist", "extension.js")}`,
  ].join(" "),
  repoRoot
);

// 2. Ship the embed page + powerbi-client bundle next to dist/.
fs.cpSync(path.join(repoRoot, "packages", "core", "media"), path.join(staging, "media"), {
  recursive: true,
});

// 3. Manifest: same contributes, only playwright-core as a runtime dep.
const manifest = {
  ...extPkg,
  main: "./dist/extension.js",
  scripts: undefined,
  devDependencies: undefined,
  dependencies: { "playwright-core": corePkg.dependencies["playwright-core"] },
  repository: { type: "git", url: "https://github.com/thenguyentrong/pbi-lens.git" },
};
fs.writeFileSync(path.join(staging, "package.json"), JSON.stringify(manifest, null, 2));
fs.copyFileSync(path.join(repoRoot, "README.md"), path.join(staging, "README.md"));
fs.copyFileSync(path.join(repoRoot, "LICENSE"), path.join(staging, "LICENSE"));
fs.copyFileSync(path.join(repoRoot, "packages", "vscode", "icon.png"), path.join(staging, "icon.png"));

// 4. Install the single runtime dep and package.
run("npm install --omit=dev --no-audit --no-fund", staging);
const vsixName = `${extPkg.name}-${extPkg.version}.vsix`;
run(`npx --yes @vscode/vsce package --allow-missing-repository -o ${path.join(repoRoot, vsixName)}`, staging);

console.log(`\nVSIX written to ${path.join(repoRoot, vsixName)}`);
