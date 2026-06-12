# PBI Lens

View Power BI reports **inside VS Code** — and let Claude (or any AI agent) actually *see* them: screenshot pages, set filters/slicers, query the data with DAX, edit the report source, and republish. Built for fully automated agent loops. Works with a **free** Power BI license via My Workspace (`-w my`); Pro only needed for shared workspaces, Premium never.

## Costs

Sign-in, tokens and all API calls used here are free of charge. The only thing Microsoft charges for is a **Pro license**, and that's only required for *shared* workspaces. The free path:

1. Sign up for a free Power BI license with any work/school account (e.g. a university account) at [powerbi.microsoft.com](https://powerbi.microsoft.com) — no credit card.
2. Use `-w my` everywhere (or `pbi-lens use -w my` once). Publish, screenshot, DAX — all against your personal My Workspace.

## How it works

Power BI Desktop can't be screenshotted reliably (WebView2 canvas) and the export-to-PNG API needs Premium. So PBI Lens uses the Pro-compatible path:

1. You publish your report to a Power BI Service workspace (one click in Desktop, or `pbi-lens publish`).
2. PBI Lens embeds it via the `powerbi-client` SDK ("user owns data", your own AAD token) in a local page.
3. **For you:** a VS Code webview shows the live interactive report.
4. **For Claude:** headless Chromium renders the same page; the CLI screenshots it to PNG, which Claude reads as an image.

```
packages/core     auth (MSAL device code), Power BI REST, embed host, Playwright capture
packages/cli      `pbi-lens` CLI — the agent-facing surface
packages/vscode   VS Code extension — the human-facing viewer
```

## Install the extension (any PC)

1. Download `pbi-lens-vscode-<version>.vsix` from this repo's [Releases](https://github.com/thenguyentrong/pbi-lens/releases) page.
2. In VS Code: **Extensions panel → `⋯` menu → Install from VSIX...** (or `code --install-extension pbi-lens-vscode-0.1.0.vsix`).
3. Run **PBI Lens: Sign in to Power BI** from the Command Palette, then **PBI Lens: Open Report**.

Requirements on the target PC: VS Code, Microsoft Edge or Chrome (Edge ships with Windows — used headlessly for captures), and a Power BI work account with a Pro license. No Node.js needed for the extension.

To rebuild the .vsix: `npm run package:vsix`.

## Dev setup (this repo, CLI included)

```bash
npm install
npm run build
node packages/cli/dist/index.js login    # device code sign-in (work account with Pro)
```

Optionally link the CLI globally: `npm link -w pbi-lens` → `pbi-lens` on PATH. Captures use your installed Edge/Chrome — no browser download.

Auth uses the Azure CLI well-known client id by default (pre-consented in most tenants). If your tenant blocks it, register an Entra app (public client, delegated Power BI scopes) and run `pbi-lens login --client-id <id> --tenant <tenantId>`.

## CLI

```bash
pbi-lens doctor                              # preflight: browser, assets, offline capture test, sign-in state
pbi-lens ls                                  # workspaces
pbi-lens ls -w "My Workspace"                # reports
pbi-lens ls -w "My Workspace" -r "Sales"     # pages
pbi-lens ls -w "My Workspace" -r "Sales" -p "Overview"   # visuals
pbi-lens use -w my -r "Sales"                # save defaults, then omit -w/-r ("my" = My Workspace, free license)

pbi-lens shot -p "Overview" -o overview.png
pbi-lens shot --all-pages -o shots/
pbi-lens shot -p "Overview" --visual "Revenue by Month" -o rev.png
pbi-lens shot -p "Overview" --filter '[{"$schema":"http://powerbi.com/product/schema#basic","target":{"table":"Region","column":"Name"},"operator":"In","values":["EMEA"],"filterType":1}]' -o emea.png

pbi-lens dax "EVALUATE SUMMARIZECOLUMNS(Region[Name], \"Rev\", [Total Revenue])"
pbi-lens publish dist/report.pbix -w "My Workspace"
```

## VS Code extension

F5 in this repo launches the Extension Development Host. Commands:

- **PBI Lens: Sign in to Power BI**
- **PBI Lens: Open Report** — live interactive report in a webview panel
- **PBI Lens: Capture Page for Claude (PNG)** — saves to `.pbi-shots/` in your workspace

Settings `pbiLens.workspace` / `pbiLens.report` skip the pickers.

## The Claude loop

See [CLAUDE.md](./CLAUDE.md). Short version: Claude runs `pbi-lens shot`, Reads the PNG, analyzes, optionally edits the PBIP/PBIR report source, republishes, re-shoots, and verifies — no human eyes required.

## Local Desktop DAX (optional)

While a .pbix is open, Power BI Desktop runs a local Analysis Services instance. For model iteration without publishing, install the open-source [mcpbi](https://github.com/jonaolden/mcpbi) MCP server, which connects to it and gives Claude direct local DAX. Port discovery: `%LocalAppData%\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\*\Data\msmdsrv.port.txt`.

## Limits to know

- `executeQueries`: 120 queries/min, 100k rows, 15 MB per query (Pro limits).
- Renders take ~4–8 s per page; the capture session reuses one browser for multi-page shots.
- Desktop does **not** hot-reload external PBIP edits — close/reopen the project after Claude edits report JSON.
