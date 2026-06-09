# PBI Lens — agent guide

This repo gives you (Claude) real eyes on Power BI reports. Never ask the user to describe what a report looks like — capture it yourself.

## Seeing a report

```bash
node packages/cli/dist/index.js shot -w <workspace> -r <report> -p "<page>" -o /tmp/page.png
```

(or just `pbi-lens ...` if linked). Then **Read the PNG** — you can view images. Defaults set via `pbi-lens use -w ... -r ...` let you omit `-w/-r`.

- All pages at once: `shot --all-pages -o shots/`
- One visual only: `shot -p "<page>" --visual "<visual title>" -o v.png`
- Discover structure first: `ls` (workspaces) → `ls -w X` (reports) → `ls -w X -r Y` (pages) → `ls -w X -r Y -p Z` (visuals, includes layout + type)

## Interacting before capture

- `--filter '<json>'` — powerbi-client filter array (basic filter: `{"$schema":"http://powerbi.com/product/schema#basic","target":{"table":"T","column":"C"},"operator":"In","values":["v"],"filterType":1}`)
- `--slicer '{"<visual title>":{"filters":[...basic filters...]}}'`

After changing state the CLI waits for the SDK `rendered` event, so the PNG always reflects the applied state.

## Exact numbers (don't squint at pixels)

```bash
pbi-lens dax "EVALUATE SUMMARIZECOLUMNS('Table'[Col], \"Measure\", [Measure])" -w <ws> -r <report>
```

Use DAX for values, screenshots for layout/visual checks. Limits: 120 q/min, 100k rows.

## Editing the report (PBIP/PBIR)

Report source lives as a `.pbip` project (JSON): `definition/pages/<page>/visuals/<visual>/visual.json`, validated by `$schema`. Workflow:

1. Edit the JSON (visual config, titles, colors, layout, page order).
2. The user reopens the project in Power BI Desktop (no hot reload) **or** you compile+publish:
   - `pbi-tools compile` → `.pbix` → `pbi-lens publish out.pbix -w <ws>` (best-effort; if compile fails on new-format projects, ask the user to click Publish in Desktop).
3. Re-shoot the changed page and **verify your edit is visible** before declaring done.

## The full loop

```
shot → Read PNG → analyze → (dax for exact values) → edit PBIR JSON → publish → shot → verify
```

## Local model work (no publish round-trip)

If Power BI Desktop is open, its local Analysis Services instance is reachable; the mcpbi MCP server (if registered) gives you `run_query` etc. against it. Find the port in `%LocalAppData%\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\*\Data\msmdsrv.port.txt`.

## Gotchas

- `Not signed in` → ask the user to run `pbi-lens login` (interactive device code; you can't complete it for them).
- 403 on executeQueries → user lacks Build permission on the dataset, or tenant disabled the executeQueries setting.
- Empty workspace list → account has no Pro license or no workspace access.
- Renders take 4–8 s; `--all-pages` reuses one browser session, prefer it over per-page invocations.
