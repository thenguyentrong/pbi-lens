# PBI Lens — agent guide

This repo gives you (Claude, Copilot, Codex — any agent) real eyes AND structured info on Power BI reports. **Never ask the user to describe a report, screenshot it, or read a formula to you — get it yourself.**

## You have an MCP server: `pbi-lens` — prefer its tools

If the `pbi-lens` MCP server is connected (it's registered user-wide on this machine), use it instead of the CLI. What you usually need → which tool:

| You need | Tool |
|---|---|
| Orient on a report (pages, visuals, bindings, schema, filters — one call) | `get_report_context` |
| Which column/measure is on which axis/role of a visual | `get_visual_fields` |
| Tables, columns, measures, relationships | `get_model` |
| Exact numbers (never squint at pixels) | `run_dax` |
| A visual's data points as CSV | `get_visual_data` |
| Active filters / slicer selections | `get_filters` |
| See a page or visual (returns a PNG you can view) | `screenshot_page`, `screenshot_visual` |
| Change state before a shot | `set_filters`, `set_slicer` |
| Ship an edited .pbix | `publish_pbix` |
| Discover targets | `list_workspaces`, `list_reports`, `list_pages`, `list_visuals` |

All tools default to the saved workspace/report (`pbi-lens use -w … -r …`), so you can usually omit `workspace`/`report` params.

## The loop

```
get_report_context → analyze (get_visual_fields / run_dax for specifics)
→ edit PBIP/PBIR JSON → publish_pbix → screenshot_page → verify → done
```

Report source lives as a `.pbip` project: `definition/pages/<page>/visuals/<visual>/visual.json`. Desktop does NOT hot-reload external edits — close/reopen, or compile+publish (`pbi-tools compile` → `publish_pbix`). Always re-screenshot and confirm the change is visible before declaring done.

## CLI fallback (when MCP isn't connected)

`node C:\pbi-lens\packages\cli\dist\index.js <cmd>` (or `pbi-lens` if linked):

| Need | Command |
|---|---|
| Orientation pack | `context [-p "<page>"] [--all-pages] [--json]` |
| Field bindings | `fields -p "<page>" [--visual "<title>"] [--json]` |
| Model schema | `model [--json]` |
| Data points CSV | `data --visual "<title>" [-p "<page>"] [--rows n]` |
| Filter state | `filters [-p "<page>"]` |
| Exact numbers | `dax "EVALUATE …"` |
| Screenshots | `shot -p "<page>" [-o x.png] [--all-pages] [--visual t] [--filter json] [--slicer json]` |
| Discover | `ls [-w X] [-r Y] [-p Z]` |
| Publish | `publish out.pbix -w <ws>` |

## Known limitations (this tenant) and fallbacks

- **Measure DAX expressions come back BLANK** via live readback (INFO.VIEW.MEASURES returns empty Expression; INFO.MEASURES is blocked by the tenant). `get_model` warns when this happens. Fallback: read the local .pbip source (`definition\**\*.tmdl` or `model.bim`) for measure definitions.
- **Field readback may flip the embed to edit mode** (report-authoring APIs) — harmless; screenshots automatically switch back to view mode. It needs edit rights on the report; cards/slicers may legitimately return no fields.
- **`get_visual_data` (exportData) may be disabled by tenant/report settings.** Fallback: `run_dax` with a `SUMMARIZECOLUMNS` built from the `get_visual_fields` targets — same data, always allowed when Build permission exists.
- Auto date/time noise (`LocalDateTable_*`, `DateTableTemplate_*`) is already filtered out of `get_model`.

## Gotchas

- First run on a new machine → `pbi-lens doctor` (checks browser, embed assets incl. the report-authoring bundle, capture pipeline, sign-in).
- `Not signed in` → the user must run `pbi-lens login --interactive` (this tenant blocks device code via Conditional Access). Work/school account with Pro; free license works only with `-w my` (My Workspace).
- 403 on executeQueries → user lacks Build permission on the dataset, or tenant disabled the setting.
- Renders take 4–8 s/page. The MCP server keeps one warm browser session per report (10 min idle), so consecutive tool calls are fast; filter/slicer state persists across calls in that session.
- Capture drives installed Edge/Chrome headlessly (no browser download).
- Local model work without publishing: if Power BI Desktop is open, its local Analysis Services port is in `%LocalAppData%\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\*\Data\msmdsrv.port.txt` (see mcpbi).
