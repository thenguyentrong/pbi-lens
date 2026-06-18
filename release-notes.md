Power BI reports for humans AND their AI agents (Claude, Copilot, Codex).

## What's new in 0.4.0

- **Local Power BI Desktop engine.** `run_dax` / `get_model` (and CLI `dax --local` / `model --local`) can now run DAX against a model open in Power BI Desktop — no publishing, no cloud round-trip. `list_local` discovers running instances (via a small bundled ADOMD sidecar; needs the .NET 8 runtime).
- **Visual branch diff.** New CLI `pbi-lens review --base <ref> --branch <ref> --dataset <id>` renders the same pages on two git refs and emits before/after + pixel-diff PNGs — a readable diff for PBIR changes that are otherwise unreadable JSON.
- **`screenshot_local` MCP tool.** Screenshot a local `.pbip` branch by publishing a throwaway thin report to a sandbox workspace (auto-cleaned).
- **Leaner outputs.** `list_reports` drops bulky embed URLs and gains `name` / `limit` / `offset` (with `total` / `truncated` so a cap is never silent); `list_pages` / `list_visuals` trimmed too.

**Extension**: "Power BI Reports" tree in the Explorer sidebar — sign in, browse workspaces → reports → pages, click to open live. Capture pages as PNG for AI.

**MCP server (18 tools)**: get_report_context, get_visual_fields, get_model, run_dax, get_visual_data, get_filters, screenshot_page/visual/local, set_filters/set_slicer, publish_pbix, refresh_report, list_workspaces/reports/pages/visuals, list_local.

**CLI**: context / model / fields / data / filters / shot / dax / publish / local / review — all `--json` where it makes sense.

**Install the extension**: download the .vsix below, then VS Code → Extensions → ⋯ → Install from VSIX.
**Connect your AI**: see the README section "Give your AI the plug (MCP)".
