#!/usr/bin/env node
import "./stdio-guard";
import * as fs from "fs";
import * as path from "path";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  CaptureSession,
  buildReportContext,
  getAccessToken,
  getModelInfo,
  getModelInfoWith,
  executeDaxLocal,
  resolveLocalPort,
  findDesktopInstances,
  compactModel,
  FiltersState,
  VisualFields,
  VisualInfo,
  ReportContext,
  Report,
  ReportPage,
  resolveReportFolder,
  buildReportDefinitionParts,
  loadConfig,
  tuneNetworkForVpn,
} from "@pbi-lens/core";
import { client, resolveWorkspace, resolveTarget, SessionManager } from "./session";

tuneNetworkForVpn();

// Read version from package.json so the dist can never drift from source.
const { version } = JSON.parse(
  fs.readFileSync(path.join(__dirname, "..", "package.json"), "utf8")
) as { version: string };

const server = new McpServer({ name: "pbi-lens", version });
const sessions = new SessionManager();

// ---------- shared result/error helpers ----------

type ToolResult = {
  content: ({ type: "text"; text: string } | { type: "image"; data: string; mimeType: string })[];
  isError?: boolean;
};

// Compact JSON, no structuredContent duplicate — every byte here lands in the
// agent's context window.
function ok(result: unknown): ToolResult {
  return { content: [{ type: "text", text: JSON.stringify(result) }] };
}

function png(buffer: Buffer, meta: Record<string, unknown>): ToolResult {
  return {
    content: [
      { type: "image", data: buffer.toString("base64"), mimeType: "image/png" },
      { type: "text", text: JSON.stringify(meta) },
    ],
  };
}

// ---------- token-diet transforms ----------

/** Bindings: drop empty roles into a name list instead of empty arrays. */
function slimFields(f: VisualFields): Record<string, unknown> {
  const bound = f.roles.filter((r) => (r.fields ?? []).some((x) => x != null));
  const empty = f.roles.filter((r) => !(r.fields ?? []).some((x) => x != null)).map((r) => r.role);
  return {
    name: f.name,
    ...(f.title ? { title: f.title } : {}),
    type: f.type,
    roles: bound.map((r) => ({ role: r.role, fields: (r.fields ?? []).filter((x) => x != null) })),
    ...(empty.length ? { emptyRoles: empty } : {}),
  };
}

/** Filter state: visuals with no filters and no slicer state are noise. */
function slimFilters(s: FiltersState): Record<string, unknown> {
  const visuals = s.visuals
    .filter((v) => (v.filters && v.filters.length > 0) || v.slicerState !== undefined)
    .map((v) => ({
      name: v.name,
      ...(v.title ? { title: v.title } : {}),
      type: v.type,
      ...(v.filters && v.filters.length ? { filters: v.filters } : {}),
      ...(v.slicerState !== undefined ? { slicerState: v.slicerState } : {}),
    }));
  return { report: s.report, page: s.page, visuals };
}

/** Context pack: compact model, slim bindings, slim filters. */
function slimContext(ctx: ReportContext): Record<string, unknown> {
  return {
    report: ctx.report,
    pages: ctx.pages.map((p) => ({ name: p.name, displayName: p.displayName })),
    model: compactModel(ctx.model),
    pagesDetail: ctx.pagesDetail.map((d) => ({
      page: { name: d.page.name, displayName: d.page.displayName },
      visuals: d.visuals.map((v) => ({
        name: v.name,
        ...(v.title ? { title: v.title } : {}),
        type: v.type,
        ...(v.layout ? { layout: v.layout } : {}),
        ...(v.fields ? { fields: "error" in v.fields ? v.fields : slimFields(v.fields).roles } : {}),
      })),
    })),
    filters: "error" in ctx.filters ? ctx.filters : slimFilters(ctx.filters),
    ...(ctx.warnings.length ? { warnings: ctx.warnings } : {}),
  };
}

/** Reports: drop the bulky embedUrl (regenerated on demand by capture) — the byte hog. */
function slimReports(reports: Report[]): Record<string, unknown>[] {
  return reports.map((r) => ({
    id: r.id,
    name: r.name,
    datasetId: r.datasetId,
    ...(r.webUrl ? { webUrl: r.webUrl } : {}),
  }));
}

/** Pages: keep order only when present. */
function slimPages(pages: ReportPage[]): Record<string, unknown>[] {
  return pages.map((p) => ({
    name: p.name,
    displayName: p.displayName,
    ...(p.order !== undefined ? { order: p.order } : {}),
  }));
}

/** Visuals: round layout to whole pixels (sub-pixel precision is just noise). */
function slimVisuals(visuals: VisualInfo[]): Record<string, unknown>[] {
  return visuals.map((v) => ({
    name: v.name,
    ...(v.title ? { title: v.title } : {}),
    type: v.type,
    ...(v.layout
      ? {
          layout: {
            x: Math.round(v.layout.x),
            y: Math.round(v.layout.y),
            width: Math.round(v.layout.width),
            height: Math.round(v.layout.height),
          },
        }
      : {}),
  }));
}

function hintFor(message: string): string {
  if (/AADSTS|no account|not signed in|interaction.?required|silent/i.test(message)) {
    return " Hint: sign in first — run `pbi-lens login --interactive` in a terminal (device code is blocked on some tenants).";
  }
  if (/403/.test(message) && /executeQueries|datasets/i.test(message)) {
    return " Hint: the signed-in user needs Build permission on the dataset (or the tenant disabled executeQueries).";
  }
  if (/exportData|export/i.test(message) && /disabled|no data|failed/i.test(message)) {
    return " Hint: data export may be disabled for this report/tenant — use run_dax instead, building the query from get_visual_fields targets.";
  }
  if (/Field readback/i.test(message)) {
    return " Hint: field readback needs edit rights on the report. Fallback: read the visual.json in the local PBIP source.";
  }
  if (/connection cannot be made|msmdsrv|local engine|ADOMD sidecar|No running Power BI Desktop/i.test(message)) {
    return " Hint: local DAX needs Power BI Desktop open with a model. Call list_local to see ports; if empty, open a .pbix/.pbip first.";
  }
  if (/capacity|Fabric|PowerBIEntityNotFound|sandbox/i.test(message)) {
    return " Hint: screenshot_local needs a Fabric-capacity sandbox workspace (My Workspace is not supported) and a valid published `dataset` id — find one with `pbi-lens find` or get_model.";
  }
  return "";
}

function tool(
  name: string,
  description: string,
  inputSchema: Record<string, z.ZodTypeAny>,
  handler: (args: Record<string, unknown>) => Promise<ToolResult>
): void {
  server.registerTool(name, { description, inputSchema }, (async (args: Record<string, unknown>) => {
    try {
      return await handler(args ?? {});
    } catch (e) {
      const message = e instanceof Error ? e.message : String(e);
      return {
        content: [{ type: "text", text: `Error: ${message}${hintFor(message)}` }],
        isError: true,
      } satisfies ToolResult;
    }
  }) as never);
}

// ---------- shared parameter shapes ----------

const targetParams = {
  workspace: z.string().optional().describe('Workspace id or name ("my" = My Workspace). Default: saved default.'),
  report: z.string().optional().describe("Report id or name. Default: saved default."),
};
const pageParam = z
  .string()
  .optional()
  .describe("Page name or display name. Default: the report's active page.");

// ---------- discovery ----------

tool("list_workspaces", "List Power BI workspaces the signed-in user can access.", {}, async () =>
  ok(await client.listWorkspaces())
);

tool(
  "list_reports",
  "List reports in a workspace (id, name, datasetId — embedUrl is omitted to stay token-lean). Use `name` to filter and `limit`/`offset` to page; the result flags `total`/`truncated` so a cap is never silent.",
  {
    workspace: targetParams.workspace,
    name: z.string().optional().describe("Case-insensitive substring filter on report name."),
    limit: z.number().int().positive().optional().describe("Max reports to return (default 50)."),
    offset: z.number().int().nonnegative().optional().describe("Skip this many matches before returning (default 0)."),
  },
  async (args) => {
    const workspace = await resolveWorkspace(args.workspace as string | undefined);
    const all = await client.listReports(workspace.id);
    const needle = (args.name as string | undefined)?.toLowerCase();
    const matched = needle ? all.filter((r) => r.name.toLowerCase().includes(needle)) : all;
    const offset = (args.offset as number | undefined) ?? 0;
    const limit = (args.limit as number | undefined) ?? 50;
    const page = matched.slice(offset, offset + limit);
    return ok({
      workspace: workspace.name,
      total: matched.length,
      offset,
      count: page.length,
      ...(offset + page.length < matched.length ? { truncated: true } : {}),
      reports: slimReports(page),
    });
  }
);

tool("list_pages", "List pages of a report.", targetParams, async (args) => {
  const { workspace, report } = await resolveTarget(args);
  return ok(slimPages(await client.listPages(workspace.id, report.id)));
});

tool(
  "list_visuals",
  "List visuals on a report page (name, title, type, layout).",
  { page: pageParam, ...targetParams },
  async (args) => {
    const { workspace, report } = await resolveTarget(args);
    return await sessions.withSession(workspace, report, async (s) => {
      if (args.page) await s.setPage(args.page as string);
      return ok(slimVisuals(await s.getVisuals()));
    });
  }
);

// ---------- structured info (what agents usually need) ----------

tool(
  "get_report_context",
  "Call this FIRST to orient on a report: pages, visuals with field bindings, model schema (tables/columns/measures/relationships) and active filters in one call. Detail covers one page by default; set all_pages for everything (~5-8 s per page).",
  {
    page: pageParam,
    all_pages: z.boolean().optional().describe("Detail every page, not just one (slow on big reports)."),
    include_fields: z.boolean().optional().describe("Include per-visual field bindings (default true)."),
    ...targetParams,
  },
  async (args) => {
    const { workspace, report } = await resolveTarget(args);
    return await sessions.withSession(workspace, report, async (s) =>
      ok(
        slimContext(
          await buildReportContext(client, workspace, report, s, {
            page: args.page as string | undefined,
            allPages: args.all_pages === true,
            includeFields: args.include_fields !== false,
          })
        )
      )
    );
  }
);

tool(
  "get_model",
  "Dataset schema: tables, columns, measures and relationships (via DAX INFO.VIEW.*). Auto date/time noise is filtered out. Default output is compact; detail=full adds per-column format/sort/category metadata. Set `local` to read from a local Power BI Desktop engine instead of the cloud.",
  {
    include_hidden: z.boolean().optional().describe("Include hidden tables/columns/measures."),
    detail: z.enum(["compact", "full"]).optional().describe("compact (default, ~5x fewer tokens) or full."),
    local: z.boolean().optional().describe("Read from a local Power BI Desktop engine (ADOMD) instead of the cloud."),
    port: z.number().int().optional().describe("With local: explicit msmdsrv port (default: newest running Desktop instance)."),
    ...targetParams,
  },
  async (args) => {
    const includeHidden = args.include_hidden === true;
    let model;
    if (args.local) {
      const port = resolveLocalPort(args.port as number | undefined);
      model = await getModelInfoWith((q) => executeDaxLocal(port, q), { includeHidden });
    } else {
      const { report } = await resolveTarget(args);
      model = await getModelInfo(client, report.datasetId, { includeHidden });
    }
    return ok(args.detail === "full" ? model : compactModel(model));
  }
);

tool(
  "get_visual_fields",
  "Field bindings of one visual: which table column or measure sits in which data role (axis, legend, values...).",
  {
    visual: z.string().describe("Visual name or title (see list_visuals)."),
    page: pageParam,
    ...targetParams,
  },
  async (args) => {
    const { workspace, report } = await resolveTarget(args);
    return await sessions.withSession(workspace, report, async (s) => {
      if (args.page) await s.setPage(args.page as string);
      return ok(slimFields(await s.getVisualFields(args.visual as string)));
    });
  }
);

tool(
  "get_visual_data",
  "Data points of one visual as CSV (summarized by default). If export is disabled, use run_dax with the visual's fields instead.",
  {
    visual: z.string().describe("Visual name or title."),
    page: pageParam,
    rows: z.number().int().min(1).max(30000).optional().describe("Max rows (default 1000)."),
    export_type: z.enum(["summarized", "underlying"]).optional(),
    ...targetParams,
  },
  async (args) => {
    const { workspace, report } = await resolveTarget(args);
    return await sessions.withSession(workspace, report, async (s) => {
      if (args.page) await s.setPage(args.page as string);
      const csv = await s.exportVisualData(args.visual as string, {
        exportType: (args.export_type as "summarized" | "underlying" | undefined) ?? "summarized",
        rows: (args.rows as number | undefined) ?? 1000,
      });
      return { content: [{ type: "text", text: csv }] } satisfies ToolResult;
    });
  }
);

tool(
  "get_filters",
  "Active filter state: report level, page level and per visual (slicer selections included; visuals without filters are omitted).",
  { page: pageParam, ...targetParams },
  async (args) => {
    const { workspace, report } = await resolveTarget(args);
    return await sessions.withSession(workspace, report, async (s) => {
      if (args.page) await s.setPage(args.page as string);
      return ok(slimFilters(await s.getFiltersState()));
    });
  }
);

tool(
  "run_dax",
  "Run a DAX query (executeQueries; Pro limits: 120/min, 100k rows). Use for exact numbers instead of reading pixels. Set `local` to run against a local Power BI Desktop engine instead of the cloud — see list_local.",
  {
    query: z.string().describe('DAX query, e.g. EVALUATE SUMMARIZECOLUMNS(...)'),
    dataset: z.string().optional().describe("Dataset id (default: the target report's dataset)."),
    local: z.boolean().optional().describe("Run against a local Power BI Desktop engine (ADOMD) instead of the cloud."),
    port: z.number().int().optional().describe("With local: explicit msmdsrv port (default: newest running Desktop instance)."),
    ...targetParams,
  },
  async (args) => {
    if (args.local) {
      const port = resolveLocalPort(args.port as number | undefined);
      return ok(await executeDaxLocal(port, args.query as string));
    }
    let datasetId = args.dataset as string | undefined;
    if (!datasetId) {
      const { report } = await resolveTarget(args);
      datasetId = report.datasetId;
    }
    const result = await client.executeDax(datasetId, args.query as string);
    return ok(result);
  }
);

tool(
  "list_local",
  "List running Power BI Desktop local engines (msmdsrv ports) usable by run_dax/get_model with local=true.",
  {},
  async () => ok(findDesktopInstances())
);

// ---------- seeing the report ----------

tool(
  "screenshot_page",
  "Render a report page headlessly and return the PNG so you can SEE it. ~4-8 s.",
  {
    page: pageParam,
    width: z.number().int().optional().describe("Viewport width (default 1600; forces a fresh render session)."),
    height: z.number().int().optional().describe("Viewport height (default 900; forces a fresh render session)."),
    ...targetParams,
  },
  async (args) => {
    const { workspace, report } = await resolveTarget(args);
    const meta = { report: report.name, page: (args.page as string) ?? "(active)" };
    if (args.width || args.height) {
      // Custom viewport: one-off session so the warm one keeps its geometry.
      const oneOff = await CaptureSession.open(
        { accessToken: await getAccessToken(), embedUrl: report.embedUrl, reportId: report.id },
        { width: (args.width as number) ?? 1600, height: (args.height as number) ?? 900 }
      );
      try {
        if (args.page) await oneOff.setPage(args.page as string);
        return png(await oneOff.screenshotBuffer(), meta);
      } finally {
        await oneOff.close();
      }
    }
    return await sessions.withSession(workspace, report, async (s) => {
      if (args.page) await s.setPage(args.page as string);
      return png(await s.screenshotBuffer(), meta);
    });
  }
);

tool(
  "screenshot_visual",
  "Render one visual (cropped) and return the PNG.",
  {
    visual: z.string().describe("Visual name or title."),
    page: pageParam,
    ...targetParams,
  },
  async (args) => {
    const { workspace, report } = await resolveTarget(args);
    return await sessions.withSession(workspace, report, async (s) => {
      if (args.page) await s.setPage(args.page as string);
      return png(await s.screenshotBuffer(args.visual as string), {
        report: report.name,
        visual: args.visual,
      });
    });
  }
);

tool(
  "screenshot_local",
  "Screenshot a LOCAL .pbip / *.Report branch: publishes a throwaway thin report (auto-deleted after) and renders it. Needs `dataset` (a published semantic model that supplies the data) and a Fabric-capacity `sandbox_ws`. NOTE: not a pure-offline render — it round-trips through a sandbox workspace. ~5-10 s.",
  {
    local_path: z.string().describe("Path to a local .pbip file or *.Report folder."),
    dataset: z.string().describe("Published semantic model (dataset) id that supplies the data."),
    sandbox_ws: z
      .string()
      .optional()
      .describe("Fabric-capacity workspace for the throwaway report. Default: saved sandbox workspace."),
    page: pageParam,
    visual: z.string().optional().describe("Crop to a single visual (name or title)."),
    filters: z
      .array(z.record(z.string(), z.unknown()))
      .optional()
      .describe("powerbi-client filter array applied before capture."),
    width: z.number().int().optional().describe("Viewport width (default 1600)."),
    height: z.number().int().optional().describe("Viewport height (default 900)."),
  },
  async (args) => {
    const localPath = args.local_path as string;
    if (!fs.existsSync(localPath)) throw new Error(`Local path not found: ${localPath}`);
    const wsArg = (args.sandbox_ws as string | undefined) ?? loadConfig().sandboxWorkspace;
    if (!wsArg) {
      throw new Error(
        'No sandbox workspace. Pass `sandbox_ws`, or save one with `pbi-lens use --sandbox-ws "<name>"`. Must be a Fabric-capacity workspace (My Workspace is not supported).'
      );
    }
    const folder = resolveReportFolder(localPath);
    const ws = await client.resolveWorkspace(wsArg);
    const parts = buildReportDefinitionParts(folder, args.dataset as string);
    const name = `_pbilens_sandbox_${path.basename(folder).replace(/\.Report$/i, "")}`;
    // Disposable report → its own one-off CaptureSession, never the warm pool.
    const reportId = await client.createReport(ws.id, name, parts);
    let session: CaptureSession | undefined;
    try {
      const rep = await client.getReport(ws.id, reportId);
      session = await CaptureSession.open(
        { accessToken: await getAccessToken(), embedUrl: rep.embedUrl, reportId },
        { width: (args.width as number) ?? 1600, height: (args.height as number) ?? 900 }
      );
      if (args.filters) await session.setFilters(args.filters as unknown[]);
      if (args.page) await session.setPage(args.page as string);
      const buf = await session.screenshotBuffer(args.visual as string | undefined);
      return png(buf, {
        source: "local",
        local: localPath,
        sandbox: ws.name,
        page: (args.page as string) ?? "(active)",
        ...(args.visual ? { visual: args.visual } : {}),
      });
    } finally {
      // Always remove the throwaway, even if capture failed.
      await session?.close();
      await client.deleteReport(ws.id, reportId).catch(() => {});
    }
  }
);

// ---------- interaction ----------

tool(
  "set_filters",
  "Replace the report-level filters (powerbi-client filter objects). Pass [] to clear. State persists in the warm session, so a following screenshot shows it.",
  {
    filters: z.array(z.record(z.string(), z.unknown())).describe("powerbi-client filter array (IBasicFilter etc.)."),
    ...targetParams,
  },
  async (args) => {
    const { workspace, report } = await resolveTarget(args);
    return await sessions.withSession(workspace, report, async (s) => {
      const filters = args.filters as unknown[];
      if (filters.length === 0) await s.clearFilters();
      else await s.setFilters(filters);
      return ok(slimFilters(await s.getFiltersState()));
    });
  }
);

tool(
  "set_slicer",
  "Set a slicer visual's state, e.g. {\"filters\":[...]}. State persists in the warm session.",
  {
    visual: z.string().describe("Slicer visual name or title."),
    state: z.record(z.string(), z.unknown()).describe("Slicer state object."),
    page: pageParam,
    ...targetParams,
  },
  async (args) => {
    const { workspace, report } = await resolveTarget(args);
    return await sessions.withSession(workspace, report, async (s) => {
      if (args.page) await s.setPage(args.page as string);
      await s.setSlicer(args.visual as string, args.state);
      return ok(slimFilters(await s.getFiltersState()));
    });
  }
);

tool(
  "publish_pbix",
  "Publish a local .pbix to a workspace (CreateOrOverwrite) and wait until the import succeeds.",
  {
    pbix_path: z.string().describe("Absolute path to the .pbix file."),
    name: z.string().optional().describe("Display name (default: file name)."),
    workspace: targetParams.workspace,
  },
  async (args) => {
    const pbixPath = args.pbix_path as string;
    if (!fs.existsSync(pbixPath)) throw new Error(`File not found: ${pbixPath}`);
    const workspace = await resolveWorkspace(args.workspace as string | undefined);
    const imp = await client.importPbix(workspace.id, pbixPath, args.name as string | undefined);
    await client.waitForImport(workspace.id, imp.id);
    // The warm embed still shows the pre-publish report — drop it so the next
    // screenshot/query re-embeds the fresh version.
    await sessions.invalidate();
    return ok({ importId: imp.id, status: "Succeeded", workspace: workspace.name, session: "refreshed" });
  }
);

tool(
  "refresh_report",
  "Reload the report in the warm session (call after the report was republished outside this server, e.g. from Power BI Desktop, so screenshots and queries reflect the new version).",
  {},
  async () => {
    await sessions.invalidate();
    return ok({ session: "refreshed", note: "Next tool call re-embeds the current published version." });
  }
);

// ---------- lifecycle ----------

async function shutdown(): Promise<void> {
  await sessions.dispose().catch(() => {});
  process.exit(0);
}
process.on("SIGINT", () => void shutdown());
process.on("SIGTERM", () => void shutdown());

async function main(): Promise<void> {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // connect() owns transport.onclose; the protocol-level hook fires when the
  // client disconnects (stdin EOF) — clean up the headless browser then.
  server.server.onclose = () => void shutdown();
  console.error("pbi-lens MCP server running (stdio)");
}

main().catch((e) => {
  console.error("Fatal:", e);
  process.exit(1);
});
