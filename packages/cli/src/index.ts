#!/usr/bin/env node
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { execFileSync } from "child_process";
import {
  login,
  loginInteractive,
  logout,
  currentAccount,
  getAccessToken,
  loadConfig,
  saveConfig,
  PowerBiClient,
  CaptureSession,
  runDoctor,
  getModelInfo,
  getModelInfoWith,
  executeDaxLocal,
  resolveLocalPort,
  findDesktopInstances,
  buildReportContext,
  Report,
  Workspace,
  ModelInfo,
  VisualFields,
  tuneNetworkForVpn,
  resolveReportFolder,
  buildReportDefinitionParts,
  diffPng,
} from "@pbi-lens/core";

tuneNetworkForVpn();

// Read version from package.json so the dist can never drift from source.
const { version } = JSON.parse(
  fs.readFileSync(path.join(__dirname, "..", "package.json"), "utf8")
) as { version: string };

const program = new Command();
program
  .name("pbi-lens")
  .description("Screenshot, query and publish Power BI reports — built for AI agent loops")
  .version(version);

function fail(err: unknown): never {
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
}

interface TargetOpts {
  workspace?: string;
  report?: string;
}

/** Resolve workspace/report from flags or saved defaults. */
async function resolveTarget(client: PowerBiClient, opts: TargetOpts) {
  const cfg = loadConfig();
  const wsArg = opts.workspace ?? cfg.defaultWorkspace;
  const repArg = opts.report ?? cfg.defaultReport;
  if (!wsArg) fail("No workspace given. Use -w <idOrName> (-w my for My Workspace, free license) or `pbi-lens use -w <name>` to set a default.");
  const workspace = await client.resolveWorkspace(wsArg!);
  if (!repArg) return { workspace, report: undefined };
  const report = await client.resolveReport(workspace.id, repArg);
  return { workspace, report };
}

program
  .command("login")
  .description("Sign in to Power BI (device code by default; --interactive for a browser flow when Conditional Access blocks device code)")
  .option("--client-id <id>", "Entra app client id (defaults to Azure CLI well-known id)")
  .option("--tenant <id>", "Entra tenant id (defaults to 'organizations')")
  .option("-i, --interactive", "Open the system browser (auth-code flow) instead of device code")
  .action(async (opts) => {
    try {
      if (opts.clientId || opts.tenant) {
        const cfg = loadConfig();
        saveConfig({ ...cfg, clientId: opts.clientId ?? cfg.clientId, tenantId: opts.tenant ?? cfg.tenantId });
      }
      const account = opts.interactive
        ? await loginInteractive({})
        : await login({}, (msg) => console.log(`\n${msg}\n`));
      console.log(`Signed in as ${account.username}`);
    } catch (e) {
      fail(e);
    }
  });

program
  .command("doctor")
  .description("Preflight checks: browser, embed assets, offline capture pipeline, sign-in state")
  .action(async () => {
    try {
      const checks = await runDoctor();
      const symbols = { ok: "ok  ", warn: "warn", fail: "FAIL" } as const;
      for (const c of checks) console.log(`${symbols[c.status]}  ${c.name.padEnd(18)} ${c.detail}`);
      if (checks.some((c) => c.status === "fail")) process.exit(1);
    } catch (e) {
      fail(e);
    }
  });

program
  .command("logout")
  .description("Remove cached credentials")
  .action(async () => {
    await logout();
    console.log("Signed out.");
  });

program
  .command("whoami")
  .description("Show the signed-in account")
  .action(async () => {
    const account = await currentAccount();
    if (!account) fail("Not signed in. Run `pbi-lens login`.");
    console.log(account!.username);
  });

program
  .command("use")
  .description("Save a default workspace/report so other commands can omit -w/-r")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
  .option("--sandbox-ws <idOrName>", "capacity workspace `shot --local` publishes throwaway reports into")
  .action((opts) => {
    const cfg = loadConfig();
    saveConfig({
      ...cfg,
      defaultWorkspace: opts.workspace ?? cfg.defaultWorkspace,
      defaultReport: opts.report ?? cfg.defaultReport,
      sandboxWorkspace: opts.sandboxWs ?? cfg.sandboxWorkspace,
    });
    console.log("Defaults saved:", JSON.stringify(loadConfig(), null, 2));
  });

program
  .command("ls")
  .description("List workspaces, or reports/pages/visuals of a target")
  .option("-w, --workspace <idOrName>")
  .option("-W, --workspaces", "list ALL workspaces (ignores the saved default)")
  .option("-r, --report <idOrName>")
  .option("-p, --page <nameOrDisplayName>", "list visuals of this page (needs a report)")
  .option("--json", "output JSON")
  .action(async (opts) => {
    try {
      const client = new PowerBiClient();
      // -W always enumerates workspaces; otherwise we only fall back to a
      // workspace listing when there's no target AND no saved default.
      if (opts.workspaces || (!opts.workspace && !loadConfig().defaultWorkspace)) {
        const workspaces = await client.listWorkspaces();
        if (opts.json) return console.log(JSON.stringify(workspaces, null, 2));
        for (const w of workspaces) console.log(`${w.id}  ${w.name}`);
        return;
      }
      const { workspace, report } = await resolveTarget(client, opts);
      if (!report) {
        const reports = await client.listReports(workspace.id);
        if (opts.json) return console.log(JSON.stringify(reports, null, 2));
        for (const r of reports) console.log(`${r.id}  ${r.name}  (dataset ${r.datasetId})`);
        return;
      }
      if (!opts.page) {
        const pages = await client.listPages(workspace.id, report.id);
        if (opts.json) return console.log(JSON.stringify(pages, null, 2));
        for (const p of pages) console.log(`${p.name}  ${p.displayName}`);
        return;
      }
      // Visuals require a live embed session.
      const token = await getAccessToken();
      const session = await CaptureSession.open({
        accessToken: token,
        embedUrl: report.embedUrl,
        reportId: report.id,
      });
      try {
        await session.setPage(opts.page);
        const visuals = await session.getVisuals();
        if (opts.json) return console.log(JSON.stringify(visuals, null, 2));
        for (const v of visuals) {
          console.log(`${v.name}  [${v.type}]  ${v.title ?? "(untitled)"}`);
        }
      } finally {
        await session.close();
      }
    } catch (e) {
      fail(e);
    }
  });

program
  .command("find <query>")
  .description("Search reports by name across ALL workspaces (case-insensitive substring)")
  .option("--json", "output JSON")
  .action(async (query, opts) => {
    try {
      const client = new PowerBiClient();
      const hits = await client.findReports(query);
      if (opts.json) return console.log(JSON.stringify(hits, null, 2));
      if (hits.length === 0) return console.log(`No reports matching "${query}".`);
      for (const { workspace, report } of hits) {
        console.log(`${workspace.name}\t${report.name}\t(report ${report.id}, dataset ${report.datasetId})`);
      }
      console.log(`\n${hits.length} match(es)`);
    } catch (e) {
      fail(e);
    }
  });

program
  .command("shot")
  .description("Screenshot a report page (or one visual) to PNG — a published report, or a LOCAL .pbip via --local")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
  .option("--local <path>", "screenshot a LOCAL .pbip / *.Report by publishing a throwaway thin report (needs --dataset)")
  .option("--dataset <id>", "with --local: published semantic model (dataset) id that supplies the data")
  .option("--sandbox-ws <idOrName>", "with --local: capacity workspace for the throwaway report (default: saved sandbox-ws)")
  .option("-p, --page <nameOrDisplayName>", "page to capture (default: active page)")
  .option("--all-pages", "capture every page (output becomes a directory)")
  .option("--visual <nameOrTitle>", "crop to a single visual")
  .option("--filter <json>", "powerbi-client filter array (JSON) applied before capture")
  .option("--slicer <json>", 'slicer states, e.g. {"Region slicer":{"filters":[...]}}')
  .option("--width <px>", "viewport width", "1600")
  .option("--height <px>", "viewport height", "900")
  .option("-o, --out <path>", "output PNG path (or directory with --all-pages)", "report.png")
  .action(async (opts) => {
    let session: CaptureSession | undefined;
    let cleanup: (() => Promise<void>) | undefined;
    let caught: unknown;
    try {
      const client = new PowerBiClient();

      // Resolve a capture target: either a throwaway from a local .pbip, or a
      // published report. Both yield an embedUrl + reportId for the session.
      let embedUrl: string;
      let reportId: string;
      if (opts.local) {
        if (!opts.dataset)
          fail("--local needs --dataset <id> (the published semantic model that supplies the data). Find it with `pbi-lens find` or `pbi-lens model --json`.");
        const wsArg = opts.sandboxWs ?? loadConfig().sandboxWorkspace;
        if (!wsArg)
          fail('No sandbox workspace. Pass --sandbox-ws <idOrName> or save one: `pbi-lens use --sandbox-ws "<Test workspace>"`. Must be a Fabric-capacity workspace (My Workspace is not supported).');
        const folder = resolveReportFolder(opts.local);
        const ws = await client.resolveWorkspace(wsArg!);
        const parts = buildReportDefinitionParts(folder, opts.dataset);
        const name = `_pbilens_sandbox_${path.basename(folder).replace(/\.Report$/i, "")}`;
        console.error(`Publishing throwaway report "${name}" to ${ws.name} (${parts.length} parts)...`);
        reportId = await client.createReport(ws.id, name, parts);
        // Always remove the throwaway, even if capture fails.
        cleanup = async () => {
          await client.deleteReport(ws.id, reportId).then(
            () => console.error(`Cleaned up sandbox report ${reportId}.`),
            (e) => console.error(`WARNING: failed to delete sandbox report ${reportId} in ${ws.name}: ${(e as Error).message}`)
          );
        };
        const rep = await client.getReport(ws.id, reportId);
        embedUrl = rep.embedUrl;
      } else {
        const { report } = await resolveTarget(client, opts);
        if (!report) fail("No report given. Use -r <idOrName>, `pbi-lens use -r <name>`, or --local <path>.");
        embedUrl = report!.embedUrl;
        reportId = report!.id;
      }

      const token = await getAccessToken();
      session = await CaptureSession.open(
        { accessToken: token, embedUrl, reportId },
        { width: parseInt(opts.width, 10), height: parseInt(opts.height, 10) }
      );

      if (opts.filter) await session.setFilters(JSON.parse(opts.filter));
      if (opts.slicer) {
        const slicers = JSON.parse(opts.slicer) as Record<string, unknown>;
        for (const [name, state] of Object.entries(slicers)) {
          await session.setSlicer(name, state);
        }
      }

      if (opts.allPages) {
        const dir = opts.out.endsWith(".png") ? path.dirname(opts.out) : opts.out;
        fs.mkdirSync(dir, { recursive: true });
        const pages = await session.getPages();
        for (const p of pages) {
          await session.setPage(p.name);
          const file = path.join(dir, `${p.displayName.replace(/[^\w-]+/g, "_")}.png`);
          await session.screenshot(file);
          console.log(file);
        }
      } else {
        if (opts.page) await session.setPage(opts.page);
        fs.mkdirSync(path.dirname(path.resolve(opts.out)), { recursive: true });
        await session.screenshot(opts.out, opts.visual);
        console.log(path.resolve(opts.out));
      }
    } catch (e) {
      // Don't fail() here: process.exit would skip the finally block and leak
      // the throwaway report. Record it and exit only after cleanup runs.
      caught = e;
    } finally {
      await session?.close();
      if (cleanup) await cleanup();
    }
    if (caught) fail(caught);
  });

program
  .command("dax <query>")
  .description("Run a DAX query against the report's dataset (cloud executeQueries, or a local Power BI Desktop engine via --local)")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>", "resolve dataset from this report")
  .option("-d, --dataset <id>", "dataset id (overrides --report)")
  .option("--local", "run against a local Power BI Desktop engine (ADOMD) instead of the cloud")
  .option("--port <port>", "with --local: explicit msmdsrv port (default: newest running instance)")
  .option("--json", "output raw JSON rows")
  .action(async (query, opts) => {
    try {
      let result;
      if (opts.local) {
        const port = resolveLocalPort(opts.port ? parseInt(opts.port, 10) : undefined);
        console.error(`Local engine on port ${port}`);
        result = await executeDaxLocal(port, query);
      } else {
        const client = new PowerBiClient();
        let datasetId = opts.dataset as string | undefined;
        if (!datasetId) {
          const { report } = await resolveTarget(client, opts);
          if (!report) fail("Need -d <datasetId>, -w/-r, or --local.");
          datasetId = report!.datasetId;
        }
        result = await client.executeDax(datasetId!, query);
      }
      if (opts.json) return console.log(JSON.stringify(result.rows, null, 2));
      if (result.rows.length === 0) return console.log("(no rows)");
      const cols = Object.keys(result.rows[0]);
      console.log(cols.join("\t"));
      for (const row of result.rows) {
        console.log(cols.map((c) => String(row[c] ?? "")).join("\t"));
      }
      console.log(`\n${result.rows.length} row(s)`);
    } catch (e) {
      fail(e);
    }
  });

/** Open a capture session against a resolved report. */
async function openSession(report: Report): Promise<CaptureSession> {
  const token = await getAccessToken();
  return CaptureSession.open({ accessToken: token, embedUrl: report.embedUrl, reportId: report.id });
}

function printModel(model: ModelInfo): void {
  console.log(`TABLES (${model.tables.length})`);
  for (const t of model.tables) {
    const cols = model.columns.filter((c) => c.table === t.name);
    console.log(`  ${t.name}  [${t.storageMode ?? "?"}]  ${cols.length} columns`);
  }
  console.log(`\nMEASURES (${model.measures.length})`);
  for (const m of model.measures) {
    console.log(`  ${m.table}[${m.name}]${m.expression ? ` = ${m.expression}` : ""}`);
  }
  console.log(`\nRELATIONSHIPS (${model.relationships.length})`);
  for (const r of model.relationships) {
    const card = `${r.fromCardinality ?? "?"}:${r.toCardinality ?? "?"}`;
    console.log(
      `  ${r.fromTable}[${r.fromColumn}] -> ${r.toTable}[${r.toColumn}]  (${card}${r.isActive ? "" : ", INACTIVE"})`
    );
  }
  for (const w of model.warnings) console.log(`\nwarn: ${w}`);
}

function printFields(f: VisualFields): void {
  console.log(`${f.title ?? f.name}  [${f.type}]`);
  for (const role of f.roles) {
    const entries = (role.fields ?? []).filter((x) => x != null && typeof x === "object");
    if (entries.length === 0) continue;
    const fields = entries
      .map((x) => {
        if (x.measure) return `${x.table ?? "?"}[${x.measure}]`;
        if (x.column) return `${x.table ?? "?"}[${x.column}]${x.aggregationFunction ? ` (${x.aggregationFunction})` : ""}`;
        if (x.hierarchy) return `${x.table ?? "?"}[${x.hierarchy}.${x.hierarchyLevel ?? "*"}]`;
        return JSON.stringify(x);
      })
      .join(", ");
    console.log(`  ${role.role}: ${fields}`);
  }
}

program
  .command("model")
  .description("Dataset schema: tables, columns, measures, relationships (DAX INFO.VIEW.*; cloud or local Desktop via --local)")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
  .option("-d, --dataset <id>", "dataset id (overrides --report)")
  .option("--local", "read from a local Power BI Desktop engine (ADOMD) instead of the cloud")
  .option("--port <port>", "with --local: explicit msmdsrv port (default: newest running instance)")
  .option("--include-hidden", "include hidden tables/columns/measures")
  .option("--json", "output JSON")
  .action(async (opts) => {
    try {
      const includeHidden = opts.includeHidden === true;
      let model;
      if (opts.local) {
        const port = resolveLocalPort(opts.port ? parseInt(opts.port, 10) : undefined);
        console.error(`Local engine on port ${port}`);
        model = await getModelInfoWith((q) => executeDaxLocal(port, q), { includeHidden });
      } else {
        const client = new PowerBiClient();
        let datasetId = opts.dataset as string | undefined;
        if (!datasetId) {
          const { report } = await resolveTarget(client, opts);
          if (!report) fail("Need -d <datasetId>, -w/-r, or --local.");
          datasetId = report!.datasetId;
        }
        model = await getModelInfo(client, datasetId!, { includeHidden });
      }
      if (opts.json) return console.log(JSON.stringify(model, null, 2));
      printModel(model);
    } catch (e) {
      fail(e);
    }
  });

program
  .command("fields")
  .description("Field bindings of visuals: which column/measure is on which axis/role")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
  .requiredOption("-p, --page <nameOrDisplayName>", "page to inspect")
  .option("--visual <nameOrTitle>", "only this visual")
  .option("--json", "output JSON")
  .action(async (opts) => {
    let session: CaptureSession | undefined;
    try {
      const client = new PowerBiClient();
      const { report } = await resolveTarget(client, opts);
      if (!report) fail("No report given. Use -r <idOrName> or `pbi-lens use -r <name>`.");
      session = await openSession(report!);
      await session.setPage(opts.page);
      if (opts.visual) {
        const fields = await session.getVisualFields(opts.visual);
        if (opts.json) return console.log(JSON.stringify(fields, null, 2));
        return printFields(fields);
      }
      const visuals = await session.getVisuals();
      const all: (VisualFields | { name: string; title?: string; type: string; error: string })[] = [];
      for (const v of visuals) {
        try {
          all.push(await session.getVisualFields(v.name));
        } catch (e) {
          all.push({ name: v.name, title: v.title, type: v.type, error: (e as Error).message.split("\n")[0] });
        }
      }
      if (opts.json) return console.log(JSON.stringify(all, null, 2));
      for (const f of all) {
        if ("error" in f) console.log(`${f.title ?? f.name}  [${f.type}]  (no fields: ${f.error})`);
        else printFields(f);
      }
    } catch (e) {
      fail(e);
    } finally {
      await session?.close();
    }
  });

program
  .command("data")
  .description("Export a visual's data points as CSV (SDK exportData)")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
  .option("-p, --page <nameOrDisplayName>", "page containing the visual")
  .requiredOption("--visual <nameOrTitle>", "visual to export")
  .option("--rows <n>", "max rows", "1000")
  .option("--underlying", "underlying rows instead of summarized")
  .option("-o, --out <path>", "write CSV to a file instead of stdout")
  .action(async (opts) => {
    let session: CaptureSession | undefined;
    try {
      const client = new PowerBiClient();
      const { report } = await resolveTarget(client, opts);
      if (!report) fail("No report given. Use -r <idOrName> or `pbi-lens use -r <name>`.");
      session = await openSession(report!);
      if (opts.page) await session.setPage(opts.page);
      const csv = await session.exportVisualData(opts.visual, {
        exportType: opts.underlying ? "underlying" : "summarized",
        rows: parseInt(opts.rows, 10),
      });
      if (opts.out) {
        fs.writeFileSync(opts.out, csv);
        console.log(path.resolve(opts.out));
      } else {
        console.log(csv);
      }
    } catch (e) {
      fail(e);
    } finally {
      await session?.close();
    }
  });

program
  .command("filters")
  .description("Active filter state: report, page and per-visual (incl. slicers)")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
  .option("-p, --page <nameOrDisplayName>", "page to inspect (default: active page)")
  .option("--json", "compact JSON (default: pretty)")
  .action(async (opts) => {
    let session: CaptureSession | undefined;
    try {
      const client = new PowerBiClient();
      const { report } = await resolveTarget(client, opts);
      if (!report) fail("No report given. Use -r <idOrName> or `pbi-lens use -r <name>`.");
      session = await openSession(report!);
      if (opts.page) await session.setPage(opts.page);
      const state = await session.getFiltersState();
      console.log(JSON.stringify(state, null, opts.json ? 0 : 2));
    } catch (e) {
      fail(e);
    } finally {
      await session?.close();
    }
  });

program
  .command("context")
  .description("One-shot orientation pack: pages, visuals + field bindings, model schema, filters")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
  .option("-p, --page <nameOrDisplayName>", "page to detail (default: active page)")
  .option("--all-pages", "detail every page (~5-8 s per page)")
  .option("--no-fields", "skip per-visual field bindings (faster)")
  .option("--json", "output JSON")
  .action(async (opts) => {
    let session: CaptureSession | undefined;
    try {
      const client = new PowerBiClient();
      const { workspace, report } = await resolveTarget(client, opts);
      if (!report) fail("No report given. Use -r <idOrName> or `pbi-lens use -r <name>`.");
      session = await openSession(report!);
      const ctx = await buildReportContext(client, workspace, report!, session, {
        page: opts.page,
        allPages: opts.allPages === true,
        includeFields: opts.fields !== false,
      });
      if (opts.json) return console.log(JSON.stringify(ctx, null, 2));
      console.log(`REPORT ${ctx.report.name}  (workspace ${ctx.report.workspaceName}, dataset ${ctx.report.datasetId})`);
      console.log(`\nPAGES (${ctx.pages.length})`);
      for (const p of ctx.pages) console.log(`  ${p.name}  ${p.displayName}`);
      for (const d of ctx.pagesDetail) {
        console.log(`\nVISUALS on "${d.page.displayName}" (${d.visuals.length})`);
        for (const v of d.visuals) {
          console.log(`  ${v.name}  [${v.type}]  ${v.title ?? "(untitled)"}`);
          if (v.fields && !("error" in v.fields)) {
            for (const role of v.fields.roles) {
              const entries = (role.fields ?? []).filter((x) => x != null && typeof x === "object");
              if (entries.length === 0) continue;
              const fields = entries
                .map((x) => (x.measure ? `${x.table ?? "?"}[${x.measure}]` : `${x.table ?? "?"}[${x.column ?? x.hierarchy}]`))
                .join(", ");
              console.log(`      ${role.role}: ${fields}`);
            }
          }
        }
      }
      console.log("");
      printModel(ctx.model);
      for (const w of ctx.warnings) console.log(`warn: ${w}`);
    } catch (e) {
      fail(e);
    } finally {
      await session?.close();
    }
  });

program
  .command("publish <pbixPath>")
  .description("Publish a .pbix to a workspace (Imports API, CreateOrOverwrite)")
  .option("-w, --workspace <idOrName>")
  .option("--name <displayName>", "dataset/report display name (default: file name)")
  .action(async (pbixPath, opts) => {
    try {
      if (!fs.existsSync(pbixPath)) fail(`File not found: ${pbixPath}`);
      const client = new PowerBiClient();
      const { workspace } = await resolveTarget(client, { workspace: opts.workspace });
      const imp = await client.importPbix(workspace.id, pbixPath, opts.name);
      console.log(`Import started (${imp.id}), waiting...`);
      await client.waitForImport(workspace.id, imp.id);
      console.log("Published.");
    } catch (e) {
      fail(e);
    }
  });

program
  .command("local")
  .description("List running Power BI Desktop local engines (msmdsrv ports) usable with `dax --local` / `model --local`")
  .option("--json", "output JSON")
  .action((opts) => {
    const instances = findDesktopInstances();
    if (opts.json) return console.log(JSON.stringify(instances, null, 2));
    if (instances.length === 0) {
      return console.log("No running Power BI Desktop instance found. Open a .pbix/.pbip in Desktop first.");
    }
    console.log(`Running local engine(s) (newest first):`);
    for (const i of instances) console.log(`  port ${i.port}  (workspace ${i.workspace})`);
  });

function sanitize(s: string): string {
  return s.replace(/[^\w.-]+/g, "_");
}

function isGitRepo(repo: string): boolean {
  try {
    execFileSync("git", ["-C", repo, "rev-parse", "--is-inside-work-tree"], { stdio: "pipe" });
    return true;
  } catch {
    return false;
  }
}

function autodetectPbip(dir: string): string {
  const hits = fs.readdirSync(dir).filter((f) => f.toLowerCase().endsWith(".pbip"));
  if (hits.length === 1) return path.join(dir, hits[0]);
  if (hits.length === 0) fail("No .pbip at this ref's repo root. Pass --pbip <relPath>.");
  return fail(`Multiple .pbip found (${hits.join(", ")}). Pass --pbip <relPath>.`);
}

/**
 * Materialize a git ref in a throwaway worktree, publish its .pbip as a thin
 * report to the sandbox, screenshot the requested pages, then clean both up.
 * Returns displayName -> PNG path.
 */
async function captureRef(
  client: PowerBiClient,
  ws: Workspace,
  repo: string,
  ref: string,
  opts: { pbip?: string; dataset: string; width: string; height: string },
  pageFilter: string[] | undefined,
  outSub: string
): Promise<Record<string, string>> {
  const wt = fs.mkdtempSync(path.join(os.tmpdir(), "pbilens-wt-"));
  let reportId: string | undefined;
  let session: CaptureSession | undefined;
  const result: Record<string, string> = {};
  try {
    execFileSync("git", ["-C", repo, "worktree", "add", "--detach", wt, ref], { stdio: "pipe" });
    const local = opts.pbip ? path.join(wt, opts.pbip) : autodetectPbip(wt);
    const folder = resolveReportFolder(local);
    const parts = buildReportDefinitionParts(folder, opts.dataset);
    const name = `_pbilens_review_${sanitize(ref)}_${path.basename(folder).replace(/\.Report$/i, "")}`;
    console.error(`[${ref}] publishing throwaway "${name}" to ${ws.name}...`);
    reportId = await client.createReport(ws.id, name, parts);
    const rep = await client.getReport(ws.id, reportId);
    session = await CaptureSession.open(
      { accessToken: await getAccessToken(), embedUrl: rep.embedUrl, reportId },
      { width: parseInt(opts.width, 10), height: parseInt(opts.height, 10) }
    );
    fs.mkdirSync(outSub, { recursive: true });
    const allPages = await session.getPages();
    const pages = pageFilter
      ? allPages.filter((p) => pageFilter.includes(p.displayName) || pageFilter.includes(p.name))
      : allPages;
    if (pageFilter && pages.length === 0)
      fail(`[${ref}] no page matched --pages. Available: ${allPages.map((p) => p.displayName).join(", ")}`);
    for (const p of pages) {
      await session.setPage(p.name);
      const file = path.join(outSub, `${sanitize(p.displayName)}.png`);
      await session.screenshot(file);
      result[p.displayName] = file;
      console.error(`[${ref}] ${p.displayName} -> ${file}`);
    }
  } finally {
    await session?.close();
    if (reportId)
      await client
        .deleteReport(ws.id, reportId)
        .catch((e) => console.error(`[${ref}] WARN: failed to delete sandbox report ${reportId}: ${(e as Error).message}`));
    try {
      execFileSync("git", ["-C", repo, "worktree", "remove", "--force", wt], { stdio: "pipe" });
    } catch {
      fs.rmSync(wt, { recursive: true, force: true });
    }
  }
  return result;
}

program
  .command("review")
  .description("Visual diff of a report across two git refs: render the SAME pages on each, then emit before/after + pixel-diff PNGs")
  .requiredOption("--base <ref>", "git ref for the baseline (e.g. main)")
  .requiredOption("--branch <ref>", "git ref for the changed version (e.g. a feature branch)")
  .requiredOption("--dataset <id>", "published semantic model (dataset) id that supplies the data")
  .option("--repo <path>", "git repo containing the .pbip (default: current dir)", ".")
  .option("--pbip <relPath>", "path to the .pbip / *.Report within the repo (default: auto-detect a single .pbip)")
  .option("--sandbox-ws <idOrName>", "Fabric-capacity workspace for throwaway renders (default: saved sandbox-ws)")
  .option("--pages <list>", "comma-separated page display names to compare (default: all pages)")
  .option("--width <px>", "viewport width", "1600")
  .option("--height <px>", "viewport height", "900")
  .option("-o, --out <dir>", "output directory", "pbi-review")
  .action(async (opts) => {
    const repo = path.resolve(opts.repo);
    const wsArg = opts.sandboxWs ?? loadConfig().sandboxWorkspace;
    if (!wsArg)
      fail('No sandbox workspace. Pass --sandbox-ws <idOrName> or save one: `pbi-lens use --sandbox-ws "<name>"`. Must be Fabric-capacity.');
    if (!isGitRepo(repo)) fail(`Not a git repository: ${repo}. Pass --repo <path>.`);
    try {
      const client = new PowerBiClient();
      const ws = await client.resolveWorkspace(wsArg!);
      const pageFilter: string[] | undefined = opts.pages
        ? String(opts.pages).split(",").map((s: string) => s.trim()).filter(Boolean)
        : undefined;
      const outDir = path.resolve(opts.out);

      const base = await captureRef(client, ws, repo, opts.base, opts, pageFilter, path.join(outDir, sanitize(opts.base)));
      const branch = await captureRef(client, ws, repo, opts.branch, opts, pageFilter, path.join(outDir, sanitize(opts.branch)));

      const diffDir = path.join(outDir, "diff");
      fs.mkdirSync(diffDir, { recursive: true });
      const pages = [...new Set([...Object.keys(base), ...Object.keys(branch)])].sort();
      const summary: Record<string, unknown>[] = [];
      for (const pg of pages) {
        const a = base[pg];
        const b = branch[pg];
        if (!a || !b) {
          summary.push({ page: pg, status: a ? "only in base" : "only in branch" });
          continue;
        }
        const out = path.join(diffDir, `${sanitize(pg)}.png`);
        const d = diffPng(a, b, out);
        summary.push({ page: pg, changedPixels: d.changedPixels, pctChanged: +(d.fraction * 100).toFixed(2), diff: out });
      }
      fs.writeFileSync(path.join(outDir, "summary.json"), JSON.stringify(summary, null, 2));

      console.log(`\nVisual diff: ${opts.base} -> ${opts.branch}`);
      for (const s of summary) {
        if (s.status) console.log(`  ${s.page}: ${s.status}`);
        else console.log(`  ${s.page}: ${s.pctChanged}% changed (${s.changedPixels}px)  ->  ${s.diff}`);
      }
      console.log(`\nOutput: ${outDir}`);
    } catch (e) {
      fail(e);
    }
  });

program.parseAsync(process.argv).catch(fail);
