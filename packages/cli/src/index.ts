#!/usr/bin/env node
import { Command } from "commander";
import * as fs from "fs";
import * as path from "path";
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
  buildReportContext,
  Report,
  ModelInfo,
  VisualFields,
  tuneNetworkForVpn,
} from "@pbi-lens/core";

tuneNetworkForVpn();

const program = new Command();
program
  .name("pbi-lens")
  .description("Screenshot, query and publish Power BI reports — built for AI agent loops")
  .version("0.2.0");

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
  .action((opts) => {
    const cfg = loadConfig();
    saveConfig({
      ...cfg,
      defaultWorkspace: opts.workspace ?? cfg.defaultWorkspace,
      defaultReport: opts.report ?? cfg.defaultReport,
    });
    console.log("Defaults saved:", JSON.stringify(loadConfig(), null, 2));
  });

program
  .command("ls")
  .description("List workspaces, or reports/pages/visuals of a target")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
  .option("-p, --page <nameOrDisplayName>", "list visuals of this page (needs a report)")
  .option("--json", "output JSON")
  .action(async (opts) => {
    try {
      const client = new PowerBiClient();
      if (!opts.workspace && !loadConfig().defaultWorkspace) {
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
  .command("shot")
  .description("Screenshot a report page (or one visual) to PNG")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
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
    try {
      const client = new PowerBiClient();
      const { report } = await resolveTarget(client, opts);
      if (!report) fail("No report given. Use -r <idOrName> or `pbi-lens use -r <name>`.");
      const token = await getAccessToken();
      session = await CaptureSession.open(
        { accessToken: token, embedUrl: report!.embedUrl, reportId: report!.id },
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
      fail(e);
    } finally {
      await session?.close();
    }
  });

program
  .command("dax <query>")
  .description("Run a DAX query against the report's dataset (executeQueries REST API)")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>", "resolve dataset from this report")
  .option("-d, --dataset <id>", "dataset id (overrides --report)")
  .option("--json", "output raw JSON rows")
  .action(async (query, opts) => {
    try {
      const client = new PowerBiClient();
      let datasetId = opts.dataset as string | undefined;
      if (!datasetId) {
        const { report } = await resolveTarget(client, opts);
        if (!report) fail("Need -d <datasetId> or -w/-r to resolve a dataset.");
        datasetId = report!.datasetId;
      }
      const result = await client.executeDax(datasetId!, query);
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
  .description("Dataset schema: tables, columns, measures, relationships (DAX INFO.VIEW.*)")
  .option("-w, --workspace <idOrName>")
  .option("-r, --report <idOrName>")
  .option("-d, --dataset <id>", "dataset id (overrides --report)")
  .option("--include-hidden", "include hidden tables/columns/measures")
  .option("--json", "output JSON")
  .action(async (opts) => {
    try {
      const client = new PowerBiClient();
      let datasetId = opts.dataset as string | undefined;
      if (!datasetId) {
        const { report } = await resolveTarget(client, opts);
        if (!report) fail("Need -d <datasetId> or -w/-r to resolve a dataset.");
        datasetId = report!.datasetId;
      }
      const model = await getModelInfo(client, datasetId!, { includeHidden: opts.includeHidden === true });
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

program.parseAsync(process.argv).catch(fail);
