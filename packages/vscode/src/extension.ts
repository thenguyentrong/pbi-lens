import * as vscode from "vscode";
import * as fs from "fs";
import * as path from "path";
import {
  login,
  loginInteractive,
  getAccessToken,
  currentAccount,
  PowerBiClient,
  CaptureSession,
  Report,
  ReportPage,
  Workspace,
  tuneNetworkForVpn,
} from "@pbi-lens/core";

// Bundled .vsix layout ships media/ next to dist/; in the dev monorepo the
// assets live in @pbi-lens/core.
const coreMediaDir = (() => {
  const bundled = path.join(__dirname, "..", "media");
  if (fs.existsSync(path.join(bundled, "embed.html"))) return bundled;
  return path.join(path.dirname(require.resolve("@pbi-lens/core/package.json")), "media");
})();

let panel: vscode.WebviewPanel | undefined;
let currentReport: Report | undefined;
let currentWorkspace: Workspace | undefined;
let reportsProvider: PbiReportsProvider | undefined;

export function activate(context: vscode.ExtensionContext) {
  tuneNetworkForVpn();
  reportsProvider = new PbiReportsProvider();
  context.subscriptions.push(
    vscode.commands.registerCommand("pbiLens.login", cmdLogin),
    vscode.commands.registerCommand("pbiLens.openReport", () => cmdOpenReport(context)),
    vscode.commands.registerCommand("pbiLens.capturePage", cmdCapturePage),
    vscode.window.registerTreeDataProvider("pbiLensReports", reportsProvider),
    vscode.commands.registerCommand("pbiLens.refreshReports", () => reportsProvider?.refresh()),
    vscode.commands.registerCommand("pbiLens.openFromTree", (node: TreeNode) => {
      if (node.kind === "report") return openTarget(context, node.workspace, node.report);
      if (node.kind === "page") return openTarget(context, node.workspace, node.report, node.page.name);
    })
  );
}

// ---------- "Power BI Reports" tree in the Explorer sidebar ----------

type TreeNode =
  | { kind: "workspace"; workspace: Workspace }
  | { kind: "report"; workspace: Workspace; report: Report }
  | { kind: "page"; workspace: Workspace; report: Report; page: ReportPage };

class PbiReportsProvider implements vscode.TreeDataProvider<TreeNode> {
  private emitter = new vscode.EventEmitter<void>();
  readonly onDidChangeTreeData = this.emitter.event;

  refresh(): void {
    this.emitter.fire();
  }

  getTreeItem(node: TreeNode): vscode.TreeItem {
    if (node.kind === "workspace") {
      const item = new vscode.TreeItem(node.workspace.name, vscode.TreeItemCollapsibleState.Collapsed);
      item.iconPath = new vscode.ThemeIcon(node.workspace.id === "my" ? "home" : "folder");
      item.contextValue = "workspace";
      return item;
    }
    if (node.kind === "report") {
      const item = new vscode.TreeItem(node.report.name, vscode.TreeItemCollapsibleState.Collapsed);
      item.iconPath = new vscode.ThemeIcon("graph");
      item.contextValue = "report";
      item.tooltip = "Click to open the live report";
      item.command = { command: "pbiLens.openFromTree", title: "Open report", arguments: [node] };
      return item;
    }
    const item = new vscode.TreeItem(node.page.displayName, vscode.TreeItemCollapsibleState.None);
    item.iconPath = new vscode.ThemeIcon("preview");
    item.contextValue = "page";
    item.tooltip = "Click to open the report on this page";
    item.command = { command: "pbiLens.openFromTree", title: "Open page", arguments: [node] };
    return item;
  }

  async getChildren(node?: TreeNode): Promise<TreeNode[]> {
    try {
      // Empty root → the viewsWelcome "Sign in" content is shown instead.
      if (!(await currentAccount())) return [];
      const client = new PowerBiClient();
      if (!node) {
        const workspaces = await client.listWorkspaces();
        return workspaces.map((workspace) => ({ kind: "workspace" as const, workspace }));
      }
      if (node.kind === "workspace") {
        const reports = await client.listReports(node.workspace.id);
        return reports.map((report) => ({ kind: "report" as const, workspace: node.workspace, report }));
      }
      if (node.kind === "report") {
        const pages = await client.listPages(node.workspace.id, node.report.id);
        return pages.map((page) => ({ kind: "page" as const, workspace: node.workspace, report: node.report, page }));
      }
      return [];
    } catch (e) {
      vscode.window.showErrorMessage(`PBI Lens: ${(e as Error).message}`);
      return [];
    }
  }
}

async function cmdLogin(): Promise<void> {
  try {
    let account;
    try {
      // Browser auth-code flow first — device code is blocked by Conditional
      // Access on some tenants (e.g. dreso.com).
      account = await loginInteractive({});
    } catch {
      account = await login({}, async (message) => {
        // The message contains the verification URL and the code to enter.
        const code = /code\s+([A-Z0-9-]+)/i.exec(message)?.[1];
        if (code) await vscode.env.clipboard.writeText(code);
        const pick = await vscode.window.showInformationMessage(
          `${message}${code ? " (code copied to clipboard)" : ""}`,
          "Open browser"
        );
        if (pick === "Open browser") {
          vscode.env.openExternal(vscode.Uri.parse("https://microsoft.com/devicelogin"));
        }
      });
    }
    vscode.window.showInformationMessage(`PBI Lens: signed in as ${account.username}`);
    reportsProvider?.refresh();
  } catch (e) {
    vscode.window.showErrorMessage(`PBI Lens login failed: ${(e as Error).message}`);
  }
}

async function ensureSignedIn(): Promise<boolean> {
  if (await currentAccount()) return true;
  const pick = await vscode.window.showWarningMessage("PBI Lens: not signed in.", "Sign in");
  if (pick === "Sign in") await cmdLogin();
  return !!(await currentAccount());
}

async function pickTarget(): Promise<{ workspace: Workspace; report: Report } | undefined> {
  const client = new PowerBiClient();
  const cfg = vscode.workspace.getConfiguration("pbiLens");
  let workspace: Workspace | undefined;
  let report: Report | undefined;

  const wsDefault = cfg.get<string>("workspace");
  if (wsDefault) {
    workspace = await client.resolveWorkspace(wsDefault);
  } else {
    const all = await client.listWorkspaces();
    const pick = await vscode.window.showQuickPick(
      all.map((w) => ({ label: w.name, ws: w })),
      { placeHolder: "Power BI workspace" }
    );
    workspace = pick?.ws;
  }
  if (!workspace) return undefined;

  const repDefault = cfg.get<string>("report");
  if (repDefault) {
    report = await client.resolveReport(workspace.id, repDefault);
  } else {
    const all = await client.listReports(workspace.id);
    const pick = await vscode.window.showQuickPick(
      all.map((r) => ({ label: r.name, rep: r })),
      { placeHolder: "Report" }
    );
    report = pick?.rep;
  }
  if (!report) return undefined;
  return { workspace, report };
}

async function cmdOpenReport(context: vscode.ExtensionContext): Promise<void> {
  if (!(await ensureSignedIn())) return;
  const target = await pickTarget();
  if (!target) return;
  await openTarget(context, target.workspace, target.report);
}

async function openTarget(
  context: vscode.ExtensionContext,
  workspace: Workspace,
  report: Report,
  pageName?: string
): Promise<void> {
  try {
    if (!(await ensureSignedIn())) return;

    // Same report already open + page click → just switch the page (instant).
    if (panel && currentReport?.id === report.id && pageName) {
      panel.reveal();
      panel.webview.postMessage({ type: "setPage", pageName });
      return;
    }

    const target = { workspace, report };
    currentWorkspace = target.workspace;
    currentReport = target.report;

    if (!panel) {
      panel = vscode.window.createWebviewPanel(
        "pbiLens",
        `Power BI: ${target.report.name}`,
        vscode.ViewColumn.One,
        {
          enableScripts: true,
          retainContextWhenHidden: true,
          localResourceRoots: [vscode.Uri.file(coreMediaDir)],
        }
      );
      panel.onDidDispose(() => (panel = undefined), null, context.subscriptions);
    } else {
      panel.title = `Power BI: ${target.report.name}`;
      panel.reveal();
    }

    const webview = panel.webview;
    const pbiJsUri = webview.asWebviewUri(vscode.Uri.file(path.join(coreMediaDir, "powerbi.min.js")));
    const authoringJsUri = webview.asWebviewUri(
      vscode.Uri.file(path.join(coreMediaDir, "powerbi-report-authoring.min.js"))
    );
    let html = fs.readFileSync(path.join(coreMediaDir, "embed.html"), "utf8");
    html = html
      .replace('<script src="./powerbi.min.js"></script>', `<script src="${pbiJsUri}"></script>`)
      .replace(
        '<script src="./powerbi-report-authoring.min.js"></script>',
        `<script src="${authoringJsUri}"></script>`
      )
      .replace(
        "<head>",
        `<head>\n  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src ${webview.cspSource} 'unsafe-inline'; style-src 'unsafe-inline'; frame-src https://app.powerbi.com https://*.powerbi.com; connect-src https://*.powerbi.com https://*.analysis.windows.net;">`
      );
    webview.html = html;

    const token = await getAccessToken();
    const sendInit = () =>
      webview.postMessage({
        type: "init",
        config: {
          accessToken: token,
          embedUrl: currentReport!.embedUrl,
          reportId: currentReport!.id,
          showNav: true,
          pageName,
        },
      });

    const sub = webview.onDidReceiveMessage((msg) => {
      if (msg.type === "ready") sendInit();
    });
    context.subscriptions.push(sub);
    // The webview may already be past its "ready" handshake (panel reuse).
    sendInit();

    // AAD tokens expire after ~1h; refresh the embed before that.
    const refresh = setInterval(async () => {
      if (!panel) return clearInterval(refresh);
      try {
        const fresh = await getAccessToken();
        webview.postMessage({ type: "setAccessToken", accessToken: fresh });
      } catch {
        /* user signed out; next open will prompt */
      }
    }, 50 * 60 * 1000);
    context.subscriptions.push({ dispose: () => clearInterval(refresh) });
  } catch (e) {
    vscode.window.showErrorMessage(`PBI Lens: ${(e as Error).message}`);
  }
}

async function cmdCapturePage(): Promise<void> {
  try {
    if (!(await ensureSignedIn())) return;
    if (!currentReport) {
      const target = await pickTarget();
      if (!target) return;
      currentWorkspace = target.workspace;
      currentReport = target.report;
    }
    const client = new PowerBiClient();
    const pages = await client.listPages(currentWorkspace!.id, currentReport!.id);
    const pagePick = await vscode.window.showQuickPick(
      pages.map((p) => ({ label: p.displayName, page: p })),
      { placeHolder: "Page to capture" }
    );
    if (!pagePick) return;

    const folder = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? require("os").tmpdir();
    const outDir = path.join(folder, ".pbi-shots");
    fs.mkdirSync(outDir, { recursive: true });
    const outPath = path.join(outDir, `${pagePick.page.displayName.replace(/[^\w-]+/g, "_")}.png`);

    await vscode.window.withProgress(
      { location: vscode.ProgressLocation.Notification, title: "PBI Lens: capturing page..." },
      async () => {
        const token = await getAccessToken();
        const session = await CaptureSession.open({
          accessToken: token,
          embedUrl: currentReport!.embedUrl,
          reportId: currentReport!.id,
        });
        try {
          await session.setPage(pagePick.page.name);
          await session.screenshot(outPath);
        } finally {
          await session.close();
        }
      }
    );
    vscode.window.showInformationMessage(`Captured ${outPath} — point Claude at it with: Read ${outPath}`);
  } catch (e) {
    vscode.window.showErrorMessage(`PBI Lens capture failed: ${(e as Error).message}`);
  }
}

export function deactivate() {}
