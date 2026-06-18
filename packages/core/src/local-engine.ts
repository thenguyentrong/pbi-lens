import * as fs from "fs";
import * as path from "path";
import { spawnSync } from "child_process";
import { DaxResult } from "./rest";

export interface LocalInstance {
  /** TCP port the local Tabular engine listens on. */
  port: number;
  /** AnalysisServicesWorkspaces subfolder name (a GUID-ish id). */
  workspace: string;
  portFile: string;
  /** Port file mtime (newest = most recently opened Desktop window). */
  modifiedMs: number;
}

/** The published ADOMD sidecar exe (framework-dependent; needs the .NET 8 runtime). */
function runnerExe(): string {
  return path.join(__dirname, "..", "native", "adomd-runner", "publish", "adomd-runner.exe");
}

/**
 * Discover running Power BI Desktop local engines by reading the msmdsrv.port.txt
 * files Desktop drops under %LOCALAPPDATA%. Newest first; empty if Desktop isn't
 * running (or off-Windows).
 */
export function findDesktopInstances(): LocalInstance[] {
  const local = process.env.LOCALAPPDATA;
  if (!local) return [];
  const base = path.join(local, "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces");
  if (!fs.existsSync(base)) return [];
  const out: LocalInstance[] = [];
  for (const ws of fs.readdirSync(base)) {
    const portFile = path.join(base, ws, "Data", "msmdsrv.port.txt");
    if (!fs.existsSync(portFile)) continue;
    // Desktop writes the port as UTF-16LE with a BOM — keep digits only.
    const raw = fs.readFileSync(portFile, "utf16le").replace(/[^0-9]/g, "");
    const port = parseInt(raw, 10);
    if (!Number.isFinite(port) || port <= 0) continue;
    out.push({ port, workspace: ws, portFile, modifiedMs: fs.statSync(portFile).mtimeMs });
  }
  return out.sort((a, b) => b.modifiedMs - a.modifiedMs);
}

/** Resolve a local engine port: the explicit one, else the newest running instance. */
export function resolveLocalPort(port?: number): number {
  if (port) return port;
  const instances = findDesktopInstances();
  if (instances.length === 0) {
    throw new Error(
      "No running Power BI Desktop instance found (no msmdsrv.port.txt under %LOCALAPPDATA%). " +
        "Open a .pbix/.pbip in Power BI Desktop first, or pass an explicit port."
    );
  }
  return instances[0].port;
}

/**
 * Run a DAX query against a local Power BI Desktop engine via the ADOMD sidecar.
 * Returns the same shape as PowerBiClient.executeDax, so it can be passed
 * wherever a DaxExecutor is expected (e.g. getModelInfoWith). Requires the
 * .NET 8 runtime and the built sidecar.
 */
export async function executeDaxLocal(port: number, query: string, database?: string): Promise<DaxResult> {
  const exe = runnerExe();
  if (!fs.existsSync(exe)) {
    throw new Error(
      `ADOMD sidecar not built at ${exe}. Build it once with \`npm run build:native -w @pbi-lens/core\` (needs the .NET SDK).`
    );
  }
  const args = database ? [String(port), database] : [String(port)];
  const res = spawnSync(exe, args, { input: query, encoding: "utf8", maxBuffer: 256 * 1024 * 1024 });
  if (res.error) throw res.error;
  if (res.status !== 0) {
    throw new Error(`Local DAX failed: ${(res.stderr || res.stdout || "").trim() || `exit ${res.status}`}`);
  }
  try {
    return JSON.parse(res.stdout) as DaxResult;
  } catch {
    throw new Error(`Local DAX returned non-JSON output: ${res.stdout.slice(0, 200)}`);
  }
}
