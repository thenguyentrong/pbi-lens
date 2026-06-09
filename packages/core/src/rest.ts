import { getAccessToken, AuthConfig } from "./auth";
import * as fs from "fs";
import * as path from "path";

const API = "https://api.powerbi.com/v1.0/myorg";

export interface Workspace {
  id: string;
  name: string;
}

export interface Report {
  id: string;
  name: string;
  embedUrl: string;
  datasetId: string;
  webUrl?: string;
}

export interface ReportPage {
  name: string; // internal name, e.g. ReportSection1
  displayName: string;
  order?: number;
}

export interface DaxResult {
  rows: Record<string, unknown>[];
}

async function api<T>(token: string, pathName: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`${API}${pathName}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init?.headers ?? {}),
    },
  });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(`Power BI API ${init?.method ?? "GET"} ${pathName} failed: ${res.status} ${res.statusText}\n${body}`);
  }
  return (await res.json()) as T;
}

export class PowerBiClient {
  constructor(private auth: AuthConfig = {}) {}

  private token(): Promise<string> {
    return getAccessToken(this.auth);
  }

  async listWorkspaces(): Promise<Workspace[]> {
    const token = await this.token();
    const data = await api<{ value: Workspace[] }>(token, "/groups");
    return data.value;
  }

  async listReports(workspaceId: string): Promise<Report[]> {
    const token = await this.token();
    const data = await api<{ value: Report[] }>(token, `/groups/${workspaceId}/reports`);
    return data.value;
  }

  async getReport(workspaceId: string, reportId: string): Promise<Report> {
    const token = await this.token();
    return api<Report>(token, `/groups/${workspaceId}/reports/${reportId}`);
  }

  async listPages(workspaceId: string, reportId: string): Promise<ReportPage[]> {
    const token = await this.token();
    const data = await api<{ value: ReportPage[] }>(token, `/groups/${workspaceId}/reports/${reportId}/pages`);
    return data.value;
  }

  /** Resolve a workspace by id or (case-insensitive) name. */
  async resolveWorkspace(idOrName: string): Promise<Workspace> {
    const all = await this.listWorkspaces();
    const found =
      all.find((w) => w.id === idOrName) ??
      all.find((w) => w.name.toLowerCase() === idOrName.toLowerCase());
    if (!found) {
      throw new Error(
        `Workspace "${idOrName}" not found. Available: ${all.map((w) => w.name).join(", ") || "(none — is the account licensed for Power BI?)"}`
      );
    }
    return found;
  }

  /** Resolve a report by id or (case-insensitive) name within a workspace. */
  async resolveReport(workspaceId: string, idOrName: string): Promise<Report> {
    const all = await this.listReports(workspaceId);
    const found =
      all.find((r) => r.id === idOrName) ??
      all.find((r) => r.name.toLowerCase() === idOrName.toLowerCase());
    if (!found) {
      throw new Error(
        `Report "${idOrName}" not found in workspace. Available: ${all.map((r) => r.name).join(", ") || "(none)"}`
      );
    }
    return found;
  }

  /**
   * Run a DAX query against a published dataset (executeQueries endpoint).
   * Pro-license compatible. Limits: 120 queries/min, 100k rows, 15 MB.
   */
  async executeDax(datasetId: string, query: string): Promise<DaxResult> {
    const token = await this.token();
    const data = await api<{
      results: { tables: { rows: Record<string, unknown>[] }[] }[];
    }>(token, `/datasets/${datasetId}/executeQueries`, {
      method: "POST",
      body: JSON.stringify({
        queries: [{ query }],
        serializerSettings: { includeNulls: true },
      }),
    });
    return { rows: data.results[0]?.tables[0]?.rows ?? [] };
  }

  /**
   * Publish a .pbix to a workspace via the Imports API (Pro compatible).
   * nameConflict=CreateOrOverwrite replaces an existing report of the same name.
   */
  async importPbix(workspaceId: string, pbixPath: string, datasetDisplayName?: string): Promise<{ id: string }> {
    const token = await this.token();
    const name = datasetDisplayName ?? path.basename(pbixPath, ".pbix");
    const buffer = fs.readFileSync(pbixPath);
    const form = new FormData();
    form.append("file", new Blob([buffer]), path.basename(pbixPath));
    const url = `${API}/groups/${workspaceId}/imports?datasetDisplayName=${encodeURIComponent(name)}&nameConflict=CreateOrOverwrite`;
    const res = await fetch(url, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}` },
      body: form,
    });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      throw new Error(`Import failed: ${res.status} ${res.statusText}\n${body}`);
    }
    return (await res.json()) as { id: string };
  }

  /** Poll an import until it succeeds or fails. */
  async waitForImport(workspaceId: string, importId: string, timeoutMs = 120_000): Promise<void> {
    const start = Date.now();
    for (;;) {
      const token = await this.token();
      const imp = await api<{ importState: string; reports?: { name: string }[] }>(
        token,
        `/groups/${workspaceId}/imports/${importId}`
      );
      if (imp.importState === "Succeeded") return;
      if (imp.importState === "Failed") throw new Error("Import failed (importState=Failed).");
      if (Date.now() - start > timeoutMs) throw new Error("Import timed out.");
      await new Promise((r) => setTimeout(r, 2000));
    }
  }
}
