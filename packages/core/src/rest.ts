import { getAccessToken, AuthConfig } from "./auth";
import * as fs from "fs";
import * as path from "path";

const API = "https://api.powerbi.com/v1.0/myorg";
// The Fabric REST API. The Power BI delegated token (analysis.windows.net
// audience) is accepted here too, so no extra scope/sign-in is needed.
const FABRIC_API = "https://api.fabric.microsoft.com/v1";

export interface Workspace {
  id: string;
  name: string;
}

/**
 * The personal "My Workspace" — works with a FREE Power BI license (no Pro).
 * Its REST routes have no /groups/{id} prefix, so it gets a sentinel id.
 */
export const MY_WORKSPACE: Workspace = { id: "my", name: "My Workspace" };

const MY_ALIASES = new Set(["my", "me", "my workspace", "myworkspace"]);

/** Route prefix for a workspace: "" for My Workspace, /groups/{id} otherwise. */
function groupPath(workspaceId: string): string {
  return workspaceId === MY_WORKSPACE.id ? "" : `/groups/${workspaceId}`;
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

/** One file of a Fabric item definition, base64-encoded. */
export interface DefinitionPart {
  path: string;
  payload: string;
  payloadType: "InlineBase64";
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
    // Page through /groups: the API caps each page (default 100), so a tenant
    // with many workspaces would otherwise be silently truncated.
    const groups: Workspace[] = [];
    const pageSize = 5000;
    for (let skip = 0; ; skip += pageSize) {
      const data = await api<{ value: Workspace[] }>(token, `/groups?$top=${pageSize}&$skip=${skip}`);
      groups.push(...data.value);
      if (data.value.length < pageSize) break;
    }
    // My Workspace is not a group; surface it explicitly so free-license
    // accounts (which have no groups at all) see their one usable target.
    return [MY_WORKSPACE, ...groups];
  }

  /**
   * Find reports whose name contains `query` (case-insensitive) across every
   * workspace the account can see. Workspaces that error (e.g. no access to
   * their reports) are skipped rather than failing the whole search.
   */
  async findReports(query: string): Promise<{ workspace: Workspace; report: Report }[]> {
    const workspaces = await this.listWorkspaces();
    const needle = query.toLowerCase();
    const perWorkspace = await Promise.all(
      workspaces.map(async (workspace) => {
        try {
          const reports = await this.listReports(workspace.id);
          return reports
            .filter((r) => r.name.toLowerCase().includes(needle))
            .map((report) => ({ workspace, report }));
        } catch {
          return [];
        }
      })
    );
    return perWorkspace.flat();
  }

  async listReports(workspaceId: string): Promise<Report[]> {
    const token = await this.token();
    const data = await api<{ value: Report[] }>(token, `${groupPath(workspaceId)}/reports`);
    return data.value;
  }

  async getReport(workspaceId: string, reportId: string): Promise<Report> {
    const token = await this.token();
    return api<Report>(token, `${groupPath(workspaceId)}/reports/${reportId}`);
  }

  async listPages(workspaceId: string, reportId: string): Promise<ReportPage[]> {
    const token = await this.token();
    const data = await api<{ value: ReportPage[] }>(token, `${groupPath(workspaceId)}/reports/${reportId}/pages`);
    return data.value;
  }

  /** Resolve a workspace by id or (case-insensitive) name. "my" → My Workspace. */
  async resolveWorkspace(idOrName: string): Promise<Workspace> {
    if (MY_ALIASES.has(idOrName.toLowerCase())) return MY_WORKSPACE;
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
    const url = `${API}${groupPath(workspaceId)}/imports?datasetDisplayName=${encodeURIComponent(name)}&nameConflict=CreateOrOverwrite`;
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
        `${groupPath(workspaceId)}/imports/${importId}`
      );
      if (imp.importState === "Succeeded") return;
      if (imp.importState === "Failed") throw new Error("Import failed (importState=Failed).");
      if (Date.now() - start > timeoutMs) throw new Error("Import timed out.");
      await new Promise((r) => setTimeout(r, 2000));
    }
  }

  /**
   * Create a Report item from a PBIR definition via the Fabric REST API.
   * The workspace must be on Fabric/Premium capacity (My Workspace is not
   * supported). Handles both the immediate (201) and long-running (202) paths.
   * Returns the new report's id.
   */
  async createReport(workspaceId: string, displayName: string, parts: DefinitionPart[]): Promise<string> {
    const token = await this.token();
    const res = await fetch(`${FABRIC_API}/workspaces/${workspaceId}/reports`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ displayName, definition: { parts } }),
    });
    if (res.status === 201) {
      const body = (await res.json()) as { id: string };
      return body.id;
    }
    if (res.status === 202) {
      const opUrl = res.headers.get("Location");
      if (!opUrl) throw new Error("Fabric accepted the report (202) but returned no operation Location header.");
      return this.waitForFabricItem(opUrl);
    }
    const body = await res.text().catch(() => "");
    throw new Error(`Fabric create report failed: ${res.status} ${res.statusText}\n${body}`);
  }

  /** Poll a Fabric long-running operation until it finishes, then return the created item id. */
  private async waitForFabricItem(operationUrl: string, timeoutMs = 120_000): Promise<string> {
    const start = Date.now();
    for (;;) {
      const token = await this.token();
      const res = await fetch(operationUrl, { headers: { Authorization: `Bearer ${token}` } });
      if (!res.ok) {
        const body = await res.text().catch(() => "");
        throw new Error(`Fabric operation poll failed: ${res.status} ${res.statusText}\n${body}`);
      }
      const op = (await res.json()) as { status: string; error?: unknown };
      if (op.status === "Succeeded") {
        const result = await fetch(`${operationUrl}/result`, { headers: { Authorization: `Bearer ${token}` } });
        const item = (await result.json()) as { id: string };
        return item.id;
      }
      if (op.status === "Failed") throw new Error(`Fabric create operation failed: ${JSON.stringify(op.error ?? op)}`);
      if (Date.now() - start > timeoutMs) throw new Error("Fabric create operation timed out.");
      await new Promise((r) => setTimeout(r, 2000));
    }
  }

  /** Delete a report (Power BI REST). 404 is treated as already-gone. */
  async deleteReport(workspaceId: string, reportId: string): Promise<void> {
    const token = await this.token();
    const res = await fetch(`${API}${groupPath(workspaceId)}/reports/${reportId}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok && res.status !== 404) {
      const body = await res.text().catch(() => "");
      throw new Error(`Delete report failed: ${res.status} ${res.statusText}\n${body}`);
    }
  }
}
