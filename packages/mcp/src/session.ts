import {
  PowerBiClient,
  Workspace,
  Report,
  CaptureSession,
  getAccessToken,
  loadConfig,
} from "@pbi-lens/core";

export const client = new PowerBiClient();

export async function resolveWorkspace(workspace?: string): Promise<Workspace> {
  const wsArg = workspace ?? loadConfig().defaultWorkspace;
  if (!wsArg) {
    throw new Error(
      'No workspace given and no default saved. Pass the `workspace` parameter (use "my" for My Workspace), ' +
        "or save a default once with `pbi-lens use -w <idOrName>`."
    );
  }
  return client.resolveWorkspace(wsArg);
}

export async function resolveTarget(opts: {
  workspace?: string;
  report?: string;
}): Promise<{ workspace: Workspace; report: Report }> {
  const workspace = await resolveWorkspace(opts.workspace);
  const repArg = opts.report ?? loadConfig().defaultReport;
  if (!repArg) {
    throw new Error(
      "No report given and no default saved. Pass the `report` parameter, " +
        "or save a default once with `pbi-lens use -r <idOrName>`."
    );
  }
  const report = await client.resolveReport(workspace.id, repArg);
  return { workspace, report };
}

const IDLE_MS = 10 * 60_000;
const RETRYABLE =
  /Target closed|browser has been closed|Protocol error|Execution context was destroyed|Target page, context or browser has been closed/i;

/**
 * One warm CaptureSession per (workspace, report): embeds cost 4-8 s, so the
 * browser stays open between tool calls. Token is refreshed on every reuse;
 * a dead browser is reopened and the call retried once; idle sessions are
 * reaped after 10 minutes.
 */
export class SessionManager {
  private session?: CaptureSession;
  private key?: string;
  private lastUsed = 0;
  private reaper: NodeJS.Timeout;

  constructor() {
    this.reaper = setInterval(() => {
      if (this.session && Date.now() - this.lastUsed > IDLE_MS) {
        void this.closeSession();
      }
    }, 60_000);
    this.reaper.unref();
  }

  private async acquire(workspace: Workspace, report: Report): Promise<CaptureSession> {
    const key = `${workspace.id}/${report.id}`;
    const token = await getAccessToken();
    if (this.session && this.key === key) {
      await this.session.refreshAccessToken(token);
    } else {
      await this.closeSession();
      this.session = await CaptureSession.open({
        accessToken: token,
        embedUrl: report.embedUrl,
        reportId: report.id,
      });
      this.key = key;
    }
    this.lastUsed = Date.now();
    return this.session;
  }

  async withSession<T>(
    workspace: Workspace,
    report: Report,
    fn: (s: CaptureSession) => Promise<T>
  ): Promise<T> {
    try {
      const session = await this.acquire(workspace, report);
      const result = await fn(session);
      this.lastUsed = Date.now();
      return result;
    } catch (e) {
      if (!RETRYABLE.test((e as Error)?.message ?? "")) throw e;
      await this.closeSession();
      const session = await this.acquire(workspace, report);
      const result = await fn(session);
      this.lastUsed = Date.now();
      return result;
    }
  }

  private async closeSession(): Promise<void> {
    const session = this.session;
    this.session = undefined;
    this.key = undefined;
    if (session) await session.close().catch(() => {});
  }

  async dispose(): Promise<void> {
    clearInterval(this.reaper);
    await this.closeSession();
  }
}
