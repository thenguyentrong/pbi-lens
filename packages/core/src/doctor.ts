import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { launchBrowser } from "./capture";
import { startEmbedHost } from "./embedHost";
import { currentAccount, getAccessToken, loadConfig } from "./auth";

export interface DoctorCheck {
  name: string;
  status: "ok" | "warn" | "fail";
  detail: string;
}

/**
 * Preflight checks for a new machine: everything except the actual Power BI
 * embed can be verified offline (no sign-in needed). The render check loads
 * the real embed page in the real headless browser and screenshots it, so a
 * pass means the only untested link left is the Power BI service itself.
 */
export async function runDoctor(): Promise<DoctorCheck[]> {
  const checks: DoctorCheck[] = [];

  const major = parseInt(process.versions.node.split(".")[0], 10);
  checks.push({
    name: "node",
    status: major >= 18 ? "ok" : "fail",
    detail: `v${process.versions.node}${major >= 18 ? "" : " — need >= 18"}`,
  });

  const mediaDir = path.join(__dirname, "..", "media");
  const missing = ["embed.html", "powerbi.min.js"].filter((f) => !fs.existsSync(path.join(mediaDir, f)));
  checks.push({
    name: "embed assets",
    status: missing.length === 0 ? "ok" : "fail",
    detail: missing.length === 0 ? mediaDir : `missing ${missing.join(", ")} in ${mediaDir} — run npm run build`,
  });

  let browserCheck: { browser: Awaited<ReturnType<typeof launchBrowser>>["browser"]; channel: string } | undefined;
  try {
    browserCheck = await launchBrowser();
    checks.push({ name: "browser", status: "ok", detail: browserCheck.channel });
  } catch (e) {
    checks.push({ name: "browser", status: "fail", detail: (e as Error).message.split("\n")[0] });
  }

  if (missing.length === 0 && browserCheck) {
    const shotPath = path.join(os.tmpdir(), `pbi-lens-doctor-${process.pid}.png`);
    const host = await startEmbedHost();
    try {
      const page = await browserCheck.browser.newPage({ viewport: { width: 800, height: 450 } });
      await page.goto(host.url);
      // __pbi only exists once powerbi.min.js has loaded and the control
      // surface IIFE ran — this is the same precondition every capture has.
      await page.waitForFunction(() => (window as any).__pbi !== undefined, undefined, { timeout: 10_000 });
      await page.screenshot({ path: shotPath });
      const size = fs.statSync(shotPath).size;
      checks.push({
        name: "capture pipeline",
        status: size > 0 ? "ok" : "fail",
        detail: `embed host + headless render + PNG write (${size} bytes)`,
      });
    } catch (e) {
      checks.push({ name: "capture pipeline", status: "fail", detail: (e as Error).message.split("\n")[0] });
    } finally {
      fs.rmSync(shotPath, { force: true });
      await host.close().catch(() => {});
    }
  } else {
    checks.push({ name: "capture pipeline", status: "fail", detail: "skipped — fix browser/assets first" });
  }
  await browserCheck?.browser.close().catch(() => {});

  const account = await currentAccount();
  if (!account) {
    checks.push({ name: "sign-in", status: "warn", detail: "not signed in — run pbi-lens login" });
  } else {
    try {
      await getAccessToken();
      checks.push({ name: "sign-in", status: "ok", detail: `${account.username} (token refresh ok)` });
    } catch (e) {
      checks.push({
        name: "sign-in",
        status: "warn",
        detail: `${account.username} — cached session unusable, run pbi-lens login (${(e as Error).message.split("\n")[0]})`,
      });
    }
  }

  const cfg = loadConfig();
  checks.push({
    name: "defaults",
    status: "ok",
    detail:
      cfg.defaultWorkspace || cfg.defaultReport
        ? `workspace=${cfg.defaultWorkspace ?? "-"} report=${cfg.defaultReport ?? "-"}`
        : "none saved (optional) — pbi-lens use -w <ws> -r <report>",
  });

  return checks;
}
