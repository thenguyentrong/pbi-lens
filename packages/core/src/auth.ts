import {
  PublicClientApplication,
  Configuration,
  AccountInfo,
  ICachePlugin,
  TokenCacheContext,
} from "@azure/msal-node";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

// Azure CLI's well-known public client id — pre-consented for the Power BI
// service in most Entra tenants, so no app registration is needed. Override
// via config/env if your tenant blocks it.
export const DEFAULT_CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46";
export const POWERBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default";

export interface AuthConfig {
  clientId?: string;
  tenantId?: string;
}

export interface StoredConfig extends AuthConfig {
  defaultWorkspace?: string;
  defaultReport?: string;
}

const CONFIG_DIR = path.join(os.homedir(), ".pbi-lens");
const CONFIG_FILE = path.join(CONFIG_DIR, "config.json");
const CACHE_FILE = path.join(CONFIG_DIR, "token-cache.json");

export function loadConfig(): StoredConfig {
  try {
    return JSON.parse(fs.readFileSync(CONFIG_FILE, "utf8"));
  } catch {
    return {};
  }
}

export function saveConfig(cfg: StoredConfig): void {
  fs.mkdirSync(CONFIG_DIR, { recursive: true });
  fs.writeFileSync(CONFIG_FILE, JSON.stringify(cfg, null, 2));
}

// Plain-file token cache. The refresh token grants Power BI access for the
// signed-in user, so the file is created user-readable only (0600 on POSIX;
// on Windows it inherits the user-profile ACL).
const cachePlugin: ICachePlugin = {
  async beforeCacheAccess(ctx: TokenCacheContext) {
    try {
      ctx.tokenCache.deserialize(fs.readFileSync(CACHE_FILE, "utf8"));
    } catch {
      /* first run */
    }
  },
  async afterCacheAccess(ctx: TokenCacheContext) {
    if (ctx.cacheHasChanged) {
      fs.mkdirSync(CONFIG_DIR, { recursive: true });
      fs.writeFileSync(CACHE_FILE, ctx.tokenCache.serialize(), { mode: 0o600 });
    }
  },
};

function buildPca(auth: AuthConfig): PublicClientApplication {
  const cfg = loadConfig();
  const clientId = auth.clientId ?? cfg.clientId ?? process.env.PBI_LENS_CLIENT_ID ?? DEFAULT_CLIENT_ID;
  const tenantId = auth.tenantId ?? cfg.tenantId ?? process.env.PBI_LENS_TENANT_ID ?? "organizations";
  const msalConfig: Configuration = {
    auth: {
      clientId,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: { cachePlugin },
  };
  return new PublicClientApplication(msalConfig);
}

async function firstAccount(pca: PublicClientApplication): Promise<AccountInfo | null> {
  const accounts = await pca.getTokenCache().getAllAccounts();
  return accounts[0] ?? null;
}

/**
 * Interactive sign-in via device code. `onCode` receives the message that
 * tells the user which URL to open and which code to enter.
 */
export async function login(
  auth: AuthConfig = {},
  onCode: (message: string) => void = (m) => console.log(m)
): Promise<AccountInfo> {
  const pca = buildPca(auth);
  const result = await pca.acquireTokenByDeviceCode({
    scopes: [POWERBI_SCOPE],
    deviceCodeCallback: (info) => onCode(info.message),
  });
  if (!result?.account) throw new Error("Login did not return an account.");
  return result.account;
}

/**
 * Interactive sign-in via the system browser (authorization-code flow with a
 * loopback redirect). Use this when the tenant's Conditional Access policy
 * blocks the device-code flow — a real browser session satisfies MFA /
 * compliant-device / location policies that device code cannot.
 */
export async function loginInteractive(auth: AuthConfig = {}): Promise<AccountInfo> {
  const pca = buildPca(auth);
  const result = await pca.acquireTokenInteractive({
    scopes: [POWERBI_SCOPE],
    openBrowser: async (url: string) => {
      const { spawn } = await import("child_process");
      // NB: do NOT route the URL through `cmd /c start` on Windows — cmd treats
      // the `&` separating query params as a command separator and truncates the
      // URL (dropping scope/redirect → AADSTS900144). rundll32's FileProtocolHandler
      // hands the full URL to the default browser without shell parsing.
      const cmd =
        process.platform === "win32"
          ? { file: "rundll32", args: ["url.dll,FileProtocolHandler", url] }
          : process.platform === "darwin"
            ? { file: "open", args: [url] }
            : { file: "xdg-open", args: [url] };
      spawn(cmd.file, cmd.args, { detached: true, stdio: "ignore" }).unref();
    },
    successTemplate: "Signed in to Power BI. You can close this tab and return to the terminal.",
  });
  if (!result?.account) throw new Error("Login did not return an account.");
  return result.account;
}

/**
 * Returns a valid Power BI access token, refreshing silently from the cache.
 * Throws with a helpful message when no cached session exists.
 */
export async function getAccessToken(auth: AuthConfig = {}): Promise<string> {
  const pca = buildPca(auth);
  const account = await firstAccount(pca);
  if (!account) {
    throw new Error("Not signed in. Run `pbi-lens login` first.");
  }
  const result = await pca.acquireTokenSilent({
    scopes: [POWERBI_SCOPE],
    account,
  });
  return result.accessToken;
}

export async function currentAccount(auth: AuthConfig = {}): Promise<AccountInfo | null> {
  return firstAccount(buildPca(auth));
}

export async function logout(auth: AuthConfig = {}): Promise<void> {
  const pca = buildPca(auth);
  const cache = pca.getTokenCache();
  for (const account of await cache.getAllAccounts()) {
    await cache.removeAccount(account);
  }
}
