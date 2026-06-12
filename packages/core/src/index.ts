import * as net from "net";

/**
 * Node's Happy-Eyeballs default tries each resolved address for only 250 ms
 * before moving on; on corporate VPNs (latency + TLS inspection) every
 * attempt "times out" and fetch/MSAL fail with UND_ERR_CONNECT_TIMEOUT while
 * browsers work fine. Give each attempt a realistic budget. Call once at
 * process start (CLI, MCP server, VS Code extension all do).
 */
export function tuneNetworkForVpn(attemptTimeoutMs = 3000): void {
  try {
    net.setDefaultAutoSelectFamilyAttemptTimeout(attemptTimeoutMs);
  } catch {
    /* older Node without the API — keep defaults */
  }
}

export * from "./auth";
export * from "./rest";
export * from "./capture";
export * from "./embedHost";
export * from "./doctor";
export * from "./model";
export * from "./context";
