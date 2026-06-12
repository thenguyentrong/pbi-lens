import express from "express";
import * as http from "http";
import * as path from "path";

export interface EmbedHost {
  url: string;
  close(): Promise<void>;
}

/**
 * Serves the embed page + powerbi-client bundle on an ephemeral localhost
 * port. The access token is never part of the served page; the Playwright
 * driver injects it at runtime via page.evaluate.
 */
export async function startEmbedHost(): Promise<EmbedHost> {
  const app = express();
  const mediaDir = path.join(__dirname, "..", "media");
  app.get("/", (_req, res) => res.sendFile(path.join(mediaDir, "embed.html")));
  app.get("/powerbi.min.js", (_req, res) => res.sendFile(path.join(mediaDir, "powerbi.min.js")));
  app.get("/powerbi-report-authoring.min.js", (_req, res) =>
    res.sendFile(path.join(mediaDir, "powerbi-report-authoring.min.js"))
  );

  const server: http.Server = await new Promise((resolve) => {
    const s = app.listen(0, "127.0.0.1", () => resolve(s));
  });
  const address = server.address();
  if (address === null || typeof address === "string") {
    throw new Error("Embed host failed to bind a port.");
  }
  return {
    url: `http://127.0.0.1:${address.port}/`,
    close: () => new Promise((resolve) => server.close(() => resolve())),
  };
}
