import * as fs from "fs";
import * as path from "path";
import { DefinitionPart } from "./rest";

/**
 * Helpers for rendering a LOCAL report (.pbip / .Report folder) without
 * publishing to a production workspace: package its PBIR definition, rebind it
 * to an already-published semantic model (which carries the data), and let the
 * caller create a throwaway report from it via the Fabric API.
 */

/** Files/dirs that are user-local or Git-system noise — never sent to Fabric. */
const SKIP_DIR_NAMES = new Set([".pbi"]);

/**
 * Resolve the `*.Report` folder from a path that may be:
 *  - a `.pbip` file (reads its report artifact),
 *  - the `*.Report` folder itself,
 *  - a project folder that contains exactly one `*.Report`.
 */
export function resolveReportFolder(localPath: string): string {
  const p = path.resolve(localPath);
  if (!fs.existsSync(p)) throw new Error(`Path not found: ${localPath}`);
  const stat = fs.statSync(p);

  if (stat.isFile()) {
    if (!p.toLowerCase().endsWith(".pbip")) {
      throw new Error(`Expected a .pbip file or a *.Report folder, got: ${localPath}`);
    }
    const proj = JSON.parse(fs.readFileSync(p, "utf8")) as {
      artifacts?: { report?: { path?: string } }[];
    };
    const rel = proj.artifacts?.find((a) => a.report?.path)?.report?.path;
    if (!rel) throw new Error(`No report artifact found in ${localPath}`);
    return path.resolve(path.dirname(p), rel);
  }

  // A folder: either the .Report itself, or a project dir containing one.
  if (fs.existsSync(path.join(p, "definition.pbir"))) return p;
  const child = fs
    .readdirSync(p)
    .find((d) => d.toLowerCase().endsWith(".report") && fs.statSync(path.join(p, d)).isDirectory());
  if (child) return path.join(p, child);

  throw new Error(`Could not locate a .Report folder from "${localPath}". Pass the .pbip file or the *.Report folder.`);
}

/**
 * Walk a `*.Report` folder into Fabric definition parts, replacing
 * `definition.pbir` with a `byConnection` reference to `semanticModelId` so the
 * deployed thin report binds to an existing dataset (REST deployment requires
 * byConnection). Alternative `*.pbir` files (e.g. definition-liveConnect.pbir)
 * are dropped to avoid ambiguity.
 */
export function buildReportDefinitionParts(reportFolder: string, semanticModelId: string): DefinitionPart[] {
  const parts: DefinitionPart[] = [];

  const walk = (dir: string): void => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (entry.isDirectory()) {
        if (SKIP_DIR_NAMES.has(entry.name)) continue;
        walk(path.join(dir, entry.name));
        continue;
      }
      const abs = path.join(dir, entry.name);
      const rel = path.relative(reportFolder, abs).split(path.sep).join("/");
      if (rel.toLowerCase().endsWith(".pbir")) continue; // definition.pbir is synthesized below
      parts.push({ path: rel, payload: fs.readFileSync(abs).toString("base64"), payloadType: "InlineBase64" });
    }
  };
  walk(reportFolder);

  const pbir = {
    $schema:
      "https://developer.microsoft.com/json-schemas/fabric/item/report/definitionProperties/2.0.0/schema.json",
    version: "4.0",
    datasetReference: { byConnection: { connectionString: `semanticmodelid=${semanticModelId}` } },
  };
  parts.push({
    path: "definition.pbir",
    payload: Buffer.from(JSON.stringify(pbir, null, 2)).toString("base64"),
    payloadType: "InlineBase64",
  });

  return parts;
}
