import { PowerBiClient, Workspace, Report, ReportPage } from "./rest";
import { CaptureSession, PageInfo, VisualInfo, VisualFields, FiltersState } from "./capture";
import { getModelInfo, ModelInfo } from "./model";

export interface PageDetail {
  page: PageInfo;
  visuals: (VisualInfo & { fields?: VisualFields | { error: string } })[];
}

/**
 * Everything an AI agent usually needs to orient on a report, in one call:
 * report identity, pages, model schema, visuals with field bindings, and the
 * active filter state.
 */
export interface ReportContext {
  report: { id: string; name: string; datasetId: string; workspaceId: string; workspaceName: string };
  pages: ReportPage[];
  model: ModelInfo;
  pagesDetail: PageDetail[];
  filters: FiltersState | { error: string };
  warnings: string[];
}

export interface ContextOptions {
  /** Page (name or display name) to detail; default: the active page. */
  page?: string;
  /** Detail every page (~5-8 s render per page). */
  allPages?: boolean;
  /** Include per-visual field bindings (default true). */
  includeFields?: boolean;
}

async function detailPage(
  session: CaptureSession,
  pageInfo: PageInfo,
  includeFields: boolean,
  warnings: string[]
): Promise<PageDetail> {
  const visuals = await session.getVisuals();
  const detailed: PageDetail["visuals"] = [];
  for (const v of visuals) {
    if (!includeFields) {
      detailed.push(v);
      continue;
    }
    try {
      const fields = await session.getVisualFields(v.name);
      detailed.push({ ...v, fields });
    } catch (e) {
      // Cards/slicers/images may refuse field readback — keep going.
      detailed.push({ ...v, fields: { error: (e as Error).message.split("\n")[0] } });
    }
  }
  const failed = detailed.filter((v) => v.fields && "error" in v.fields);
  if (includeFields && detailed.length > 0 && failed.length === detailed.length) {
    warnings.push(`Field readback failed for every visual on page "${pageInfo.displayName}".`);
  }
  return { page: pageInfo, visuals: detailed };
}

export async function buildReportContext(
  client: PowerBiClient,
  workspace: Workspace,
  report: Report,
  session: CaptureSession,
  opts: ContextOptions = {}
): Promise<ReportContext> {
  const warnings: string[] = [];
  const includeFields = opts.includeFields !== false;

  // Model comes over REST while the embed session does browser work.
  const modelPromise = getModelInfo(client, report.datasetId).catch((e) => {
    warnings.push(`Model schema unavailable: ${(e as Error).message.split("\n")[0]}`);
    return { tables: [], columns: [], measures: [], relationships: [], warnings: [] } as ModelInfo;
  });
  const restPagesPromise = client.listPages(workspace.id, report.id).catch((e) => {
    warnings.push(`Page list (REST) unavailable: ${(e as Error).message.split("\n")[0]}`);
    return [] as ReportPage[];
  });

  if (opts.page) await session.setPage(opts.page);

  const pagesDetail: PageDetail[] = [];
  if (opts.allPages) {
    for (const p of await session.getPages()) {
      await session.setPage(p.name);
      pagesDetail.push(await detailPage(session, p, includeFields, warnings));
    }
  } else {
    const pages = await session.getPages();
    const active = pages.find((p) => p.isActive) ?? pages[0];
    if (active) {
      pagesDetail.push(await detailPage(session, active, includeFields, warnings));
    }
  }

  let filters: ReportContext["filters"];
  try {
    filters = await session.getFiltersState();
  } catch (e) {
    filters = { error: (e as Error).message.split("\n")[0] };
  }

  const model = await modelPromise;
  const pages = await restPagesPromise;
  warnings.push(...model.warnings);

  return {
    report: {
      id: report.id,
      name: report.name,
      datasetId: report.datasetId,
      workspaceId: workspace.id,
      workspaceName: workspace.name,
    },
    pages,
    model: { ...model, warnings: [] },
    pagesDetail,
    filters,
    warnings,
  };
}
