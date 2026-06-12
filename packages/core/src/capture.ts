import * as fs from "fs";
import { chromium, Browser, Page } from "playwright-core";
import { startEmbedHost, EmbedHost } from "./embedHost";

/**
 * playwright-core ships no browser binaries. Drive an installed Edge or
 * Chrome (Edge is preinstalled on Windows 10/11), falling back to a
 * playwright-managed Chromium if the user has one.
 */
export async function launchBrowser(): Promise<{ browser: Browser; channel: string }> {
  const channels: (string | undefined)[] = ["msedge", "chrome", undefined];
  let lastError: unknown;
  for (const channel of channels) {
    try {
      const browser = await chromium.launch({ headless: true, channel });
      return { browser, channel: channel ?? "playwright-chromium" };
    } catch (e) {
      lastError = e;
    }
  }
  throw new Error(
    `No Chromium-based browser found. Install Microsoft Edge or Google Chrome.\n${(lastError as Error)?.message ?? ""}`
  );
}

export interface EmbedTarget {
  accessToken: string;
  embedUrl: string;
  reportId: string;
}

export interface PageInfo {
  name: string;
  displayName: string;
  isActive: boolean;
  defaultSize?: { width?: number; height?: number };
}

export interface VisualInfo {
  name: string;
  title?: string;
  type: string;
  layout?: { x: number; y: number; width: number; height: number };
}

export type EmbedMode = "view" | "edit";

/** One column/measure/hierarchy bound to a data role of a visual. */
export interface FieldTarget {
  table?: string;
  column?: string;
  measure?: string;
  hierarchy?: string;
  hierarchyLevel?: string;
  aggregationFunction?: string;
  [key: string]: unknown;
}

/** Field bindings of one visual: which fields sit in which data role. */
export interface VisualFields {
  name: string;
  title?: string;
  type: string;
  roles: { role: string; displayName?: string; fields: FieldTarget[] }[];
}

/** Active filters at report/page/visual level (slicer state for slicers). */
export interface FiltersState {
  report: unknown[];
  page: unknown[];
  visuals: { name: string; title?: string; type: string; filters: unknown[] | null; slicerState?: unknown }[];
}

export interface CaptureOptions {
  /** Internal page name (ReportSectionXXXX) or display name. */
  page?: string;
  /** powerbi-client filter objects to apply before capture. */
  filters?: unknown[];
  /** Slicer states to apply: visual name (or title) -> state object. */
  slicers?: Record<string, unknown>;
  /** Crop to a single visual (by name or title). */
  visual?: string;
  /** Viewport; defaults to the report page's own size. */
  width?: number;
  height?: number;
  /** Max ms to wait for each render. */
  renderTimeoutMs?: number;
}

const DEFAULT_RENDER_TIMEOUT = 60_000;

/**
 * Drives a headless Chromium with the embed page loaded, so callers can
 * inspect pages/visuals, apply filters and take screenshots. One session per
 * report; reuse it across multiple captures (each browser launch costs ~1s,
 * each report embed ~4-8s).
 */
export class CaptureSession {
  private mode: EmbedMode;

  private constructor(
    private browser: Browser,
    private host: EmbedHost,
    private page: Page,
    private target: EmbedTarget,
    mode: EmbedMode
  ) {
    this.mode = mode;
  }

  static async open(
    target: EmbedTarget,
    opts: { width?: number; height?: number; mode?: EmbedMode } = {}
  ): Promise<CaptureSession> {
    const host = await startEmbedHost();
    const { browser } = await launchBrowser();
    const page = await browser.newPage({
      viewport: { width: opts.width ?? 1600, height: opts.height ?? 900 },
      deviceScaleFactor: 2,
    });
    await page.goto(host.url);
    const session = new CaptureSession(browser, host, page, target, opts.mode ?? "view");
    await session.embed();
    return session;
  }

  private async embed(pageName?: string): Promise<void> {
    await this.page.evaluate(
      ([target, pn, mode]) => {
        (window as any).__pbi.init({ ...(target as object), pageName: pn, showNav: false, mode });
      },
      [this.target as unknown, pageName as unknown, this.mode as unknown] as const
    );
    await this.waitForRender();
  }

  /** Switch the embed between view and edit mode (re-renders). */
  async ensureMode(mode: EmbedMode): Promise<void> {
    if (this.mode === mode) return;
    await this.page.evaluate((m) => (window as any).__pbi.switchMode(m), mode);
    await this.waitForRender();
    this.mode = mode;
  }

  /** Swap in a fresh AAD token without re-embedding (long-lived sessions). */
  async refreshAccessToken(token: string): Promise<void> {
    this.target = { ...this.target, accessToken: token };
    await this.page.evaluate((t) => (window as any).__pbi.setAccessToken(t), token);
  }

  private async waitForRender(timeoutMs = DEFAULT_RENDER_TIMEOUT): Promise<void> {
    await this.page.waitForFunction(
      () => (window as any).__state.rendered || (window as any).__state.error,
      undefined,
      { timeout: timeoutMs }
    );
    const error = await this.page.evaluate(() => (window as any).__state.error);
    if (error) throw new Error(`Power BI embed error: ${error}`);
  }

  async getPages(): Promise<PageInfo[]> {
    return (await this.page.evaluate(() => (window as any).__pbi.getPages())) as PageInfo[];
  }

  async setPage(nameOrDisplayName: string): Promise<void> {
    const pages = await this.getPages();
    const target =
      pages.find((p) => p.name === nameOrDisplayName) ??
      pages.find((p) => p.displayName.toLowerCase() === nameOrDisplayName.toLowerCase());
    if (!target) {
      throw new Error(
        `Page "${nameOrDisplayName}" not found. Pages: ${pages.map((p) => p.displayName).join(", ")}`
      );
    }
    if (target.isActive) return;
    await this.page.evaluate((n) => (window as any).__pbi.setPage(n), target.name);
    await this.waitForRender();
  }

  async getVisuals(): Promise<VisualInfo[]> {
    return (await this.page.evaluate(() => (window as any).__pbi.getVisuals())) as VisualInfo[];
  }

  async setFilters(filters: unknown[]): Promise<void> {
    await this.page.evaluate((f) => (window as any).__pbi.setFilters(f), filters);
    await this.waitForRender();
  }

  async clearFilters(): Promise<void> {
    await this.page.evaluate(() => (window as any).__pbi.clearFilters());
    await this.waitForRender();
  }

  /** Find a visual on the active page by internal name or (case-insensitive) title. */
  private async resolveVisual(nameOrTitle: string): Promise<VisualInfo> {
    const visuals = await this.getVisuals();
    const target =
      visuals.find((v) => v.name === nameOrTitle) ??
      visuals.find((v) => v.title?.toLowerCase() === nameOrTitle.toLowerCase());
    if (!target) {
      throw new Error(
        `Visual "${nameOrTitle}" not found. Visuals: ${visuals.map((v) => v.title ?? v.name).join(", ")}`
      );
    }
    return target;
  }

  async setSlicer(visualNameOrTitle: string, state: unknown): Promise<void> {
    const target = await this.resolveVisual(visualNameOrTitle);
    await this.page.evaluate(
      ([name, s]) => (window as any).__pbi.setSlicer(name, s),
      [target.name, state] as const
    );
    await this.waitForRender();
  }

  /**
   * Field bindings of one visual (which column/measure on which role/axis).
   * The report-authoring APIs may refuse in view mode; on failure the session
   * flips to edit mode once and retries (screenshots flip it back).
   */
  async getVisualFields(visualNameOrTitle: string): Promise<VisualFields> {
    const target = await this.resolveVisual(visualNameOrTitle);
    try {
      return (await this.page.evaluate(
        (n) => (window as any).__pbi.getVisualFields(n),
        target.name
      )) as VisualFields;
    } catch (first) {
      if (this.mode === "edit") throw first;
      await this.ensureMode("edit");
      try {
        return (await this.page.evaluate(
          (n) => (window as any).__pbi.getVisualFields(n),
          target.name
        )) as VisualFields;
      } catch (second) {
        throw new Error(
          `Field readback failed in view mode (${(first as Error).message}) and edit mode ` +
            `(${(second as Error).message}). Field readback needs edit rights on the report.`
        );
      }
    }
  }

  /** Data points of one visual as CSV (SDK exportData; tenant may disable export). */
  async exportVisualData(
    visualNameOrTitle: string,
    opts: { exportType?: "summarized" | "underlying"; rows?: number } = {}
  ): Promise<string> {
    const target = await this.resolveVisual(visualNameOrTitle);
    const result = (await this.page.evaluate(
      ([name, type, rows]) => (window as any).__pbi.exportVisualData(name, type, rows),
      [target.name, opts.exportType ?? "summarized", opts.rows] as const
    )) as { data?: string };
    if (typeof result?.data !== "string") {
      throw new Error("exportData returned no data (export may be disabled for this report/tenant).");
    }
    return result.data;
  }

  /** Active filter state of the report, active page and every visual on it. */
  async getFiltersState(): Promise<FiltersState> {
    return (await this.page.evaluate(() => (window as any).__pbi.getFilters())) as FiltersState;
  }

  /**
   * Screenshot the active page (or one visual cropped out of it) as a PNG buffer.
   * Always drops back to view mode first so edit chrome never leaks into shots.
   */
  async screenshotBuffer(visualNameOrTitle?: string): Promise<Buffer> {
    await this.ensureMode("view");
    // Let any in-flight animations settle.
    await this.page.waitForTimeout(300);
    const container = this.page.locator("#container iframe").first();
    if (!visualNameOrTitle) {
      return await container.screenshot();
    }
    const target = await this.resolveVisual(visualNameOrTitle);
    if (!target.layout) {
      throw new Error(`Visual "${visualNameOrTitle}" has no layout to crop to.`);
    }
    // Visual layout is in report-page coordinates; scale to rendered iframe size.
    const pageSize = (await this.page.evaluate(() => (window as any).__pbi.getActivePageSize())) as {
      width?: number;
      height?: number;
    };
    const box = await container.boundingBox();
    if (!box || !pageSize.width || !pageSize.height) {
      throw new Error("Could not determine geometry for visual crop.");
    }
    const scale = Math.min(box.width / pageSize.width, box.height / pageSize.height);
    const offsetX = box.x + (box.width - pageSize.width * scale) / 2;
    const offsetY = box.y + (box.height - pageSize.height * scale) / 2;
    return await this.page.screenshot({
      clip: {
        x: offsetX + target.layout.x * scale,
        y: offsetY + target.layout.y * scale,
        width: target.layout.width * scale,
        height: target.layout.height * scale,
      },
    });
  }

  /**
   * Screenshot the active page (or one visual cropped out of it) to a PNG file.
   */
  async screenshot(outPath: string, visualNameOrTitle?: string): Promise<void> {
    fs.writeFileSync(outPath, await this.screenshotBuffer(visualNameOrTitle));
  }

  async close(): Promise<void> {
    await this.browser.close().catch(() => {});
    await this.host.close().catch(() => {});
  }
}
