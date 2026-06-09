import { chromium, Browser, Page } from "playwright";
import { startEmbedHost, EmbedHost } from "./embedHost";

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
  private constructor(
    private browser: Browser,
    private host: EmbedHost,
    private page: Page,
    private target: EmbedTarget
  ) {}

  static async open(target: EmbedTarget, opts: { width?: number; height?: number } = {}): Promise<CaptureSession> {
    const host = await startEmbedHost();
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({
      viewport: { width: opts.width ?? 1600, height: opts.height ?? 900 },
      deviceScaleFactor: 2,
    });
    await page.goto(host.url);
    const session = new CaptureSession(browser, host, page, target);
    await session.embed();
    return session;
  }

  private async embed(pageName?: string): Promise<void> {
    await this.page.evaluate(
      ([target, pn]) => {
        (window as any).__pbi.init({ ...(target as object), pageName: pn, showNav: false });
      },
      [this.target as unknown, pageName as unknown] as const
    );
    await this.waitForRender();
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

  async setSlicer(visualNameOrTitle: string, state: unknown): Promise<void> {
    const visuals = await this.getVisuals();
    const target =
      visuals.find((v) => v.name === visualNameOrTitle) ??
      visuals.find((v) => v.title?.toLowerCase() === visualNameOrTitle.toLowerCase());
    if (!target) {
      throw new Error(
        `Slicer "${visualNameOrTitle}" not found. Visuals: ${visuals.map((v) => v.title ?? v.name).join(", ")}`
      );
    }
    await this.page.evaluate(
      ([name, s]) => (window as any).__pbi.setSlicer(name, s),
      [target.name, state] as const
    );
    await this.waitForRender();
  }

  /**
   * Screenshot the active page (or one visual cropped out of it) to a PNG.
   */
  async screenshot(outPath: string, visualNameOrTitle?: string): Promise<void> {
    // Let any in-flight animations settle.
    await this.page.waitForTimeout(300);
    const container = this.page.locator("#container iframe").first();
    if (!visualNameOrTitle) {
      await container.screenshot({ path: outPath });
      return;
    }
    const visuals = await this.getVisuals();
    const target =
      visuals.find((v) => v.name === visualNameOrTitle) ??
      visuals.find((v) => v.title?.toLowerCase() === visualNameOrTitle.toLowerCase());
    if (!target?.layout) {
      throw new Error(
        `Visual "${visualNameOrTitle}" not found or has no layout. Visuals: ${visuals
          .map((v) => v.title ?? v.name)
          .join(", ")}`
      );
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
    await this.page.screenshot({
      path: outPath,
      clip: {
        x: offsetX + target.layout.x * scale,
        y: offsetY + target.layout.y * scale,
        width: target.layout.width * scale,
        height: target.layout.height * scale,
      },
    });
  }

  async close(): Promise<void> {
    await this.browser.close().catch(() => {});
    await this.host.close().catch(() => {});
  }
}
