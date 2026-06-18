import * as fs from "fs";
import { PNG } from "pngjs";
import pixelmatch from "pixelmatch";

export interface PngDiff {
  width: number;
  height: number;
  changedPixels: number;
  totalPixels: number;
  /** Fraction of pixels that differ, 0..1. */
  fraction: number;
}

/**
 * Pixel-diff two PNGs and write a highlighted diff image to `outPath`.
 * The two renders are padded to the larger common canvas first, so a page that
 * changed size still compares (the extra area simply registers as changed).
 */
export function diffPng(aPath: string, bPath: string, outPath: string): PngDiff {
  const a = PNG.sync.read(fs.readFileSync(aPath));
  const b = PNG.sync.read(fs.readFileSync(bPath));
  const width = Math.max(a.width, b.width);
  const height = Math.max(a.height, b.height);
  const ca = fitCanvas(a, width, height);
  const cb = fitCanvas(b, width, height);
  const diff = new PNG({ width, height });
  const changedPixels = pixelmatch(ca.data, cb.data, diff.data, width, height, { threshold: 0.1 });
  fs.writeFileSync(outPath, PNG.sync.write(diff));
  const totalPixels = width * height;
  return { width, height, changedPixels, totalPixels, fraction: totalPixels ? changedPixels / totalPixels : 0 };
}

/** Return `src` unchanged if it already fills the canvas, else copy it into a fresh one. */
function fitCanvas(src: PNG, width: number, height: number): PNG {
  if (src.width === width && src.height === height) return src;
  const out = new PNG({ width, height });
  PNG.bitblt(src, out, 0, 0, src.width, src.height, 0, 0);
  return out;
}
