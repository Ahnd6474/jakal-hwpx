import { pageScale, type Zoom } from './viewer-zoom.ts';

export type Sheet = { index: number; top: number; height: number; width: number; paperHeight: number; scale: number };
export function sheetLayout(pages: readonly { width: number; height: number }[], zoom: Zoom, available: { width: number; height: number }): Sheet[] {
  let top = 0;
  return pages.map((page, index) => {
    const scale = pageScale(zoom, available, page);
    const sheet = { index, top, height: page.height * scale + 52, width: page.width * scale, paperHeight: page.height * scale, scale };
    top += sheet.height;
    return sheet;
  });
}

export function pageAtOffset(sheets: readonly Sheet[], offset: number): number {
  let low = 0, high = sheets.length - 1;
  while (low < high) {
    const mid = Math.ceil((low + high) / 2);
    if (sheets[mid].top <= offset) low = mid;
    else high = mid - 1;
  }
  return Math.max(0, low);
}

export function pageWindow(sheets: readonly Sheet[], top: number, height: number) {
  const first = pageAtOffset(sheets, Math.max(0, top));
  const last = pageAtOffset(sheets, Math.max(0, top + height));
  // Keep the live SVG DOM bounded even for hundreds of unusually short pages.
  return { start: Math.max(0, first - 1), end: Math.min(sheets.length - 1, last + 1, first + 22) };
}

export function readingPage(sheets: readonly Sheet[], top: number, height: number): number {
  if (!sheets.length) return 0;
  const first = pageAtOffset(sheets, top);
  return pageAtOffset(sheets, top + Math.min(height / 2, sheets[first].height / 2));
}
