export type Zoom = number | 'fit' | 'page';
export const clampZoom = (percent: number) => Math.min(400, Math.max(10, Math.round(percent)));
export function pageScale(zoom: Zoom, available: { width: number; height: number }, page: { width: number; height: number }) {
  if (typeof zoom === 'number') return clampZoom(zoom) / 100;
  const width = Math.max(1, available.width) / page.width;
  return Math.min(4, Math.max(.01, zoom === 'page' ? Math.min(width, Math.max(1, available.height) / page.height) : width));
}
export function stepZoom(percent: number, direction: number) {
  const step = percent < 100 ? 10 : 25;
  return clampZoom(direction > 0 ? (Math.floor(percent / step) + 1) * step : (Math.ceil(percent / step) - 1) * step);
}
