export type PageInfo = { pageIndex: number; width: number; height: number; sectionIndex: number; text: string };
export type Figure = { id: string; page: number; type: string; x: number; y: number; w: number; h: number; section: number };
export type VisualTable = Figure & { rows: number; columns: number };
export type RenderedPage = { page: number; svg: string };
export type ViewerMessage = { type: 'load'; bytes: ArrayBuffer } | { type: 'render'; page: number } | { type: 'activity'; active: boolean };
export type ViewerResult = { type: 'ready'; pages: PageInfo[]; figures: Figure[]; tables: VisualTable[] } | { type: 'page'; page: number; svg: string } | { type: 'indexed-batch'; entries: { page: number; text: string; limited: boolean }[] } | { type: 'error'; message: string } | { type: 'progress'; message: string };
