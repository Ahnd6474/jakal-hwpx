import init, { HwpDocument } from '@rhwp/core';
import wasmUrl from '@rhwp/core/rhwp_bg.wasm?url';
import type { Figure, PageInfo, VisualTable, ViewerMessage, ViewerResult } from './viewer-types';
import { renderedText } from './viewer-text';
import { ActivityGate, IndexBatch } from './viewer-index';

let document: HwpDocument | null = null;
let initialize: Promise<unknown> | null = null;
const activity = new ActivityGate();
const canvas = new OffscreenCanvas(1, 1);
const context = canvas.getContext('2d')!;
(globalThis as unknown as { measureTextWidth: (font: string, text: string) => number }).measureTextWidth = (font, text) => {
  context.font = font;
  return context.measureText(text).width;
};
const send = (message: ViewerResult) => postMessage(message);
const yieldTask = () => new Promise(resolve => setTimeout(resolve, 0));
const cache = new Map<number, string>();
let cacheSize = 0;
function pageSvg(page: number): string {
  const cached = cache.get(page);
  if (cached) { cache.delete(page); cache.set(page, cached); return cached; }
  const svg = document!.renderPageSvg(page);
  if (svg.length > 12_000_000) throw new Error('페이지 렌더링 크기 제한');
  cache.set(page, svg); cacheSize += svg.length;
  while (cacheSize > 24_000_000 && cache.size > 1) { const first = cache.keys().next().value!; cacheSize -= cache.get(first)!.length; cache.delete(first); }
  return svg;
}

onmessage = async (event: MessageEvent<ViewerMessage>) => {
  if (event.data.type === 'activity') { activity.setActive(event.data.active); return; }
  try {
    if (event.data.type === 'load') {
      if (!initialize) initialize = init({ module_or_path: wasmUrl });
      await initialize;
      send({ type: 'progress', message: '원래 페이지 배치를 읽고 있습니다' });
      document?.free();
      cache.clear(); cacheSize = 0;
      document = new HwpDocument(new Uint8Array(event.data.bytes));
      const count = document.pageCount();
      if (count < 1 || count > 500) throw new Error('페이지 수 제한');
      const pages: PageInfo[] = [];
      const figures: Figure[] = [];
      const tables: VisualTable[] = [];
      for (let page = 0; page < count; page++) {
        const info = JSON.parse(document.getPageInfo(page));
        if (!Number.isFinite(info.width) || !Number.isFinite(info.height) || info.width <= 0 || info.height <= 0 || info.width > 20000 || info.height > 20000) throw new Error('페이지 크기 제한');
        pages.push({ pageIndex: page, width: info.width, height: info.height, sectionIndex: info.sectionIndex, text: '' });
        const controls = JSON.parse(document.getPageControlLayout(page)).controls || [];
        for (const [index, c] of controls.entries()) {
          if (c.type === 'table' && [c.x, c.y, c.w, c.h].every(Number.isFinite) && c.w > 0 && c.h > 0) {
            tables.push({ id: `page-${page}-table-${index}`, page, type: 'table', x: c.x, y: c.y, w: c.w, h: c.h, section: c.secIdx ?? info.sectionIndex, rows: c.rowCount || 0, columns: c.colCount || 0 });
            if (tables.length > 5000) throw new Error('표 수 제한');
          }
          // Lines and header ornaments are not standalone figures. A graph may
          // be a chart, embedded image, grouped drawing, or freeform shape.
          if (!['chart', 'image', 'shape', 'group', 'ole'].includes(c.type) || c.headerFooter || ![c.x, c.y, c.w, c.h].every(Number.isFinite) || c.w < 8 || c.h < 8) continue;
          figures.push({ id: `page-${page}-figure-${index}`, page, type: c.type, x: c.x, y: c.y, w: c.w, h: c.h, section: c.secIdx ?? info.sectionIndex });
          if (figures.length > 5000) throw new Error('개체 수 제한');
        }
        if (page % 5 === 0) { send({ type: 'progress', message: `페이지와 그래프·그림을 정리하고 있습니다 (${page + 1}/${count})` }); await yieldTask(); }
      }
      send({ type: 'ready', pages, figures, tables });
      // Build search from the text actually painted on each page. Paragraph-
      // based extraction can include text that flows onto an adjacent page.
      let textBudget = 2_000_000;
      const batch = new IndexBatch(entries => send({ type: 'indexed-batch', entries }));
      for (let page = 0; page < count; page++) {
        await yieldTask();
        if (!activity.active) batch.flush();
        await activity.wait();
        // Search is explicitly bounded; do not render pages whose text cannot
        // be retained. They remain available for normal on-demand viewing.
        if (textBudget <= 0) { batch.add({ page, text: '', limited: true }); continue; }
        try {
          const text = renderedText(pageSvg(page));
          const limit = Math.min(100_000, textBudget);
          const indexed = text.slice(0, limit);
          textBudget -= indexed.length;
          batch.add({ page, text: indexed, limited: indexed.length < text.length });
        } catch { batch.add({ page, text: '', limited: true }); }
      }
      batch.flush();
    } else if (document) {
      const page = event.data.page;
      if (!Number.isInteger(page) || page < 0 || page >= document.pageCount()) return;
      const svg = pageSvg(page);
      send({ type: 'page', page, svg });
    }
  } catch {
    send({ type: 'error', message: '이 문서의 일부 배치를 표시하지 못했습니다. 원본 파일을 내려받거나 텍스트·표 보기에서 내용을 확인할 수 있습니다.' });
  }
};
