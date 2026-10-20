import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { ArrowDownToLine, ArrowLeft, ArrowRight, ChevronDown, ChevronLeft, ChevronRight, ChevronUp, FileText, Images, LoaderCircle, Maximize2, Minus, Plus, Search, X } from 'lucide-react';
import type { Document } from './api';
import type { Figure, PageInfo, VisualTable, ViewerResult } from './viewer-types';
import { cropPng, highlightSvg, safeSvg } from './viewer-svg';
import { clampZoom, stepZoom, type Zoom } from './viewer-zoom';
import { pageWindow, readingPage, sheetLayout } from './reader-layout';

type Props = { doc: Document; active: boolean; mode: 'original' | 'figures' | 'tables'; dirty: number; onMode: (mode: 'original' | 'figures') => void; onEdit: () => void; onEditTable: () => void; toolsCollapsed: boolean; onToggleTools: () => void; onReadingScroll: () => void };
const figureKind = (type: string) => type === 'table' ? '표' : type === 'chart' ? '차트' : type === 'image' ? '그림' : type === 'ole' ? '삽입 개체' : '도형';

function PageArtwork({ svg, query, info, scale }: { svg: string; query: string; info: PageInfo; scale: number }) {
  const content = useMemo(() => highlightSvg(svg, query), [svg, query]);
  return <div className="page-svg" role="document" aria-label={`${info.pageIndex + 1}쪽 원문`} style={{ width: info.width, height: info.height, transform: `scale(${scale})`, transformOrigin: 'top left' }} dangerouslySetInnerHTML={{ __html: content }} />;
}

function FigureImage({ svg, figure, label }: { svg?: string; figure: Figure; label: string }) {
  const [url, setUrl] = useState('');
  const [error, setError] = useState(false);
  const imageUrl = useRef('');
  useEffect(() => { setUrl(''); setError(false); return () => { if (imageUrl.current) URL.revokeObjectURL(imageUrl.current); imageUrl.current = ''; }; }, [figure.id]);
  useEffect(() => {
    if (!svg) return;
    let cancelled = false;
    setError(false);
    void cropPng(svg, figure).then(blob => { if (!cancelled) { if (imageUrl.current) URL.revokeObjectURL(imageUrl.current); imageUrl.current = URL.createObjectURL(blob); setUrl(imageUrl.current); } }).catch(() => { if (!cancelled) setError(true); });
    return () => { cancelled = true; };
  }, [svg, figure]);
  return url ? <img src={url} alt={label} draggable={false} /> : <span className="figure-loading">{error ? '원문 위치에서 확인해 주세요' : <><LoaderCircle size={19} className="spin" />영역을 불러오는 중</>}</span>;
}

export function DocumentReader({ doc, active, mode, dirty, onMode, onEdit, onEditTable, toolsCollapsed, onToggleTools, onReadingScroll }: Props) {
  const [original, setOriginal] = useState(false);
  const revision = original ? 0 : doc.revision;
  const [pages, setPages] = useState<PageInfo[]>([]);
  const [pageTexts, setPageTexts] = useState<Record<number, string>>({});
  const [range, setRange] = useState({ start: 0, end: 1 });
  const svgCache = useRef(new Map<number, string>());
  const wantedPages = useRef<number[]>([]);
  const pendingPage = useRef<number | null>(null);
  const scrollFrame = useRef(0);
  const ignoreScrollUntil = useRef(0);
  const lastScroll = useRef(0);
  const downwardTravel = useRef(0);
  const originalPosition = useRef({ top: 0, left: 0 });
  const previousMode = useRef(mode);
  const [figures, setFigures] = useState<Figure[]>([]);
  const [tables, setTables] = useState<VisualTable[]>([]);
  const [tableId, setTableId] = useState('');
  const [rendered, setRendered] = useState<Record<number, string>>({});
  const [page, setPage] = useState(0);
  const currentPageRef = useRef(page);
  currentPageRef.current = page;
  const [pageInput, setPageInput] = useState('1');
  const [zoom, setZoom] = useState<Zoom>('fit');
  const [available, setAvailable] = useState({ width: 736, height: 600 });
  const [query, setQuery] = useState('');
  const [loading, setLoading] = useState('원문 보기를 준비하고 있습니다');
  const [error, setError] = useState('');
  const [retry, setRetry] = useState(0);
  const [figureFilter, setFigureFilter] = useState('all');
  const [figureSection, setFigureSection] = useState(-1);
  const [visibleCount, setVisibleCount] = useState(12);
  const [selectedFigure, setSelectedFigure] = useState<Figure | null>(null);
  const workerRef = useRef<Worker | null>(null);
  const requested = useRef(new Set<number>());
  const readerRef = useRef<HTMLDivElement>(null);
  const surfaceRef = useRef<HTMLDivElement>(null);
  const framesRef = useRef(new Map<number, HTMLDivElement>());
  const zoomAnchor = useRef<{ page: number; x: number; y: number; screenX: number; screenY: number } | null>(null);
  const pinchRef = useRef<{ distance: number; percent: number } | null>(null);
  const searchRef = useRef<HTMLInputElement>(null);
  const pageInputRef = useRef<HTMLInputElement>(null);
  const timers = useRef(new Map<number, ReturnType<typeof setTimeout>>());
  const [announcement, setAnnouncement] = useState('');
  const [indexedCount, setIndexedCount] = useState(0);
  const [searchLimited, setSearchLimited] = useState(false);
  const layout = useMemo(() => sheetLayout(pages, zoom, available), [pages, zoom, available]);
  const current = pages[page];
  const scale = layout[page]?.scale ?? 1;
  const updateViewport = useCallback(() => {
    const surface = surfaceRef.current;
    if (!surface || !layout.length || mode !== 'original' || !active) return;
    const top = Math.max(0, surface.scrollTop - parseFloat(getComputedStyle(surface).paddingTop));
    originalPosition.current = { top: surface.scrollTop, left: surface.scrollLeft };
    const next = pageWindow(layout, top, surface.clientHeight);
    setRange(previous => previous.start === next.start && previous.end === next.end ? previous : next);
    const atEnd = surface.scrollHeight > surface.clientHeight + 2 && surface.scrollHeight - surface.clientHeight - surface.scrollTop < 2;
    setPage(atEnd ? layout.length - 1 : readingPage(layout, top, surface.clientHeight));
  }, [layout, mode, active]);
  const jumpToPage = useCallback((index: number) => {
    const entry = layout[index];
    const surface = surfaceRef.current;
    if (!entry || !surface) return;
    if (mode !== 'original' || !active) { pendingPage.current = index; return; }
    ignoreScrollUntil.current = performance.now() + 400;
    downwardTravel.current = 0;
    const top = entry.top + parseFloat(getComputedStyle(surface).paddingTop);
    surface.scrollTo({ top });
    lastScroll.current = surface.scrollTop;
    setPage(index);
    setRange(pageWindow(layout, entry.top, surface.clientHeight));
  }, [layout, mode, active]);
  const captureAnchor = useCallback((point?: { x: number; y: number }) => {
    const surface = surfaceRef.current;
    if (surface) {
      const bounds = surface.getBoundingClientRect();
      const screenX = point?.x ?? bounds.left + surface.clientWidth / 2;
      const screenY = point?.y ?? bounds.top + Math.min(surface.clientHeight / 2, 200);
      const target = [...framesRef.current.entries()].find(([, frame]) => { const r = frame.getBoundingClientRect(); return r.top <= screenY && r.bottom >= screenY; });
      const index = target?.[0] ?? page;
      const frame = target?.[1] ?? framesRef.current.get(index);
      if (frame) {
        const paper = frame.getBoundingClientRect();
        const oldScale = layout[index]?.scale ?? 1;
        zoomAnchor.current = { page: index, x: (screenX - paper.left) / oldScale, y: (screenY - paper.top) / oldScale, screenX, screenY };
      }
    }
  }, [layout, page]);
  const captureAnchorRef = useRef(captureAnchor);
  captureAnchorRef.current = captureAnchor;
  const availableRef = useRef(available);
  const changeZoom = useCallback((value: Zoom, point?: { x: number; y: number }) => {
    captureAnchor(point);
    ignoreScrollUntil.current = performance.now() + 400;
    setZoom(typeof value === 'number' ? clampZoom(value) : value);
  }, [captureAnchor]);
  useLayoutEffect(() => {
    const anchor = zoomAnchor.current;
    const surface = surfaceRef.current;
    const frame = anchor ? framesRef.current.get(anchor.page) : undefined;
    if (anchor && surface && frame) {
      const bounds = frame.getBoundingClientRect();
      const nextScale = layout[anchor.page]?.scale ?? 1;
      surface.scrollLeft += bounds.left + anchor.x * nextScale - anchor.screenX;
      surface.scrollTop += bounds.top + anchor.y * nextScale - anchor.screenY;
      zoomAnchor.current = null;
    }
    if (mode === 'original' && active && pendingPage.current !== null && layout.length) {
      const target = pendingPage.current;
      pendingPage.current = null;
      jumpToPage(target);
    } else {
      if (mode === 'original' && previousMode.current !== 'original' && surface) {
        surface.scrollTo(originalPosition.current);
        ignoreScrollUntil.current = performance.now() + 400;
      }
      updateViewport();
    }
    previousMode.current = mode;
  }, [layout, mode, active, jumpToPage, updateViewport]);
  function handleScroll() {
    if (scrollFrame.current) return;
    scrollFrame.current = requestAnimationFrame(() => {
      scrollFrame.current = 0;
      const surface = surfaceRef.current;
      if (!surface || mode !== 'original') return;
      const delta = surface.scrollTop - lastScroll.current;
      lastScroll.current = surface.scrollTop;
      if (performance.now() >= ignoreScrollUntil.current && !toolsCollapsed && !document.querySelector('[role="dialog"]') && !(document.activeElement instanceof HTMLInputElement || document.activeElement instanceof HTMLTextAreaElement || document.activeElement instanceof HTMLSelectElement)) {
        downwardTravel.current = delta > 0 ? downwardTravel.current + delta : 0;
        if (downwardTravel.current > 100) { downwardTravel.current = 0; onReadingScroll(); }
      }
      updateViewport();
    });
  }
  useEffect(() => () => cancelAnimationFrame(scrollFrame.current), []);

  useEffect(() => {
    const abort = new AbortController();
    const worker = new Worker(new URL('./document.worker.ts', import.meta.url), { type: 'module' });
    workerRef.current = worker;
    requested.current.clear(); svgCache.current.clear(); framesRef.current.clear(); setPageTexts({}); setRange({ start: 0, end: 1 }); pendingPage.current = 0;
    setIndexedCount(0); setSearchLimited(false);
    setPages([]); setFigures([]); setTables([]); setTableId(''); setRendered({}); setPage(0); setPageInput('1'); setQuery(''); setSelectedFigure(null); setError(''); setLoading('원문 보기를 준비하고 있습니다');
    const loadTimeout = setTimeout(() => { worker.terminate(); setLoading(''); setError('원문을 여는 데 시간이 오래 걸립니다. 다시 시도하거나 텍스트·표 보기로 이동해 주세요.'); }, 60_000);
    worker.onmessage = (event: MessageEvent<ViewerResult>) => {
      const message = event.data;
      if (message.type === 'progress') setLoading(message.message);
      else if (message.type === 'ready') { clearTimeout(loadTimeout); setPages(message.pages); setFigures(message.figures); setTables(message.tables); setLoading(''); }
      else if (message.type === 'page') {
        clearTimeout(timers.current.get(message.page)); timers.current.delete(message.page);
        try {
          const cache = svgCache.current;
          cache.delete(message.page); cache.set(message.page, safeSvg(message.svg));
          let size = [...cache.values()].reduce((total, svg) => total + svg.length, 0);
          while ((size > 32_000_000 || cache.size > 24) && cache.size > 1) {
            const victim = [...cache.keys()].find(index => !wantedPages.current.includes(index) && index !== currentPageRef.current) ?? [...cache.keys()].find(index => index !== currentPageRef.current)!;
            size -= cache.get(victim)!.length; cache.delete(victim); requested.current.delete(victim);
          }
          setRendered(Object.fromEntries(cache));
        }
        catch { setError('이 페이지를 안전하게 표시하지 못했습니다. 원본 파일에서 확인해 주세요.'); }
      } else if (message.type === 'indexed') {
        setPageTexts(previous => ({ ...previous, [message.page]: message.text }));
        setIndexedCount(n => n + 1);
        if (message.limited) setSearchLimited(true);
      } else if (message.type === 'error') { clearTimeout(loadTimeout); setError(message.message); setLoading(''); }
    };
    worker.onerror = () => { clearTimeout(loadTimeout); setLoading(''); setError('원문 뷰어를 시작하지 못했습니다. 다시 시도하거나 텍스트·표 보기로 이동해 주세요.'); };
    void fetch(`/api/documents/${doc.id}/source?revision=${revision}`, { signal: abort.signal }).then(async response => {
      if (!response.ok) throw new Error('원문 파일을 불러오지 못했습니다. 문서를 다시 열어 주세요.');
      const bytes = await response.arrayBuffer();
      if (!abort.signal.aborted) worker.postMessage({ type: 'load', bytes }, [bytes]);
    }).catch(e => { if (!abort.signal.aborted) { clearTimeout(loadTimeout); setLoading(''); setError(e.message); } });
    return () => { abort.abort(); clearTimeout(loadTimeout); worker.terminate(); workerRef.current = null; timers.current.forEach(clearTimeout); timers.current.clear(); };
  }, [doc.id, revision, retry]);

  const requestPage = useCallback((index: number) => {
    if (requested.current.has(index) || !workerRef.current) return;
    requested.current.add(index);
    workerRef.current.postMessage({ type: 'render', page: index });
    timers.current.set(index, setTimeout(() => { setError('페이지 표시 시간이 초과되었습니다. 범위를 줄이거나 다시 시도해 주세요.'); }, 45_000));
  }, []);
  const matchingPages = useMemo(() => query.trim() ? pages.filter(p => (pageTexts[p.pageIndex] || '').toLocaleLowerCase().includes(query.replace(/\s/g, '').toLocaleLowerCase())).map(p => p.pageIndex) : [], [pages, pageTexts, query]);
  const filteredFigures = useMemo(() => figures.filter(f => (figureSection === -1 || f.section === figureSection) && (figureFilter === 'all' || (figureFilter === 'chart' ? f.type === 'chart' : f.type !== 'chart'))), [figures, figureFilter, figureSection]);
  const currentTable = tables.find(t => t.id === tableId) || tables[0];

  useEffect(() => {
    if (!pages.length || !active || error) return;
    if (mode === 'original') { wantedPages.current = layout.slice(range.start, range.end + 1).map(entry => entry.index); wantedPages.current.forEach(requestPage); }
    else if (mode === 'tables') { if (currentTable) requestPage(currentTable.page); }
    else if (selectedFigure) requestPage(selectedFigure.page);
    else [...new Set(filteredFigures.slice(0, visibleCount).map(f => f.page))].forEach(requestPage);
  }, [pages, layout, range, active, mode, page, currentTable, selectedFigure, filteredFigures, visibleCount, requestPage, error]);
  useEffect(() => { if (document.activeElement !== pageInputRef.current) setPageInput(String(page + 1)); }, [page]);
  useEffect(() => { if (mode !== 'original') surfaceRef.current?.scrollTo({ top: 0, left: 0 }); }, [mode, selectedFigure?.id, currentTable?.id]);
  useEffect(() => { setVisibleCount(12); setSelectedFigure(null); }, [figureFilter, figureSection]);
  useEffect(() => {
    const surface = surfaceRef.current;
    if (!surface) return;
    const observer = new ResizeObserver(() => {
      if (!surface.clientWidth || !surface.clientHeight) return;
      const css = getComputedStyle(surface);
      const next = { width: surface.clientWidth - parseFloat(css.paddingLeft) - parseFloat(css.paddingRight), height: surface.clientHeight - parseFloat(css.paddingTop) - parseFloat(css.paddingBottom) - 28 };
      if (availableRef.current.width === next.width && availableRef.current.height === next.height) return;
      captureAnchorRef.current();
      availableRef.current = next;
      ignoreScrollUntil.current = performance.now() + 300;
      setAvailable(previous => previous.width === next.width && previous.height === next.height ? previous : next);
    });
    observer.observe(surface);
    return () => observer.disconnect();
  }, []);
  useEffect(() => {
    if (!active) return;
    function keyboard(event: KeyboardEvent) {
      if (document.querySelector('[role="dialog"]')) return;
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f' && mode === 'original') { event.preventDefault(); if (toolsCollapsed) onToggleTools(); requestAnimationFrame(() => searchRef.current?.focus()); }
      if ((event.ctrlKey || event.metaKey) && mode === 'original' && current && !(event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement || event.target instanceof HTMLSelectElement)) {
        if (['+', '=', '-', '0'].includes(event.key)) { event.preventDefault(); changeZoom(event.key === '0' ? 'fit' : stepZoom(scale * 100, event.key === '-' ? -1 : 1)); }
      }
      if (event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement || event.target instanceof HTMLSelectElement || event.ctrlKey || event.metaKey || event.altKey) return;
      if (mode !== 'original' || !pages.length) return;
      if (event.key === 'ArrowLeft') { event.preventDefault(); jumpToPage(Math.max(0, page - 1)); }
      if (event.key === 'ArrowRight') { event.preventDefault(); jumpToPage(Math.min(pages.length - 1, page + 1)); }
    }
    window.addEventListener('keydown', keyboard);
    return () => window.removeEventListener('keydown', keyboard);
  }, [active, mode, pages.length, current, scale, changeZoom, jumpToPage, page, toolsCollapsed, onToggleTools]);
  useEffect(() => {
    const surface = surfaceRef.current;
    if (!surface || !active || mode !== 'original' || !current) return;
    const distance = (touches: TouchList) => Math.hypot(touches[0].clientX - touches[1].clientX, touches[0].clientY - touches[1].clientY);
    const wheel = (event: WheelEvent) => {
      if (!event.ctrlKey && !event.metaKey) return;
      event.preventDefault();
      changeZoom(scale * 100 * Math.exp(-event.deltaY * .005), { x: event.clientX, y: event.clientY });
    };
    const start = (event: TouchEvent) => {
      if (event.touches.length !== 2) return;
      event.preventDefault();
      pinchRef.current = { distance: distance(event.touches), percent: scale * 100 };
    };
    const move = (event: TouchEvent) => {
      const pinch = pinchRef.current;
      if (!pinch || event.touches.length !== 2) return;
      event.preventDefault();
      changeZoom(pinch.percent * distance(event.touches) / Math.max(1, pinch.distance), { x: (event.touches[0].clientX + event.touches[1].clientX) / 2, y: (event.touches[0].clientY + event.touches[1].clientY) / 2 });
    };
    const end = () => { pinchRef.current = null; };
    surface.addEventListener('wheel', wheel, { passive: false });
    surface.addEventListener('touchstart', start, { passive: false });
    surface.addEventListener('touchmove', move, { passive: false });
    surface.addEventListener('touchend', end);
    surface.addEventListener('touchcancel', end);
    return () => { surface.removeEventListener('wheel', wheel); surface.removeEventListener('touchstart', start); surface.removeEventListener('touchmove', move); surface.removeEventListener('touchend', end); surface.removeEventListener('touchcancel', end); };
  }, [active, mode, current, scale, changeZoom]);

  function navigate(value: string) {
    const next = Number(value);
    if (Number.isInteger(next) && next >= 1 && next <= pages.length) { jumpToPage(next - 1); setPageInput(String(next)); }
    else setPageInput(String(page + 1));
  }
  function nextMatch(direction: number) {
    if (!matchingPages.length) return;
    const next = direction > 0 ? matchingPages.find(n => n > page) ?? matchingPages[0] : [...matchingPages].reverse().find(n => n < page) ?? matchingPages.at(-1)!;
    jumpToPage(next);
  }
  function showPage(index: number) { setSelectedFigure(null); pendingPage.current = index; onMode('original'); }
  async function downloadFigure(figure: Figure) {
    const svg = rendered[figure.page];
    if (!svg) return;
    try {
      const url = URL.createObjectURL(await cropPng(svg, figure));
      const a = document.createElement('a'); a.href = url; a.download = `${doc.name.replace(/\.[^.]+$/, '')}_${figure.page + 1}쪽_${figureKind(figure.type)}.png`; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
      setAnnouncement('선택한 원문 영역을 PNG 이미지로 저장했습니다.');
    } catch { setAnnouncement('이미지를 저장하지 못했습니다. 원문 위치에서 확인해 주세요.'); }
  }
  async function fullscreen() {
    try { if (document.fullscreenElement) await document.exitFullscreen(); else await readerRef.current?.requestFullscreen(); }
    catch { setAnnouncement('이 브라우저에서는 전체 화면을 사용할 수 없습니다.'); }
  }
  const sourceUrl = `/api/documents/${doc.id}/original`;


  return <div ref={readerRef} className="document-reader" hidden={!active}>
    <div className="reader-toolbar" aria-label="원문 읽기 도구">
      {mode === 'original' ? <>
        <div className="page-navigation"><button className="icon-button" aria-label="이전 페이지" disabled={!pages.length || page === 0} onClick={() => jumpToPage(Math.max(0, page - 1))}><ChevronLeft size={17} /></button><label><span className="sr-only">페이지 번호</span><input ref={pageInputRef} inputMode="numeric" aria-label="페이지 번호" value={pageInput} disabled={!pages.length} onChange={e => setPageInput(e.target.value)} onBlur={e => navigate(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') navigate(pageInput); }} /></label><span>/ {pages.length || '—'}</span><button className="icon-button" aria-label="다음 페이지" disabled={!pages.length || page >= pages.length - 1} onClick={() => jumpToPage(Math.min(pages.length - 1, page + 1))}><ChevronRight size={17} /></button></div>
        <div className="reader-zoom"><button className="icon-button" aria-label="축소" title="축소 (Ctrl −)" disabled={!current || scale <= .1} onClick={() => changeZoom(stepZoom(scale * 100, -1))}><Minus size={15} /></button><select aria-label="원문 확대 비율" disabled={!current} value={zoom} onChange={e => changeZoom(e.target.value === 'fit' || e.target.value === 'page' ? e.target.value : Number(e.target.value))}><option value="fit">너비 맞춤</option><option value="page">한 쪽 맞춤</option>{[...new Set([10, 25, 50, 75, 100, 125, 150, 175, 200, 250, 300, 400, ...(typeof zoom === 'number' ? [zoom] : [])])].sort((a,b) => a-b).map(n => <option value={n} key={n}>{n}%</option>)}</select><button className="icon-button" aria-label="확대" title="확대 (Ctrl +)" disabled={!current || scale >= 4} onClick={() => changeZoom(stepZoom(scale * 100, 1))}><Plus size={15} /></button><output aria-label="현재 문서 배율">{current ? `${Math.round(scale * 100)}%` : '—'}</output></div>
        <div className="reader-search"><Search size={15} /><input ref={searchRef} aria-label="원문에서 검색" placeholder="원문에서 찾기" maxLength={200} value={query} onChange={e => setQuery(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') nextMatch(e.shiftKey ? -1 : 1); if (e.key === 'Escape') setQuery(''); }} />{query && <><span title={searchLimited ? '일부 내용은 검색에서 제외되었습니다' : undefined}>{indexedCount < pages.length ? '검색 중 · ' : ''}{matchingPages.length}쪽</span><button aria-label="이전 검색 결과" onClick={() => nextMatch(-1)} disabled={!matchingPages.length}><ChevronLeft size={14} /></button><button aria-label="다음 검색 결과" onClick={() => nextMatch(1)} disabled={!matchingPages.length}><ChevronRight size={14} /></button><button aria-label="원문 검색 지우기" onClick={() => setQuery('')}><X size={14} /></button></>}</div>
        <button className="icon-button fullscreen-button" aria-label="원문 전체 화면" title="전체 화면" onClick={() => void fullscreen()}><Maximize2 size={16} /></button>
        <button className="reader-tools-toggle icon-button" aria-label={toolsCollapsed ? '도구 모음 펼치기' : '도구 모음 접기'} aria-expanded={!toolsCollapsed} title={toolsCollapsed ? '도구 모음 펼치기 · 펼친 상태 유지' : '도구 모음 접기'} onClick={onToggleTools}>{toolsCollapsed ? <ChevronDown size={17} /> : <ChevronUp size={17} />}<span>{toolsCollapsed ? '도구' : '접기'}</span></button>
      </> : mode === 'tables' ? <><div className="figure-toolbar-label"><FileText size={16} /><strong>원문 표</strong><span>{tables.length}</span></div><select aria-label="원문 표 선택" value={currentTable?.id || ''} onChange={e => setTableId(e.target.value)}>{tables.map((t, i) => <option key={t.id} value={t.id}>표 {i + 1} · {t.page + 1}쪽 ({t.rows}행 × {t.columns}열)</option>)}</select><button className="text-button" onClick={onEditTable}>셀 편집 목록<ArrowRight size={14} /></button></> : <><div className="figure-toolbar-label"><Images size={16} /><strong>그래프·그림</strong><span>{filteredFigures.length}</span></div><select aria-label="그래프·그림 종류" value={figureFilter} onChange={e => setFigureFilter(e.target.value)}><option value="all">전체 종류</option><option value="chart">차트만</option><option value="drawing">그림·도형</option></select><select aria-label="그래프·그림 섹션" value={figureSection} onChange={e => setFigureSection(Number(e.target.value))}><option value={-1}>전체 섹션</option>{doc.sections.map(s => <option key={s.index} value={s.index}>섹션 {s.index + 1}</option>)}</select></>}
      {mode !== 'original' && <button className="icon-button fullscreen-button" aria-label="개체 전체 화면" title="전체 화면" onClick={() => void fullscreen()}><Maximize2 size={16} /></button>}
    </div>
    <div className="reader-version-bar"><span><FileText size={13} />{revision === 0 ? doc.conversion ? 'HWPX 변환 원문' : '업로드 원본' : `저장된 버전 ${doc.revision + 1}`}<span className="reader-readonly">읽기 전용 보기</span></span><div>{doc.revision > 0 && <label><input type="checkbox" checked={original} onChange={e => setOriginal(e.target.checked)} />{doc.conversion ? '변환 원문 보기' : '업로드 원본 보기'}</label>}<a href={sourceUrl} download>{doc.conversion ? 'HWP 원본 받기' : '원본 파일 받기'}<ArrowDownToLine size={13} /></a></div></div>
    {dirty > 0 && <div className="reader-draft-note">저장 전 변경 {dirty}건은 아직 원문 보기에 반영되지 않았습니다.<button onClick={onEdit}>편집으로 돌아가기<ArrowRight size={13} /></button></div>}
    <div tabIndex={0} className={`reader-surface ${mode !== 'original' ? 'figures-surface' : ''}`} ref={surfaceRef} onScroll={handleScroll} aria-label={mode === 'original' ? '연속 문서 스크롤' : '개체 미리보기'}>
      {loading ? <div className="reader-state" role="status"><LoaderCircle className="spin" size={24} /><h3>{loading}</h3><p>문서를 브라우저 안에서 읽고 있습니다.</p></div> : error ? <div className="reader-state" role="alert"><FileText size={30} /><h3>원문을 표시하지 못했습니다</h3><p>{error}</p><div><button className="outline-button" onClick={() => setRetry(retry + 1)}>다시 시도</button><button className="outline-button" onClick={onEdit}>텍스트 편집 보기</button><a href={sourceUrl} download className="outline-button">원본 파일 받기</a></div></div> : mode === 'original' && current ? <div className="reader-pages" style={{ width: Math.max(available.width, ...layout.map(entry => entry.width)) }}>
        {layout.map(entry => {
          const info = pages[entry.index];
          const near = entry.index >= range.start && entry.index <= range.end;
          return <section key={entry.index} className="reader-sheet" data-page={entry.index + 1} style={{ height: entry.height }} aria-label={`${entry.index + 1}쪽`}>
            <div className="reader-page-label">{entry.index + 1}쪽 · 섹션 {info.sectionIndex + 1}{query && matchingPages.includes(entry.index) && <span>검색어 강조 표시</span>}</div>
            <div ref={frame => { if (frame) framesRef.current.set(entry.index, frame); else framesRef.current.delete(entry.index); }} className="page-frame" style={{ width: entry.width, height: entry.paperHeight }}>
              {near && (rendered[entry.index] ? <PageArtwork svg={rendered[entry.index]} query={query} info={info} scale={entry.scale} /> : <div className="page-placeholder" role="status"><LoaderCircle size={20} className="spin" /><span>{entry.index + 1}쪽을 불러오는 중</span></div>)}
            </div>
          </section>;
        })}
      </div> : mode === 'tables' ? currentTable ? <section className="figure-detail"><div className="figure-detail-heading"><strong>표 {tables.indexOf(currentTable) + 1} · {currentTable.rows}행 × {currentTable.columns}열</strong><button className="text-button" onClick={() => showPage(currentTable.page)}>{currentTable.page + 1}쪽 원문 위치 보기<ArrowRight size={15} /></button></div><div className="figure-detail-image table-original-image"><FigureImage svg={rendered[currentTable.page]} figure={currentTable} label={`${currentTable.page + 1}쪽 원문 표`} /></div><div className="figure-detail-footer"><p>수식·서식·병합을 포함해 원문에 배치된 표 영역을 표시합니다.</p><button className="outline-button" disabled={!rendered[currentTable.page]} onClick={() => downloadFigure(currentTable)}><ArrowDownToLine size={15} />이미지 저장</button></div></section> : <div className="reader-state"><FileText size={30} /><h3>문서에 표가 없습니다</h3><p>표가 그림으로 삽입된 경우 그래프·그림 탭에서 확인할 수 있습니다.</p></div> : selectedFigure ? <section className="figure-detail"><div className="figure-detail-heading"><button className="text-button" onClick={() => setSelectedFigure(null)}><ArrowLeft size={15} />목록으로</button><strong>{selectedFigure.page + 1}쪽 · {figureKind(selectedFigure.type)}</strong><button className="text-button" onClick={() => showPage(selectedFigure.page)}>원문 위치 보기<ArrowRight size={15} /></button></div><div className="figure-detail-image"><FigureImage svg={rendered[selectedFigure.page]} figure={selectedFigure} label={`${selectedFigure.page + 1}쪽 ${figureKind(selectedFigure.type)} 확대`} /></div><div className="figure-detail-footer"><p>원문에 배치된 영역을 확대해 표시합니다.</p><button className="outline-button" disabled={!rendered[selectedFigure.page]} onClick={() => downloadFigure(selectedFigure)}><ArrowDownToLine size={15} />이미지 저장</button></div></section> : <>
        <div className="figure-gallery-heading"><h2>문서 속 그래프와 그림</h2><p>차트·삽입 그림·묶음 도형을 모았습니다. 항목을 누르면 확대하거나 원문 위치로 이동할 수 있습니다.</p></div>
        {filteredFigures.length ? <><div className="figure-grid">{filteredFigures.slice(0, visibleCount).map((figure, index) => <article className="figure-card" key={figure.id}><button className="figure-preview" onClick={() => setSelectedFigure(figure)} aria-label={`${figure.page + 1}쪽 ${figureKind(figure.type)} ${index + 1} 확대`}><FigureImage svg={rendered[figure.page]} figure={figure} label={`${figure.page + 1}쪽 ${figureKind(figure.type)}`} /></button><div><span><strong>{figureKind(figure.type)} {index + 1}</strong><small>{figure.page + 1}쪽 · 섹션 {figure.section + 1}</small></span><button className="icon-button" aria-label={`${index + 1}번째 개체 원문 위치`} title="원문 위치 보기" onClick={() => showPage(figure.page)}><ArrowRight size={16} /></button></div></article>)}</div>{filteredFigures.length > visibleCount && <button className="outline-button show-more-figures" onClick={() => setVisibleCount(visibleCount + 12)}>더 보기 ({Math.min(12, filteredFigures.length - visibleCount)}개)</button>}</> : <div className="reader-state"><Images size={32} /><h3>이 범위에는 그래프·그림이 없습니다</h3><p>표는 표 탭에서, 수식과 선은 원문 보기에서 확인하세요.</p></div>}
      </>}
    </div>
    <div className="reader-bottom"><span>{mode === 'original' ? `${pages.length ? `${page + 1} / ${pages.length}쪽 · ` : ''}스크롤하여 다음 쪽으로 이동 · 텍스트 선택 가능` : '원문 영역 미리보기'} </span><span title="원문 뷰어는 공개 렌더러를 사용합니다. 설치된 글꼴과 지원하지 않는 개체에 따라 한글 프로그램의 표시와 차이가 있을 수 있습니다.">일부 글꼴·배치는 한글과 다를 수 있습니다</span></div>
    {announcement && <div className="reader-announcement" role="status">{announcement}<button aria-label="이미지 저장 알림 닫기" onClick={() => setAnnouncement('')}><X size={14} /></button></div>}
  </div>;
}
