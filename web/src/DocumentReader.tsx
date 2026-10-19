import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ArrowDownToLine, ArrowLeft, ArrowRight, ChevronLeft, ChevronRight, FileText, Images, LoaderCircle, Maximize2, Minus, Plus, Search, X } from 'lucide-react';
import type { Document } from './api';
import type { Figure, PageInfo, VisualTable, ViewerResult } from './viewer-types';
import { cropPng, highlightSvg, safeSvg } from './viewer-svg';

type Props = { doc: Document; active: boolean; mode: 'original' | 'figures' | 'tables'; dirty: number; onMode: (mode: 'original' | 'figures') => void; onEdit: () => void; onEditTable: () => void };
const figureKind = (type: string) => type === 'table' ? '표' : type === 'chart' ? '차트' : type === 'image' ? '그림' : type === 'ole' ? '삽입 개체' : '도형';

function FigureImage({ svg, figure, label }: { svg?: string; figure: Figure; label: string }) {
  const [url, setUrl] = useState('');
  const [error, setError] = useState(false);
  useEffect(() => {
    let cancelled = false;
    let preview = '';
    setUrl(''); setError(false);
    if (svg) void cropPng(svg, figure).then(blob => { if (!cancelled) { preview = URL.createObjectURL(blob); setUrl(preview); } }).catch(() => { if (!cancelled) setError(true); });
    return () => { cancelled = true; if (preview) URL.revokeObjectURL(preview); };
  }, [svg, figure]);
  return url ? <img src={url} alt={label} draggable={false} /> : <span className="figure-loading">{error ? '원문 위치에서 확인해 주세요' : <><LoaderCircle size={19} className="spin" />영역을 불러오는 중</>}</span>;
}

export function DocumentReader({ doc, active, mode, dirty, onMode, onEdit, onEditTable }: Props) {
  const [original, setOriginal] = useState(false);
  const revision = original ? 0 : doc.revision;
  const [pages, setPages] = useState<PageInfo[]>([]);
  const [figures, setFigures] = useState<Figure[]>([]);
  const [tables, setTables] = useState<VisualTable[]>([]);
  const [tableId, setTableId] = useState('');
  const [rendered, setRendered] = useState<Record<number, string>>({});
  const [page, setPage] = useState(0);
  const [pageInput, setPageInput] = useState('1');
  const [zoom, setZoom] = useState<number | 'fit'>('fit');
  const [width, setWidth] = useState(800);
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
  const searchRef = useRef<HTMLInputElement>(null);
  const timers = useRef(new Map<number, ReturnType<typeof setTimeout>>());
  const [announcement, setAnnouncement] = useState('');
  const [indexedCount, setIndexedCount] = useState(0);
  const [searchLimited, setSearchLimited] = useState(false);

  useEffect(() => {
    const abort = new AbortController();
    const worker = new Worker(new URL('./document.worker.ts', import.meta.url), { type: 'module' });
    workerRef.current = worker;
    requested.current.clear();
    setIndexedCount(0); setSearchLimited(false);
    setPages([]); setFigures([]); setTables([]); setTableId(''); setRendered({}); setPage(0); setPageInput('1'); setQuery(''); setSelectedFigure(null); setError(''); setLoading('원문 보기를 준비하고 있습니다');
    const loadTimeout = setTimeout(() => { worker.terminate(); setLoading(''); setError('원문을 여는 데 시간이 오래 걸립니다. 다시 시도하거나 텍스트·표 보기로 이동해 주세요.'); }, 60_000);
    worker.onmessage = (event: MessageEvent<ViewerResult>) => {
      const message = event.data;
      if (message.type === 'progress') setLoading(message.message);
      else if (message.type === 'ready') { clearTimeout(loadTimeout); setPages(message.pages); setFigures(message.figures); setTables(message.tables); setLoading(''); }
      else if (message.type === 'page') {
        clearTimeout(timers.current.get(message.page)); timers.current.delete(message.page);
        try { const svg = safeSvg(message.svg); setRendered(previous => ({ ...previous, [message.page]: svg })); }
        catch { setError('이 페이지를 안전하게 표시하지 못했습니다. 원본 파일에서 확인해 주세요.'); }
      } else if (message.type === 'indexed') {
        setPages(previous => previous.map(p => p.pageIndex === message.page ? { ...p, text: message.text } : p));
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
  const matchingPages = useMemo(() => query.trim() ? pages.filter(p => p.text.toLocaleLowerCase().includes(query.replace(/\s/g, '').toLocaleLowerCase())).map(p => p.pageIndex) : [], [pages, query]);
  const filteredFigures = useMemo(() => figures.filter(f => (figureSection === -1 || f.section === figureSection) && (figureFilter === 'all' || (figureFilter === 'chart' ? f.type === 'chart' : f.type !== 'chart'))), [figures, figureFilter, figureSection]);
  const currentTable = tables.find(t => t.id === tableId) || tables[0];

  useEffect(() => {
    if (!pages.length || !active || error) return;
    if (mode === 'original') requestPage(page);
    else if (mode === 'tables') { if (currentTable) requestPage(currentTable.page); }
    else if (selectedFigure) requestPage(selectedFigure.page);
    else [...new Set(filteredFigures.slice(0, visibleCount).map(f => f.page))].forEach(requestPage);
  }, [pages, active, mode, page, currentTable, selectedFigure, filteredFigures, visibleCount, requestPage, error]);
  useEffect(() => { setPageInput(String(page + 1)); surfaceRef.current?.scrollTo({ top: 0, left: 0 }); }, [page]);
  useEffect(() => { surfaceRef.current?.scrollTo({ top: 0, left: 0 }); }, [mode, selectedFigure?.id, currentTable?.id]);
  useEffect(() => { setVisibleCount(12); setSelectedFigure(null); }, [figureFilter, figureSection]);
  useEffect(() => {
    const surface = surfaceRef.current;
    if (!surface) return;
    const observer = new ResizeObserver(() => setWidth(surface.clientWidth));
    observer.observe(surface);
    return () => observer.disconnect();
  }, []);
  useEffect(() => {
    if (!active) return;
    function keyboard(event: KeyboardEvent) {
      if (document.querySelector('[role="dialog"]')) return;
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f' && mode === 'original') { event.preventDefault(); searchRef.current?.focus(); }
      if (event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement || event.target instanceof HTMLSelectElement || event.ctrlKey || event.metaKey || event.altKey) return;
      if (mode !== 'original' || !pages.length) return;
      if (event.key === 'ArrowLeft') { event.preventDefault(); setPage(p => Math.max(0, p - 1)); }
      if (event.key === 'ArrowRight') { event.preventDefault(); setPage(p => Math.min(pages.length - 1, p + 1)); }
    }
    window.addEventListener('keydown', keyboard);
    return () => window.removeEventListener('keydown', keyboard);
  }, [active, mode, pages.length]);

  function navigate(value: string) {
    const next = Number(value);
    if (Number.isInteger(next) && next >= 1 && next <= pages.length) setPage(next - 1);
    else setPageInput(String(page + 1));
  }
  function nextMatch(direction: number) {
    if (!matchingPages.length) return;
    const next = direction > 0 ? matchingPages.find(n => n > page) ?? matchingPages[0] : [...matchingPages].reverse().find(n => n < page) ?? matchingPages.at(-1)!;
    setPage(next);
  }
  function showPage(index: number) { setSelectedFigure(null); setPage(index); onMode('original'); }
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
  const current = pages[page];
  const scale = current ? (zoom === 'fit' ? Math.max(.15, (width - (width < 600 ? 24 : 64)) / current.width) : zoom / 100) : 1;
  const sourceUrl = `/api/documents/${doc.id}/source?revision=0`;
  const displayedSvg = useMemo(() => rendered[page] ? highlightSvg(rendered[page], query) : '', [rendered, page, query]);

  return <div ref={readerRef} className="document-reader" hidden={!active}>
    <div className="reader-toolbar">
      {mode === 'original' ? <>
        <div className="page-navigation"><button className="icon-button" aria-label="이전 페이지" disabled={!pages.length || page === 0} onClick={() => setPage(p => Math.max(0, p - 1))}><ChevronLeft size={17} /></button><label><span className="sr-only">페이지 번호</span><input inputMode="numeric" aria-label="페이지 번호" value={pageInput} disabled={!pages.length} onChange={e => setPageInput(e.target.value)} onBlur={e => navigate(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') navigate(pageInput); }} /></label><span>/ {pages.length || '—'}</span><button className="icon-button" aria-label="다음 페이지" disabled={!pages.length || page >= pages.length - 1} onClick={() => setPage(p => Math.min(pages.length - 1, p + 1))}><ChevronRight size={17} /></button></div>
        <div className="reader-zoom"><button className="icon-button" aria-label="축소" disabled={!current} onClick={() => setZoom(Math.min(300, Math.max(25, Math.round((zoom === 'fit' ? scale * 100 : zoom) / 25) * 25 - 25)))}><Minus size={15} /></button><select aria-label="원문 확대 비율" value={zoom} onChange={e => setZoom(e.target.value === 'fit' ? 'fit' : Number(e.target.value))}><option value="fit">너비 맞춤</option>{[25, 50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300].map(n => <option value={n} key={n}>{n}%</option>)}</select><button className="icon-button" aria-label="확대" disabled={!current} onClick={() => setZoom(Math.min(300, Math.round((zoom === 'fit' ? scale * 100 : zoom) / 25) * 25 + 25))}><Plus size={15} /></button></div>
        <div className="reader-search"><Search size={15} /><input ref={searchRef} aria-label="원문에서 검색" placeholder="원문에서 찾기" maxLength={200} value={query} onChange={e => setQuery(e.target.value)} onKeyDown={e => { if (e.key === 'Enter') nextMatch(e.shiftKey ? -1 : 1); if (e.key === 'Escape') setQuery(''); }} />{query && <><span>{indexedCount < pages.length ? '검색 중 · ' : ''}{matchingPages.length}쪽</span><button aria-label="이전 검색 결과" onClick={() => nextMatch(-1)} disabled={!matchingPages.length}><ChevronLeft size={14} /></button><button aria-label="다음 검색 결과" onClick={() => nextMatch(1)} disabled={!matchingPages.length}><ChevronRight size={14} /></button><button aria-label="원문 검색 지우기" onClick={() => setQuery('')}><X size={14} /></button></>}</div>
        <button className="icon-button fullscreen-button" aria-label="원문 전체 화면" title="전체 화면" onClick={() => void fullscreen()}><Maximize2 size={16} /></button>
      </> : mode === 'tables' ? <><div className="figure-toolbar-label"><FileText size={16} /><strong>원문 표</strong><span>{tables.length}</span></div><select aria-label="원문 표 선택" value={currentTable?.id || ''} onChange={e => setTableId(e.target.value)}>{tables.map((t, i) => <option key={t.id} value={t.id}>표 {i + 1} · {t.page + 1}쪽 ({t.rows}행 × {t.columns}열)</option>)}</select><button className="text-button" onClick={onEditTable}>셀 편집 목록<ArrowRight size={14} /></button></> : <><div className="figure-toolbar-label"><Images size={16} /><strong>그래프·그림</strong><span>{filteredFigures.length}</span></div><select aria-label="그래프·그림 종류" value={figureFilter} onChange={e => setFigureFilter(e.target.value)}><option value="all">전체 종류</option><option value="chart">차트만</option><option value="drawing">그림·도형</option></select><select aria-label="그래프·그림 섹션" value={figureSection} onChange={e => setFigureSection(Number(e.target.value))}><option value={-1}>전체 섹션</option>{doc.sections.map(s => <option key={s.index} value={s.index}>섹션 {s.index + 1}</option>)}</select></>}
      {mode !== 'original' && <button className="icon-button fullscreen-button" aria-label="개체 전체 화면" title="전체 화면" onClick={() => void fullscreen()}><Maximize2 size={16} /></button>}
    </div>
    <div className="reader-version-bar"><span><FileText size={13} />{revision === 0 ? '업로드 원본' : `저장된 버전 ${doc.revision + 1}`}<span className="reader-readonly">읽기 전용 보기</span></span><div>{doc.revision > 0 && <label><input type="checkbox" checked={original} onChange={e => setOriginal(e.target.checked)} />업로드 원본 보기</label>}<a href={sourceUrl} download>원본 파일 받기<ArrowDownToLine size={13} /></a></div></div>
    {dirty > 0 && <div className="reader-draft-note">저장 전 변경 {dirty}건은 아직 원문 보기에 반영되지 않았습니다.<button onClick={onEdit}>편집으로 돌아가기<ArrowRight size={13} /></button></div>}
    <div tabIndex={0} className={`reader-surface ${mode !== 'original' ? 'figures-surface' : ''}`} ref={surfaceRef}>
      {loading ? <div className="reader-state" role="status"><LoaderCircle className="spin" size={24} /><h3>{loading}</h3><p>문서를 브라우저 안에서 읽고 있습니다.</p></div> : error ? <div className="reader-state" role="alert"><FileText size={30} /><h3>원문을 표시하지 못했습니다</h3><p>{error}</p><div><button className="outline-button" onClick={() => setRetry(retry + 1)}>다시 시도</button><button className="outline-button" onClick={onEdit}>텍스트 편집 보기</button><a href={sourceUrl} download className="outline-button">원본 파일 받기</a></div></div> : mode === 'original' && current ? <>
        <div className="reader-page-label">{page + 1}쪽 · 섹션 {current.sectionIndex + 1}{query && <span>{matchingPages.includes(page) ? '검색어가 강조 표시되어 있습니다' : indexedCount < pages.length ? '검색을 준비하고 있습니다' : searchLimited ? '일부 내용은 검색에서 제외되었습니다' : '이 페이지에는 검색어가 없습니다'}</span>}</div>
        <div className="page-frame" style={{ width: current.width * scale, height: current.height * scale }}>
          {rendered[page] ? <div className="page-svg" role="document" aria-label={`${page + 1}쪽 원문`} style={{ width: current.width, height: current.height, transform: `scale(${scale})`, transformOrigin: 'top left' }} dangerouslySetInnerHTML={{ __html: displayedSvg }} /> : <div className="reader-state" role="status"><LoaderCircle size={22} className="spin" /><p>페이지를 표시하고 있습니다.</p></div>}
        </div>
      </> : mode === 'tables' ? currentTable ? <section className="figure-detail"><div className="figure-detail-heading"><strong>표 {tables.indexOf(currentTable) + 1} · {currentTable.rows}행 × {currentTable.columns}열</strong><button className="text-button" onClick={() => showPage(currentTable.page)}>{currentTable.page + 1}쪽 원문 위치 보기<ArrowRight size={15} /></button></div><div className="figure-detail-image table-original-image"><FigureImage svg={rendered[currentTable.page]} figure={currentTable} label={`${currentTable.page + 1}쪽 원문 표`} /></div><div className="figure-detail-footer"><p>수식·서식·병합을 포함해 원문에 배치된 표 영역을 표시합니다.</p><button className="outline-button" disabled={!rendered[currentTable.page]} onClick={() => downloadFigure(currentTable)}><ArrowDownToLine size={15} />이미지 저장</button></div></section> : <div className="reader-state"><FileText size={30} /><h3>문서에 표가 없습니다</h3><p>표가 그림으로 삽입된 경우 그래프·그림 탭에서 확인할 수 있습니다.</p></div> : selectedFigure ? <section className="figure-detail"><div className="figure-detail-heading"><button className="text-button" onClick={() => setSelectedFigure(null)}><ArrowLeft size={15} />목록으로</button><strong>{selectedFigure.page + 1}쪽 · {figureKind(selectedFigure.type)}</strong><button className="text-button" onClick={() => showPage(selectedFigure.page)}>원문 위치 보기<ArrowRight size={15} /></button></div><div className="figure-detail-image"><FigureImage svg={rendered[selectedFigure.page]} figure={selectedFigure} label={`${selectedFigure.page + 1}쪽 ${figureKind(selectedFigure.type)} 확대`} /></div><div className="figure-detail-footer"><p>원문에 배치된 영역을 확대해 표시합니다.</p><button className="outline-button" disabled={!rendered[selectedFigure.page]} onClick={() => downloadFigure(selectedFigure)}><ArrowDownToLine size={15} />이미지 저장</button></div></section> : <>
        <div className="figure-gallery-heading"><h2>문서 속 그래프와 그림</h2><p>차트·삽입 그림·묶음 도형을 모았습니다. 항목을 누르면 확대하거나 원문 위치로 이동할 수 있습니다.</p></div>
        {filteredFigures.length ? <><div className="figure-grid">{filteredFigures.slice(0, visibleCount).map((figure, index) => <article className="figure-card" key={figure.id}><button className="figure-preview" onClick={() => setSelectedFigure(figure)} aria-label={`${figure.page + 1}쪽 ${figureKind(figure.type)} ${index + 1} 확대`}><FigureImage svg={rendered[figure.page]} figure={figure} label={`${figure.page + 1}쪽 ${figureKind(figure.type)}`} /></button><div><span><strong>{figureKind(figure.type)} {index + 1}</strong><small>{figure.page + 1}쪽 · 섹션 {figure.section + 1}</small></span><button className="icon-button" aria-label={`${index + 1}번째 개체 원문 위치`} title="원문 위치 보기" onClick={() => showPage(figure.page)}><ArrowRight size={16} /></button></div></article>)}</div>{filteredFigures.length > visibleCount && <button className="outline-button show-more-figures" onClick={() => setVisibleCount(visibleCount + 12)}>더 보기 ({Math.min(12, filteredFigures.length - visibleCount)}개)</button>}</> : <div className="reader-state"><Images size={32} /><h3>이 범위에는 그래프·그림이 없습니다</h3><p>표는 표 탭에서, 수식과 선은 원문 보기에서 확인하세요.</p></div>}
      </>}
    </div>
    <div className="reader-bottom"><span>{mode === 'original' ? `${pages.length ? `${page + 1} / ${pages.length}쪽 · ` : ''}텍스트를 드래그해 선택할 수 있습니다` : '원문 영역 미리보기'} </span><span title="원문 뷰어는 공개 렌더러를 사용합니다. 설치된 글꼴과 지원하지 않는 개체에 따라 한글 프로그램의 표시와 차이가 있을 수 있습니다.">일부 글꼴·배치는 한글과 다를 수 있습니다</span></div>
    {announcement && <div className="reader-announcement" role="status">{announcement}<button aria-label="이미지 저장 알림 닫기" onClick={() => setAnnouncement('')}><X size={14} /></button></div>}
  </div>;
}
