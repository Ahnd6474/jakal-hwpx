import React, { useEffect, useReducer, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { ArrowDownToLine, ArrowRight, ArrowUpFromLine, BookOpen, Check, ChevronRight, CircleCheck, Clock3, FileCheck2, FileText, HelpCircle, LayoutPanelLeft, LoaderCircle, LockKeyhole, PanelLeftClose, Plus, RotateCcw, RotateCw, Save, Search, ShieldCheck, Table2, Trash2, TriangleAlert, Type, X } from 'lucide-react';
import { api, post, waitForJob, type Document, type DocumentSummary, type Operation, type Paragraph, type Table } from './api';
import './styles.css';
import { AdvancedSettings } from './AdvancedSettings';
import { textTargets, type Proposal } from './editing';
import { BookOpenCheck, Images, Info, SlidersHorizontal } from 'lucide-react';
import { DocumentReader } from './DocumentReader';

type Drafts = Record<string, Operation>;
type History = { past: Drafts[]; present: Drafts; future: Drafts[] };
type HistoryAction = { type: 'change'; key: string; operation: Operation; original: string } | { type: 'batch'; changes: { key: string; operation: Operation; original: string }[] } | { type: 'undo' | 'redo' | 'reset' };
function draftReducer(state: History, action: HistoryAction): History {
  if (action.type === 'reset') return { past: [], present: {}, future: [] };
  if (action.type === 'undo') return state.past.length ? { past: state.past.slice(0, -1), present: state.past.at(-1)!, future: [state.present, ...state.future] } : state;
  if (action.type === 'redo') return state.future.length ? { past: [...state.past, state.present], present: state.future[0], future: state.future.slice(1) } : state;
  if (action.type === 'batch') {
    const present = { ...state.present };
    for (const change of action.changes) {
      if (change.operation.text === change.original) delete present[change.key];
      else present[change.key] = change.operation;
    }
    return { past: [...state.past.slice(-59), state.present], present, future: [] };
  }
  if (action.type !== 'change') return state;
  const present = { ...state.present };
  if (action.operation.text === action.original) delete present[action.key];
  else present[action.key] = action.operation;
  return { past: [...state.past.slice(-59), state.present], present, future: [] };
}

function App() {
  const [documents, setDocuments] = useState<DocumentSummary[]>([]);
  const [doc, setDoc] = useState<Document | null>(null);
  const [tab, setTab] = useState<'original' | 'paragraphs' | 'tables' | 'figures'>('original');
  const [section, setSection] = useState(-1);
  const [search, setSearch] = useState('');
  const [fileSearch, setFileSearch] = useState('');
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [dragging, setDragging] = useState(false);
  const [reportOpen, setReportOpen] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [advancedOpen, setAdvancedOpen] = useState(false);
  const [infoOpen, setInfoOpen] = useState(false);
  const [showEmpty, setShowEmpty] = useState(false);
  const [showReadonly, setShowReadonly] = useState(false);
  const [tableId, setTableId] = useState('');
  const [tableEditing, setTableEditing] = useState(false);
  const [selected, setSelected] = useState<{ label: string; text: string } | null>(null);
  const [history, dispatch] = useReducer(draftReducer, { past: [], present: {}, future: [] });
  const inputRef = useRef<HTMLInputElement>(null);
  const dirty = Object.keys(history.present).length;
  const disabled = Boolean(busy);
  const saved = doc?.validated === doc?.revision;

  async function refreshList() { setDocuments(await api<DocumentSummary[]>('/documents')); }
  async function load(id: string) {
    const value = await api<Document>(`/documents/${id}`);
    if (value.status !== 'ready') throw new Error(value.error || '문서 분석이 진행 중입니다. 잠시 후 다시 열어 주세요.');
    setDoc(value); setTab('original'); setTableId(''); setSelected(null); setSection(-1); setSearch(''); dispatch({ type: 'reset' });
    localStorage.setItem('jakal-document', id);
    setSidebarOpen(false);
    await refreshList();
  }
  async function task(label: string, action: () => Promise<void>) {
    if (disabled) return;
    setBusy(label); setError(''); setNotice('');
    try { await action(); }
    catch (e) { setError(e instanceof Error ? e.message : '요청을 처리하지 못했습니다.'); }
    finally { setBusy(''); }
  }
  useEffect(() => {
    void (async () => {
      try {
        const list = await api<DocumentSummary[]>('/documents');
        setDocuments(list);
        const previous = localStorage.getItem('jakal-document');
        if (previous && list.some(d => d.id === previous && d.status === 'ready')) await load(previous);
      } catch { setError('서버에 연결할 수 없습니다. 서버 실행 상태를 확인해 주세요.'); }
    })();
  }, []);
  useEffect(() => {
    function beforeUnload(event: BeforeUnloadEvent) { if (dirty || busy) { event.preventDefault(); event.returnValue = ''; } }
    window.addEventListener('beforeunload', beforeUnload);
    return () => window.removeEventListener('beforeunload', beforeUnload);
  }, [dirty, busy]);
  useEffect(() => {
    if (!notice) return;
    const timer = setTimeout(() => setNotice(''), 6500);
    return () => clearTimeout(timer);
  }, [notice]);
  useEffect(() => {
    if (!reportOpen && !helpOpen) return;
    const dialog = document.querySelector<HTMLElement>('[role="dialog"]');
    const previous = document.activeElement as HTMLElement | null;
    const oldOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    const controls = () => [...(dialog?.querySelectorAll<HTMLElement>('button:not([disabled]), a[href], input, select, textarea') || [])];
    controls()[0]?.focus();
    function keyboard(event: KeyboardEvent) {
      if (event.key === 'Escape') { setReportOpen(false); setHelpOpen(false); }
      if (event.key !== 'Tab') return;
      const items = controls();
      if (!items.length) return;
      if (event.shiftKey && document.activeElement === items[0]) { event.preventDefault(); items.at(-1)?.focus(); }
      else if (!event.shiftKey && document.activeElement === items.at(-1)) { event.preventDefault(); items[0].focus(); }
    }
    document.addEventListener('keydown', keyboard);
    return () => { document.removeEventListener('keydown', keyboard); document.body.style.overflow = oldOverflow; previous?.focus(); };
  }, [reportOpen, helpOpen]);
  function mayLeave() { return !dirty || window.confirm('저장하지 않은 수정 내용이 있습니다. 변경을 버리고 이동할까요?'); }
  async function upload(file?: File) {
    if (!file || !mayLeave()) return;
    if (!/\.(hwp|hwpx)$/i.test(file.name)) { setError('.hwp 또는 .hwpx 파일을 선택해 주세요.'); return; }
    if (file.size > 20 * 1024 * 1024) { setError('파일 크기는 20MB 이하여야 합니다.'); return; }
    await task('파일을 업로드하고 있습니다', async () => {
      const form = new FormData(); form.append('file', file);
      const result = await api<{ document_id: string; job_id: string }>('/documents', { method: 'POST', body: form });
      await refreshList();
      await waitForJob(result.job_id, job => setBusy(job.status === 'queued' ? '문서 분석을 기다리고 있습니다' : '문단과 표를 분석하고 있습니다'));
      await load(result.document_id); setTab('original');
      setNotice('문서를 열었습니다. 원본 파일은 별도로 보관됩니다.');
    });
    await refreshList().catch(() => {});
  }
  async function sample(format: 'hwp' | 'hwpx' = 'hwpx') {
    if (!mayLeave()) return;
    await task('샘플 문서를 준비하고 있습니다', async () => {
      const response = await fetch(`/sample.${format}`);
      if (!response.ok) throw new Error('샘플 문서를 불러오지 못했습니다.');
      const form = new FormData(); form.append('file', new File([await response.blob()], `워크스페이스_시작하기.${format}`));
      const result = await api<{ document_id: string; job_id: string }>('/documents', { method: 'POST', body: form });
      await waitForJob(result.job_id, () => setBusy('샘플 문서를 분석하고 있습니다'));
      await load(result.document_id); setTab('original'); setNotice('샘플을 자유롭게 수정해 보세요.');
    });
  }
  async function save() {
    if (!doc || !dirty) return;
    await task('수정 내용을 저장하고 있습니다', async () => {
      const result = await post<{ job_id: string }>(`/documents/${doc.id}/operations`, { revision: doc.revision, operations: Object.values(history.present) });
      await waitForJob(result.job_id, () => setBusy('수정 내용을 적용하고 파일을 검사하고 있습니다'));
      const fresh = await api<Document>(`/documents/${doc.id}`);
      setDoc(fresh); dispatch({ type: 'reset' }); await refreshList();
      setNotice(`버전 ${fresh.revision + 1} 저장 완료. 다운로드하려면 현재 버전을 검증하세요.`);
    });
  }
  async function validate() {
    if (!doc) return;
    await task('문서 검증을 시작하고 있습니다', async () => {
      const result = await post<{ job_id: string }>(`/documents/${doc.id}/validations`, { revision: doc.revision });
      await waitForJob(result.job_id, () => setBusy('문서 구조·참조·저장 후 다시 열기를 검사하고 있습니다'));
      setDoc(await api<Document>(`/documents/${doc.id}`)); setReportOpen(true); setNotice('현재 수정 버전의 검증이 완료되었습니다.');
    });
  }
  async function download() {
    if (!doc) return;
    await task('다운로드를 준비하고 있습니다', async () => {
      const result = await post<{ url: string }>(`/documents/${doc.id}/exports`, { revision: doc.revision });
      const link = document.createElement('a'); link.href = result.url; link.download = ''; link.click();
      setNotice('검증한 수정 버전의 다운로드를 시작했습니다.');
    });
  }
  async function deleteDocument(item: DocumentSummary) {
    if (!window.confirm(`“${item.name}”의 원본과 모든 수정본을 삭제할까요?`)) return;
    await task('문서를 삭제하고 있습니다', async () => {
      await api(`/documents/${item.id}`, { method: 'DELETE' });
      if (doc?.id === item.id) { setDoc(null); dispatch({ type: 'reset' }); localStorage.removeItem('jakal-document'); }
      await refreshList(); setNotice('문서와 저장된 파일을 삭제했습니다.');
    });
  }
  const visibleSections = doc?.sections.filter(s => section === -1 || s.index === section) || [];
  const paragraphCount = doc?.sections.reduce((n, s) => n + s.paragraphs.length, 0) || 0;
  const tableCount = doc?.sections.reduce((n, s) => n + s.tables.length, 0) || 0;
  const value = (key: string, original: string) => history.present[key]?.text ?? original;
  const visibleTables = visibleSections.flatMap(s => s.tables).filter(t => !search || t.cells.some(c => value(`${t.id}-${c.row}-${c.column}`, c.text).toLocaleLowerCase().includes(search.toLocaleLowerCase())));
  const activeTable = visibleTables.find(t => t.id === tableId) || visibleTables[0];
  const visibleParagraphs = visibleSections.flatMap(s => s.paragraphs).filter(p => (showEmpty || !!value(p.id, p.text).trim() || !!history.present[p.id]) && (showReadonly || p.editable) && (!search || value(p.id, p.text).toLocaleLowerCase().includes(search.toLocaleLowerCase())));
  function openHome() {
    if (!mayLeave()) return;
    setDoc(null); dispatch({ type: 'reset' }); setSidebarOpen(false);
    localStorage.removeItem('jakal-document');
  }
  const listedDocuments = documents.filter(item => item.name.toLocaleLowerCase().includes(fileSearch.toLocaleLowerCase()));
  const formatDate = (timestamp: number) => new Date(timestamp * 1000).toLocaleString('ko-KR', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false });
  function update(key: string, operation: Operation, original: string) { dispatch({ type: 'change', key, operation, original }); }
  function applyProposals(proposals: Proposal[]) {
    if (!doc) return;
    const originals = new Map(textTargets(doc, {}).map(t => [t.id, t.original]));
    dispatch({ type: 'batch', changes: proposals.map(p => ({ key: p.id, operation: p.operation, original: originals.get(p.id)! })) });
    setTab(proposals.every(p => p.operation.type === 'cell') ? 'tables' : 'paragraphs'); setTableEditing(true);
    setNotice(`${proposals.length}개 수정안을 편집에 추가했습니다. 수정 저장을 눌러 반영하세요.`);
  }
  function paragraph(p: Paragraph) {
    const text = value(p.id, p.text);
    if (!showEmpty && !text.trim() && !history.present[p.id]) return null;
    if (!showReadonly && !p.editable) return null;
    if (search && !text.toLocaleLowerCase().includes(search.toLocaleLowerCase())) return null;
    return <article className={`paragraph ${history.present[p.id] ? 'changed' : ''} ${!p.editable ? 'readonly' : ''}`} key={p.id}>
      <div className="paragraph-label"><span>{String(p.index + 1).padStart(2, '0')}</span>{history.present[p.id] && <i title="수정됨" />}</div>
      {p.editable ? <textarea aria-label={`${p.section + 1}섹션 ${p.index + 1}문단`} rows={Math.min(12, Math.max(1, text.split('\n').length, Math.ceil(text.length / 60)))} value={text} placeholder="빈 문단입니다. 내용을 입력하세요." maxLength={100000} disabled={disabled} onFocus={() => setSelected({ label: `${p.section + 1}섹션 · ${p.index + 1}문단`, text: '문단의 텍스트를 수정합니다. 기존 문서의 구조를 유지하며 저장합니다.' })} onChange={e => update(p.id, { type: 'paragraph', section: p.section, index: p.index, text: e.target.value }, p.text)} /> : <div className="readonly-content"><LockKeyhole size={14} /><span>{text || '복합 개체가 있는 문단입니다. 원문 보기에서 확인하세요.'}</span></div>}
    </article>;
  }
  function tableView(table: Table) {
    if (search && !table.cells.some(c => value(`${table.id}-${c.row}-${c.column}`, c.text).toLocaleLowerCase().includes(search.toLocaleLowerCase()))) return null;
    const cells = new Map(table.cells.map(c => [`${c.row}-${c.column}`, c]));
    return <article className="table-card" key={table.id}><div className="table-title"><div><Table2 size={17} /><strong>표 {table.index + 1}</strong></div><span>{table.rows}행 × {table.columns}열</span></div>
      <div className="table-scroll"><table><tbody>{Array.from({ length: table.rows }, (_, row) => <tr key={row}>{Array.from({ length: table.columns }, (_, column) => {
        const cell = cells.get(`${row}-${column}`); if (!cell) return null;
        const key = `${table.id}-${row}-${column}`;
        return <td key={column} rowSpan={cell.row_span} colSpan={cell.col_span} className={history.present[key] ? 'changed' : ''}>{cell.editable ? <textarea aria-label={`${table.section + 1}섹션 표 ${table.index + 1} ${row + 1}행 ${column + 1}열`} rows={3} maxLength={100000} value={value(key, cell.text)} disabled={disabled} onFocus={() => setSelected({ label: `표 ${table.index + 1} · ${row + 1}행 ${column + 1}열`, text: `셀의 텍스트를 수정합니다.${cell.row_span > 1 || cell.col_span > 1 ? ' 병합된 셀의 구조는 유지됩니다.' : ''}` })} onChange={e => update(key, { type: 'cell', section: table.section, table: table.index, row, column, text: e.target.value }, cell.text)} /> : <div className="locked-cell"><LockKeyhole size={13} />{cell.text || '복합 개체 · 읽기 전용'}</div>}</td>;
      })}</tr>)}</tbody></table></div></article>;
  }

  return <div className={`app ${doc ? 'editing' : ''}`} onDragOver={e => { if (!e.dataTransfer.types.includes('Files')) return; e.preventDefault(); if (!disabled) setDragging(true); }} onDragLeave={e => { if (!e.relatedTarget) setDragging(false); }} onDrop={e => { if (!e.dataTransfer.files.length) return; e.preventDefault(); setDragging(false); if (!disabled) void upload(e.dataTransfer.files[0]); }}>
    <input ref={inputRef} type="file" accept=".hwp,.hwpx" hidden aria-label="HWP 또는 HWPX 파일 선택" onChange={e => { void upload(e.target.files?.[0]); e.target.value = ''; }} />
    <aside className={`sidebar ${sidebarOpen ? 'open' : ''}`}>
      <button className="brand" disabled={disabled} onClick={openHome} aria-label="Jakal 내 문서"><span className="brand-mark">j</span><strong>jakal</strong><span className="brand-caption">문서 도구</span></button>
      <button className="new-button" disabled={disabled} onClick={() => inputRef.current?.click()}><Plus size={17} />파일 열기</button>
      <nav aria-label="문서 탐색"><button className={`nav-item ${!doc ? 'active' : ''}`} disabled={disabled} onClick={openHome}><LayoutPanelLeft size={17} />내 문서</button></nav>
      <div className="nav-heading recent-heading">최근에 연 문서<span>{documents.length}</span></div>
      <div className="document-list">{documents.length ? documents.map(item => <div key={item.id} className={`document-item ${doc?.id === item.id ? 'active' : ''}`}><button disabled={disabled} title={item.name} onClick={() => { if (mayLeave()) void task('문서를 여는 중', () => load(item.id)); }}><FileText size={16} /><span><b>{item.name}</b><small>{item.status === 'failed' ? '분석 실패' : item.status === 'processing' ? '분석 중' : `버전 ${item.revision + 1}`}</small></span></button><button className="delete-button" disabled={disabled} aria-label={`${item.name} 삭제`} onClick={() => void deleteDocument(item)}><Trash2 size={14} /></button></div>) : <p className="no-documents">최근 문서가 없습니다.</p>}</div>
      <div className="sidebar-bottom"><button onClick={() => setHelpOpen(true)}><HelpCircle size={17} />도움말</button><div className="session-note"><LockKeyhole size={13} /><span>이 브라우저에 연결된 문서<br />업로드 후 24시간 보관</span></div></div>
    </aside>
    {sidebarOpen && <button className="sidebar-scrim" aria-label="메뉴 닫기" onClick={() => setSidebarOpen(false)} />}
    <main className="main">
      <header className="topbar"><div className="breadcrumb"><button className="menu-toggle icon-button" aria-label="메뉴 열기" onClick={() => setSidebarOpen(true)}><PanelLeftClose size={19} /></button><button disabled={disabled} onClick={openHome}>내 문서</button>{doc && <><ChevronRight size={13} /><span title={doc.name}>{doc.name}</span></>}</div><button className="help-button" aria-label="사용 안내" onClick={() => setHelpOpen(true)}><BookOpen size={16} /><span>사용 안내</span></button></header>
      {error && <div className="message error" role="alert"><TriangleAlert size={17} /><span>{error}</span><button aria-label="오류 닫기" onClick={() => setError('')}><X size={17} /></button></div>}
      {notice && <div className="message success" role="status"><CircleCheck size={17} /><span>{notice}</span><button aria-label="알림 닫기" onClick={() => setNotice('')}><X size={17} /></button></div>}
      {busy && <div className="busy-bar" role="status"><LoaderCircle size={16} className="spin" />{busy}</div>}
      {!doc ? <div className="library">
        <section className="library-heading"><div><h1>내 문서</h1><p>HWP · HWPX 문서를 열어 문단과 표를 수정합니다.</p></div><button className="primary-button" disabled={disabled} onClick={() => inputRef.current?.click()}><Plus size={16} />파일 열기</button></section>
        <section className={`upload-zone ${dragging ? 'dragging' : ''}`} onDragLeave={() => setDragging(false)} aria-label="파일 업로드">
          <ArrowUpFromLine size={24} strokeWidth={1.4} /><div><h2>HWP 또는 HWPX 파일을 끌어 놓으세요</h2><p>.hwp, .hwpx · 최대 20MB</p></div><button className="outline-button" disabled={disabled} onClick={() => inputRef.current?.click()}>파일 선택</button>
        </section>
        <section className="library-files" aria-labelledby="recent-title"><div className="list-toolbar"><div><h2 id="recent-title">최근 문서</h2><span>{documents.length}</span></div><label className="file-search"><Search size={16} /><input aria-label="파일 이름 검색" placeholder="파일 이름 검색" value={fileSearch} onChange={e => setFileSearch(e.target.value)} /></label></div>
          <div className="file-table-scroll"><table className="file-table"><thead><tr><th>이름</th><th>업로드</th><th>상태</th><th><span className="sr-only">문서 삭제</span></th></tr></thead><tbody>{listedDocuments.map(item => <tr key={item.id}><td><button className="file-row-button" disabled={disabled} onClick={() => { if (mayLeave()) void task('문서를 여는 중', () => load(item.id)); }}><FileText size={20} strokeWidth={1.4} /><span><strong>{item.name}</strong><small>{item.name.toLowerCase().endsWith('.hwp') ? 'HWP' : 'HWPX'} 문서</small></span></button></td><td><time dateTime={new Date(item.created * 1000).toISOString()}>{formatDate(item.created)}</time></td><td><span className={`file-status ${item.status}`}><i />{item.status === 'failed' ? '분석 실패' : item.status === 'processing' ? '분석 중' : `버전 ${item.revision + 1}`}</span></td><td><button className="icon-button" aria-label={`${item.name} 목록에서 삭제`} disabled={disabled} onClick={() => void deleteDocument(item)}><Trash2 size={15} /></button></td></tr>)}</tbody></table></div>
          {!listedDocuments.length && <div className="library-empty"><FileText size={31} strokeWidth={1.1} /><h3>{fileSearch ? '일치하는 문서가 없습니다' : '아직 열어 본 문서가 없습니다'}</h3><p>{fileSearch ? '다른 파일 이름으로 검색하세요.' : '파일을 열면 이 목록에서 작업을 이어갈 수 있습니다.'}</p>{!fileSearch && <button className="text-button" disabled={disabled} onClick={() => void sample()}>예제 문서로 시작하기<ArrowRight size={15} /></button>}</div>}
        </section>
        <div className="library-footer"><span><Clock3 size={14} />문서는 업로드 후 24시간이 지나면 삭제됩니다.</span><div className="sample-links"><button className="text-button" disabled={disabled} onClick={() => void sample('hwp')}>HWP 예제<ArrowRight size={14} /></button><button className="text-button" disabled={disabled} onClick={() => void sample()}>HWPX 예제<ArrowRight size={14} /></button></div></div>
        <section className="library-guide"><h2>처음 사용하시나요?</h2><p>파일을 열면 원문이 먼저 표시됩니다.<br /><b>텍스트 편집</b>이나 <b>표</b> 탭에서 수정하고 저장하세요.<br /><b>검증하기</b>로 저장 결과를 확인하면 수정본을 다운로드할 수 있습니다.</p><button className="text-button" onClick={() => setHelpOpen(true)}>사용 안내<ArrowRight size={14} /></button></section>
      </div> : <>
        <section className="document-header"><div className="file-heading"><FileText size={24} strokeWidth={1.4} /><div><h1 title={doc.name}>{doc.name}</h1><p>{doc.format.toUpperCase()}<span>·</span>버전 {doc.revision + 1}<span>·</span><span className={dirty ? 'unsaved-label' : ''}>{dirty ? `${dirty}개 변경 · 저장 전` : '저장됨'}</span></p></div></div><div className="document-actions"><button className="outline-button" disabled={disabled || !!dirty} onClick={() => void validate()} title={dirty ? '수정 내용을 먼저 저장하세요' : '현재 버전 검사'}><ShieldCheck size={16} />검증하기</button><button className="primary-button" disabled={disabled || !!dirty || !saved} onClick={() => void download()} title={!saved ? '현재 버전을 먼저 검증하세요' : '검증한 파일 다운로드'}><ArrowDownToLine size={16} />다운로드</button></div></section>
        <div className="workspace-modebar"><div className="workspace-tabs" role="tablist" aria-label="문서 보기 방식">
          <button role="tab" aria-selected={tab === 'original'} className={tab === 'original' ? 'active' : ''} onClick={() => setTab('original')}><BookOpenCheck size={17} />원문 보기</button>
          <button role="tab" aria-selected={tab === 'paragraphs'} className={tab === 'paragraphs' ? 'active' : ''} onClick={() => setTab('paragraphs')}><Type size={16} />텍스트 편집</button>
          <button role="tab" aria-selected={tab === 'tables'} className={tab === 'tables' ? 'active' : ''} onClick={() => { setTab('tables'); setTableEditing(false); }}><Table2 size={16} />표<span>{tableCount}</span></button>
          <button role="tab" aria-selected={tab === 'figures'} className={tab === 'figures' ? 'active' : ''} onClick={() => setTab('figures')}><Images size={16} />그래프·그림</button>
        </div><div className="workspace-tools"><button className="icon-button" aria-label="문서 정보" title="문서 정보" aria-pressed={infoOpen} onClick={() => setInfoOpen(!infoOpen)}><Info size={17} /></button><button className="advanced-trigger" aria-label="고급 설정" title="고급 설정" disabled={disabled} onClick={() => setAdvancedOpen(true)}><SlidersHorizontal size={15} /><span>고급 설정</span></button></div></div>
        <div className={`editor-layout ${!infoOpen ? 'wide-layout' : ''}`}><section className="editor">
          <DocumentReader key={doc.id} doc={doc} active={tab === 'original' || tab === 'figures' || (tab === 'tables' && !tableEditing)} mode={tab === 'figures' ? 'figures' : tab === 'tables' ? 'tables' : 'original'} dirty={dirty} onMode={setTab} onEdit={() => setTab('paragraphs')} onEditTable={() => setTableEditing(true)} />
          {(tab === 'paragraphs' || (tab === 'tables' && tableEditing)) && <>
          <div className="edit-actions-bar"><div>{tab === 'paragraphs' ? <><label><input type="checkbox" checked={showEmpty} onChange={e => setShowEmpty(e.target.checked)} />빈 문단 표시</label><label><input type="checkbox" checked={showReadonly} onChange={e => setShowReadonly(e.target.checked)} />읽기 전용 문단 표시</label></> : <label className="table-picker-label"><button className="text-button" onClick={() => setTableEditing(false)}>원문 표 보기</button><select aria-label="표 선택" value={activeTable?.id || ''} onChange={e => setTableId(e.target.value)}>{visibleTables.map(t => <option key={t.id} value={t.id}>섹션 {t.section + 1} · 표 {t.index + 1} ({t.rows}행 × {t.columns}열)</option>)}</select></label>}</div><div className="history-actions"><button className="icon-button" aria-label="실행 취소" title="실행 취소" disabled={disabled || !history.past.length} onClick={() => dispatch({ type: 'undo' })}><RotateCcw size={16} /></button><button className="icon-button" aria-label="다시 실행" title="다시 실행" disabled={disabled || !history.future.length} onClick={() => dispatch({ type: 'redo' })}><RotateCw size={16} /></button><span /><button className="save-button" disabled={disabled || !dirty} onClick={() => void save()}><Save size={15} />수정 저장{dirty > 0 && <b>{dirty}</b>}</button></div></div>
          <div className="filter-bar"><label className="search-field"><Search size={15} /><input placeholder="문서에서 찾기" aria-label="문서 내용 검색" value={search} onChange={e => setSearch(e.target.value)} />{search && <button aria-label="검색 초기화" onClick={() => setSearch('')}><X size={14} /></button>}</label><select aria-label="섹션 선택" value={section} onChange={e => setSection(Number(e.target.value))}><option value={-1}>전체 섹션</option>{doc.sections.map(s => <option key={s.index} value={s.index}>섹션 {s.index + 1}</option>)}</select></div>
          {doc.format_note && <div className="format-note"><span className="format-badge">HWP 5</span><p>{doc.format_note}</p></div>}<div className="document-canvas"><div className="canvas-note"><span>내용 편집</span><span>인쇄 배치는 한글 프로그램에서 확인하세요.</span></div><div className={`document-paper ${tab === 'tables' ? 'table-paper' : ''}`}>
          {tab === 'paragraphs' ? visibleSections.map(s => <section className="document-section" key={s.index}>{s.paragraphs.some(p => visibleParagraphs.includes(p)) && <div className="section-label"><span>섹션 {s.index + 1}</span><span>텍스트 편집</span></div>}{s.paragraphs.map(paragraph)}</section>) : activeTable ? <><div className="section-label"><span>섹션 {activeTable.section + 1} · 표 {activeTable.index + 1}</span><span>병합 구조 유지</span></div>{tableView(activeTable)}</> : null}
          {(tab === 'paragraphs' ? !visibleParagraphs.length : !activeTable) && <div className="empty-state"><FileText size={27} /><p>{search ? '검색 결과가 없습니다.' : tab === 'tables' ? '이 범위에는 표가 없습니다.' : '이 범위에는 편집 가능한 텍스트가 없습니다.'}</p>{tab === 'paragraphs' && <button className="text-button" onClick={() => setTab('original')}>수식·그림은 원문에서 보기<ArrowRight size={14} /></button>}</div>}
          </div></div>
          <div className="editor-bottom"><span><span className={`status-dot ${dirty ? 'unsaved' : ''}`} />{dirty ? `${dirty}개 요소 변경 · 저장 전` : '저장됨'}</span><button onClick={() => setReportOpen(true)}>{saved && !dirty ? <><CircleCheck size={13} />현재 버전 검증 완료</> : '검사 결과'}<ChevronRight size={12} /></button></div>
          </>}
        </section>{infoOpen && <aside className="inspector" aria-label="문서 정보 패널"><div className="inspector-title"><span>문서 정보</span><button className="icon-button" aria-label="문서 정보 닫기" onClick={() => setInfoOpen(false)}><X size={16} /></button></div><div className="inspector-body"><h2>{doc.title}</h2><dl className="document-details"><div><dt>형식</dt><dd>{doc.format.toUpperCase()}</dd></div><div><dt>버전</dt><dd>{doc.revision + 1}</dd></div><div><dt>문단</dt><dd>{paragraphCount}</dd></div><div><dt>표</dt><dd>{tableCount}</dd></div><div><dt>섹션</dt><dd>{doc.sections.length}</dd></div></dl><hr /><h3 className="inspector-label">선택한 요소</h3>{selected ? <div className="selection-info"><h3>{selected.label}</h3><p>{selected.text}</p></div> : <p className="selection-placeholder">문단 또는 표의 셀을 선택하세요.</p>}<hr /><h3 className="inspector-label">검증</h3><div className="validation-state"><span className={`status-dot ${saved && !dirty ? '' : 'pending'}`} /><strong>{saved && !dirty ? '현재 버전 검증 완료' : dirty ? '저장 후 검증 필요' : '검증 필요'}</strong></div><p className="inspector-description">저장 결과를 검사한 뒤 다운로드할 수 있습니다.</p><button className="text-button" disabled={disabled || !!dirty} onClick={() => saved ? setReportOpen(true) : void validate()}>{saved ? '검사 결과 보기' : '검증하기'}<ChevronRight size={14} /></button><hr /><a className="text-download" href={`/api/documents/${doc.id}/text`} download><ArrowDownToLine size={14} />본문 텍스트 다운로드</a><div className="retention"><Clock3 size={14} /><p>{formatDate(doc.expires)} 자동 삭제</p></div></div></aside>}</div>
      </>}
    </main>
    {advancedOpen && doc && <AdvancedSettings doc={doc} drafts={history.present} onClose={() => setAdvancedOpen(false)} onApply={applyProposals} onSettings={setDoc} />}
    {dragging && <div className="drop-overlay" onDragLeave={() => setDragging(false)}><ArrowUpFromLine size={32} strokeWidth={1.4} /><h2>파일을 놓아 열기</h2><p>HWP · HWPX · 최대 20MB</p></div>}
    {reportOpen && doc && <div className="modal-backdrop" onClick={() => setReportOpen(false)}><section className="modal report-modal" role="dialog" aria-modal="true" aria-labelledby="report-title" onClick={e => e.stopPropagation()}><div className="modal-heading"><div><h2 id="report-title">검증 결과</h2><p>버전 {doc.revision + 1} · {saved ? '검증 완료' : '자동 검사 결과'}</p></div><button className="icon-button" aria-label="검증 결과 닫기" onClick={() => setReportOpen(false)}><X size={20} /></button></div><div className={`report-summary ${doc.report.issues.length ? 'warning' : ''}`}><FileCheck2 size={23} /><div><h3>{doc.report.issues.length ? '원본의 기존 오류가 남아 있습니다' : '새로 발생한 오류가 없습니다'}</h3><p>{saved ? '현재 버전을 다운로드할 수 있습니다.' : '다운로드하려면 검증하기를 실행하세요.'}</p></div></div><div className="check-list">{doc.report.checks.map(check => <div key={check}><CircleCheck size={16} /><span>{check}</span><b>완료</b></div>)}</div><p className="report-explanation">이 검사는 문서 구조를 확인합니다. 인쇄 배치는 한글 프로그램에서 확인하세요.</p>{doc.report.issues.length > 0 && <div className="issue-list"><h3>원본의 기존 오류 {doc.report.issues.length}개</h3>{doc.report.issues.map((issue, i) => <div key={i}><b>{issue.code || issue.kind}</b><p>{issue.message}</p><small>{issue.location}</small></div>)}</div>}{doc.report.repairs.length > 0 && <p className="repair-note">오래된 줄 배치 정보 {doc.report.repairs.length}건을 정리했습니다.</p>}<div className="modal-actions"><a className="outline-button" href={`/api/documents/${doc.id}/report`} download><ArrowDownToLine size={15} />검사 기록 JSON</a><button className="primary-button" onClick={() => setReportOpen(false)}>확인</button></div></section></div>}
    {helpOpen && <div className="modal-backdrop" onClick={() => setHelpOpen(false)}><section className="modal help-modal" role="dialog" aria-modal="true" aria-labelledby="help-title" onClick={e => e.stopPropagation()}><div className="modal-heading"><h2 id="help-title">사용 안내</h2><button className="icon-button" aria-label="사용 안내 닫기" onClick={() => setHelpOpen(false)}><X size={20} /></button></div><ol className="guide-steps"><li><b>파일 열기</b><p>HWP 5 또는 HWPX 파일을 선택하거나 화면으로 끌어 놓으세요. 최대 20MB까지 지원합니다.</p></li><li><b>내용 수정 및 저장</b><p>원문 보기에서 문서를 읽고, 텍스트 편집·표 탭에서 내용을 바꾼 뒤 수정 저장을 누르세요. 그래프·그림 탭에서는 삽입된 개체를 따로 볼 수 있습니다. 복합 개체가 있는 문단과 셀은 읽기 전용입니다.</p></li><li><b>검증 및 다운로드</b><p>현재 버전을 검증한 뒤 수정본을 다운로드하세요. 추가 수정 후에는 다시 검증해야 합니다.</p></li></ol><p className="help-note">문서는 이 브라우저 세션에서만 접근할 수 있으며 업로드 후 24시간이 지나면 삭제됩니다. 원본과 수정본은 별도로 보관합니다.</p><button className="primary-button" onClick={() => setHelpOpen(false)}>확인</button></section></div>}
  </div>;
}

createRoot(document.getElementById('root')!).render(<React.StrictMode><App /></React.StrictMode>);
