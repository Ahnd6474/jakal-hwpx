import React, { useCallback, useEffect, useMemo, useReducer, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { ArrowDownToLine, ArrowRight, ArrowUpFromLine, BookOpen, Check, ChevronDown, ChevronRight, ChevronUp, CircleCheck, Clock3, FileCheck2, FileText, HelpCircle, LayoutPanelLeft, LoaderCircle, LockKeyhole, PanelLeftClose, Plus, RotateCcw, RotateCw, Save, Search, ShieldCheck, Table2, Trash2, TriangleAlert, Type, X } from 'lucide-react';
import { api, post, waitForJob, type Document, type DocumentSummary, type Operation, type Paragraph, type Table } from './api';
import './styles.css';
import { textTargets, type Proposal } from './editing';
import { BookOpenCheck, Images, Info, SlidersHorizontal } from 'lucide-react';
import { PageEditor, type PageEditorHandle } from './PageEditor';
import { WorkspaceMenu } from './WorkspaceMenu';
import { TextFlowEditor, type TextFlowHandle } from './TextFlowEditor';
import { draftReducer } from './draft-history';
import { lastDocument } from './preferences';
import { backupText, clearDraftBackup, readDraftBackup, writeDraftBackup, type DraftBackup } from './draft-backup';
import { useDraftBackup } from './useDraftBackup';
import { prepareDownload } from './download-workflow';
import type { TextTarget } from './editing';
import type { TemplateInput } from './NewDocument';
import { ToolBoundary } from './ToolBoundary';
import { ApplicationBoundary } from './ApplicationBoundary';
import { ConnectionStatus } from './ConnectionStatus';
import { trapModalFocus } from './modal-focus';
import { requestBytes } from './request';
import { documentCommands } from './document-commands';
import { DocumentSyncNotice } from './DocumentSyncNotice';
import { studioBackup } from './studio-backup';
import type { LayoutApplication } from './PageLayoutTemplates';
import { PrivacyGate } from './PrivacyGate';
const PrivacyPanel = React.lazy(() => import('./PrivacyPanel').then(module => ({ default: module.PrivacyPanel })));
const NewDocument = React.lazy(() => import('./NewDocument').then(module => ({ default: module.NewDocument })));
const AdvancedSettings = React.lazy(() => import('./AdvancedSettings').then(module => ({ default: module.AdvancedSettings })));
const DocumentReader = React.lazy(() => import('./DocumentReader').then(module => ({ default: module.DocumentReader })));
const VersionHistory = React.lazy(() => import('./VersionHistory').then(module => ({ default: module.VersionHistory })));
const ChangeReview = React.lazy(() => import('./ChangeReview').then(module => ({ default: module.ChangeReview })));
const DocumentActions = React.lazy(() => import('./DocumentActions').then(module => ({ default: module.DocumentActions })));
const BackupShelf = React.lazy(() => import('./BackupShelf').then(module => ({ default: module.BackupShelf })));
const PageLayoutTemplates = React.lazy(() => import('./PageLayoutTemplates').then(module => ({ default: module.PageLayoutTemplates })));

function App() {
  const [documents, setDocuments] = useState<DocumentSummary[]>([]);
  const [booting, setBooting] = useState(true);
  const [doc, setDoc] = useState<Document | null>(null);
  const [tab, setTab] = useState<'studio' | 'original' | 'paragraphs' | 'tables' | 'figures'>('studio');
  const [studioDirty, setStudioDirty] = useState(false);
  const [readerOpened, setReaderOpened] = useState(false);
  const pageEditor = useRef<PageEditorHandle>(null);
  const [editorEpoch, setEditorEpoch] = useState(0);
  const [section, setSection] = useState(-1);
  const [search, setSearch] = useState('');
  const [fileSearch, setFileSearch] = useState('');
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [dragging, setDragging] = useState(false);
  const [reportOpen, setReportOpen] = useState(false);
  const [helpOpen, setHelpOpen] = useState(false);
  const [privacyOpen, setPrivacyOpen] = useState(false);
  const closePrivacy = useCallback(() => setPrivacyOpen(false), []);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [advancedOpen, setAdvancedOpen] = useState(false);
  const [infoOpen, setInfoOpen] = useState(false);
  const [reviewOpen, setReviewOpen] = useState(false);
  const [versionsOpen, setVersionsOpen] = useState(false);
  const [newDocumentOpen, setNewDocumentOpen] = useState(false);
  const [pageLayoutOpen, setPageLayoutOpen] = useState(false);
  const closePageLayout = useCallback(() => setPageLayoutOpen(false), []);
  const [fileAction, setFileAction] = useState<'rename' | 'copy' | null>(null);
  const closeFileAction = useCallback(() => setFileAction(null), []);
  const [backupShelf, setBackupShelf] = useState(false);
  const closeBackupShelf = useCallback(() => setBackupShelf(false), []);
  const closeNewDocument = useCallback(() => setNewDocumentOpen(false), []);
  const closeVersions = useCallback(() => setVersionsOpen(false), []);
  const closeReview = useCallback(() => setReviewOpen(false), []);
  const [staleBackup, setStaleBackup] = useState<DraftBackup | null>(null);
  const [confirmDropBackup, setConfirmDropBackup] = useState(false);
  const [focusTarget, setFocusTarget] = useState<string | null>(null);
  const taskRunning = useRef(false);
  const reloadWithBackup = useRef(false);
  const [toolsCollapsed, setToolsCollapsed] = useState(false);
  const [toolsPinned, setToolsPinned] = useState(false);
  useEffect(() => { setToolsCollapsed(window.matchMedia('(max-width: 760px)').matches); setToolsPinned(false); }, [doc?.id]);
  function toggleReaderTools() {
    setToolsCollapsed(!toolsCollapsed);
    setToolsPinned(toolsCollapsed);
  }
  const [textView, setTextView] = useState<'flow' | 'list'>('flow');
  const flowEditor = useRef<TextFlowHandle>(null);
  const continuousText = tab === 'paragraphs' && textView === 'flow';
  const textSearch = useRef<HTMLInputElement>(null);
  useEffect(() => { if (continuousText && window.matchMedia('(max-width: 760px)').matches) setToolsCollapsed(true); }, [continuousText]);
  useEffect(() => {
    if (!continuousText) return;
    function keyboard(event: KeyboardEvent) {
      if (document.querySelector('[role="dialog"]')) return;
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') {
        event.preventDefault(); setToolsCollapsed(false); setToolsPinned(true);
        requestAnimationFrame(() => textSearch.current?.focus());
      }
    }
    window.addEventListener('keydown', keyboard);
    return () => window.removeEventListener('keydown', keyboard);
  }, [continuousText]);
  const [showEmpty, setShowEmpty] = useState(false);
  const [showReadonly, setShowReadonly] = useState(false);
  const [tableId, setTableId] = useState('');
  const [tableEditing, setTableEditing] = useState(false);
  const [selected, setSelected] = useState<{ label: string; text: string } | null>(null);
  const [history, dispatch] = useReducer(draftReducer, { past: [], present: {}, future: [] });
  const backupStatus = useDraftBackup(doc, history.present, !!staleBackup);
  const inputRef = useRef<HTMLInputElement>(null);
  const dirty = Object.keys(history.present).length + (studioDirty ? 1 : 0);
  const disabled = Boolean(busy);
  useEffect(() => { document.title = `${dirty ? '● ' : ''}${doc ? doc.name : '내 문서'} · Jakal`; }, [doc?.name, dirty]);
  const saved = doc?.validated === doc?.revision;

  async function refreshList() { setDocuments(await api<DocumentSummary[]>('/documents')); }
  async function load(id: string, keepDrafts = false) {
    const value = await api<Document>(`/documents/${id}`);
    if (value.status !== 'ready') throw new Error(value.error || '문서 분석이 진행 중입니다. 잠시 후 다시 열어 주세요.');
    if (doc && dirty && !staleBackup && !keepDrafts) clearDraftBackup(doc.id);
    const recovery = readDraftBackup(value);
    setDoc(value); setStudioDirty(false); setReaderOpened(false); setTab('studio'); setTableId(''); setSelected(null); setSection(-1); setSearch(''); dispatch({ type: 'reset' });
    setEditorEpoch(n => n + 1);
    setReviewOpen(false); setStaleBackup(recovery.kind === 'conflict' ? recovery.backup : null); setConfirmDropBackup(false);
    if (recovery.kind === 'restore' && Object.keys(recovery.drafts).length) {
      dispatch({ type: 'batch', changes: textTargets(value, recovery.drafts).filter(item => !!recovery.drafts[item.id]).map(item => ({ key: item.id, original: item.original, operation: recovery.drafts[item.id] })) });
      setTab(Object.values(recovery.drafts).every(operation => operation.type === 'cell') ? 'tables' : 'paragraphs');
      setTableEditing(true);
      setNotice(`미저장 편집 ${Object.keys(recovery.drafts).length}개를 이 탭의 임시본에서 복원했습니다.`);
    }
    lastDocument.write(id);
    setSidebarOpen(false);
    await refreshList();
  }
  async function task(label: string, action: () => Promise<void>) {
    if (taskRunning.current) return;
    taskRunning.current = true;
    setBusy(label); setError(''); setNotice('');
    try { await action(); }
    catch (e) { setError(e instanceof Error ? e.message : '요청을 처리하지 못했습니다.'); }
    finally { taskRunning.current = false; setBusy(''); }
  }
  useEffect(() => {
    void (async () => {
      try {
        const list = await api<DocumentSummary[]>('/documents');
        setDocuments(list);
        const previous = lastDocument.read();
        if (previous && list.some(d => d.id === previous && d.status === 'ready')) await load(previous);
      } catch (e) { setError(e instanceof Error ? e.message : '서버에 연결할 수 없습니다. 서버 실행 상태를 확인해 주세요.'); }
      finally { setBooting(false); }
    })();
  }, []);
  useEffect(() => {
    function beforeUnload(event: BeforeUnloadEvent) { if ((dirty || busy) && !reloadWithBackup.current) { event.preventDefault(); event.returnValue = ''; } }
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
    const dialog = document.getElementById(helpOpen ? 'help-title' : 'report-title')?.closest<HTMLElement>('[role="dialog"]');
    if (dialog) return trapModalFocus(dialog, () => { setReportOpen(false); setHelpOpen(false); });
  }, [reportOpen, helpOpen]);
  function mayLeave() {
    if (studioDirty) { setError('원문 편집 내용을 먼저 수정 저장한 뒤 다른 문서를 열어 주세요.'); return false; }
    return !dirty || window.confirm('저장하지 않은 수정 내용이 있습니다. 변경을 버리고 이동할까요?');
  }
  function changeView(next: typeof tab) {
    if (next === tab) return;
    if (studioDirty || (next === 'studio' && dirty)) { setError('수정 내용을 저장한 뒤 보기 방식을 바꿔 주세요.'); return; }
    if (next === 'original' || next === 'tables' || next === 'figures') setReaderOpened(true);
    setTab(next);
  }
  async function upload(file?: File) {
    if (!file) return;
    if (!file.size) { setError('빈 파일은 업로드할 수 없습니다.'); return; }
    if (!/\.(hwp|hwpx)$/i.test(file.name)) { setError('.hwp 또는 .hwpx 파일을 선택해 주세요.'); return; }
    if (file.size > 20 * 1024 * 1024) { setError('파일 크기는 20MB 이하여야 합니다.'); return; }
    if (!mayLeave()) return;
    await task('파일을 업로드하고 있습니다', async () => {
      const form = new FormData(); form.append('file', file);
      const result = await api<{ document_id: string; job_id: string }>('/documents', { method: 'POST', body: form });
      await refreshList();
      await waitForJob(result.job_id, job => setBusy(job.status === 'queued' ? '문서 분석을 기다리고 있습니다' : /\.hwp$/i.test(file.name) ? 'HWP를 HWPX로 변환하고 문서를 분석하고 있습니다' : '문단과 표를 분석하고 있습니다'));
      await load(result.document_id); setTab('studio');
      setNotice(/\.hwp$/i.test(file.name) ? 'HWP를 HWPX로 변환해 열었습니다. 수정·검증 후 HWPX로 다운로드할 수 있습니다.' : '문서를 열었습니다. 원본 파일은 별도로 보관됩니다.');
    });
    await refreshList().catch(() => {});
  }
  async function sample(format: 'hwp' | 'hwpx' = 'hwpx') {
    if (!mayLeave()) return;
    await task('샘플 문서를 준비하고 있습니다', async () => {
      const bytes = await requestBytes(`/sample.${format}`);
      const form = new FormData(); form.append('file', new File([bytes], `워크스페이스_시작하기.${format}`));
      const result = await api<{ document_id: string; job_id: string }>('/documents', { method: 'POST', body: form });
      await waitForJob(result.job_id, () => setBusy('샘플 문서를 분석하고 있습니다'));
      await load(result.document_id); setTab('studio'); setNotice('샘플을 자유롭게 수정해 보세요.');
    });
  }
  async function createDocument(value: TemplateInput) {
    if (!mayLeave()) return;
    await task('새 문서를 만들고 있습니다', async () => {
      const result = await post<{ document_id: string; job_id: string }>('/documents/from-template', value);
      await waitForJob(result.job_id, () => setBusy('문서 양식을 준비하고 검사하고 있습니다'));
      await load(result.document_id); setNewDocumentOpen(false);
      setNotice('새 문서를 만들었습니다. 문서 위에서 내용을 입력하거나 작업 도구를 사용하세요.');
    });
  }
  async function saveRevision(current: Document, operations: Operation[]) {
      setBusy('수정 내용을 저장하고 있습니다');
      const fresh = await documentCommands.execute(current, { kind: 'edit', operations }, {
        onStage: () => setBusy('수정 내용을 적용하고 파일을 검사하고 있습니다'),
        onReconnect: () => setBusy('연결을 다시 확인하고 기존 저장 작업을 찾고 있습니다'),
      });
      if (!staleBackup) clearDraftBackup(current.id);
      setDoc(fresh); setEditorEpoch(n => n + 1); dispatch({ type: 'reset' }); await refreshList();
      return fresh;
  }
  async function applyPageLayout(value: LayoutApplication) {
    if (taskRunning.current) throw new Error('진행 중인 작업이 끝난 뒤 다시 시도해 주세요.');
    taskRunning.current = true; setBusy('쪽 양식을 준비하고 있습니다'); setError('');
    try {
      let source = doc;
      if (source && dirty) source = studioDirty ? await saveStudioRevision(source) : await saveRevision(source, Object.values(history.present));
      let fresh: Document;
      if (value.create) {
        const created = await post<{ document_id: string; job_id: string }>('/documents/from-page-layout', { title: value.title, layout: value.layout, ...(value.background_pages !== undefined ? { background_pages: value.background_pages } : {}) });
        await waitForJob(created.job_id, () => setBusy('쪽 양식으로 빈 문서를 준비하고 검사하고 있습니다'));
        fresh = await api<Document>(`/documents/${created.document_id}`);
      } else {
        if (!source) throw new Error('쪽 양식을 적용할 문서를 먼저 열어 주세요.');
        fresh = await documentCommands.execute(source, { kind: 'layout', section: value.section, layout: value.layout, ...(value.background_pages !== undefined ? { background_pages: value.background_pages } : {}) }, {
          onStage: () => setBusy('쪽 양식을 적용하고 본문·표·그림 보존을 검사하고 있습니다'),
          onReconnect: () => setBusy('연결을 다시 확인하고 기존 쪽 양식 작업을 찾고 있습니다'),
        });
      }
      await load(fresh.id);
      setNotice('쪽 양식을 적용했습니다. 단 나누기와 머리말·꼬리말을 확인하세요. 이전 버전은 기록에 보관했습니다.');
    } finally { taskRunning.current = false; setBusy(''); }
  }
  async function checkRevision(current: Document) {
      setBusy('문서 검증을 시작하고 있습니다');
      const fresh = await documentCommands.execute(current, { kind: 'validate' }, {
        onStage: () => setBusy('문서 구조·참조·저장 후 다시 열기를 검사하고 있습니다'),
        onReconnect: () => setBusy('연결을 다시 확인하고 기존 검증 작업을 찾고 있습니다'),
      });
      setDoc(fresh);
      return fresh;
  }
  async function saveStudioRevision(current: Document) {
    if (!pageEditor.current) throw new Error('편집기를 준비하고 있습니다. 잠시 후 다시 시도하세요.');
    const handle = pageEditor.current;
    const bytes = await handle.export();
    setBusy('원문 편집을 저장하고 문서 구조를 검사하고 있습니다');
    const fresh = await documentCommands.execute(current, { kind: 'studio', bytes }, {
      onStage: () => setBusy('편집 파일의 구조·참조·보호 문구를 검사하고 있습니다'),
      onReconnect: () => setBusy('연결을 다시 확인하고 기존 편집 저장 작업을 찾고 있습니다'),
    });
    await handle.saved(fresh.revision);
    setDoc(fresh);
    await refreshList();
    return fresh;
  }
  async function exportRevision(current: Document) {
      setBusy('다운로드를 준비하고 있습니다');
      const result = await post<{ url: string }>(`/documents/${current.id}/exports`, { revision: current.revision });
      return result.url;
  }
  async function save() {
    if (!doc || !dirty) return;
    await task('수정 내용을 저장하고 있습니다', async () => {
      const fresh = studioDirty ? await saveStudioRevision(doc) : await saveRevision(doc, Object.values(history.present));
      setNotice(`버전 ${fresh.revision + 1} 저장 완료. 다운로드하려면 현재 버전을 검증하세요.`);
    });
  }
  const saveCurrent = useRef(save);
  saveCurrent.current = save;
  const onPageSave = useCallback(() => { void saveCurrent.current(); }, []);
  async function validate() {
    if (!doc) return;
    await task('문서 검증을 시작하고 있습니다', async () => {
      await checkRevision(doc); setReportOpen(true); setNotice('현재 수정 버전의 검증이 완료되었습니다.');
    });
  }
  async function download() {
    if (!doc) return;
    await task('다운로드를 준비하고 있습니다', async () => {
      const current = studioDirty ? await saveStudioRevision(doc) : doc;
      const url = await prepareDownload(current, Object.values(history.present), { save: saveRevision, validate: checkRevision, export: exportRevision });
      const link = document.createElement('a'); link.href = url; link.download = ''; link.click();
      setNotice('검증한 수정 버전의 다운로드를 시작했습니다.');
    });
  }
  async function recoverBackup(file: File) {
    if (!mayLeave()) throw new Error('현재 편집 내용을 저장한 뒤 다시 열어 주세요.');
    if (taskRunning.current) throw new Error('진행 중인 작업이 끝난 뒤 다시 시도해 주세요.');
    if (!file.size || file.size > 20 * 1024 * 1024) throw new Error('복구할 파일은 0바이트보다 크고 20MB 이하여야 합니다. 임시본을 내려받아 확인해 주세요.');
    taskRunning.current = true; setBusy('임시본을 새 문서로 가져오고 있습니다'); setError('');
    try {
      const form = new FormData(); form.append('file', file);
      const result = await api<{ document_id: string; job_id: string }>('/documents', { method: 'POST', body: form });
      await waitForJob(result.job_id, () => setBusy('임시본의 문서 구조를 분석하고 있습니다'));
      await load(result.document_id); setBackupShelf(false); setNotice('임시본을 별도의 새 문서로 열었습니다. 보관함의 임시본도 유지됩니다.');
    } finally { taskRunning.current = false; setBusy(''); }
  }
  async function manageDocument(mode: 'rename' | 'copy', name: string) {
    if (!doc || taskRunning.current) throw new Error('진행 중인 작업이 끝난 뒤 다시 시도해 주세요.');
    taskRunning.current = true; setError(''); setNotice('');
    setBusy(mode === 'rename' ? '문서 이름을 변경하고 있습니다' : '복사본을 만들고 있습니다');
    try {
      if (mode === 'rename') {
        const fresh = await api<Document>(`/documents/${doc.id}/name`, { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ revision: doc.revision, expected_name: doc.name, name }) });
        setDoc(fresh); await refreshList(); setNotice('파일 이름을 변경했습니다. 편집 내용은 유지됩니다.');
      } else {
        const source = studioDirty ? await saveStudioRevision(doc) : dirty ? await saveRevision(doc, Object.values(history.present)) : doc;
        const result = await post<{ document_id: string; job_id: string }>(`/documents/${source.id}/copies`, { revision: source.revision, name });
        await waitForJob(result.job_id, () => setBusy('복사본을 분석하고 문서 구조를 검사하고 있습니다'));
        await load(result.document_id); setNotice('독립된 복사본을 열었습니다. 이전 문서와 저장 기록은 내 문서에 남아 있습니다.');
      }
      setFileAction(null);
    } finally { taskRunning.current = false; setBusy(''); }
  }
  const hostAction = useRef((action: string, file?: File) => {});
  hostAction.current = (action, file) => {
    if (taskRunning.current) return;
    if (action === 'new') setNewDocumentOpen(true);
    else if (action === 'open') inputRef.current?.click();
    else if (action === 'documents') setSidebarOpen(true);
    else if (action === 'save') void save();
    else if (action === 'copy') setFileAction('copy');
    else if (action === 'download') void download();
    else if (action === 'history') setVersionsOpen(true);
    else if (action === 'drop-many') setError('한 번에 한 파일씩 열어 주세요.');
    else if (action === 'drop' && file) void upload(file);
  };
  const onHostFileAction = useCallback((action: string, file?: File) => hostAction.current(action, file), []);
  async function deleteDocument(item: DocumentSummary) {
    if (!window.confirm(`“${item.name}”의 원본과 모든 수정본을 삭제할까요?`)) return;
    await task('문서를 삭제하고 있습니다', async () => {
      await api(`/documents/${item.id}`, { method: 'DELETE' });
      clearDraftBackup(item.id);
      await studioBackup.clear(item.id);
      if (doc?.id === item.id) { setDoc(null); dispatch({ type: 'reset' }); lastDocument.write(null); }
      await refreshList(); setNotice('문서와 저장된 파일을 삭제했습니다.');
    });
  }
  async function restoreVersion(source: number) {
    if (!doc || dirty || taskRunning.current) throw new Error('미저장 편집을 저장하고 진행 중인 작업이 끝난 뒤 복원해 주세요.');
    taskRunning.current = true;
    setBusy('이전 버전을 복원하고 문서 구조를 검사하고 있습니다'); setError(''); setNotice('');
    try {
      await documentCommands.execute(doc, { kind: 'restore', source_revision: source }, {
        onStage: () => setBusy('이전 버전의 구조·참조·현재 고급 설정을 검사하고 있습니다'),
        onReconnect: () => setBusy('연결을 다시 확인하고 기존 복원 작업을 찾고 있습니다'),
      });
      await load(doc.id);
      setNotice(`버전 ${source + 1}을 새 버전 ${doc.revision + 2}로 복원했습니다. 이전 기록은 그대로 유지됩니다.`);
    } finally { taskRunning.current = false; setBusy(''); }
  }
  const visibleSections = useMemo(() => doc?.sections.filter(s => section === -1 || s.index === section) || [], [doc?.sections, section]);
  const paragraphCount = doc?.sections.reduce((n, s) => n + s.paragraphs.length, 0) || 0;
  const tableCount = doc?.sections.reduce((n, s) => n + s.tables.length, 0) || 0;
  const value = (key: string, original: string) => history.present[key]?.text ?? original;
  const normalizedSearch = search.toLocaleLowerCase();
  const visibleTables = useMemo(() => visibleSections.flatMap(s => s.tables).filter(t => !normalizedSearch || t.cells.some(c => value(`${t.id}-${c.row}-${c.column}`, c.text).toLocaleLowerCase().includes(normalizedSearch))), [visibleSections, normalizedSearch, history.present]);
  const activeTable = visibleTables.find(t => t.id === tableId) || visibleTables[0];
  const visibleParagraphs = useMemo(() => visibleSections.flatMap(s => s.paragraphs).filter(p => (showEmpty || !!value(p.id, p.text).trim() || !!history.present[p.id]) && (showReadonly || p.editable) && (!normalizedSearch || value(p.id, p.text).toLocaleLowerCase().includes(normalizedSearch))), [visibleSections, showEmpty, showReadonly, normalizedSearch, history.present]);
  const visibleParagraphIds = useMemo(() => new Set(visibleParagraphs.map(p => p.id)), [visibleParagraphs]);
  function openHome() {
    if (!mayLeave()) return;
    if (doc && dirty && !staleBackup) clearDraftBackup(doc.id);
    setDoc(null); dispatch({ type: 'reset' }); setSidebarOpen(false);
    setStaleBackup(null);
    lastDocument.write(null);
  }
  const listedDocuments = documents.filter(item => item.name.toLocaleLowerCase().includes(fileSearch.toLocaleLowerCase()));
  const formatDate = (timestamp: number) => new Date(timestamp * 1000).toLocaleString('ko-KR', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false });
  const update = useCallback((key: string, operation: Operation, original: string) => { dispatch({ type: 'change', key, operation, original }); }, []);
  const selectFlowParagraph = useCallback((p: Paragraph) => setSelected({ label: `${p.section + 1}섹션 · ${p.index + 1}문단`, text: '본문의 텍스트를 수정합니다. 수정한 부분만 문서에 저장합니다.' }), []);
  const showOriginal = useCallback(() => { setReaderOpened(true); setTab('original'); }, []);
  function revertChanges(items: TextTarget[]) {
    dispatch({ type: 'batch', changes: items.map(item => ({ key: item.id, original: item.original, operation: { ...item.operation, text: item.original } })) });
  }
  function goToChange(operation: Operation) {
    setReviewOpen(false); setSection(-1); setSearch('');
    if (operation.type === 'paragraph') {
      setTab('paragraphs'); setShowEmpty(true);
      setFocusTarget(`${operation.section + 1}섹션 ${operation.index + 1}문단`);
    } else {
      setTab('tables'); setTableEditing(true); setTableId(`s${operation.section}-t${operation.table}`);
      setFocusTarget(`${operation.section + 1}섹션 표 ${operation.table + 1} ${operation.row + 1}행 ${operation.column + 1}열`);
    }
  }
  useEffect(() => {
    if (!focusTarget || reviewOpen) return;
    const frame = requestAnimationFrame(() => {
      const field = document.querySelector<HTMLTextAreaElement>(`textarea[aria-label="${CSS.escape(focusTarget)}"]`);
      field?.focus(); field?.scrollIntoView({ block: 'center' }); setFocusTarget(null);
    });
    return () => cancelAnimationFrame(frame);
  }, [focusTarget, reviewOpen, tab, tableId, section, search]);
  useEffect(() => {
    function keyboard(event: KeyboardEvent) {
      if (doc && (event.ctrlKey || event.metaKey) && !event.altKey && !event.shiftKey && event.key.toLowerCase() === 'p') {
        event.preventDefault();
        if (disabled || document.querySelector('[role="dialog"]')) return;
        if (tab !== 'studio') { setError('현재 수정 내용을 저장하고 문서 편집 보기에서 인쇄해 주세요.'); return; }
        void pageEditor.current?.print().catch(e => setError(e instanceof Error ? e.message : '인쇄를 준비하지 못했습니다.'));
        return;
      }
      if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== 's' || !doc) return;
      event.preventDefault();
      if (disabled || document.querySelector('[role="dialog"]')) return;
      if (event.shiftKey) { setFileAction('copy'); return; }
      if (dirty) void save();
    }
    window.addEventListener('keydown', keyboard);
    return () => window.removeEventListener('keydown', keyboard);
  }, [doc, dirty, disabled, staleBackup, tab]);
  function downloadOldBackup() {
    if (!staleBackup) return;
    const url = URL.createObjectURL(new Blob([backupText(staleBackup)], { type: 'text/plain;charset=utf-8' }));
    const link = document.createElement('a'); link.href = url; link.download = `${doc?.name || '문서'}_이전_임시편집.txt`; link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }
  function reopenWithBackup() {
    if (!doc || disabled || staleBackup) return;
    if (!writeDraftBackup(doc, history.present)) { setError('임시보관을 완료하지 못해 문서를 다시 열지 않았습니다. 수정 저장을 먼저 해 주세요.'); return; }
    reloadWithBackup.current = true;
    window.location.reload();
  }
  async function reopenLatest() {
    if (!doc) return;
    await task('편집을 임시보관하고 최신 문서를 열고 있습니다', async () => {
      if (studioDirty) {
        if (!pageEditor.current) throw new Error('편집기를 준비한 뒤 다시 시도해 주세요.');
        const bytes = await pageEditor.current.export();
        if (!bytes.byteLength || bytes.byteLength > 20 * 1024 * 1024) throw new Error('임시보관할 편집 파일은 0바이트보다 크고 20MB 이하여야 합니다. 현재 편집을 확인해 주세요.');
        await studioBackup.write({ id: doc.id, revision: doc.revision, name: doc.name, bytes, saved: Date.now() });
      }
      if (Object.keys(history.present).length) {
        if (staleBackup) throw new Error('이전 임시본문을 먼저 내려받고 정리해 주세요. 현재 편집도 임시본으로 내려받아 보관할 수 있습니다.');
        if (!writeDraftBackup(doc, history.present)) throw new Error('임시보관을 완료하지 못해 최신 문서를 열지 않았습니다. 현재 편집 임시본을 내려받아 주세요.');
      }
      await load(doc.id, true);
      setNotice(dirty ? '최신 문서를 열었습니다. 이전 편집은 임시본으로 보관했습니다. 내용을 비교해 반영하세요.' : '최신 문서를 열었습니다.');
    });
  }
  async function downloadCurrentDraft() {
    if (!doc || !dirty) return;
    await task('현재 편집 임시본을 준비하고 있습니다', async () => {
      let blob: Blob;
      let extension: string;
      if (studioDirty) {
        if (!pageEditor.current) throw new Error('편집기를 준비한 뒤 다시 시도해 주세요.');
        blob = new Blob([new Uint8Array(await pageEditor.current.export())], { type: 'application/hwp+zip' });
        extension = 'hwpx';
      } else {
        const changes = textTargets(doc, history.present).filter(target => !!history.present[target.id])
          .map(target => ({ id: target.id, original: target.original, operation: history.present[target.id] }));
        blob = new Blob([backupText({ version: 1, documentId: doc.id, revision: doc.revision, expires: doc.expires, updatedAt: Date.now(), changes })], { type: 'text/plain;charset=utf-8' });
        extension = 'txt';
      }
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a'); link.href = url; link.download = `${doc.name.replace(/\.[^.]+$/, '')}_미저장편집.${extension}`; link.click();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
      setNotice('미저장 편집의 임시본을 내려받았습니다. 서버에 저장·검증한 파일과 구분해 보관하세요.');
    });
  }
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

  return <div className={`app ${doc ? 'editing editor-workspace' : ''} ${doc && tab === 'original' ? 'reading-document' : ''} ${continuousText ? 'text-flow-document' : ''} ${doc && (tab === 'original' || continuousText) && toolsCollapsed ? 'reading-tools-collapsed' : ''}`} onDragOver={e => { if (!e.dataTransfer.types.includes('Files')) return; e.preventDefault(); if (!disabled) setDragging(true); }} onDragLeave={e => { if (!e.relatedTarget) setDragging(false); }} onDrop={e => { if (!e.dataTransfer.files.length) return; e.preventDefault(); setDragging(false); if (e.dataTransfer.files.length > 1) { setError('한 번에 한 파일씩 열어 주세요.'); return; } if (!disabled) void upload(e.dataTransfer.files[0]); }}>
    <a className="skip-link" href="#main-content">본문으로 건너뛰기</a>
    <input ref={inputRef} type="file" accept=".hwp,.hwpx" hidden aria-label="HWP 또는 HWPX 파일 선택" onChange={e => { void upload(e.target.files?.[0]); e.target.value = ''; }} />
    <aside className={`sidebar ${sidebarOpen ? 'open' : ''}`} inert={!!doc && !sidebarOpen} aria-hidden={!!doc && !sidebarOpen}>
      <button className="brand" disabled={disabled} onClick={openHome} aria-label="Jakal 내 문서"><span className="brand-mark">j</span><strong>jakal</strong><span className="brand-caption">문서 도구</span></button>
      <button className="new-button" disabled={disabled} onClick={() => inputRef.current?.click()}><Plus size={17} />파일 열기</button>
      <button className="nav-item" disabled={disabled} onClick={() => { setError(''); setNewDocumentOpen(true); }}><FileText size={17} />새 문서 만들기</button>
      <button className="nav-item" disabled={disabled} onClick={() => setPageLayoutOpen(true)}><LayoutPanelLeft size={17} />쪽 양식·템플릿</button>
      <nav aria-label="문서 탐색"><button className={`nav-item ${!doc ? 'active' : ''}`} disabled={disabled} onClick={openHome}><LayoutPanelLeft size={17} />내 문서</button></nav>
      <div className="nav-heading recent-heading">최근에 연 문서<span>{documents.length}</span></div>
      <div className="document-list">{documents.length ? documents.map(item => <div key={item.id} className={`document-item ${doc?.id === item.id ? 'active' : ''}`}><button disabled={disabled} title={item.name} onClick={() => { if (mayLeave()) void task('문서를 여는 중', () => load(item.id)); }}><FileText size={16} /><span><b>{item.name}</b><small>{item.status === 'failed' ? '분석 실패' : item.status === 'processing' ? '분석 중' : `버전 ${item.revision + 1}`}</small></span></button><button className="delete-button" disabled={disabled} aria-label={`${item.name} 삭제`} onClick={() => void deleteDocument(item)}><Trash2 size={14} /></button></div>) : <p className="no-documents">최근 문서가 없습니다.</p>}</div>
      <div className="sidebar-bottom"><button disabled={disabled} onClick={() => setBackupShelf(true)}><RotateCcw size={17} />임시 편집 보관함</button><button disabled={disabled} onClick={() => setPrivacyOpen(true)}><ShieldCheck size={17} />보안·사용 종료</button><button onClick={() => setHelpOpen(true)}><HelpCircle size={17} />도움말</button><div className="session-note"><LockKeyhole size={13} /><span>이 브라우저에 연결된 문서<br />업로드 후 24시간 보관</span></div></div>
    </aside>
    {sidebarOpen && <button className="sidebar-scrim" aria-label="메뉴 닫기" onClick={() => setSidebarOpen(false)} />}
    <main className="main" id="main-content" tabIndex={-1}>
      <ConnectionStatus />
      <header className="topbar"><div className="breadcrumb"><button className="menu-toggle icon-button" aria-label="메뉴 열기" onClick={() => setSidebarOpen(true)}><PanelLeftClose size={19} /></button><button disabled={disabled} onClick={openHome}>내 문서</button>{doc && <><ChevronRight size={13} /><span title={doc.name}>{doc.name}</span></>}</div><button className="help-button" aria-label="사용 안내" onClick={() => setHelpOpen(true)}><BookOpen size={16} /><span>사용 안내</span></button></header>
      {error && <div className="message error" role="alert"><TriangleAlert size={17} /><span>{error}</span>{!doc && <button className="text-button" disabled={disabled || booting} onClick={() => void task('문서 목록을 확인하고 있습니다', refreshList)}>목록 다시 확인</button>}<button aria-label="오류 닫기" onClick={() => setError('')}><X size={17} /></button></div>}
      {doc && <DocumentSyncNotice doc={doc} disabled={disabled} dirty={!!dirty} onReload={() => void reopenLatest()} onBackup={() => void downloadCurrentDraft()} />}
      {notice && <div className="message success" role="status"><CircleCheck size={17} /><span>{notice}</span><button aria-label="알림 닫기" onClick={() => setNotice('')}><X size={17} /></button></div>}
      {busy && <div className="busy-bar" role="status"><LoaderCircle size={16} className="spin" />{busy}</div>}
      {staleBackup && doc && <div className="draft-conflict" role="status"><div><strong>다른 버전의 임시 편집이 있습니다.</strong><span>덮어쓰지 않고 보관했습니다. 이전 임시본을 정리하면 현재 편집의 임시보관을 다시 시작합니다.</span></div><div><button className="text-button" onClick={downloadOldBackup}>임시본문 받기</button>{confirmDropBackup ? <><span>이전 임시 편집을 버릴까요?</span><button className="text-button" onClick={() => { if (clearDraftBackup(doc.id)) { setStaleBackup(null); setConfirmDropBackup(false); } else setError('브라우저 임시보관함에 접근할 수 없습니다.'); }}>버리기 확인</button><button className="text-button" onClick={() => setConfirmDropBackup(false)}>취소</button></> : <button className="text-button" onClick={() => setConfirmDropBackup(true)}>이전 임시본 버리기</button>}</div></div>}
      {!doc && booting ? <div className="workspace-loading" role="status"><LoaderCircle className="spin" size={22} /><p>최근 문서를 불러오고 있습니다</p></div> : !doc ? <div className="library">
        <section className="library-heading"><div><h1>내 문서</h1><p>HWP · HWPX 문서를 열어 문단과 표를 수정합니다.</p></div><div className="library-heading-actions"><button className="outline-button" disabled={disabled} onClick={() => { setError(''); setNewDocumentOpen(true); }}><FileText size={16} />새 문서</button><button className="primary-button" disabled={disabled} onClick={() => inputRef.current?.click()}><Plus size={16} />파일 열기</button></div></section>
        <section className={`upload-zone ${dragging ? 'dragging' : ''}`} onDragLeave={() => setDragging(false)} aria-label="파일 업로드">
          <ArrowUpFromLine size={24} strokeWidth={1.4} /><div><h2>HWP 또는 HWPX 파일을 끌어 놓으세요</h2><p>HWP는 HWPX로 변환 · 최대 20MB</p></div><button className="outline-button" disabled={disabled} onClick={() => inputRef.current?.click()}>파일 선택</button>
        </section>
        <section className="library-files" aria-labelledby="recent-title"><div className="list-toolbar"><div><h2 id="recent-title">최근 문서</h2><span>{documents.length}</span></div><label className="file-search"><Search size={16} /><input aria-label="파일 이름 검색" placeholder="파일 이름 검색" value={fileSearch} onChange={e => setFileSearch(e.target.value)} /></label></div>
          <div className="file-table-scroll"><table className="file-table"><thead><tr><th>이름</th><th>업로드</th><th>상태</th><th><span className="sr-only">문서 삭제</span></th></tr></thead><tbody>{listedDocuments.map(item => <tr key={item.id}><td><button className="file-row-button" disabled={disabled} onClick={() => { if (mayLeave()) void task('문서를 여는 중', () => load(item.id)); }}><FileText size={20} strokeWidth={1.4} /><span><strong>{item.name}</strong><small>{item.name.toLowerCase().endsWith('.hwp') ? 'HWP' : 'HWPX'} 문서</small></span></button>{item.status === 'failed' && <a className="failed-original-link" href={`/api/documents/${item.id}/original`} download>업로드 원본 받기</a>}</td><td><time dateTime={new Date(item.created * 1000).toISOString()}>{formatDate(item.created)}</time></td><td><span className={`file-status ${item.status}`}><i />{item.status === 'failed' ? '분석 실패' : item.status === 'processing' ? '분석 중' : `버전 ${item.revision + 1}`}</span></td><td><button className="icon-button" aria-label={`${item.name} 목록에서 삭제`} disabled={disabled} onClick={() => void deleteDocument(item)}><Trash2 size={15} /></button></td></tr>)}</tbody></table></div>
          {!listedDocuments.length && <div className="library-empty"><FileText size={31} strokeWidth={1.1} /><h3>{fileSearch ? '일치하는 문서가 없습니다' : '아직 열어 본 문서가 없습니다'}</h3><p>{fileSearch ? '다른 파일 이름으로 검색하세요.' : '파일을 열면 이 목록에서 작업을 이어갈 수 있습니다.'}</p>{!fileSearch && <button className="text-button" disabled={disabled} onClick={() => void sample()}>예제 문서로 시작하기<ArrowRight size={15} /></button>}</div>}
        </section>
        <div className="library-footer"><span><Clock3 size={14} />문서는 업로드 후 24시간이 지나면 삭제됩니다.</span><div className="sample-links"><button className="text-button" disabled={disabled} onClick={() => void sample('hwp')}>HWP 예제<ArrowRight size={14} /></button><button className="text-button" disabled={disabled} onClick={() => void sample()}>HWPX 예제<ArrowRight size={14} /></button></div></div>
        <section className="library-guide"><h2>처음 사용하시나요?</h2><p>파일을 열면 원문 서식으로 편집기가 표시됩니다.<br />문서 위를 클릭해 수정하거나 상단 메뉴에서 글자·문단·표·쪽 설정을 바꾸세요. 하단 작업 도구(Ctrl+Shift+K)에서 찾기·바꾸기·서식·표·쪽 설정을 검색하고 실행할 수 있습니다.<br /><b>검증하기</b>로 저장 결과를 확인하면 수정본을 다운로드할 수 있습니다.</p><button className="text-button" onClick={() => setHelpOpen(true)}>사용 안내<ArrowRight size={14} /></button></section>
      </div> : <>
        <section className="document-header"><button className="icon-button document-menu-toggle" aria-label="문서 목록 열기" title="문서 목록 열기" onClick={() => setSidebarOpen(true)}><LayoutPanelLeft size={19} /></button><div className="file-heading"><FileText size={24} strokeWidth={1.4} /><div><h1 title={doc.name}><button className="file-name-trigger" disabled={disabled} aria-label="문서 이름 변경" title="문서 이름 변경" onClick={() => setFileAction('rename')}>{doc.name}</button></h1><p>{doc.conversion ? 'HWP → HWPX' : doc.format.toUpperCase()}<span>·</span>버전 {doc.revision + 1}<span>·</span><span className={dirty ? 'unsaved-label' : ''}>{studioDirty ? '편집 중 · 저장 전' : dirty ? `${dirty}개 변경 · 저장 전` : '저장됨'}</span></p></div></div><WorkspaceMenu view={tab} disabled={disabled} dirty={dirty} studioDirty={studioDirty} tableCount={tableCount} onView={next => { changeView(next); if (next === 'tables') setTableEditing(false); }} onValidate={() => void validate()} onInfo={() => setInfoOpen(!infoOpen)} onSettings={() => setAdvancedOpen(true)} onHelp={() => setHelpOpen(true)} onReview={() => setReviewOpen(true)} onHistory={() => setVersionsOpen(true)} onFiles={() => setFileAction('rename')} onPageLayout={() => setPageLayoutOpen(true)} /><div className="document-actions"><button className="outline-button header-save" disabled={disabled || !dirty} onClick={() => void save()} title="수정 저장 (Ctrl S)"><Save size={15} />수정 저장</button><button className="primary-button" disabled={disabled} onClick={() => void download()} aria-label={dirty ? '저장·검증 후 다운로드' : saved ? '다운로드' : '검증 후 다운로드'} title="필요한 수정 저장과 현재 버전 검증을 마친 뒤 다운로드합니다"><ArrowDownToLine size={16} /><span className="download-label">{dirty ? '저장·검증 후 다운로드' : saved ? '다운로드' : '검증 후 다운로드'}</span><span className="download-label-short">{dirty ? '저장 후 받기' : saved ? '다운로드' : '검증 후 받기'}</span></button></div></section>
        <div className={`editor-layout ${!infoOpen ? 'wide-layout' : ''}`}><section className="editor">
          <PageEditor key={`editor-${doc.id}-${editorEpoch}`} ref={pageEditor} doc={doc} active={tab === 'studio'} disabled={disabled || versionsOpen} onDirty={setStudioDirty} onSave={onPageSave} onFileAction={onHostFileAction} />{(readerOpened || tab === 'original' || tab === 'figures' || tab === 'tables') && <ToolBoundary onClose={() => { setReaderOpened(false); setTab('studio'); }}><React.Suspense fallback={<div role="status">원문 보기 도구를 불러오고 있습니다.</div>}><DocumentReader key={doc.id} doc={doc} active={tab === 'original' || tab === 'figures' || (tab === 'tables' && !tableEditing)} mode={tab === 'figures' ? 'figures' : tab === 'tables' ? 'tables' : 'original'} dirty={dirty} onMode={setTab} onEdit={() => setTab('paragraphs')} onEditTable={() => setTableEditing(true)} toolsCollapsed={toolsCollapsed} onToggleTools={toggleReaderTools} onReadingScroll={() => { if (!toolsPinned) setToolsCollapsed(true); }} /></React.Suspense></ToolBoundary>}
          {(tab === 'paragraphs' || (tab === 'tables' && tableEditing)) && <>
          <div className={`edit-actions-bar ${continuousText ? 'flow-actions-bar' : ''}`}>{continuousText && <span className="flow-compact-label">본문 편집</span>}<div className="edit-view-controls">{tab === 'paragraphs' ? <><div className="text-view-switch" role="group" aria-label="텍스트 보기 방식"><button aria-pressed={textView === 'flow'} onClick={() => setTextView('flow')}>본문 이어 보기</button><button aria-pressed={textView === 'list'} onClick={() => setTextView('list')}>문단 목록</button></div>{textView === 'list' && <><label><input type="checkbox" checked={showEmpty} onChange={e => setShowEmpty(e.target.checked)} />빈 문단 표시</label><label><input type="checkbox" checked={showReadonly} onChange={e => setShowReadonly(e.target.checked)} />읽기 전용 문단 표시</label></>}</> : <label className="table-picker-label"><button className="text-button" onClick={() => setTableEditing(false)}>원문 표 보기</button><select aria-label="표 선택" value={activeTable?.id || ''} onChange={e => setTableId(e.target.value)}>{visibleTables.map(t => <option key={t.id} value={t.id}>섹션 {t.section + 1} · 표 {t.index + 1} ({t.rows}행 × {t.columns}열)</option>)}</select></label>}</div><div className="history-actions"><button className="icon-button" aria-label="실행 취소" title="실행 취소" disabled={disabled || !history.past.length} onClick={() => dispatch({ type: 'undo' })}><RotateCcw size={16} /></button><button className="icon-button" aria-label="다시 실행" title="다시 실행" disabled={disabled || !history.future.length} onClick={() => dispatch({ type: 'redo' })}><RotateCw size={16} /></button><span /><button className="save-button" disabled={disabled || !dirty} onClick={() => void save()} title="수정 저장 (Ctrl/Cmd S)"><Save size={15} />수정 저장{dirty > 0 && <b>{dirty}</b>}</button>{continuousText && <button className="icon-button flow-tools-toggle" aria-label={toolsCollapsed ? '편집 도구 펼치기' : '편집 도구 접기'} aria-expanded={!toolsCollapsed} title={toolsCollapsed ? '편집 도구 펼치기' : '편집 도구 접기'} onClick={toggleReaderTools}>{toolsCollapsed ? <ChevronDown size={17} /> : <ChevronUp size={17} />}<span>도구</span></button>}</div></div>
          <div className="filter-bar"><label className="search-field"><Search size={15} /><input ref={textSearch} placeholder="문서에서 찾기" aria-label="문서 내용 검색" value={search} onChange={e => setSearch(e.target.value)} onKeyDown={e => { if (continuousText && e.key === 'Enter') { e.preventDefault(); flowEditor.current?.find(e.shiftKey ? -1 : 1); } }} />{search && <button aria-label="검색 초기화" onClick={() => setSearch('')}><X size={14} /></button>}</label><select aria-label="섹션 선택" value={section} onChange={e => setSection(Number(e.target.value))}><option value={-1}>전체 섹션</option>{doc.sections.map(s => <option key={s.index} value={s.index}>섹션 {s.index + 1}</option>)}</select></div>
          {continuousText ? <TextFlowEditor key={doc.id} ref={flowEditor} sections={visibleSections} drafts={history.present} search={search} disabled={disabled} onChange={update} onSelect={selectFlowParagraph} onOriginal={showOriginal} /> : <>
          {doc.format_note && <div className="format-note"><span className="format-badge">{doc.conversion ? 'HWPX 변환' : 'HWP 5'}</span><p>{doc.format_note}</p></div>}<div className="document-canvas"><div className="canvas-note"><span>내용 편집</span><span>인쇄 배치는 한글 프로그램에서 확인하세요.</span></div><div className={`document-paper ${tab === 'tables' ? 'table-paper' : ''}`}>
          {tab === 'paragraphs' ? visibleSections.map(s => <section className="document-section" key={s.index}>{s.paragraphs.some(p => visibleParagraphIds.has(p.id)) && <div className="section-label"><span>섹션 {s.index + 1}</span><span>텍스트 편집</span></div>}{s.paragraphs.map(paragraph)}</section>) : activeTable ? <><div className="section-label"><span>섹션 {activeTable.section + 1} · 표 {activeTable.index + 1}</span><span>병합 구조 유지</span></div>{tableView(activeTable)}</> : null}
          {(tab === 'paragraphs' ? !visibleParagraphs.length : !activeTable) && <div className="empty-state"><FileText size={27} /><p>{search ? '검색 결과가 없습니다.' : tab === 'tables' ? '이 범위에는 표가 없습니다.' : '이 범위에는 편집 가능한 텍스트가 없습니다.'}</p>{tab === 'paragraphs' && <button className="text-button" onClick={() => setTab('original')}>수식·그림은 원문에서 보기<ArrowRight size={14} /></button>}</div>}
          </div></div></>}
          <div className="editor-bottom"><button disabled={!dirty || disabled} onClick={() => setReviewOpen(true)} title="변경 전후를 확인하고 항목별로 되돌리기"><span className={`status-dot ${dirty ? 'unsaved' : ''}`} />{dirty ? `${dirty}개 변경 · 검토` : '저장됨'}</button><span className={`draft-backup-state ${backupStatus === 'unavailable' ? 'unavailable' : ''}`} title="임시보관은 이 탭에서만 유지됩니다. 실제 파일은 수정 저장을 눌러 반영하세요.">{dirty ? backupStatus === 'stored' ? '이 탭에 임시보관됨' : backupStatus === 'pending' ? '임시보관 중' : backupStatus === 'paused' ? '임시보관 일시 중지' : '임시보관 불가 · 수정 저장 필요' : ''}{!!dirty && backupStatus === 'stored' && <button className="draft-reopen" aria-label="임시보관 후 문서 다시 불러오기" title="임시 편집을 보관하고 최신 문서를 다시 엽니다" disabled={disabled} onClick={reopenWithBackup}><RotateCw size={12} /></button>}</span><button onClick={() => setReportOpen(true)}>{saved && !dirty ? <><CircleCheck size={13} />현재 버전 검증 완료</> : '검사 결과'}<ChevronRight size={12} /></button></div>
          </>}
        </section>{infoOpen && <aside className="inspector" aria-label="문서 정보 패널"><div className="inspector-title"><span>문서 정보</span><button className="icon-button" aria-label="문서 정보 닫기" onClick={() => setInfoOpen(false)}><X size={16} /></button></div><div className="inspector-body"><h2>{doc.title}</h2><dl className="document-details"><div><dt>형식</dt><dd>{doc.conversion ? 'HWP → HWPX' : doc.format.toUpperCase()}</dd></div><div><dt>버전</dt><dd>{doc.revision + 1}</dd></div><div><dt>문단</dt><dd>{paragraphCount}</dd></div><div><dt>표</dt><dd>{tableCount}</dd></div><div><dt>섹션</dt><dd>{doc.sections.length}</dd></div></dl><hr /><h3 className="inspector-label">선택한 요소</h3>{selected ? <div className="selection-info"><h3>{selected.label}</h3><p>{selected.text}</p></div> : <p className="selection-placeholder">문단 또는 표의 셀을 선택하세요.</p>}<hr /><h3 className="inspector-label">검증</h3><div className="validation-state"><span className={`status-dot ${saved && !dirty ? '' : 'pending'}`} /><strong>{saved && !dirty ? '현재 버전 검증 완료' : dirty ? '저장 후 검증 필요' : '검증 필요'}</strong></div><p className="inspector-description">저장 결과를 검사한 뒤 다운로드할 수 있습니다.</p><button className="text-button" disabled={disabled || !!dirty} onClick={() => saved ? setReportOpen(true) : void validate()}>{saved ? '검사 결과 보기' : '검증하기'}<ChevronRight size={14} /></button><hr /><a className="text-download" href={`/api/documents/${doc.id}/text`} download><ArrowDownToLine size={14} />본문 텍스트 다운로드</a><div className="retention"><Clock3 size={14} /><p>{formatDate(doc.expires)} 자동 삭제</p></div></div></aside>}</div>
      </>}
    </main>
    {advancedOpen && doc && <ToolBoundary onClose={() => setAdvancedOpen(false)}><React.Suspense fallback={<div className="workbench-loading" role="status">도구를 불러오고 있습니다.</div>}><AdvancedSettings doc={doc} drafts={history.present} onClose={() => setAdvancedOpen(false)} onApply={applyProposals} onSettings={setDoc} /></React.Suspense></ToolBoundary>}
    {newDocumentOpen && <ToolBoundary onClose={closeNewDocument}><React.Suspense fallback={<div className="workbench-loading" role="status">문서 양식을 열고 있습니다.</div>}><NewDocument busy={disabled} error={error} onClose={closeNewDocument} onCreate={value => void createDocument(value)} /></React.Suspense></ToolBoundary>}
    {pageLayoutOpen && <ToolBoundary onClose={closePageLayout}><React.Suspense fallback={<div className="workbench-loading" role="status">쪽 양식 도구를 열고 있습니다.</div>}><PageLayoutTemplates doc={doc} dirty={!!dirty} busy={disabled} onClose={closePageLayout} onApply={applyPageLayout} /></React.Suspense></ToolBoundary>}
    {reviewOpen && doc && <ToolBoundary onClose={closeReview}><React.Suspense fallback={<div className="workbench-loading" role="status">도구를 불러오고 있습니다.</div>}><ChangeReview doc={doc} drafts={history.present} onClose={closeReview} onRevert={revertChanges} onGo={goToChange} /></React.Suspense></ToolBoundary>}
    {versionsOpen && doc && <ToolBoundary onClose={closeVersions}><React.Suspense fallback={<div className="workbench-loading" role="status">도구를 불러오고 있습니다.</div>}><VersionHistory doc={doc} dirty={!!dirty} onClose={closeVersions} onRestore={restoreVersion} /></React.Suspense></ToolBoundary>}
    {backupShelf && <ToolBoundary onClose={closeBackupShelf}><React.Suspense fallback={<div className="workbench-loading" role="status">임시보관함을 열고 있습니다.</div>}><BackupShelf activeId={doc?.id} names={Object.fromEntries(documents.map(item => [item.id, item.name]))} onOpen={recoverBackup} onClose={closeBackupShelf} /></React.Suspense></ToolBoundary>}
    {fileAction && doc && <ToolBoundary onClose={closeFileAction}><React.Suspense fallback={<div className="workbench-loading" role="status">파일 작업을 열고 있습니다.</div>}><DocumentActions doc={doc} initialMode={fileAction} dirty={!!dirty} onClose={closeFileAction} onSubmit={manageDocument} /></React.Suspense></ToolBoundary>}
    {dragging && <div className="drop-overlay" onDragLeave={() => setDragging(false)}><ArrowUpFromLine size={32} strokeWidth={1.4} /><h2>파일을 놓아 열기</h2><p>HWP · HWPX · 최대 20MB</p></div>}
    {reportOpen && doc && <div className="modal-backdrop" onClick={() => setReportOpen(false)}><section className="modal report-modal" role="dialog" aria-modal="true" aria-labelledby="report-title" onClick={e => e.stopPropagation()}><div className="modal-heading"><div><h2 id="report-title">검증 결과</h2><p>버전 {doc.revision + 1} · {saved ? '검증 완료' : '자동 검사 결과'}</p></div><button className="icon-button" aria-label="검증 결과 닫기" onClick={() => setReportOpen(false)}><X size={20} /></button></div><div className={`report-summary ${doc.report.issues.length ? 'warning' : ''}`}><FileCheck2 size={23} /><div><h3>{doc.report.issues.length ? '원본의 기존 오류가 남아 있습니다' : '새로 발생한 오류가 없습니다'}</h3><p>{saved ? '현재 버전을 다운로드할 수 있습니다.' : '다운로드하려면 검증하기를 실행하세요.'}</p></div></div><div className="check-list">{doc.report.checks.map(check => <div key={check}><CircleCheck size={16} /><span>{check}</span><b>완료</b></div>)}</div><p className="report-explanation">이 검사는 문서 구조를 확인합니다. 인쇄 배치는 한글 프로그램에서 확인하세요.</p>{doc.report.issues.length > 0 && <div className="issue-list"><h3>원본의 기존 오류 {doc.report.issues.length}개</h3>{doc.report.issues.map((issue, i) => <div key={i}><b>{issue.code || issue.kind}</b><p>{issue.message}</p><small>{issue.location}</small></div>)}</div>}{doc.report.repairs.length > 0 && <p className="repair-note">오래된 줄 배치 정보 {doc.report.repairs.length}건을 정리했습니다.</p>}<div className="modal-actions"><a className="outline-button" href={`/api/documents/${doc.id}/report`} download><ArrowDownToLine size={15} />검사 기록 JSON</a><button className="primary-button" onClick={() => setReportOpen(false)}>확인</button></div></section></div>}
    {privacyOpen && <ToolBoundary onClose={closePrivacy}><React.Suspense fallback={<div className="workbench-loading" role="status">보안 설정을 열고 있습니다.</div>}><PrivacyPanel dirty={!!dirty} onClose={closePrivacy} /></React.Suspense></ToolBoundary>}
    {helpOpen && <div className="modal-backdrop" onClick={() => setHelpOpen(false)}><section className="modal help-modal" role="dialog" aria-modal="true" aria-labelledby="help-title" onClick={e => e.stopPropagation()}><div className="modal-heading"><h2 id="help-title">사용 안내</h2><button className="icon-button" aria-label="사용 안내 닫기" onClick={() => setHelpOpen(false)}><X size={20} /></button></div><ol className="guide-steps"><li><b>문서 시작하기</b><p>내 문서의 새 문서에서 빈 문서·보고서·회의록·표 문서를 만들 수 있습니다. HWP 5 또는 HWPX 파일을 선택하거나 화면으로 끌어 놓으세요. 최대 20MB까지 지원합니다.</p></li><li><b>내용 수정 및 저장</b><p>문서 편집에서 원문 위를 클릭해 입력하고 글자·문단·표·그림·쪽 도구를 사용하세요. 서식 도구를 접거나 전체 화면을 켜면 문서를 더 넓게 볼 수 있습니다. 문서 목록은 왼쪽 위 버튼으로 엽니다. 텍스트 편집은 본문 이어 보기가 기본입니다. 문단 목록에서 번호별로 확인할 수도 있습니다. 내용을 바꾼 뒤 수정 저장(Ctrl/Cmd S)을 누르세요. 수정 내용에서 전후를 비교하거나 항목별로 되돌릴 수 있습니다. 미저장 텍스트 편집은 이 탭에, 원문 편집 임시본은 이 브라우저에 보관합니다. 같은 문서 버전을 다시 열면 복원할 수 있습니다. 임시보관은 파일 다운로드를 대신하지 않습니다. 그래프·그림 탭에서는 삽입된 개체를 따로 볼 수 있습니다. 복합 개체가 있는 문단과 셀은 읽기 전용입니다.</p></li><li><b>검증 및 다운로드</b><p>저장·검증 후 다운로드로 한 번에 완료하거나, 수정 저장과 검증을 각각 실행할 수 있습니다. 검증에 실패하면 다운로드를 중단합니다.</p></li><li><b>이전 버전 비교 및 복원</b><p>보기 및 도구의 버전 기록에서 저장한 내용을 비교하고 검색할 수 있습니다. 문서 이름을 클릭하거나 보기 및 도구의 이름 변경·복사에서 파일 이름을 바꾸고 독립된 복사본을 만들 수 있습니다. 작업 도구의 검토·출력에서 인쇄·PDF 저장을 열 수 있습니다. PDF는 브라우저 인쇄 창에서 PDF로 저장을 선택하세요. 이전 버전을 복원하면 현재 내용과 기록을 유지하면서 새 버전으로 저장합니다. 미저장 편집을 먼저 저장하고, 복원한 버전은 검증 후 다운로드하세요.</p></li></ol><p className="help-note">파일은 업로드 시 이 사이트의 서버로 전송됩니다. 서버의 원본·수정본·작업 메모는 업로드 후 24시간 동안 보관하며, 내 문서에서 삭제할 수 있습니다. 접근 권한은 이 브라우저의 세션 쿠키로 구분합니다. 브라우저에 저장된 원문 편집 임시본은 서버 보관 기간과 별개로 남을 수 있습니다. 문서 목록 메뉴의 임시 편집 보관함에서 내려받거나 새 문서로 복구·삭제할 수 있습니다. 공용 기기에서는 작업 후 다운로드하고 사이트 데이터를 지워 주세요. AI 수정이나 외부 AI 서비스 전송은 사용하지 않습니다.</p><button className="primary-button" onClick={() => setHelpOpen(false)}>확인</button></section></div>}
  </div>;
}

createRoot(document.getElementById('root')!).render(<React.StrictMode><ApplicationBoundary><PrivacyGate><App /></PrivacyGate></ApplicationBoundary></React.StrictMode>);
