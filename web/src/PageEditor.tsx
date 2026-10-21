import { forwardRef, lazy, memo, Suspense, useCallback, useEffect, useImperativeHandle, useRef, useState } from 'react';
import { createStudio, type RhwpEditor } from '@rhwp/editor';
import { ChevronDown, ChevronUp, LoaderCircle, Maximize2, Minimize2, Minus, Plus, Save, Search } from 'lucide-react';
const EditorWorkbench = lazy(() => import('./EditorWorkbench').then(module => ({ default: module.EditorWorkbench })));
import type { Document } from './api';
import { studioBackup, type StudioBackup } from './studio-backup';
import { IdleBackup } from './idle-backup';
import { ToolBoundary } from './ToolBoundary';
import { requestBytes } from './request';

export type PageEditorHandle = {
  export: () => Promise<Uint8Array>;
  saved: (revision: number) => Promise<void>;
};
type Props = { doc: Document; active: boolean; disabled: boolean; onDirty: (dirty: boolean) => void; onSave: () => void; onFileAction: (action: string) => void };

export const PageEditor = memo(forwardRef<PageEditorHandle, Props>(function PageEditor({ doc, active, disabled, onDirty, onSave, onFileAction }, ref) {
  const container = useRef<HTMLDivElement>(null);
  const shell = useRef<HTMLDivElement>(null);
  const editor = useRef<RhwpEditor | null>(null);
  const current = useRef({ doc, onDirty, onSave, onFileAction });
  current.current = { doc, onDirty, onSave, onFileAction };
  const base = useRef(doc.revision);
  const sequence = useRef(0);
  const exported = useRef('');
  const dirty = useRef(false);
  const loading = useRef(true);
  const autosave = useRef<IdleBackup | null>(null);
  const generation = useRef(0);
  const [state, setState] = useState('원래 문서 배치로 편집기를 열고 있습니다');
  const [error, setError] = useState('');
  const [retry, setRetry] = useState(0);
  const [backup, setBackup] = useState<StudioBackup | null>(null);
  const [backupStatus, setBackupStatus] = useState('');
  const [tools, setTools] = useState<'essential' | 'full' | 'focus'>('essential');
  const [ready, setReady] = useState(false);
  const [workbench, setWorkbench] = useState(false);
  const toolRunning = useRef(false);
  const closeWorkbench = useCallback(() => setWorkbench(false), []);
  const listCommands = useCallback(async () => editor.current ? editor.current.commands.list() : [], []);
  const availability = useRef({ active, disabled, ready, backup });
  availability.current = { active, disabled, ready, backup };
  useEffect(() => {
    const key = (event: KeyboardEvent) => {
      const value = availability.current;
      if ((event.ctrlKey || event.metaKey) && event.shiftKey && !event.altKey && event.code === 'KeyK' && value.active && value.ready && !value.disabled && !value.backup && !document.querySelector('[role="dialog"]')) {
        event.preventDefault(); setWorkbench(true);
      }
    };
    window.addEventListener('keydown', key);
    return () => window.removeEventListener('keydown', key);
  }, []);
  useEffect(() => { if (!active || disabled) setWorkbench(false); }, [active, disabled]);
  async function runTool(id: string, params?: Record<string, unknown>) {
    const instance = editor.current;
    const value = availability.current;
    if (!instance || !value.active || !value.ready || value.disabled || value.backup || toolRunning.current) throw new Error('문서 준비를 마친 후 사용해 주세요.');
    toolRunning.current = true;
    let panelClosed = false;
    try {
      // Recheck at execution: enabled state can change after opening the panel.
      if (!await instance.commands.isEnabled(id)) throw new Error('현재 커서 위치에서는 사용할 수 없습니다. 본문을 클릭한 뒤 다시 열어 주세요.');
      setWorkbench(false);
      panelClosed = true;
      await new Promise<void>(resolve => requestAnimationFrame(() => resolve()));
      if (editor.current !== instance || availability.current.disabled || loading.current) throw new Error('문서 상태가 바뀌었습니다. 준비가 끝난 뒤 다시 사용해 주세요.');
      const result = await instance.commands.execute(id, params, { allowDialog: true });
      if (!result.ok) throw new Error(result.message || '도구를 실행하지 못했습니다. 커서 위치를 확인해 주세요.');
      instance.element.contentWindow?.focus();
    } catch (e) { if (!panelClosed) throw e; setError(e instanceof Error ? e.message : '도구 실행에 실패했습니다.'); }
    finally { toolRunning.current = false; }
  }

  async function digest(bytes: Uint8Array) {
    const result = await crypto.subtle.digest('SHA-256', new Uint8Array(bytes));
    return Array.from(new Uint8Array(result), value => value.toString(16).padStart(2, '0')).join('');
  }

  async function storeBackup(isCurrent: () => boolean) {
    const instance = editor.current;
    const epoch = generation.current;
    if (!instance || !dirty.current || loading.current) return;
    const at = sequence.current;
    const began = performance.now();
    try {
      shell.current?.setAttribute('data-backup-count', String(Number(shell.current?.getAttribute('data-backup-count') || '0') + 1));
      const bytes = await instance.exportHwpx();
      if (epoch !== generation.current || !dirty.current || at !== sequence.current || !isCurrent()) return;
      if (bytes.byteLength > 20 * 1024 * 1024) throw new Error('size');
      await studioBackup.write({ id: current.current.doc.id, revision: base.current, bytes, saved: Date.now(), name: current.current.doc.name });
      shell.current?.setAttribute('data-backup-ms', String(Math.round(performance.now() - began)));
      if (epoch === generation.current && dirty.current && at === sequence.current && isCurrent()) setBackupStatus('이 브라우저에 임시보관됨');
    } catch { if (epoch === generation.current) setBackupStatus('임시보관 실패 · 수정 저장 필요'); }
  }
  useImperativeHandle(ref, () => ({
    async export() {
      if (!editor.current || loading.current || backup) throw new Error('문서 편집기 준비를 마친 뒤 다시 시도하세요.');
      autosave.current?.cancel();
      editor.current.element.inert = true;
      const bytes = await editor.current.exportHwpx();
      exported.current = await digest(bytes);
      return bytes;
    },
    async saved(revision) {
      // Export flushes pending IME/table layout and emits render events. Compare
      // actual bytes rather than mistaking those notifications for new input.
      if (!editor.current || exported.current !== await digest(await editor.current.exportHwpx())) throw new Error('저장 중 문서가 갱신되었습니다. 최신 편집을 다시 저장해 주세요.');
      await editor.current?.notifySaved();
      base.current = revision;
      dirty.current = false; current.current.onDirty(false);
      autosave.current?.cancel();
      await studioBackup.clear(current.current.doc.id).catch(() => {});
      setBackupStatus('');
    },
  }), [backup]);

  useEffect(() => {
    let cancelled = false;
    const epoch = ++generation.current;
    const abort = new AbortController();
    const scheduler = new IdleBackup(storeBackup);
    autosave.current = scheduler;
    let instance: RhwpEditor | null = null;
    loading.current = true; dirty.current = false; setReady(false); setError('');
    shell.current?.setAttribute('data-backup-count', '0');
    setState('원래 문서 배치로 편집기를 열고 있습니다');
    const host = container.current!;
    // Detached cleanup is important when StrictMode remounts during startup.
    const mount = document.createElement('div'); mount.className = 'page-editor-mount'; host.appendChild(mount);
    const initialization = async () => {
      try {
        const began = performance.now();
        const source = requestBytes(`/api/documents/${doc.id}/source?revision=${doc.revision}`, { signal: abort.signal });
        const recovered = studioBackup.read(doc.id).catch(() => undefined);
        const studioUrl = new URL('/studio/?chrome=embed', location.origin);
        if (new URLSearchParams(location.search).has('profile-input')) studioUrl.searchParams.set('profile-input', '1');
        const created = createStudio(mount, { studioUrl: studioUrl.href, renderer: 'canvas2d', requestTimeoutMs: 60_000 }).then(value => {
          instance = value;
          if (cancelled) value.destroy();
          return value;
        });
        const [runtime, bytes, stored] = await Promise.all([created, source, recovered]);
        instance = runtime;
        if (cancelled) { instance.destroy(); return; }
        instance.element.title = '원문 서식 문서 편집기';
        editor.current = instance;
        await instance.loadFile(bytes, doc.name.replace(/\.hwp$/i, '.hwpx'));
        if (cancelled) return;
        if (window.matchMedia('(max-width: 760px)').matches) await instance.commands.execute('view:zoom-fit-width');
        base.current = doc.revision; loading.current = false; setState(''); setReady(true);
        shell.current?.setAttribute('data-load-ms', String(Math.round(performance.now() - began)));
        if (stored && !cancelled) setBackup(stored);
      } catch (e) { if (!cancelled) { setError(e instanceof Error ? e.message : '편집기를 열지 못했습니다.'); setState(''); } }
    };
    function message(event: MessageEvent) {
      if (event.origin !== location.origin || event.source !== instance?.element.contentWindow || event.data?.channel !== 'jakal-studio-v1') return;
      if (event.data.type === 'file-action') {
        if (availability.current.active && !availability.current.disabled && !document.querySelector('[role="dialog"]') && typeof event.data.action === 'string') current.current.onFileAction(event.data.action);
        return;
      }
      if (event.data.type === 'save') { current.current.onSave(); return; }
      if (event.data.type === 'workbench') { const value = availability.current; if (value.active && value.ready && !value.disabled && !value.backup && !document.querySelector('[role="dialog"]')) setWorkbench(true); return; }
      if (event.data.type === 'composition') { scheduler.setComposing(Boolean(event.data.composing)); return; }
      if (loading.current || cancelled) return;
      if (event.data.type === 'changed') sequence.current = event.data.sequence;
      const next = Boolean(event.data.dirty);
      if (dirty.current !== next) { dirty.current = next; current.current.onDirty(next); }
      if (next) {
        setBackupStatus('편집 중 · 잠시 후 임시보관'); scheduler.schedule();
      }
    }
    const hidden = () => { if (document.hidden) void scheduler.flush(); };
    window.addEventListener('message', message);
    document.addEventListener('visibilitychange', hidden);
    void initialization();
    return () => { cancelled = true; generation.current++; abort.abort(); scheduler.cancel(); document.removeEventListener('visibilitychange', hidden); window.removeEventListener('message', message); instance?.destroy(); mount.remove(); if (editor.current === instance) editor.current = null; };
  }, [doc.id, retry]);

  useEffect(() => {
    if (doc.revision !== base.current && editor.current && !dirty.current) {
      // A text-list edit was saved while the visual editor was hidden.
      setRetry(n => n + 1);
    }
  }, [doc.revision]);
  useEffect(() => { if (editor.current) editor.current.element.inert = disabled || !!backup || workbench; }, [disabled, backup, ready, workbench]);
  useEffect(() => {
    if (ready) editor.current?.element.contentWindow?.postMessage({ channel: 'jakal-studio-v1', type: 'tools', mode: tools }, location.origin);
  }, [tools, ready]);
  useEffect(() => {
    if (ready) editor.current?.element.contentWindow?.postMessage({ channel: 'jakal-studio-v1', type: 'filename', name: doc.name }, location.origin);
  }, [doc.name, ready]);
  useEffect(() => {
    const narrow = window.matchMedia('(max-width: 760px)');
    const resize = () => { if (narrow.matches && editor.current) void editor.current.commands.execute('view:zoom-fit-width').catch(() => {}); };
    narrow.addEventListener('change', resize);
    return () => narrow.removeEventListener('change', resize);
  }, []);
  async function restore() {
    if (!backup || !editor.current) return;
    try {
      loading.current = true; setReady(false); setState('임시 편집을 복원하고 있습니다');
      editor.current.element.inert = true;
      await editor.current.loadFile(new Uint8Array(backup.bytes).buffer, doc.name.replace(/\.hwp$/i, '.hwpx'));
      dirty.current = true; sequence.current++; current.current.onDirty(true); setBackup(null); setBackupStatus('임시 편집 복원됨 · 수정 저장 필요');
    } catch (e) { setError(e instanceof Error ? e.message : '임시 편집을 복원하지 못했습니다.'); }
    finally { loading.current = false; setState(''); setReady(true); }
  }
  function downloadBackup() {
    if (!backup) return;
    const url = URL.createObjectURL(new Blob([new Uint8Array(backup.bytes)], { type: 'application/hwp+zip' }));
    const a = document.createElement('a'); a.href = url; a.download = `${doc.name.replace(/\.[^.]+$/, '')}_임시편집.hwpx`; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 30_000);
  }
  async function fullscreen() {
    try { if (document.fullscreenElement) await document.exitFullscreen(); else await shell.current?.requestFullscreen(); }
    catch { setError('이 브라우저에서는 전체 화면을 사용할 수 없습니다.'); }
  }
  async function zoom(command: string) {
    try { await editor.current?.commands.execute(command); }
    catch { setError('확대 비율을 바꾸지 못했습니다.'); }
  }
  return <div ref={shell} className="page-editor" hidden={!active}>
    {workbench && <ToolBoundary onClose={closeWorkbench}><Suspense fallback={<div className="workbench-loading" role="status">작업 도구를 열고 있습니다.</div>}><EditorWorkbench onClose={closeWorkbench} commands={listCommands} onRun={id => runTool(id)} /></Suspense></ToolBoundary>}
    {backup && <div className="studio-recovery" role="status"><strong>{backup.revision === doc.revision ? '저장하지 않은 편집이 있습니다.' : '이전 버전의 임시 편집이 있습니다.'}</strong><span>{backup.revision === doc.revision ? '복원하면 마지막 편집을 이어갈 수 있습니다.' : '최신 문서와 충돌하므로 파일로 내려받아 확인해 주세요.'}</span><div>{backup.revision === doc.revision && <button className="primary-button" onClick={() => void restore()}>임시 편집 복원</button>}<button className="outline-button" onClick={downloadBackup}>임시본 받기</button><button className="text-button" onClick={() => { void studioBackup.clear(doc.id).then(() => setBackup(null)).catch(() => setError('임시본을 지우지 못했습니다. 브라우저 저장 공간을 확인해 주세요.')); }}>임시본 버리고 계속</button></div></div>}
    {error && <div className="studio-error" role="alert">{error}{!ready && <button className="outline-button" onClick={() => setRetry(n => n + 1)}>다시 시도</button>}<button className="icon-button" aria-label="편집기 알림 닫기" onClick={() => setError('')}>×</button></div>}
    <div className="page-editor-host" ref={container} />
    {state && <div className="studio-loading" role="status"><LoaderCircle size={24} className="spin" /><p>{state}</p></div>}
    <div className="page-editor-controls"><span className="studio-save-state" title="원본 별도 보관 · 일부 글꼴과 배치는 한글과 다를 수 있습니다."><i className={dirty.current ? 'pending' : ''} />{backupStatus || (dirty.current ? '저장 전' : '저장됨')}</span><div className="studio-zoom"><button className="icon-button" aria-label="편집 문서 축소" disabled={!ready || disabled} onClick={() => void zoom('view:zoom-out')}><Minus size={14} /></button><button className="text-button" aria-label="편집 문서 너비 맞춤" disabled={!ready || disabled} onClick={() => void zoom('view:zoom-fit-width')}>너비 맞춤</button><button className="icon-button" aria-label="편집 문서 확대" disabled={!ready || disabled} onClick={() => void zoom('view:zoom-in')}><Plus size={14} /></button></div><div className="studio-display-controls"><button className="text-button workbench-trigger" aria-label="작업 도구 열기" title="작업 도구 (Ctrl+Shift+K)" disabled={!ready || disabled || !!backup} onClick={() => setWorkbench(true)}><Search size={14} /><span>작업 도구</span></button><button className="text-button all-tools-toggle" aria-pressed={tools === 'full'} disabled={!ready || disabled} onClick={() => setTools(tools === 'full' ? 'essential' : 'full')}>{tools === 'full' ? <ChevronUp size={13} /> : <ChevronDown size={13} />}<span>{tools === 'full' ? '도구 줄이기' : '모든 도구'}</span></button><button className="text-button focus-mode-toggle" aria-label={tools === 'focus' ? '편집 도구 복귀' : '집중 보기'} aria-pressed={tools === 'focus'} disabled={!ready || disabled} onClick={() => setTools(tools === 'focus' ? 'essential' : 'focus')}><Minimize2 size={14} /><span>{tools === 'focus' ? '편집 도구 복귀' : '집중 보기'}</span></button><button className="fullscreen-save icon-button" aria-label="전체 화면 수정 저장" disabled={!dirty.current || disabled || !ready} onClick={onSave}><Save size={14} /></button><button className="icon-button" aria-label="편집기 전체 화면" onClick={() => void fullscreen()}><Maximize2 size={15} /></button></div></div>
  </div>;
}));
