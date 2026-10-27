import { useEffect, useRef, useState, useCallback } from 'react';
import { FileText, LoaderCircle, X } from 'lucide-react';
import { api } from './api';
import { useModalFocus } from './modal-focus';

export type Template = { id: string; name: string; description: string; title: string };
export type TemplateInput = { template: string; title: string };
export function NewDocument({ busy, error, onClose, onCreate }: { busy: boolean; error: string; onClose: () => void; onCreate: (value: TemplateInput) => void }) {
  const [templates, setTemplates] = useState<Template[]>([]);
  const [kind, setKind] = useState('blank');
  const [title, setTitle] = useState('새 문서');
  const [failure, setFailure] = useState('');
  const root = useRef<HTMLElement>(null);
  const close = useCallback(() => { if (!busy) onClose(); }, [busy, onClose]);
  useModalFocus(root, close);
  useEffect(() => {
    const abort = new AbortController();
    void api<Template[]>('/templates', { signal: abort.signal }).then(setTemplates).catch(e => { if (!abort.signal.aborted) setFailure(e.message); });
    return () => abort.abort();
  }, []);
  return <div className="modal-backdrop creation-backdrop" onMouseDown={e => { if (e.target === e.currentTarget) close(); }}><section ref={root} className="modal creation-modal" role="dialog" aria-modal="true" aria-labelledby="creation-title">
    <div className="modal-heading"><div><h2 id="creation-title">새 문서</h2><p>기본 양식으로 시작하고 문서 위에서 자유롭게 편집하세요.</p></div><button className="icon-button" aria-label="새 문서 닫기" disabled={busy} onClick={close}><X size={19} /></button></div>
    <div className="template-grid" role="group" aria-label="문서 양식">{templates.map(item => <button key={item.id} disabled={busy} aria-pressed={kind === item.id} className="template-card" onClick={() => { setKind(item.id); setTitle(item.title); }}><span className={`template-paper template-${item.id}`} aria-hidden="true"><FileText size={30} strokeWidth={1.2} /><i /><i /><i /></span><strong>{item.name}</strong><small>{item.description}</small></button>)}</div>
    {!templates.length && !failure && <p role="status">양식을 불러오고 있습니다.</p>}
    <form onSubmit={e => { e.preventDefault(); if (title.trim() && templates.length && !busy) onCreate({ template: kind, title: title.trim() }); }}>
      <label className="creation-field">문서 제목<input data-autofocus required maxLength={100} disabled={busy} value={title} onChange={e => setTitle(e.target.value)} /></label>
      {(error || failure) && <p className="tool-error" role="alert">{error || failure}</p>}
      <div className="creation-footer"><span>HWPX · A4 · 내용 직접 편집</span><button className="primary-button" disabled={busy || !title.trim() || !templates.length}>{busy && <LoaderCircle size={15} className="spin" />}{busy ? '문서 준비 중' : '만들기'}</button></div>
    </form>
  </section></div>;
}
