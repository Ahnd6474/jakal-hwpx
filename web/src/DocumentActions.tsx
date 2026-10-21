import { useRef, useState } from 'react';
import { Copy, Pencil, X } from 'lucide-react';
import type { Document } from './api';
import { useModalFocus } from './modal-focus';

export function DocumentActions({ doc, initialMode, dirty, onClose, onSubmit }: {
  doc: Document; initialMode: 'rename' | 'copy'; dirty: boolean; onClose: () => void;
  onSubmit: (mode: 'rename' | 'copy', name: string) => Promise<void>;
}) {
  const root = useRef<HTMLElement>(null);
  const [mode, setMode] = useState(initialMode);
  const stem = doc.name.replace(/\.(hwp|hwpx)$/i, '');
  const [name, setName] = useState(initialMode === 'copy' ? `${stem}_복사본`.slice(0, 150) : stem);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  useModalFocus(root, () => { if (!busy) onClose(); });
  const extension = mode === 'copy' ? 'hwpx' : doc.name.split('.').at(-1);
  async function submit() {
    if (busy) return;
    if (!name.trim() || /[\\/:*?"<>|\x00-\x1f\x7f]/.test(name)) { setError('사용할 수 있는 파일 이름을 입력해 주세요.'); return; }
    setBusy(true); setError('');
    try { await onSubmit(mode, name.trim()); }
    catch (e) { setError(e instanceof Error ? e.message : '문서 작업을 완료하지 못했습니다.'); }
    finally { setBusy(false); }
  }
  return <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget && !busy) onClose(); }}><section ref={root} className="modal document-actions-modal" role="dialog" aria-modal="true" aria-labelledby="document-actions-title" aria-busy={busy}>
    <header className="modal-heading"><h2 id="document-actions-title">문서 파일 관리</h2><button className="icon-button" aria-label="문서 파일 관리 닫기" disabled={busy} onClick={onClose}><X size={20} /></button></header>
    <div className="document-action-options" role="group" aria-label="파일 작업">
      <button className="outline-button" aria-pressed={mode === 'rename'} disabled={busy} onClick={() => { setMode('rename'); setName(stem); setError(''); }}><Pencil size={15} />이름 변경</button>
      <button className="outline-button" aria-pressed={mode === 'copy'} disabled={busy} onClick={() => { setMode('copy'); setName(`${stem}_복사본`.slice(0, 150)); setError(''); }}><Copy size={15} />복사본 만들기</button>
    </div>
    <form onSubmit={e => { e.preventDefault(); void submit(); }}>
      <label className="creation-field">파일 이름<span className="document-name-field"><input data-autofocus required maxLength={160} disabled={busy} value={name} onChange={e => setName(e.target.value)} /><span>.{extension}</span></span></label>
      <p className="document-action-note">{mode === 'rename' ? '문서 목록과 다운로드 파일의 이름을 변경합니다. 본문 제목과 편집 내용은 그대로 유지합니다.' : dirty ? '먼저 현재 문서의 수정 내용을 저장한 다음, 그 버전으로 독립된 HWPX 복사본을 만듭니다. 현재 문서에도 수정 내용이 저장됩니다.' : '현재 저장 버전으로 독립된 HWPX 문서를 만듭니다. 원래 문서의 파일과 버전 기록은 유지합니다. 편집 규칙도 함께 복사합니다.'}</p>
      {error && <p className="tool-error" role="alert">{error}</p>}
      <div className="modal-actions"><button type="button" className="outline-button" disabled={busy} onClick={onClose}>취소</button><button className="primary-button" disabled={busy || !name.trim()}>{busy ? '처리 중…' : mode === 'rename' ? '이름 변경' : dirty ? '현재 문서 저장 후 복사' : '복사본 만들기'}</button></div>
    </form>
  </section></div>;
}
