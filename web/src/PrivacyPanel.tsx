import { useRef, useState } from 'react';
import { ShieldCheck, X } from 'lucide-react';
import { api } from './api';
import { clearPrivateStorage, lockWorkspace, workspaceContext } from './privacy-context';
import { useModalFocus } from './modal-focus';

export function PrivacyPanel({ dirty, onClose }: { dirty: boolean; onClose: () => void }) {
  const root = useRef<HTMLElement>(null);
  const [confirm, setConfirm] = useState(false), [busy, setBusy] = useState(false), [error, setError] = useState('');
  useModalFocus(root, () => { if (!busy) onClose(); });
  async function end() {
    if (busy || !confirm) return;
    setBusy(true); setError('');
    try {
      await api('/session', { method: 'DELETE' });
      try { const channel = new BroadcastChannel('jakal-private-session'); channel.postMessage({ ended: workspaceContext() }); channel.close(); } catch { /* Other tabs also check on focus. */ }
      lockWorkspace();
      await clearPrivateStorage();
    } catch (e) { setError(e instanceof Error ? e.message : '사용 종료를 완료하지 못했습니다.'); }
    finally { setBusy(false); }
  }
  return <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget && !busy) onClose(); }}><section className="modal privacy-panel" ref={root} role="dialog" aria-modal="true" aria-labelledby="privacy-title">
    <div className="modal-heading"><h2 id="privacy-title"><ShieldCheck size={20} />문서 보안·사용 종료</h2><button className="icon-button" disabled={busy} aria-label="문서 보안 닫기" onClick={onClose}><X size={18} /></button></div>
    <p>문서는 서버가 발급한 이 브라우저의 비공개 세션에 연결됩니다. 파일 주소만으로 다른 세션에서 열 수 없습니다. 편집 임시본과 배경 양식 보관함도 세션별로 구분합니다.</p>
    <p>사용 종료 시 이 세션의 문서 접근을 즉시 차단하고 서버 파일을 삭제 대상으로 지정합니다. 처리 중인 작업이 끝난 뒤 정리되며, 이 브라우저의 임시본·내 양식도 지웁니다. 다른 탭에도 종료를 알립니다.</p>
    {dirty && <p className="tool-error">미저장 편집이 있습니다. 필요한 파일을 먼저 저장하고 다운로드하세요.</p>}
    <p>같은 브라우저 프로필을 사용하는 사람은 사용 종료 전까지 같은 세션을 공유합니다. 공용 기기에서는 작업을 마치고 사용 종료를 눌러 주세요.</p>
    <label className="privacy-confirm"><input type="checkbox" checked={confirm} disabled={busy} onChange={e => setConfirm(e.target.checked)} />필요한 파일을 내려받았으며 이 브라우저의 임시본·내 양식 삭제에 동의합니다.</label>
    {error && <p className="tool-error" role="alert">{error}</p>}
    <div className="modal-actions"><button className="outline-button" disabled={busy} onClick={onClose}>편집 계속</button><button className="primary-button" disabled={busy || !confirm} onClick={() => void end()}>{busy ? '사용 종료 중…' : '문서 접근 종료·보관 데이터 삭제'}</button></div>
  </section></div>;
}
