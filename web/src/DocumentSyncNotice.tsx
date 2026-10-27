import { useEffect, useState } from 'react';
import type { Document } from './api';
import { ApiError, requestJson } from './request';

type State = Pick<Document, 'id' | 'revision' | 'settings_version' | 'name' | 'status'> & { pending_proposals?: number };
type Props = { doc: State; disabled: boolean; dirty: boolean; onReload: () => void; onBackup: () => void; onConnections: () => void };

/** Observe saved state only; never replace a live editor or its local drafts. */
export function DocumentSyncNotice({ doc, disabled, dirty, onReload, onBackup, onConnections }: Props) {
  const [changed, setChanged] = useState<State | 'missing' | null>(null);
  const [pending, setPending] = useState(0);
  useEffect(() => {
    setChanged(null);
    if (disabled) return;
    const abort = new AbortController();
    let reading = false;
    async function check() {
      if (reading || document.visibilityState === 'hidden' || !navigator.onLine) return;
      reading = true;
      try {
        const fresh = await requestJson<State>(`/api/documents/${doc.id}/state`, { signal: abort.signal, timeoutMs: 5_000 });
        if (!abort.signal.aborted && fresh.id === doc.id) {
          setChanged(fresh.revision !== doc.revision || fresh.settings_version !== doc.settings_version || fresh.name !== doc.name ? fresh : null);
          setPending(fresh.pending_proposals ?? 0);
        }
      } catch (error) {
        if (!abort.signal.aborted && error instanceof ApiError && error.status === 404) setChanged('missing');
      } finally { reading = false; }
    }
    void check();
    const interval = setInterval(() => void check(), 5000);
    window.addEventListener('focus', check); window.addEventListener('online', check);
    document.addEventListener('visibilitychange', check);
    return () => {
      abort.abort(); clearInterval(interval);
      window.removeEventListener('focus', check); window.removeEventListener('online', check);
      document.removeEventListener('visibilitychange', check);
    };
  }, [doc.id, doc.revision, doc.settings_version, doc.name, disabled]);
  if (!changed) return pending ? <div className="draft-conflict" role="status"><div><strong>외부 도구에서 수정 제안 {pending}개를 보냈습니다.</strong><span>변경 전후를 확인하고 적용할 항목을 선택할 수 있습니다.</span></div><button className="text-button" disabled={disabled} onClick={onConnections}>수정 제안 확인</button></div> : null;
  return <div className="draft-conflict" role="status"><div>
    <strong>{changed === 'missing' ? '서버의 문서를 찾을 수 없습니다.' : '다른 작업에서 문서나 설정이 변경되었습니다.'}</strong>
    <span>{changed === 'missing' ? (dirty ? '현재 편집은 유지됩니다. 미저장 임시본을 내려받아 별도로 보관해 주세요.' : '보관 기간이 만료되었거나 다른 탭에서 삭제했을 수 있습니다.')
      : dirty ? '현재 편집은 유지됩니다. 최신 문서를 열 때 이전 편집을 임시보관합니다. 임시본을 파일로 내려받을 수도 있습니다.' : '최신 문서를 다시 열면 변경 사항을 확인할 수 있습니다.'}</span>
  </div><div>{dirty && <button className="text-button" disabled={disabled} onClick={onBackup}>현재 편집 임시본 받기</button>}
    {changed !== 'missing' && <button className="text-button" disabled={disabled} onClick={onReload}>{dirty ? '편집 보관 후 최신 열기' : '최신 문서 열기'}</button>}</div></div>;
}
