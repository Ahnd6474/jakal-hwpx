import { useEffect, useRef, useState } from 'react';
import { ArrowDownToLine, FileText, RefreshCw, Trash2, X } from 'lucide-react';
import { studioBackup, type StudioBackupInfo } from './studio-backup';
import { useModalFocus } from './modal-focus';

export function BackupShelf({ activeId, names, onOpen, onClose }: {
  activeId?: string; names: Record<string, string>; onOpen: (file: File) => Promise<void>; onClose: () => void;
}) {
  const root = useRef<HTMLElement>(null);
  const [items, setItems] = useState<StudioBackupInfo[]>([]);
  const [loading, setLoading] = useState(true);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [reload, setReload] = useState(0);
  const [limit, setLimit] = useState(20);
  const [deleting, setDeleting] = useState('');
  useModalFocus(root, () => { if (!busy) onClose(); });
  useEffect(() => {
    let cancelled = false;
    setLoading(true); setError('');
    void studioBackup.list().then(value => { if (!cancelled) setItems(value); }).catch(() => { if (!cancelled) setError('브라우저 임시보관함을 읽지 못했습니다. 사이트 저장 공간과 다른 탭의 상태를 확인해 주세요.'); }).finally(() => { if (!cancelled) setLoading(false); });
    return () => { cancelled = true; };
  }, [reload]);
  const nameOf = (item: StudioBackupInfo) => item.name || names[item.id] || `이전 임시본_${item.id.slice(0, 8)}.hwpx`;
  async function act(item: StudioBackupInfo, action: 'download' | 'open' | 'delete') {
    if (busy) return;
    setBusy(true); setError('');
    try {
      if (action === 'delete') {
        if (item.id === activeId) throw new Error('현재 열린 문서의 임시본은 문서 안에서 관리해 주세요.');
        await studioBackup.clear(item.id); setDeleting(''); setItems(values => values.filter(value => value.id !== item.id)); return;
      }
      const stored = await studioBackup.read(item.id);
      if (!stored?.bytes?.byteLength) throw new Error('임시본이 변경되거나 삭제되었습니다. 목록을 새로 확인해 주세요.');
      const file = new File([new Uint8Array(stored.bytes)], `${nameOf(item).replace(/\.(hwp|hwpx)$/i, '')}_임시편집.hwpx`, { type: 'application/hwp+zip' });
      if (action === 'open') { await onOpen(file); return; }
      const url = URL.createObjectURL(file);
      const link = document.createElement('a'); link.href = url; link.download = file.name; link.click();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
    } catch (e) { setError(e instanceof Error ? e.message : '임시본 작업을 완료하지 못했습니다.'); }
    finally { setBusy(false); }
  }
  return <div className="modal-backdrop" onMouseDown={e => { if (e.target === e.currentTarget && !busy) onClose(); }}><section ref={root} className="modal backup-shelf" role="dialog" aria-modal="true" aria-labelledby="backup-shelf-title" aria-busy={busy || loading}>
    <header className="modal-heading"><div><h2 id="backup-shelf-title">임시 편집 보관함</h2><p>이 브라우저에 남아 있는 원문 편집 임시본입니다.</p></div><button className="icon-button" aria-label="임시 편집 보관함 닫기" disabled={busy} onClick={onClose}><X size={20} /></button></header>
    <p className="document-action-note">서버 보관 기간이 지났거나 세션이 바뀌어도 임시본을 내려받을 수 있습니다. 새 문서로 열면 서버에서 다시 분석하며 이전 문서를 덮어쓰지 않습니다. 임시본은 마지막 저장 이후의 모든 입력을 포함하지 않을 수 있습니다.</p>
    <button className="text-button" disabled={busy || loading} onClick={() => setReload(value => value + 1)}><RefreshCw size={14} />목록 새로 확인</button>
    {error && <p className="tool-error" role="alert">{error}</p>}
    {loading ? <p role="status" className="backup-shelf-empty">임시보관함을 확인하고 있습니다.</p> : !items.length ? <p className="backup-shelf-empty">원문 편집 임시본이 없습니다. 텍스트 편집의 임시본은 해당 탭에서 문서를 다시 열어 확인할 수 있습니다.</p> : <ul className="backup-shelf-list">{items.slice(0, limit).map(item => <li key={item.id}>
      <div className="backup-shelf-name"><FileText size={17} /><div><strong>{nameOf(item)}</strong><small>{item.saved ? new Date(item.saved).toLocaleString('ko-KR') : '이전 버전에서 보관한 임시본'}{item.size ? ` · ${(item.size / 1024).toLocaleString('ko-KR', { maximumFractionDigits: 0 })}KB` : ''}{item.id === activeId ? ' · 현재 문서' : ''}</small></div></div>
      <div className="backup-shelf-actions"><button className="text-button" disabled={busy} onClick={() => void act(item, 'download')}><ArrowDownToLine size={14} />임시본 받기</button><button className="text-button" disabled={busy} onClick={() => void act(item, 'open')}>새 문서로 열기</button>{deleting === item.id ? <><span>삭제할까요?</span><button className="text-button" disabled={busy} onClick={() => void act(item, 'delete')}>삭제 확인</button><button className="text-button" disabled={busy} onClick={() => setDeleting('')}>취소</button></> : <button className="icon-button" disabled={busy || item.id === activeId} aria-label={`${nameOf(item)} 임시본 삭제`} onClick={() => setDeleting(item.id)}><Trash2 size={14} /></button>}</div>
    </li>)}</ul>}
    {items.length > limit && <button className="outline-button" disabled={busy} onClick={() => setLimit(value => value + 20)}>임시본 더 보기</button>}
    <footer className="modal-actions"><button className="primary-button" disabled={busy} onClick={onClose}>닫기</button></footer>
  </section></div>;
}
