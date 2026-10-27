import { useEffect, useRef, useState } from 'react';
import { Copy, Link2, LoaderCircle, RefreshCw, Unplug, X } from 'lucide-react';
import { api, post, waitForJob, type Document } from './api';
import { useModalFocus } from './modal-focus';

type Connection = { id: string; label: string; permission: 'read' | 'propose' | 'write'; expires: number; revoked: number; last_used: number | null };
type Proposal = { id: string; connection: string; label: string; title: string; revision: number; settings_version: number; status: 'pending' | 'accepted' | 'rejected'; job_id: string | null };
type Event = { action: string; label: string; created: number; job_id: string | null };
type Listing = { connections: Connection[]; proposals: Proposal[]; events: Event[]; state?: { revision: number; settings_version: number }; editing?: { tab: string; dirty: boolean; last_seen: number }[] };
type Review = Pick<Proposal, 'id' | 'title' | 'revision' | 'settings_version' | 'status'> & { changes: { id: string; label: string; before: string; after: string }[] };
type Issued = { id: string; token: string; expires: number; http_config: unknown; stdio_config: unknown | null };
type Props = { doc: Document; dirty: boolean; onClose: () => void; onApplied: () => Promise<void>; initialTab?: 'connections' | 'proposals' };
const permissionLabels = { read: '읽기', propose: '수정 제안', write: '직접 수정' };
const actions: Record<string, string> = { connected: '연결 발급', revoked: '연결 해제', proposed: '수정 제안', approved: '제안 적용 요청', edited: '직접 수정 요청', validated: '검증 요청', rejected: '제안 거절' };
const date = (seconds: number) => new Date(seconds * 1000).toLocaleString('ko-KR', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' });

export function ExternalConnections({ doc, dirty, onClose, onApplied, initialTab = 'connections' }: Props) {
  const root = useRef<HTMLElement>(null);
  const locked = useRef(false);
  const alive = useRef(true);
  const [data, setData] = useState<Listing>({ connections: [], proposals: [], events: [] });
  const [tab, setTab] = useState<'connections' | 'proposals' | 'events'>(initialTab);
  const [label, setLabel] = useState('외부 도구');
  const [permission, setPermission] = useState<Connection['permission']>('propose');
  const [hours, setHours] = useState(1);
  const [writeConsent, setWriteConsent] = useState(false);
  const [issued, setIssued] = useState<Issued | null>(null);
  const [transport, setTransport] = useState<'stdio' | 'http'>('http');
  const [review, setReview] = useState<Review | null>(null);
  const [selected, setSelected] = useState<string[]>([]);
  const [busy, setBusy] = useState('');
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');
  const [confirmRevoke, setConfirmRevoke] = useState<string | null>(null);
  const [unlock, setUnlock] = useState<string | null>(null);
  useModalFocus(root, () => { if (!locked.current) onClose(); });
  useEffect(() => { alive.current = true; return () => { alive.current = false; }; }, []);

  async function refresh(signal?: AbortSignal) {
    const value = await api<Listing>(`/documents/${doc.id}/connections`, { signal });
    if (alive.current && !signal?.aborted) setData(value);
  }
  useEffect(() => {
    const abort = new AbortController();
    let reading = false;
    const check = async () => {
      if (reading || locked.current || document.visibilityState === 'hidden') return;
      reading = true;
      try { await refresh(abort.signal); }
      catch (e) { if (!abort.signal.aborted) setError(e instanceof Error ? e.message : '연결을 불러오지 못했습니다.'); }
      finally { reading = false; }
    };
    void check();
    const timer = setInterval(() => void check(), 5000);
    return () => { abort.abort(); clearInterval(timer); };
  }, [doc.id]);

  async function run(message: string, action: () => Promise<void>) {
    if (locked.current) return;
    locked.current = true; setBusy(message); setError(''); setNotice('');
    try { await action(); }
    catch (e) { if (alive.current) setError(e instanceof Error ? e.message : '요청을 처리하지 못했습니다.'); }
    finally { locked.current = false; if (alive.current) setBusy(''); }
  }
  async function create() {
    if (permission === 'write' && (dirty || !writeConsent)) return;
    await run('연결 설정을 만들고 있습니다', async () => {
      const value = await post<Issued>('/connections', { document_id: doc.id, label: label.trim(), permission, hours });
      setIssued(value); setTransport(value.stdio_config ? 'stdio' : 'http'); await refresh();
    });
  }
  async function inspect(item: Proposal) {
    await run('제안을 불러오고 있습니다', async () => {
      const value = await api<Review>(`/proposals/${item.id}`);
      setReview(value); setSelected(value.changes.map(change => change.id));
    });
  }
  async function approve() {
    if (!review || dirty || !selected.length) return;
    await run('선택한 수정을 적용하고 검증하고 있습니다', async () => {
      const result = await post<{ job_id: string }>(`/proposals/${review.id}/approve`, { revision: review.revision, settings_version: review.settings_version, targets: selected });
      await waitForJob(result.job_id, () => {});
      await onApplied(); setReview(null); await refresh();
      setNotice('선택한 수정을 새 버전으로 저장했습니다. 버전 기록에서 이전 내용으로 복원할 수 있습니다.');
    });
  }
  const config = issued ? JSON.stringify(transport === 'stdio' ? issued.stdio_config : issued.http_config, null, 2) : '';
  const stale = !!review && (review.revision !== (data.state?.revision ?? doc.revision) || review.settings_version !== (data.state?.settings_version ?? doc.settings_version));
  return <div className="modal-backdrop external-backdrop" onMouseDown={e => { if (e.target === e.currentTarget && !busy) onClose(); }}>
    <section className="modal external-modal" ref={root} role="dialog" aria-modal="true" aria-labelledby="external-title">
      <div className="modal-heading"><div><h2 id="external-title"><Link2 size={19} />외부 도구 연결</h2><p title={doc.name}>{doc.name} · 저장 버전 {doc.revision + 1}</p></div><button className="icon-button" disabled={!!busy} aria-label="외부 도구 연결 닫기" onClick={onClose}><X size={20} /></button></div>
      <div className="external-tabs" role="tablist" aria-label="외부 연결 관리" onKeyDown={event => {
        const keys = ['connections', 'proposals', 'events'] as const;
        if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
        event.preventDefault();
        const index = event.key === 'Home' ? 0 : event.key === 'End' ? 2 : (keys.indexOf(tab) + (event.key === 'ArrowLeft' ? 2 : 1)) % 3;
        setTab(keys[index]); document.getElementById(`external-tab-${keys[index]}`)?.focus();
      }}>
        {([['connections', '연결'], ['proposals', `수정 제안 ${data.proposals.filter(p => p.status === 'pending').length}`], ['events', '작업 기록']] as const).map(([key, title]) => <button key={key} id={`external-tab-${key}`} role="tab" tabIndex={tab === key ? 0 : -1} aria-selected={tab === key} aria-controls={`external-panel-${key}`} onClick={() => setTab(key)}>{title}</button>)}
        <button className="icon-button" disabled={!!busy} aria-label="외부 연결 새로고침" onClick={() => void run('새로고침 중', () => refresh())}><RefreshCw size={15} /></button>
      </div>
      {error && <p className="external-message tool-error" role="alert">{error}</p>}{notice && <p className="external-message" role="status">{notice}</p>}
      {busy && <p className="external-message" role="status"><LoaderCircle className="spin" size={14} />{busy}</p>}
      {!!data.editing?.length && <div className="external-editing"><p>저장 전 편집이 있는 탭 {data.editing.length}개 · 외부 수정은 저장 후 진행됩니다.</p>{data.editing.filter(item => item.last_seen * 1000 < Date.now() - 90_000).map(item => <div key={item.tab}>{unlock === item.tab ? <><span>닫힌 탭인지 확인하세요. 임시본은 유지되지만 외부 수정이 가능해집니다.</span><button className="text-button" disabled={!!busy} onClick={() => setUnlock(null)}>취소</button><button className="outline-button" disabled={!!busy} onClick={() => void run('닫힌 탭의 잠금을 해제하고 있습니다', async () => { await api(`/documents/${doc.id}/presence/${item.tab}`, { method: 'DELETE' }); setUnlock(null); await refresh(); })}>잠금 해제 확인</button></> : <button className="text-button" disabled={!!busy} onClick={() => setUnlock(item.tab)}>응답 없는 탭의 잠금 해제</button>}</div>)}</div>}
      {tab === 'connections' && <div className="external-body" id="external-panel-connections" role="tabpanel" aria-labelledby="external-tab-connections">
        <p className="external-intro">이 문서만 외부 프로그램에 연결합니다. 연결된 프로그램은 허용한 내용을 읽을 수 있습니다. 원문 보기와 웹 편집은 계속 사용할 수 있습니다.</p>
        {issued ? <div className="external-config"><div><h3>연결 설정이 준비됐습니다</h3><small>{date(issued.expires)} 만료 · 이 창을 닫으면 토큰을 다시 표시하지 않습니다.</small></div>
          <label>연결 방식<select value={transport} onChange={e => setTransport(e.target.value as typeof transport)}>{!!issued.stdio_config && <option value="stdio">로컬 실행 · stdio</option>}<option value="http">사이트 주소 · Streamable HTTP</option></select></label>
          <p>{transport === 'stdio' ? 'MCP 클라이언트의 서버 설정에 아래 내용을 등록하세요. 편집기는 실행된 상태여야 합니다.' : 'URL과 Authorization 헤더를 지원하는 MCP 클라이언트에 등록하세요. 이 버전은 별도 토큰 방식이며 OAuth 자동 로그인은 제공하지 않습니다.'}</p>
          <textarea aria-label="MCP 연결 설정" value={config} readOnly spellCheck={false} onFocus={e => e.currentTarget.select()} />
          <div className="external-buttons"><button className="outline-button" onClick={() => void run('복사 중', async () => { await navigator.clipboard.writeText(config); setNotice('설정을 복사했습니다. 토큰은 다른 사람과 공유하지 마세요.'); })}><Copy size={14} />설정 복사</button><button className="text-button" onClick={() => setIssued(null)}>설정 표시 닫기</button></div>
        </div> : <form className="external-create" onSubmit={e => { e.preventDefault(); void create(); }}>
          <h3>새 연결</h3><div className="external-fields"><label>프로그램 이름<input value={label} maxLength={80} required disabled={!!busy} placeholder="예: Cursor" onChange={e => setLabel(e.target.value)} /></label>
          <label>허용할 작업<select value={permission} disabled={!!busy} onChange={e => { setPermission(e.target.value as typeof permission); setWriteConsent(false); }}><option value="read">읽기</option><option value="propose">수정 제안 · 웹에서 승인</option><option value="write">직접 수정</option></select></label>
          <label>연결 기간<select value={hours} disabled={!!busy} onChange={e => setHours(Number(e.target.value))}>{[1, 4, 12, 24].map(hour => <option key={hour} value={hour}>{hour}시간</option>)}</select></label></div>
          {permission === 'write' && <label className="privacy-confirm"><input type="checkbox" checked={writeConsent} disabled={!!busy || dirty} onChange={e => setWriteConsent(e.target.checked)} />외부 프로그램이 웹 승인 없이 내용을 수정하도록 허용합니다. {dirty ? '현재 편집을 먼저 저장해 주세요.' : '작업은 버전 기록에 남습니다.'}</label>}
          <p>연결은 선택한 문서와 이 브라우저 세션의 유효기간 안에서만 유지됩니다. 내보내기는 검증을 마친 저장본에만 허용됩니다.</p>
          <button className="primary-button" type="submit" disabled={!!busy || !label.trim() || (permission === 'write' && (!writeConsent || dirty))}><Link2 size={14} />연결 설정 만들기</button>
        </form>}
        <h3 className="external-section-heading">발급한 연결</h3><div className="external-list">{data.connections.map(connection => {
          const active = !connection.revoked && connection.expires * 1000 > Date.now();
          return <article key={connection.id} className="external-connection"><div><strong>{connection.label}</strong><span>{permissionLabels[connection.permission]} · {active ? `${date(connection.expires)}까지` : connection.revoked ? '해제됨' : '만료됨'}</span></div>
            {active && (confirmRevoke === connection.id ? <div className="external-buttons"><span>이 연결을 해제할까요?</span><button className="text-button" disabled={!!busy} onClick={() => setConfirmRevoke(null)}>취소</button><button className="outline-button" disabled={!!busy} onClick={() => void run('연결을 해제하고 있습니다', async () => { await api(`/connections/${connection.id}`, { method: 'DELETE' }); if (issued?.id === connection.id) setIssued(null); setConfirmRevoke(null); await refresh(); })}>해제</button></div> : <button className="text-button" disabled={!!busy} onClick={() => setConfirmRevoke(connection.id)}><Unplug size={14} />해제</button>)}
          </article>;
        })}{!data.connections.length && <p className="external-empty">이 문서에 발급한 연결이 없습니다.</p>}</div>
      </div>}
      {tab === 'proposals' && <div className="external-body" id="external-panel-proposals" role="tabpanel" aria-labelledby="external-tab-proposals">
        {review ? <><button className="text-button" disabled={!!busy} onClick={() => setReview(null)}>← 제안 목록</button><h3>{review.title}</h3><p>버전 {review.revision + 1} 기준 · 적용할 항목을 선택하세요.</p>
          {(dirty || stale) && <p className="tool-error" role="status">{dirty ? '미저장 편집을 먼저 저장해 주세요.' : '문서나 설정이 변경되었습니다. 최신 문서를 열고 외부 도구에 새 제안을 요청해 주세요.'}</p>}
          <div className="external-comparisons">{review.changes.map(change => <article className="change-item" key={change.id}><label><input type="checkbox" checked={selected.includes(change.id)} disabled={!!busy || review.status !== 'pending'} onChange={e => setSelected(current => e.target.checked ? [...current, change.id] : current.filter(id => id !== change.id))} />{change.label}</label><div className="change-comparison"><div><span>변경 전</span><p>{change.before || '(빈 내용)'}</p></div><div><span>변경 후</span><p>{change.after || '(빈 내용)'}</p></div></div></article>)}</div>
          {review.status === 'pending' ? <div className="external-buttons"><button className="outline-button" disabled={!!busy} onClick={() => void run('제안을 거절하고 있습니다', async () => { await post(`/proposals/${review.id}/reject`, {}); setReview(null); await refresh(); })}>제안 거절</button><button className="primary-button" disabled={!!busy || dirty || stale || !selected.length} onClick={() => void approve()}>선택한 {selected.length}개 적용</button></div> : <p>이미 {review.status === 'accepted' ? '적용 요청' : '거절'}한 제안입니다. 적용 결과는 작업 기록이나 버전 기록에서 확인하세요.</p>}
        </> : <div className="external-list">{data.proposals.map(proposal => <button className="external-proposal" key={proposal.id} disabled={!!busy} onClick={() => void inspect(proposal)}><span><strong>{proposal.title}</strong><small>{proposal.label} · 버전 {proposal.revision + 1}</small></span><span>{proposal.status === 'pending' ? '검토 대기' : proposal.status === 'accepted' ? '적용 요청됨' : '거절됨'}</span></button>)}{!data.proposals.length && <p className="external-empty">수정 제안이 도착하면 이곳에서 변경 전후를 확인할 수 있습니다.</p>}</div>}
      </div>}
      {tab === 'events' && <div className="external-body" id="external-panel-events" role="tabpanel" aria-labelledby="external-tab-events"><p>최근 연결 작업을 표시합니다. 적용 결과와 오류는 해당 작업에서 확인할 수 있습니다.</p><div className="external-list">{data.events.map((event, index) => <div className="external-event" key={index}><span><strong>{actions[event.action] || event.action}</strong><small>{event.label} · {date(event.created)}</small></span>{event.job_id && <button className="text-button" disabled={!!busy} onClick={() => void run('작업 결과를 확인하고 있습니다', async () => { const job = await waitForJob(event.job_id!, () => {}); setNotice(job.status === 'completed' ? '작업이 완료되었습니다. 최신 문서와 버전 기록에서 결과를 확인해 주세요.' : '작업을 확인했습니다.'); })}>결과 확인</button>}</div>)}{!data.events.length && <p className="external-empty">아직 연결 작업이 없습니다.</p>}</div></div>}
    </section>
  </div>;
}
