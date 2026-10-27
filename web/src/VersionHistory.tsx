import { useModalFocus } from './modal-focus';
import { useEffect, useMemo, useRef, useState } from 'react';
import { History, LoaderCircle, RefreshCw, RotateCcw, X } from 'lucide-react';
import { api, type Document, type RevisionPage, type RevisionSnapshot, type SavedRevision } from './api';
import { compareRevisions } from './revision-comparison';

type Props = { doc: Document; dirty: boolean; onClose: () => void; onRestore: (source: number) => Promise<void> };
const labels = { analyze: '업로드 원본', edit: '텍스트 수정', studio: '문서 편집', restore: '이전 버전 복원', layout: '쪽 양식 적용' };
const date = (value: number) => new Date(value * 1000).toLocaleString('ko-KR', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false });

export function VersionHistory({ doc, dirty, onClose, onRestore }: Props) {
  const dialog = useRef<HTMLElement>(null);
  const [items, setItems] = useState<SavedRevision[]>([]);
  const [cursor, setCursor] = useState<number | null>(null);
  const [selected, setSelected] = useState(Math.max(0, doc.revision - 1));
  const [snapshot, setSnapshot] = useState<RevisionSnapshot | null>(null);
  const [listLoading, setListLoading] = useState(true);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [stale, setStale] = useState(false);
  const [reload, setReload] = useState(0);
  const [search, setSearch] = useState('');
  const [limit, setLimit] = useState(30);
  const [confirm, setConfirm] = useState(false);
  const [restoring, setRestoring] = useState(false);
  const restoringRef = useRef(false);
  const comparison = useMemo(() => snapshot ? compareRevisions(snapshot, doc) : null, [snapshot, doc]);
  const filtered = useMemo(() => {
    const needle = search.trim().toLocaleLowerCase();
    return comparison?.items.filter(item => !needle || `${item.label}\n${item.before}\n${item.after}`.toLocaleLowerCase().includes(needle)) || [];
  }, [comparison, search]);

  useEffect(() => {
    const abort = new AbortController();
    setListLoading(true); setError('');
    api<RevisionPage>(`/documents/${doc.id}/revisions`, { signal: abort.signal }).then(page => {
      setItems(page.items); setCursor(page.next_before); setStale(page.current_revision !== doc.revision);
    }).catch(e => { if (!abort.signal.aborted) setError(e.message); }).finally(() => { if (!abort.signal.aborted) setListLoading(false); });
    return () => abort.abort();
  }, [doc.id, doc.revision, reload]);

  useEffect(() => {
    const abort = new AbortController();
    setSnapshot(null); setLoading(true); setConfirm(false); setSearch(''); setLimit(30);
    api<RevisionSnapshot>(`/documents/${doc.id}/revisions/${selected}`, { signal: abort.signal }).then(setSnapshot).catch(e => { if (!abort.signal.aborted) setError(e.message); }).finally(() => { if (!abort.signal.aborted) setLoading(false); });
    return () => abort.abort();
  }, [doc.id, selected, reload]);

  useModalFocus(dialog, () => { if (!restoringRef.current) onClose(); });

  async function more() {
    if (cursor === null || listLoading) return;
    setListLoading(true); setError('');
    try {
      const page = await api<RevisionPage>(`/documents/${doc.id}/revisions?before=${cursor}`);
      setItems(current => [...current, ...page.items.filter(item => !current.some(other => other.revision === item.revision))]);
      setCursor(page.next_before); setStale(page.current_revision !== doc.revision);
    } catch (e) { setError(e instanceof Error ? e.message : '기록을 불러오지 못했습니다.'); }
    finally { setListLoading(false); }
  }
  async function restore() {
    if (restoringRef.current || dirty || stale || !snapshot || selected === doc.revision) return;
    restoringRef.current = true; setRestoring(true); setError('');
    try { await onRestore(selected); onClose(); }
    catch (e) { setError(e instanceof Error ? e.message : '이전 버전을 복원하지 못했습니다.'); setConfirm(false); }
    finally { restoringRef.current = false; setRestoring(false); }
  }
  return <div className="modal-backdrop version-backdrop" onClick={() => { if (!restoring) onClose(); }}><section className="modal version-modal" ref={dialog} role="dialog" aria-modal="true" aria-labelledby="version-title" onClick={e => e.stopPropagation()}>
    <div className="modal-heading"><div><h2 id="version-title"><History size={18} />버전 기록</h2><p>저장된 버전을 비교하고 이전 내용으로 복원할 수 있습니다.</p></div><div className="version-heading-actions"><button className="icon-button" aria-label="버전 기록 새로고침" disabled={restoring || listLoading} onClick={() => setReload(n => n + 1)}><RefreshCw size={16} /></button><button className="icon-button" aria-label="버전 기록 닫기" disabled={restoring} onClick={onClose}><X size={20} /></button></div></div>
    {error && <p className="version-alert" role="alert">{error}</p>}
    {stale && <p className="version-alert" role="status">다른 탭에서 저장한 새 버전이 있습니다. 창을 닫고 최신 문서를 다시 연 뒤 복원해 주세요.</p>}
    <div className="version-body"><aside className="version-timeline" aria-label="저장된 버전"><button className="version-original text-button" disabled={restoring} aria-pressed={selected === 0} onClick={() => { setSelected(0); setError(''); }}>업로드 원본과 비교</button><div className="version-list">{items.map(item => <button key={item.revision} className={`version-entry ${selected === item.revision ? 'selected' : ''}`} aria-pressed={selected === item.revision} disabled={restoring} onClick={() => { setSelected(item.revision); setError(''); }}><span><b>버전 {item.revision + 1}</b>{item.revision === doc.revision && <i>현재</i>}</span><small>{item.kind === 'restore' ? `버전 ${(item.source_revision ?? 0) + 1}에서 복원` : labels[item.kind]}</small><time dateTime={new Date(item.created * 1000).toISOString()}>{date(item.created)}</time></button>)}</div>{listLoading && <p className="version-status" role="status"><LoaderCircle size={14} className="spin" />기록 불러오는 중</p>}{cursor !== null && <button className="text-button version-more" disabled={listLoading || restoring} onClick={() => void more()}>이전 기록 더 보기</button>}</aside>
    <div className="version-preview" aria-busy={loading}><div className="version-preview-heading"><h3>버전 {selected + 1}<span>→</span>현재 버전 {doc.revision + 1}</h3><p>본문·표의 내용과 표 구조를 비교합니다. 글꼴·그림·쪽 배치는 비교에 포함되지 않으며, 복원할 때는 파일 전체가 적용됩니다.</p></div>
    {loading ? <p className="version-status" role="status"><LoaderCircle size={18} className="spin" />선택한 버전 불러오는 중</p> : comparison && <><div className="version-comparison-tools"><span>본문 {comparison.paragraphs}구간 · 표 {comparison.tables}구간 변경</span><input aria-label="버전 변경 내용 검색" placeholder="변경 내용에서 찾기" value={search} onChange={e => { setSearch(e.target.value); setLimit(30); }} /></div><div className="version-changes">{filtered.slice(0, limit).map(item => <article className="change-item" key={item.id}><header><strong>{item.label}</strong><small>{!item.beforeCount ? '내용 추가' : !item.afterCount ? '내용 삭제' : '내용 변경'}</small></header><div className="change-comparison"><div><span>선택한 버전 {selected + 1}</span><p>{item.beforeCount ? item.before || '(빈 내용)' : '(내용 없음)'}</p></div><div><span>현재 버전 {doc.revision + 1}</span><p>{item.afterCount ? item.after || '(빈 내용)' : '(내용 없음)'}</p></div></div>{item.clipped && <p className="version-clipped">긴 내용은 앞부분 16,000자까지 표시합니다.</p>}</article>)}{!filtered.length && <div className="version-no-changes"><History size={25} /><p>{search ? '검색어와 일치하는 변경 내용이 없습니다.' : selected === doc.revision ? '현재 저장된 버전입니다.' : '본문·표 내용에 차이가 없습니다.'}</p>{!search && selected !== doc.revision && <small>서식이나 그림의 변경은 이 비교에 표시되지 않을 수 있습니다.</small>}</div>}{filtered.length > limit && <button className="outline-button changes-more" onClick={() => setLimit(n => n + 30)}>변경 내용 더 보기 ({filtered.length - limit}개 남음)</button>}{comparison.omitted > 0 && <p className="version-clipped">전체 {comparison.total}구간 중 앞의 500구간을 표시합니다.</p>}</div></>}
    </div></div>
    <div className="version-footer">{confirm ? <div className="version-confirm" role="status"><p><strong>버전 {selected + 1}의 내용으로 복원할까요?</strong><span>새 버전 {doc.revision + 2}로 저장합니다. 현재 버전과 이전 기록은 유지되며 다운로드 전에 검증이 필요합니다.</span></p><div><button className="outline-button" disabled={restoring} onClick={() => setConfirm(false)}>취소</button><button className="primary-button" disabled={restoring || dirty || stale} onClick={() => void restore()}>{restoring ? <LoaderCircle size={14} className="spin" /> : <RotateCcw size={14} />}{restoring ? '복원 중' : '새 버전으로 복원'}</button></div></div> : <><p>{dirty ? '미저장 편집을 먼저 수정 저장하면 이전 버전을 복원할 수 있습니다.' : '선택한 버전의 내용과 개체를 새 저장본으로 복원합니다. 현재 고급 설정도 다시 검사합니다.'}</p><button className="primary-button" disabled={dirty || stale || loading || !snapshot || selected === doc.revision} onClick={() => setConfirm(true)}><RotateCcw size={14} />이 버전 복원</button></>}</div>
  </section></div>;
}
